codeunit 81772 "SSC SharePoint S2S Certificate" implements "SharePoint Authorization"
{
    SingleInstance = true;

    var
        SharepointSetup: Record "SSC Sharepoint Setup";
        ClientId: Text;
        CertificatePassword: Text;
        EntraTenantId: Text;
        Scope: Text;
        CertificateText: Text;
        AuthorityTxt: Label 'https://login.microsoftonline.com/%1/oauth2/v2.0/token', Comment = '%1 = Microsoft Entra tenant ID', Locked = true;
        BearerTxt: Label 'Bearer %1', Comment = '%1 - Token', Locked = true;

    procedure SetParameters(NewEntraTenantId: Text; NewClientId: Text; NewScope: Text; NewCertificateText: Text; NewCertificatePassword: Text; var NewSharepointSetup: Record "SSC Sharepoint Setup")
    begin
        EntraTenantId := NewEntraTenantId;
        ClientId := NewClientId;
        CertificatePassword := NewCertificatePassword;
        Scope := NewScope;
        CertificateText := NewCertificateText;
        SharepointSetup := NewSharepointSetup;
    end;

    procedure Authorize(var HttpRequestMessage: HttpRequestMessage);
    var
        Headers: HttpHeaders;
    begin
        HttpRequestMessage.GetHeaders(Headers);
        Headers.Add('Authorization', SecretStrSubstNo(BearerTxt, GetToken()));
    end;

    local procedure GetToken(): SecretText
    var
        AccessToken: SecretText;
    begin
        AcquireToken(AccessToken);
        exit(AccessToken);
    end;

    local procedure AcquireToken(var AccessToken: SecretText)
    var
        APIMgt: Codeunit "SSC API Mgt.";
        ResponseTempBlob: Codeunit "Temp Blob";
        DictionaryContentHeaders: Codeunit "Dictionary Wrapper";
        DictionaryDefaultHeaders: Codeunit "Dictionary Wrapper";
        TypeHelper: Codeunit "Type Helper";
        contentToSend: TextBuilder;
        ResponseInstream: InStream;
        TokenRequest: Text;
        JObject: JsonObject;
        JToken: JsonToken;
        Response: Text;
        Buffer: Text;
        AssertionKey: Text;
        HourDuration: Duration;
        MinuteErr: Duration;
    begin
        HourDuration := 1 * 60 * 60 * 1000; //1 hour
        MinuteErr := 1 * 60 * 1000; //1 minute

        if SharepointSetup."Token Expiration DT" <> 0DT then
            if SharepointSetup."Token Expiration DT" >= (CurrentDateTime() + MinuteErr) then begin
                AccessToken := SharepointSetup.GetSecret(Enum::"SSC Secret Type"::CustomAccessToken);
                exit;
            end;

        JObject.Add('clientId', ClientId);
        JObject.Add('tenantId', EntraTenantId);
        JObject.Add('certificatePassword', CertificatePassword);
        JObject.Add('base64Cert', CertificateText);
        JObject.WriteTo(TokenRequest);

        APIMgt.SendRequest(TokenRequest, Enum::"Http Request Type"::POST,
             StrSubstNo('%1%2', SharepointSetup."Azure Authrization URL", SharepointSetup."Azure Authrization Key"),
             '', 0, ResponseTempBlob, DictionaryContentHeaders, DictionaryDefaultHeaders);

        ResponseTempBlob.CreateInStream(ResponseInstream);
        while not ResponseInstream.EOS() do begin
            ResponseInstream.Read(Buffer);
            Response += Buffer;
        end;
        AssertionKey := Response;
        Clear(Response);
        Clear(ResponseInstream);
        Clear(ResponseTempBlob);
        Clear(JObject);

        contentToSend.Append(StrSubstNo('client_id=%1', ClientId));
        contentToSend.Append(StrSubstNo('&client_assertion=%1', TypeHelper.UriEscapeDataString(AssertionKey)));
        contentToSend.Append(StrSubstNo('&scope=%1&', TypeHelper.UriEscapeDataString(Scope)));
        contentToSend.Append(StrSubstNo('client_assertion_type=%1&', TypeHelper.UriEscapeDataString('urn:ietf:params:oauth:client-assertion-type:jwt-bearer')));
        contentToSend.Append('grant_type=client_credentials');

        APIMgt.SendRequest(contentToSend.ToText(), Enum::"Http Request Type"::POST, StrSubstNo(AuthorityTxt, EntraTenantId),
            'application/x-www-form-urlencoded', 0, ResponseTempBlob, DictionaryContentHeaders, DictionaryDefaultHeaders);

        ResponseTempBlob.CreateInStream(ResponseInstream);

        while not ResponseInstream.EOS() do begin
            ResponseInstream.Read(Buffer);
            Response += Buffer;
        end;
        JObject.ReadFrom(Response);

        JObject.Get('access_token', JToken);
        AccessToken := JToken.AsValue().AsText();

        SharepointSetup.SetSecret(Enum::"SSC Secret Type"::CustomAccessToken, AccessToken);
        SharepointSetup."Token Expiration DT" := CurrentDateTime() + HourDuration;
        SharepointSetup.Modify();
    end;
}