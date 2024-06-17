page 81771 "SSC Sharepoint Setup"
{
    ApplicationArea = All;
    Caption = 'Sharepoint Setup';
    PageType = Card;
    SourceTable = "SSC Sharepoint Setup";
    UsageCategory = Administration;

    layout
    {
        area(Content)
        {
            group(General)
            {
                Caption = 'General';
                field("Authorizaton Type"; Rec."Authorizaton Type")
                {
                    ToolTip = 'Specifies the value of the Authorizaton Type field.', Comment = '%';
                }
                field("Client Id"; Rec."Client Id")
                {
                    ToolTip = 'Specifies the value of the Client Id field.', Comment = '%';
                }
                group(ClientSecretGroup)
                {
                    ShowCaption = false;
                    Visible = Rec."Authorizaton Type" = Rec."Authorizaton Type"::"Authorization Code";
                    field(ClientSecret; ClientSecret)
                    {
                        ApplicationArea = All;
                        Caption = 'Client Secret';
                        ToolTip = 'Client Secret';
                        ExtendedDatatype = Masked;
                        trigger OnValidate()
                        begin
                            Rec.SetSecret(Enum::"SSC Secret Type"::ClientSecret, ClientSecret);
                        end;
                    }
                }
                field(Tenant; Rec.Tenant)
                {
                    ToolTip = 'Specifies the value of the Tenant field.', Comment = '%';
                }
                field("Sharepoint URL"; Rec."Sharepoint URL")
                {
                    ToolTip = 'Specifies the value of the Sharepoint URL field.', Comment = '%';
                }
                field("Sharepoint Folder"; Rec."Sharepoint Folder")
                {
                    ToolTip = 'Specifies the value of the Sharepoint Folder field.', Comment = '%';
                }
                field(Scope; Rec.Scope)
                {
                    ToolTip = 'Specifies the value of the Scope field.', Comment = '%';
                }
                group(CertificateGroup)
                {
                    ShowCaption = false;
                    Visible = Rec."Authorizaton Type" <> Rec."Authorizaton Type"::"Authorization Code";
                    field("CertificateUploaded"; Rec.Certificate.HasValue())
                    {
                        ApplicationArea = All;
                        Caption = 'Certificate Uploaded';
                        ToolTip = 'Certificate Uploaded';
                    }
                    field(CertificatePassword; CertificatePassword)
                    {
                        ApplicationArea = All;
                        Caption = 'Certificate Password';
                        ToolTip = 'Certificate Password';
                        ExtendedDatatype = Masked;
                        trigger OnValidate()
                        begin
                            Rec.SetCertificatePassword(CertificatePassword);
                        end;
                    }
                }
                group(CustomCertificateGroup)
                {
                    ShowCaption = false;
                    Visible = Rec."Authorizaton Type" = Rec."Authorizaton Type"::"Custom Certificate";
                    field("Azure Authrization URL"; Rec."Azure Authrization URL")
                    {
                        ToolTip = 'Specifies the value of the Azure Authrization URL field.', Comment = '%';
                    }
                    field("Azure Authrization Key"; Rec."Azure Authrization Key")
                    {
                        ToolTip = 'Specifies the value of the Azure Authrization Key field.', Comment = '%';
                    }
                }
            }
        }
    }
    actions
    {
        area(Promoted)
        {
            actionref(UploadCerificate_promoted; UploadCerificate)
            {

            }
            actionref(SharepointContent_promoted; SharepointContent)
            {

            }
        }
        area(Processing)
        {
            action(UploadCerificate)
            {
                ApplicationArea = All;
                Caption = 'Upload Certificate';
                ToolTip = 'Upload Certificate';
                Image = Import;
                trigger OnAction()
                var
                    TempBlob: Codeunit "Temp Blob";
                    Base64Convert: Codeunit "Base64 Convert";
                    FileInStream: InStream;
                    FileOutStream: OutStream;
                begin
                    TempBlob.CreateInStream(FileInStream, TextEncoding::UTF8);
                    if not UploadIntoStream('', FileInStream) then
                        exit;

                    Rec.Certificate.CreateOutStream(FileOutStream, TextEncoding::UTF8);
                    FileOutStream.Write(Base64Convert.ToBase64(FileInStream));
                    Rec.Modify();
                end;
            }
            action(SharepointContent)
            {
                ApplicationArea = All;
                Caption = 'Sharepoint Content';
                ToolTip = 'Sharepoint Content';
                Image = AllLines;
                RunObject = page "SSC Sharepoint Content";
            }
        }
    }
    trigger OnOpenPage()
    begin
        Rec.Reset();
        if not Rec.Get() then begin
            Rec.Init();
            Rec.Insert();
        end;
    end;

    trigger OnAfterGetCurrRecord()
    begin
        if Rec.GetCertificatePassword() <> '' then
            CertificatePassword := '****'
        else
            CertificatePassword := '';

        if not Rec.GetSecret(Enum::"SSC Secret Type"::ClientSecret).IsEmpty() then
            ClientSecret := '****'
        else
            ClientSecret := '';
    end;

    var
        [NonDebuggable]
        CertificatePassword: Text;
        [NonDebuggable]
        ClientSecret: Text;
}
