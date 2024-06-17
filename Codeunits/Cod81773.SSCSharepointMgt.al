codeunit 81773 "SSC Sharepoint Mgt."
{
    local procedure InitializeConnection()
    var
        TempSharePointList: Record "SharePoint List" temporary;
        Diagnostics: Interface "HTTP Diagnostics";
    begin
        SharepointSetup.Get();
        SharepointSetup.TestField("Sharepoint URL");
        SharepointSetup.TestField("Sharepoint Folder");
        SharepointSetup.TestField(Tenant);
        SharepointSetup.TestField(Scope);

        SharePointClient.Initialize(SharepointSetup."Sharepoint URL", GetSharePointAuthorization());

        SharePointClient.GetLists(TempSharePointList);
        Diagnostics := SharePointClient.GetDiagnostics();

        if not Diagnostics.IsSuccessStatusCode() then
            Error(DiagErr, Diagnostics.GetHttpStatusCode(), Diagnostics.GetErrorMessage(), Diagnostics.GetResponseReasonPhrase());
    end;

    local procedure GetSharePointAuthorization(): Interface "SharePoint Authorization"
    var
        TempBlob: Codeunit "Temp Blob";
        SharePointAuth: Codeunit "SharePoint Auth.";
        SharepointCustomCertificate: Codeunit "SSC SharePoint S2S Certificate";
        FileInStream: InStream;
        TxtBuffer: Text;
        CertificateBase64: Text;
    begin
        SharepointSetup.CalcFields(Certificate);
        TempBlob.FromRecord(SharepointSetup, SharepointSetup.FieldNo(Certificate));
        TempBlob.CreateInStream(FileInStream, TextEncoding::UTF8);
        while not FileInStream.EOS() do begin
            FileInStream.Read(TxtBuffer);
            CertificateBase64 += TxtBuffer;
        end;

        case SharepointSetup."Authorizaton Type" of
            SharepointSetup."Authorizaton Type"::"Authorization Code":
                exit(SharePointAuth.CreateAuthorizationCode(SharepointSetup.Tenant, SharepointSetup."Client Id",
                    SharepointSetup.GetSecret(Enum::"SSC Secret Type"::ClientSecret), SharepointSetup.Scope));

            SharepointSetup."Authorizaton Type"::Certificate:
                exit(SharePointAuth.CreateClientCredentials(SharepointSetup.Tenant, SharepointSetup."Client Id",
                    CertificateBase64, SharepointSetup.GetCertificatePassword(), SharepointSetup.Scope));

            SharepointSetup."Authorizaton Type"::"Custom Certificate":
                begin
                    SharepointSetup.TestField("Azure Authrization URL");
                    SharepointSetup.TestField("Azure Authrization URL");
                    SharepointCustomCertificate.SetParameters(SharepointSetup.Tenant, SharepointSetup."Client ID", SharepointSetup.Scope, CertificateBase64,
                        SharepointSetup.GetCertificatePassword(), SharepointSetup);
                    exit(SharepointCustomCertificate);
                end;

        end;
    end;

    procedure DownloadFile(FilePath: Text)
    var
        TempSharePointFile: Record "SharePoint File" temporary;
        FileMgt: Codeunit "File Management";
        FileName: Text;
    begin
        InitializeConnection();

        FileName := FileMgt.GetFileName(FilePath);
        if FileName = '' then
            Error(FileNotFoundErr);

        FilePath := '\' + FileMgt.GetDirectoryName(FilePath);

        if not SharePointClient.GetFolderFilesByServerRelativeUrl(FilePath, TempSharePointFile) then
            Error(FileNotFoundErr);

        TempSharePointFile.SetRange(Name, FileName);
        if not TempSharePointFile.FindFirst() then
            Error(FileNotFoundErr);

        SharePointClient.DownloadFileContent(TempSharePointFile.OdataId, FileName);
    end;

    procedure UploadFile(FolderPath: Text; FileName: Text; FileInstream: InStream): Boolean
    var
        TempSharepointFile: Record "SharePoint File" temporary;
        Diagnostics: Interface "HTTP Diagnostics";
    begin
        InitializeConnection();

        CreateFolder(FolderPath);

        if SharepointClient.AddFileToFolder(FolderPath, FileName, FileInstream, TempSharepointFile) then
            exit(true);

        Diagnostics := SharepointClient.GetDiagnostics();
        if (not Diagnostics.IsSuccessStatusCode()) then
            Error(DiagErr, Diagnostics.GetHttpStatusCode(), Diagnostics.GetErrorMessage(), Diagnostics.GetResponseReasonPhrase());
    end;

    procedure CreateFolder(FolderPath: Text): Boolean
    var
        TempSharepointFolder: Record "SharePoint Folder" temporary;
        Diagnostics: Interface "HTTP Diagnostics";
    begin
        InitializeConnection();

        if SharepointClient.CreateFolder(FolderPath, TempSharepointFolder) then
            exit(true);

        Diagnostics := SharepointClient.GetDiagnostics();
        if not Diagnostics.IsSuccessStatusCode() then
            Error(DiagErr, Diagnostics.GetHttpStatusCode(), Diagnostics.GetErrorMessage(), Diagnostics.GetResponseReasonPhrase());
    end;

    procedure DeleteFile(FilePath: Text): Boolean
    var
        TempSharepointFile: Record "SharePoint File" temporary;
        Diagnostics: Interface "HTTP Diagnostics";
    begin
        InitializeConnection();

        if SharepointClient.GetFileByServerRelativeUrl(FilePath, TempSharepointFile, false) then
            if SharepointClient.DeleteFileByServerRelativeUrl(FilePath) then
                exit(true);

        Diagnostics := SharepointClient.GetDiagnostics();
        if (not Diagnostics.IsSuccessStatusCode()) then
            Error(DiagErr, Diagnostics.GetHttpStatusCode(), Diagnostics.GetErrorMessage(), Diagnostics.GetResponseReasonPhrase());
    end;

    var
        SharepointSetup: Record "SSC Sharepoint Setup";
        SharePointClient: Codeunit "SharePoint Client";
        DiagErr: Label 'Sharepoint Management error:\%1\%2\%3', Comment = '%1 = HttpStatusCode, %2 = ErrorMessage, %3 = ResponseReasonPhrase';
        FileNotFoundErr: Label 'File not found.';
}
