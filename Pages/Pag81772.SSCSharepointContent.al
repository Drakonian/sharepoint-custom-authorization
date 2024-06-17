page 81772 "SSC Sharepoint Content"
{
    ApplicationArea = All;
    Caption = 'Sharepoint Content';
    PageType = List;
    SourceTable = "SSC Sharepoint Content";
    UsageCategory = Lists;
    ModifyAllowed = false;
    InsertAllowed = false;
    DeleteAllowed = false;

    layout
    {
        area(Content)
        {
            repeater(General)
            {
                field("Relative File Path"; Rec."Relative File Path")
                {
                    ToolTip = 'Specifies the value of the Relative File Path field.', Comment = '%';
                }
                field("File Name"; Rec."File Name")
                {
                    ToolTip = 'Specifies the value of the File Name field.', Comment = '%';
                }
            }
        }
    }
    actions
    {
        area(Promoted)
        {
            actionref(UploadFileToSharepoint_promoted; UploadFileToSharepoint)
            {

            }
            actionref(DownloadFileFromSharepoint_promoted; DownloadFileFromSharepoint)
            {

            }
            actionref(DeleteFileFromSharepoint_promoted; DeleteFileFromSharepoint)
            {

            }
        }
        area(Processing)
        {
            action(UploadFileToSharepoint)
            {
                ApplicationArea = All;
                Caption = 'Upload File to Sharepoint';
                ToolTip = 'Upload File to Sharepoint';
                Image = Import;
                trigger OnAction()
                var
                    SharepointSetup: Record "SSC Sharepoint Setup";
                    SharepointMgt: Codeunit "SSC Sharepoint Mgt.";
                    TempBlob: Codeunit "Temp Blob";
                    FileInStream: InStream;
                    FileName: Text;
                begin
                    SharepointSetup.Get();
                    TempBlob.CreateInStream(FileInStream, TextEncoding::UTF8);
                    if not UploadIntoStream('Upload File to Sharepoint', '', '', FileName, FileInStream) then
                        exit;

                    SharepointMgt.UploadFile(SharepointSetup."Sharepoint Folder", FileName, FileInStream);

                    Rec.Init();
                    Rec."File Name" := CopyStr(FileName, 1, MaxStrLen(Rec."File Name"));
                    Rec."Relative File Path" := CopyStr(StrSubstNo('%1/%2', SharepointSetup."Sharepoint Folder", FileName), 1, MaxStrLen(Rec."Relative File Path"));
                    Rec.Insert(true);

                    CurrPage.Update(false);
                end;
            }
            action(DownloadFileFromSharepoint)
            {
                ApplicationArea = All;
                Caption = 'Download File From Sharepoint';
                ToolTip = 'Download File From Sharepoint';
                Image = Download;
                trigger OnAction()
                var
                    SharepointMgt: Codeunit "SSC Sharepoint Mgt.";
                begin
                    if Rec."Relative File Path" = '' then
                        exit;
                    SharepointMgt.DownloadFile(Rec."Relative File Path");
                end;
            }
            action(DeleteFileFromSharepoint)
            {
                ApplicationArea = All;
                Caption = 'Delete File From Sharepoint';
                ToolTip = 'Delete File From Sharepoint';
                Image = Delete;
                trigger OnAction()
                var
                    SharepointMgt: Codeunit "SSC Sharepoint Mgt.";
                begin
                    if Rec."Relative File Path" = '' then
                        exit;
                    SharepointMgt.DeleteFile(Rec."Relative File Path");

                    Rec.Delete(true);
                end;
            }
        }
    }
}
