permissionset 81771 "SSC Sharepoint S2S"
{
    Assignable = true;
    Permissions = tabledata "SSC Sharepoint Setup" = RIMD,
        table "SSC Sharepoint Setup" = X,
        codeunit "SSC SharePoint S2S Certificate" = X,
        tabledata "SSC Sharepoint Content" = RIMD,
        table "SSC Sharepoint Content" = X,
        codeunit "SSC API Mgt." = X,
        codeunit "SSC Sharepoint Mgt." = X,
        page "SSC Sharepoint Content" = X,
        page "SSC Sharepoint Setup" = X;
}