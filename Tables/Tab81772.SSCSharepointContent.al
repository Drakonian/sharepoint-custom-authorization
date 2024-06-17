table 81772 "SSC Sharepoint Content"
{
    Caption = 'Sharepoint Content';
    DataClassification = CustomerContent;

    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            AutoIncrement = true;
        }
        field(2; "Relative File Path"; Text[2048])
        {
            Caption = 'Relative File Path';
        }
        field(3; "File Name"; Text[500])
        {
            Caption = 'File Name';
        }
    }
    keys
    {
        key(PK; "Entry No.")
        {
            Clustered = true;
        }
    }
}
