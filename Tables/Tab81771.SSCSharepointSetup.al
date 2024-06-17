table 81771 "SSC Sharepoint Setup"
{
    Caption = 'Sharepoint Setup';
    DataClassification = CustomerContent;

    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            Caption = 'Primary Key';
        }
        field(2; "Client Id"; Text[50])
        {
            Caption = 'Client Id';
        }
        field(3; Certificate; Blob)
        {
            Caption = 'Certificate';
        }
        field(4; Tenant; Text[100])
        {
            Caption = 'Tenant';
        }
        field(5; "Sharepoint URL"; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Sharepoint URL';
        }
        field(6; Scope; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Scope';
        }
        field(7; "Authorizaton Type"; enum "SSC Auth Type")
        {
            DataClassification = CustomerContent;
            Caption = 'Authorizaton Type';
        }
        field(8; "Token Expiration DT"; DateTime)
        {
            DataClassification = CustomerContent;
            Caption = 'Token Expiration DT';
        }
        field(9; "Azure Authrization URL"; Text[250])
        {
            DataClassification = CustomerContent;
            ExtendedDatatype = URL;

        }
        field(10; "Azure Authrization Key"; Text[250])
        {
            DataClassification = CustomerContent;
            ExtendedDatatype = Masked;
        }
        field(11; "Sharepoint Folder"; Text[250])
        {
            DataClassification = CustomerContent;
            Caption = 'Sharepoint Folder';
        }
    }
    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }

    procedure GetCertificatePassword(): Text
    var
        Secret: Text;
    begin
        if IsolatedStorage.Contains(GetCertificateKey(), DataScope::Company) then begin
            IsolatedStorage.Get(GetCertificateKey(), DataScope::Company, Secret);
            exit(Secret);
        end;
    end;

    procedure SetCertificatePassword(CertificatePassword: Text)
    begin
        if IsolatedStorage.Contains(GetCertificateKey(), DataScope::Company) then
            IsolatedStorage.Delete(GetCertificateKey(), DataScope::Company);

        IsolatedStorage.set(GetCertificateKey(), CertificatePassword, DataScope::Company);
    end;

    procedure SetSecret(SecretType: Enum "SSC Secret Type"; Secret: SecretText)
    var
        StorageKey: Guid;
    begin
        case SecretType of
            SecretType::ClientSecret:
                StorageKey := GetClientSecretKey();
            SecretType::CustomAccessToken:
                StorageKey := GetCustomAccessTokenSecretKey();
        end;

        if IsolatedStorage.Contains(StorageKey, DataScope::Company) then
            IsolatedStorage.Delete(StorageKey, DataScope::Company);

        IsolatedStorage.set(StorageKey, Secret, DataScope::Company);
    end;

    procedure GetSecret(SecretType: Enum "SSC Secret Type"): SecretText
    var
        Secret: SecretText;
        StorageKey: Guid;
    begin
        case SecretType of
            SecretType::ClientSecret:
                StorageKey := GetClientSecretKey();
            SecretType::CustomAccessToken:
                StorageKey := GetCustomAccessTokenSecretKey();
        end;

        if IsolatedStorage.Contains(StorageKey, DataScope::Company) then begin
            IsolatedStorage.Get(StorageKey, DataScope::Company, Secret);
            exit(Secret);
        end;
    end;

    local procedure GetCertificateKey(): Guid
    begin
        exit('b7c0e0d5-b425-4911-a169-4178678df231');
    end;

    local procedure GetClientSecretKey(): Guid
    begin
        exit('0371e543-6b3d-4ad0-9c7c-32c9130d47f2');
    end;

    local procedure GetCustomAccessTokenSecretKey(): Guid
    begin
        exit('56ebe875-2a55-4292-8655-06660a051823');
    end;
}