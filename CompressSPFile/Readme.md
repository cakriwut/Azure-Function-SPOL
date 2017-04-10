# CompressSPFile Azure Function
This function app will retrieve file from SharePoint online and return zip stream to the client.
In order to install, configure the deployment from source control to Azure Function app.

## Configuration
1. CompressSPFile.ClientId : Azure App client ID
2. CompressSPFile.Cert : Certificate file with private key. (*.pfx)
3. CompressSPFile.CertPassword: Certificate password
4. CompressSPFile.Authority : Azure App authentication authority
5. CompressSPFile.Resource : Azure App authentication resource