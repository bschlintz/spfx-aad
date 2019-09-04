# spfx-aad
Sample project with an Azure Function Http Trigger (e.g. API) that validates and decodes an Azure AD authorization header before returning it to the caller. The project also includes an SPFx solution that reads the decoded token and renders it on the page.

To setup this project, create an Azure AD app registration and use the Application (Client) ID in the webpart settings. Run the Azure Function locally and set the URL in the webpart property settings.

Be sure to grant access to your custom app registration from SPO by using the commands below.
```
Connect-SPOService -Url https://mytenant-admin.sharepoint.com
Approve-SPOTenantServicePrincipalPermissionGrant -Resource "App Registration Display Name" -Scope "target_scope"
```

## Disclaimer

Microsoft provides programming examples for illustration only, without warranty either expressed or implied, including, but not limited to, the implied warranties of merchantability and/or fitness for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys' fees, that arise or result from the use or distribution of the Sample Code.
