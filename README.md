# TeamsChannelArchiver
Tool to archive Microsoft Teams channels as Text or HTML.

NuGet:
- Microsoft.Graph (4.52.0)
- Microsoft.Graph.Core (2.0.14)
- Microsoft.Identity.Client.Extensions.Msal (2.25.3)
- Newtonsoft.Json (13.0.2)

Getting started
1. Register a new Application in your Azure Tenant
https://learn.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory

2. Set Permissions for the App
https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis

- User.Read
- User.Read.All
- Channel.ReadBasic.All
- Channel.Message.Read.All
- Channel.Message.Send
- ChannelMember.Read.All
- Files.ReadWrite.All

3. Insert TENANTID, APPLICATIONID and REDIRECTURI in MainWindow.xaml.cs

4. Create the App

