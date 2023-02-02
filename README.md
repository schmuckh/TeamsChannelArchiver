# TeamsChannelArchiver
Tool to archive Microsoft Teams channels as Text or HTML.

## NuGet:
- Microsoft.Graph (4.52.0)
- Microsoft.Graph.Core (2.0.14)
- Microsoft.Identity.Client.Extensions.Msal (2.25.3)
- Newtonsoft.Json (13.0.2)

## Getting started
1. Register a new Application in your Azure Tenant
https://learn.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory

2. Set Permissions for the App
https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-configure-app-access-web-apis
![image](https://user-images.githubusercontent.com/124037247/216256673-63e8c540-6726-4732-b0d5-ab912b28b375.png)

- User.Read
- User.Read.All
- Channel.ReadBasic.All
- Channel.Message.Read.All
- Channel.Message.Send
- ChannelMember.Read.All
- Files.ReadWrite.All

3. Insert TENANTID, APPLICATIONID and REDIRECTURI in MainWindow.xaml.cs
![image](https://user-images.githubusercontent.com/124037247/216256168-0d0cfab3-b517-4ed8-9d92-e519af3698be.png)

4. Create the App

## Usage
![image](https://user-images.githubusercontent.com/124037247/216258438-b57451d6-4b81-4600-b5a1-64a397e923a5.png)

Click on Login will open the Azure Login Dialog in your Browser

Choose Timeframe to Export (Default: Last 90 days)

Select Team

Select Channel

Select format (Text/HTML or both)

Select if a Message will appear in the Teams channel that the conversation was exported
![image](https://user-images.githubusercontent.com/124037247/216258132-4055dbd9-6664-4351-b67d-1f6801607208.png)

Preview will show you the Channel Messages

Click on Export

A new folder "Chatprotokol" will apear in your channels files folder
![image](https://user-images.githubusercontent.com/124037247/216258813-bb1d5c40-9083-4e90-9928-9fa869bbd394.png)

## Known Issues

The Export is plain Text or HTML, so some of the channel messages - like Praise - will apear in raw form (JSON most times)
The App was not testet on any possible channel message format!

