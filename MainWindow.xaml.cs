using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Windows.Interop;
using System;
using System.Text;
using System.ComponentModel;
using System.IO;
using System.Collections.Generic;
using System.Security;
using System.Threading.Tasks;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace TeamsChannelArchiver
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        const string APPLICATIONID = "<YOUR APPLICATIONID>";
        const string REDIRECTURI = "<YOUR REDIRECTURI>"; //-> http://localhost:<PORT>
        const string TENANTID = "<YOUR TENANTID>";
        string AUTHORITY = $"https://login.microsoftonline.com/{TENANTID}/v2.0/adminconsent?client_id={APPLICATIONID}";

        public GraphServiceClient authClient = null;

        //Set default timeframe to last 90 days
        public DateTime startTime = DateTime.Now.AddDays(-90);
        public DateTime endTime = DateTime.Now;

        private List<string> messages = new List<string>();
        private List<string> channelMembers = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            btnExport.IsEnabled = false;
        }

        private async void btnLogin_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Authentication
                AuthenticationResult authResult = null;

                var cca = PublicClientApplicationBuilder.Create(APPLICATIONID)
                                                            .WithAuthority(AUTHORITY)
                                                            .WithRedirectUri(REDIRECTURI)
                                                            .Build();

                tbxSysMessages.Text = string.Empty;

                var accounts = await cca.GetAccountsAsync();
                var firstAccount = accounts.FirstOrDefault();

                //Set permissions
                List<string> scopes = new List<string>();
                scopes.Add("User.Read");
                scopes.Add("User.Read.All");
                scopes.Add("Channel.ReadBasic.All");
                scopes.Add("ChannelMessage.Read.All");
                scopes.Add("ChannelMessage.Send");
                scopes.Add("ChannelMember.Read.All");
                scopes.Add("Files.ReadWrite.All");

                //Get token
                try
                {
                    authResult = await cca.AcquireTokenSilent(scopes, firstAccount)
                        .ExecuteAsync();
                }
                catch (MsalUiRequiredException ex)
                {
                    try
                    {
                        authResult = await cca.AcquireTokenInteractive(scopes)
                            .WithAccount(firstAccount)
                            .WithParentActivityOrWindow(new WindowInteropHelper(this).Handle) // optional, used to center the browser on the window
                            .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                            .ExecuteAsync();
                    }
                    catch (MsalException msalex)
                    {
                        tbxSysMessages.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                    }
                }
                catch (Exception ex)
                {
                    tbxSysMessages.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                    return;
                }

                //Create authentication provider
                var authProvider = new DelegateAuthenticationProvider(async (request) => {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                });

                //create client
                authClient = new GraphServiceClient(authProvider);

                //Current Tenant
                var requestCurrentTenant = await authClient.Organization.Request().GetAsync();
                var resultTenant = requestCurrentTenant.FirstOrDefault().DisplayName;

                //Current User
                var requestCurrentUser = await authClient.Me.Request().GetAsync();
                var resultCurrent = requestCurrentUser.DisplayName;

                tbxSysMessages.Text += "Current User: " + resultCurrent;
                lblLoginUser.Content = "User: " + resultCurrent;
                lblLoginTenant.Content = "Tenant: " + resultTenant;

                //List my Teams
                var myTeams = await authClient.Me.JoinedTeams.Request().GetAsync();
                var resultTeams = myTeams.CurrentPage;

                //Add teams to Combobox
                cbxTeamSelect.ItemsSource = resultTeams;
                cbxTeamSelect.DisplayMemberPath = "DisplayName";
                cbxTeamSelect.SelectedValuePath = "Id";
                cbxTeamSelect.SelectedIndex = 0;

                //Set timeframe
                dtpFrom.SelectedDate = startTime;
                dtpTo.SelectedDate = endTime;

                btnExport.IsEnabled = false;
                tbxSysMessages.Text = "Finished." + Environment.NewLine + tbxSysMessages.Text;

                //Activate Button Close
                btnLogin.IsEnabled = false;
                btnClose.IsEnabled = true;
            }
            catch (Exception ex)
            {
                tbxSysMessages.Text = ex.Message;
            }
        }

        private void dtpTo_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            endTime = dtpTo.SelectedDate.Value;
            cbxChannelSelect_SelectionChanged(sender, e);
            prgOutput.IsIndeterminate = false;
        }

        private void dtpFrom_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            startTime = dtpFrom.SelectedDate.Value;
            cbxChannelSelect_SelectionChanged(sender, e);
            prgOutput.IsIndeterminate = false;
        }

        private async void cbxTeamSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            prgOutput.IsIndeterminate = true;
            btnExport.IsEnabled = false;

            //Selected Team
            Team selectedTeam = (Team)cbxTeamSelect.SelectedItem;

            if (selectedTeam != null)
            {
                var channels = await authClient.Teams[selectedTeam.Id].Channels.Request().GetAsync();
                var channelResult = channels.CurrentPage;

                //List Channels for selected Team
                cbxChannelSelect.ItemsSource = channelResult;
                cbxChannelSelect.DisplayMemberPath = "DisplayName";
                cbxChannelSelect.SelectedValuePath = "Id";
                cbxChannelSelect.SelectedIndex = -1;

                Channel selectedChannel = (Channel)cbxChannelSelect.SelectedItem;
            }
            prgOutput.IsIndeterminate = false;
        }

        private async void cbxChannelSelect_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            prgOutput.IsIndeterminate = true;
            btnExport.IsEnabled = false;

            //Selected Team
            Team selectedTeam = (Team)cbxTeamSelect.SelectedItem;

            //Selected Channel
            Channel selectedChannel = (Channel)cbxChannelSelect.SelectedItem;

            //Clear Preview
            messages.Clear();
            lsbPreview.ItemsSource = null;
            lsbPreview.Items.Clear();
            

            if (selectedTeam != null && selectedChannel != null)
            {
                //List Members for selected Channel
                var channelMember = await authClient.Teams[selectedTeam.Id].Channels[selectedChannel.Id].Members.Request().GetAsync();
                var channelMemberresult = channelMember.CurrentPage;

                channelMembers.Clear();
                foreach (Microsoft.Graph.ConversationMember s in channelMemberresult)
                {
                    channelMembers.Add(s.DisplayName);
                }

                //Get Messages
                try
                {
                    var channelMessages = await authClient.Teams[selectedTeam.Id].Channels[selectedChannel.Id].Messages.Request().Top(50).GetAsync();


                    List<ChatMessage> chatmessages = new List<ChatMessage>();

                    var pageIterator = PageIterator<ChatMessage>.CreatePageIterator(authClient, channelMessages, (m) =>
                    {
                        chatmessages.Add(m);
                        return true;
                    });

                    await pageIterator.IterateAsync();

                    int itemCount = chatmessages.Where(m => m.CreatedDateTime >= startTime && m.CreatedDateTime <= endTime).OrderBy(m => m.CreatedDateTime).Count();
                    prgOutput.IsIndeterminate = false;

                    tbxSysMessages.Text = "Processing " + itemCount + " Items ..." + Environment.NewLine + tbxSysMessages.Text;

                    //ProgressBar
                    prgOutput.Minimum = 1;
                    prgOutput.Maximum = itemCount;
                    prgOutput.Value = 1;

                    //Write messages to list
                    foreach (var msg in chatmessages.Where(m => m.CreatedDateTime >= startTime && m.CreatedDateTime <= endTime).OrderBy(m => m.CreatedDateTime))
                    {
                        string message = "******************MESSAGE*********************************";
                        message = message + Environment.NewLine + " Date: " + msg.CreatedDateTime;
                        if (msg.From != null) //User can be null for some reason
                        {
                            message = message + Environment.NewLine + " User: " + msg.From.User.DisplayName;
                        }
                        else
                        {
                            message = message + Environment.NewLine + " User: Unknowwn";
                        }
                        message = message + Environment.NewLine + " Body: " + msg.Body.Content;

                        //Get Attachments
                        var attachedFiles = msg.Attachments;
                        if (attachedFiles.Count() > 0)
                        {
                            foreach (var a in attachedFiles)
                            {
                                message = message + Environment.NewLine + " Attachment: " + a.Name;
                                message = message + Environment.NewLine + " - Content Type: " + a.ContentType;
                                message = message + Environment.NewLine + " - Content Url: " + a.ContentUrl;
                                message = message + Environment.NewLine + " - Content: " + Environment.NewLine + a.Content;
                            }
                        }

                        //Get Reactions
                        if (msg.Reactions.ToList().Count > 0)
                        {
                            foreach (var react in msg.Reactions.ToList())
                            {
                                message = message + Environment.NewLine + " Reaction: " + react.ReactionType;
                                message = message + Environment.NewLine + " From: " + react.User.User.DisplayName;
                            }
                        }
                        message = message + Environment.NewLine + "************************************************************";

                        messages.Add(message); //Add message to list


                        //Get replies for message
                        var messagesReplies = await authClient.Teams[selectedTeam.Id].Channels[selectedChannel.Id].Messages[msg.Id].Replies.Request().GetAsync();

                        foreach (var reply in messagesReplies.OrderBy(m => m.CreatedDateTime))
                        {
                            string msgReply = " ******************REPLY*********************************";
                            msgReply = msgReply + Environment.NewLine + "  Date: " + reply.CreatedDateTime;
                            if (reply.From != null) //User can be null for some reason
                            {
                                msgReply = msgReply + Environment.NewLine + " User: " + reply.From.User.DisplayName;
                            }
                            else
                            {
                                msgReply = msgReply + Environment.NewLine + " User: Unknowwn";
                            }
                            msgReply = msgReply + Environment.NewLine + "  Body: " + reply.Body.Content;

                            //Attachments reply
                            var attachedFilesR = reply.Attachments;
                            if (attachedFilesR.Count() > 0)
                            {
                                foreach (var a in attachedFilesR)
                                {
                                    msgReply = msgReply + Environment.NewLine + " Attachment: " + a.Name;
                                    msgReply = msgReply + Environment.NewLine + " - Content Type: " + a.ContentType;
                                    msgReply = msgReply + Environment.NewLine + " - Content Url: " + a.ContentUrl;
                                    msgReply = msgReply + Environment.NewLine + " - Content: " + Environment.NewLine + a.Content;
                                    //msgReply = msgReply + Environment.NewLine + " Content: " + Environment.NewLine + getContent(a.ContentType, a.Content);
                                }
                            }

                            //Reactions reply
                            if (reply.Reactions.ToList().Count > 0)
                            {
                                foreach (var react in reply.Reactions.ToList())
                                {
                                    msgReply = msgReply + Environment.NewLine + "  Reaction: " + react.ReactionType;
                                    msgReply = msgReply + Environment.NewLine + "  From: " + react.User.User.DisplayName;
                                }
                            }
                            msgReply = msgReply + Environment.NewLine + " *********************************************************";

                            messages.Add(msgReply);  //Add reply to list                      
                        }

                        //Increment progressbar
                        prgOutput.Value++;
                    }

                    //Finished
                    prgOutput.IsIndeterminate = false;
                    tbxSysMessages.Text = "Protocol ready to export." + Environment.NewLine + tbxSysMessages.Text;
                    lsbPreview.ItemsSource = messages; 
                    btnExport.IsEnabled = true;
                }
                catch (Exception ex)
                {
                    tbxSysMessages.Text = $"Error getting channel messages:{System.Environment.NewLine}{ex}";
                    return;
                }
            }
        }

        private async void btnExport_Click(object sender, RoutedEventArgs e)
        {
            prgOutput.IsIndeterminate = true;
            btnExport.IsEnabled = false;

            //Selected Team
            Team selectedTeam = (Team)cbxTeamSelect.SelectedItem;

            //Selected Channel
            Channel selectedChannel = (Channel)cbxChannelSelect.SelectedItem;

            //Current User
            var requestCurrentUser = await authClient.Me.Request().GetAsync();
            var resultCurrent = requestCurrentUser.DisplayName;

            //Export html
            if (chbHtml.IsChecked.Value == true)
            {
                //Create HTML File
                string fileNameHtml = "Chatprotocol_" + selectedTeam.DisplayName + "_" + selectedChannel.DisplayName + ".html";
                string htmlFile = createProtocol_html(selectedTeam.DisplayName, selectedChannel.DisplayName, resultCurrent, startTime.ToString(), endTime.ToString(), chbHtml.IsChecked.Value);
                tbxSysMessages.Text = "Export of Protocol " + fileNameHtml + " finished." + Environment.NewLine + tbxSysMessages.Text;

                //Upload to Teams Channel
                tbxSysMessages.Text = "Uploading " + fileNameHtml + " to Channel ..." + Environment.NewLine + tbxSysMessages.Text;
                string fileHtml = System.IO.File.ReadAllText(htmlFile);
                DriveItem newFile = await uploadFile(fileNameHtml, fileHtml, "Chatprotocol", selectedTeam.Id, selectedChannel.Id, false, true);
                if (newFile != null)
                {
                    tbxSysMessages.Text = "Upload of Protocol " + fileNameHtml + " finished." + Environment.NewLine + tbxSysMessages.Text;
                }
                else
                {
                    tbxSysMessages.Text = "Upload of Protocol " + fileNameHtml + " failed." + Environment.NewLine + tbxSysMessages.Text;
                }
            }

            //Export txt
            if (chbText.IsChecked.Value == true)
            {
                //Create TXT File
                string fileNameTxt = "Chatprotocol_" + selectedTeam.DisplayName + "_" + selectedChannel.DisplayName + ".txt";
                string txtFile = createProtocol_txt(selectedTeam.DisplayName, selectedChannel.DisplayName, resultCurrent, startTime.ToString(), endTime.ToString(), chbText.IsChecked.Value);
                tbxSysMessages.Text = "Export of Protocol " + fileNameTxt + " finished." + Environment.NewLine + tbxSysMessages.Text;

                //Upload to Teams Channel
                tbxSysMessages.Text = "Uploading " + fileNameTxt + " to Channel ..." + Environment.NewLine + tbxSysMessages.Text;
                string fileHtml = System.IO.File.ReadAllText(txtFile);
                DriveItem newFile = await uploadFile(fileNameTxt, fileHtml, "Chatprotocol", selectedTeam.Id, selectedChannel.Id, false, true);
                if (newFile != null)
                {
                    tbxSysMessages.Text = "Upload of Protocol " + fileNameTxt + " finished." + Environment.NewLine + tbxSysMessages.Text;
                }
                else
                {
                    tbxSysMessages.Text = "Upload of Protocol " + fileNameTxt + " failed." + Environment.NewLine + tbxSysMessages.Text;
                }
            }

            //Send Channel message
            if (chbMessage.IsChecked.Value == true)
            {
                tbxSysMessages.Text = "Sending Protocol Message ... " + Environment.NewLine + tbxSysMessages.Text;
                sendProtocolMessage(selectedTeam.DisplayName, selectedChannel.DisplayName, selectedTeam.Id, selectedChannel.Id, resultCurrent, startTime.ToString(), endTime.ToString(), chbMessage.IsChecked.Value);
                tbxSysMessages.Text = "Protocol Message in Chat " + selectedTeam.DisplayName + "-" + selectedChannel.DisplayName + " finished." + Environment.NewLine + tbxSysMessages.Text;
            }

            btnExport.IsEnabled = true;
            prgOutput.IsIndeterminate = false;
        }

        private void lsbPreview_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ToDo: Select Items to delete from output
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            prgOutput.IsIndeterminate = true;
            btnExport.IsEnabled = false;

            //Clear Preview
            messages.Clear();
            lsbPreview.ItemsSource = null;
            lsbPreview.Items.Clear();

            prgOutput.IsIndeterminate = false;
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (System.IO.Directory.Exists(@".\tmp"))
            {
                System.IO.Directory.Delete(@".\tmp", true);
            }
        }

        private string createProtocol_html(string teamName, string channelName, string userName, string startDate, string endDate, bool active) //Returns the filepath of the protocol
        {
            //Create Protocol
            string disclaimer = @"<font face=""Arial""><div><strong>Disclaimer</strong><div>" +
                                    "<div>Information contained in this document and any attachments maybe privileged or confidential and intended for the exclusive use of the owner. " +
                                    "If you have received this document by mistake, please delete it and advise the owner immediately.</font></div>";

            string header = "Chat Protocol - " + teamName + "-" + channelName + Environment.NewLine +
                        userName + " Date from: " + startDate + " to " + endDate;

            if (active)
            {
                string folder = @".\tmp";
                if (!System.IO.Directory.Exists(folder))
                {
                    System.IO.Directory.CreateDirectory(folder);
                }

                string filePath = folder + @"\Protocol_" + teamName + "_" + channelName + ".html";

                using (StreamWriter protocol = new StreamWriter(filePath))
                {
                    protocol.WriteLine(disclaimer);
                    protocol.WriteLine("<br>");
                    protocol.WriteLine(@"<div><font face=""Arial""><strong>" + header.Replace("Current", "</div><div>").Replace("Date", "</div><div>Date ") + "</strong></div>");
                    protocol.WriteLine("<br>");
                    protocol.WriteLine("<div>" + "<div><strong>Channel Member:</strong></div>" + "</div>");
                    protocol.WriteLine("<br>");
                    foreach (string s in channelMembers)
                    {
                        protocol.WriteLine("<li>" + s + "</li>");
                    }
                    protocol.WriteLine("<br>");
                    protocol.WriteLine("<div><strong>Chat:</strong></font></div>");
                    foreach (string s in messages)
                    {
                        if (s.Contains("*MESSAGE*")) { protocol.WriteLine("<br>"); }
                        protocol.WriteLine(@"<div><font face = ""Arial"">" + s.ToString().Replace(Environment.NewLine, " </div><div>") + "</font></div>");
                    }
                }
                return filePath;
            }
            else
            {
                return "";
            }
        }

        private string createProtocol_txt(string teamName, string channelName, string userName, string startDate, string endDate, bool active) //Returns the filepath of the protocol
        {
            string folder = @".\tmp";
            if (!System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }

            if (active)
            {
                //Create Protocol
                string disclaimer = "Disclaimer" + Environment.NewLine +
                                    "Information contained in this document and any attachments maybe privileged or confidential and intended for the exclusive use of the owner." +
                                    "If you have received this document by mistake, please delete it and advise the owner immediately.";

                string header = "Chat Protocol - " + teamName + "-" + channelName + Environment.NewLine +
                            userName + " Date from: " + startDate + " to " + endDate;

                string filePath = folder + @"\Protocol_" + teamName + "_" + channelName + ".txt";

                using (StreamWriter protocol = new StreamWriter(filePath))
                {
                    protocol.WriteLine(disclaimer);
                    protocol.WriteLine("");
                    protocol.WriteLine(header);
                    protocol.WriteLine("");
                    protocol.WriteLine("Channel Member:");
                    foreach (string s in channelMembers)
                    {
                        protocol.WriteLine(s);
                    }
                    protocol.WriteLine("");
                    protocol.WriteLine("Chat:");
                    foreach (string s in messages)
                    {
                        protocol.WriteLine(s.ToString());
                    }

                }
                return filePath;
            }
            else
            {
                return "";
            }
        }

        private async Task<DriveItem> createFolder(string name, string teamId, string channelId, bool active)
        {
            try
            {
                //Create Folder
                var driveItem = new DriveItem
                {
                    Name = name,
                    Folder = new Folder { },
                    AdditionalData = new Dictionary<string, object>()
                    { {"@microsoft.graph.conflictBehavior", "fail" } }
                };

                Team team = await authClient.Teams[teamId].Request().GetAsync();
                Channel channel = await authClient.Teams[teamId].Channels[channelId].Request().GetAsync();

                DriveItem chFolder = new DriveItem();
                try
                {
                    chFolder = await authClient.Teams[teamId].Channels[channelId].FilesFolder.Request().GetAsync(); // ToDo -> Prüfen ob der Folder gefunden wurde
                }
                catch
                {
                    chFolder = null;
                }

                if (chFolder != null)
                {
                    try
                    {
                        return (DriveItem)await authClient.Drives[chFolder.ParentReference.DriveId].Items[chFolder.Id].ItemWithPath(name).Request().GetAsync();
                    }
                    catch (Exception)
                    {
                        return await authClient.Drives[chFolder.ParentReference.DriveId].Root.ItemWithPath(chFolder.Name).Children.Request().AddAsync(driveItem);
                    }
                }
                else
                {
                    MessageBoxButton buttons = MessageBoxButton.YesNo;
                    MessageBoxResult result;
                    result = MessageBox.Show("Channel folder not found! Write to General Folder?", "Channel folder missing", buttons);
                    if (result == MessageBoxResult.Yes)
                    {
                        return await authClient.Drives[chFolder.ParentReference.DriveId].Root.ItemWithPath("General").Children.Request().AddAsync(driveItem);
                    }
                    else
                    {
                        return null;
                    }
                }

            }
            catch (Exception e)
            {
                tbxSysMessages.Text = e.Message;
                return null;
            }
        }

        private async Task<DriveItem> uploadFile(string fileName, string file, string uploadFolder, string teamId, string channelId, bool rootFolder, bool active)
        {
            DriveItem folder = await createFolder(uploadFolder, teamId, channelId, true);

            if (folder != null)
            {
                using (var stream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(file)))
                {
                    return await authClient.Drives[folder.ParentReference.DriveId].Items[folder.Id].ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(stream);
                }
            }
            else
            {
                return null;
            }

        }

        private async void sendProtocolMessage(string teamName, string channelName, string teamId, string channelId, string userName, string startDate, string endDate, bool active)
        {
            if (active)
            {
                //Send Protocol Message
                string attId = Guid.NewGuid().ToString();

                string chmsg = "<div><strong>Chat Protocol - " + teamName + "-" + channelName + " created.</div><div>" +
                                 userName + " </div><div>Date from: " + startDate + " to " + endDate + "</strong></div>" +
                                 "<attachment id =\"" + attId + "\"></attachment>";

                var chatMessage = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        ContentType = BodyType.Html,
                        Content = chmsg
                    },
                };

                await authClient.Teams[teamId].Channels[channelId].Messages.Request().AddAsync(chatMessage);
            }
        }
    }
}
