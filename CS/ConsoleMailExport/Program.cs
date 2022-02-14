using DevExpress.DashboardCommon;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;

namespace ConsoleMailExport {
    class Program {
        const string smtpHost = "ENTER YOUR HOST NAME";
        const int smtpPort = 25;
        const string userName = "ENTER YOUR USER NAME";
        const string password = "ENTER YOUR PASSWORD";
        static void Main(string[] args) {
            //Get information about teams
            List<TeamInformation> teams = GetTeamsInformation();
            //Send Dashboard to each team
            foreach(var team in teams) {
                SendDashboardToTeamAsync(team);
            }
        }
        static List<TeamInformation> GetTeamsInformation() {
            List<TeamInformation> teams = new List<TeamInformation>();
            teams.Add(new TeamInformation() {
                TeamName = "Management Team",
                TeamMail = "management@somewhere.com",
                TeamRegions = new List<string> { "Eastern", "Northern", "Southern", "Western" }
            });
            teams.Add(new TeamInformation() {
                TeamName = "Eastern Team",
                TeamMail = "eastern@somewhere.com",
                TeamRegions = new List<string> { "Eastern" }
            });
            teams.Add(new TeamInformation() {
                TeamName = "Northern Team",
                TeamMail = "northern@somewhere.com",
                TeamRegions = new List<string> { "Northern" }
            });
            teams.Add(new TeamInformation() {
                TeamName = "Southern Team",
                TeamMail = "southern@somewhere.com",
                TeamRegions = new List<string> { "Southern" }
            });
            teams.Add(new TeamInformation() {
                TeamName = "Western Team",
                TeamMail = "western@somewhere.com",
                TeamRegions = new List<string> { "Western" }
            });
            return teams;
        }
        private static async void SendDashboardToTeamAsync(TeamInformation team) {
            //Create Mail Message
            using(MimeMessage mail = CreateMimeMessage(team)) {
                if(mail == null) {
                    Console.WriteLine($"Sending Dashboard for the {team.TeamName} fail.");
                    return;
                }
                //Send Mail Message
                using(var client = new SmtpClient()) {
                    try {
                        client.Connect(smtpHost, smtpPort, SecureSocketOptions.Auto);
                        client.Authenticate(userName, password);
                        await client.SendAsync(mail);
                        Console.WriteLine($"Dashboard for the {team.TeamName} was sent.");
                    }
                    catch(Exception ex) {
                        Console.WriteLine(
                            $"Sending Dashboard for the {team.TeamName} fail. The following error occurs: {ex.Message}");
                    }
                    client.Disconnect(true);
                }
            }
        }
        private static MimeMessage CreateMimeMessage(TeamInformation team) {
            try {
                // Create Mail Message
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("Someone", "someone@somewhere.com"));
                message.To.Add(new MailboxAddress(team.TeamName, team.TeamMail));
                message.Subject = "Freight Dashboard";
                var builder = new BodyBuilder();
                builder.TextBody = "This is a test e-mail message sent by an application.";
                //Create Filter State based on Regions
                DashboardState regionFilterState = new DashboardState();
                DashboardItemState itemState = new DashboardItemState("listBoxDashboardItem1");
                foreach(string region in team.TeamRegions) {
                    itemState.MasterFilterValues.Add(new object[] { region });
                }
                regionFilterState.Items.Add(itemState);
                //Create Exporter
                DashboardExporter exporter = new DashboardExporter();
                exporter.ConnectionError += Exporter_ConnectionError;
                exporter.DashboardItemDataLoadingError += Exporter_DashboardItemDataLoadingError;
                exporter.DataLoadingError += Exporter_DataLoadingError;
                // Export Dashboard to PDf and attach to the Mail Message.
                using(MemoryStream stream = new MemoryStream()) {
                    exporter.ExportToPdf(
                        dashboardXmlPath: "Data/MailDashboard.xml",
                        outputStream: stream,
                        dashboardSize: new System.Drawing.Size(2000, 1000),
                        state: regionFilterState);
                    stream.Seek(0, System.IO.SeekOrigin.Begin);
                    builder.Attachments.Add("Dashboard.pdf", stream.ToArray(), new ContentType("application", "pdf"));
                }
                message.Body = builder.ToMessageBody();
                return message;
            }
            catch{ return null;}
        }
        static void Exporter_ConnectionError(object sender,
            DashboardExporterConnectionErrorEventArgs e) {
            Console.WriteLine(
                $"The following error occurs in {e.DataSourceName}: {e.Exception.Message}");
            throw new Exception();
        }
        static void Exporter_DataLoadingError(object sender,
            DataLoadingErrorEventArgs e) {
            foreach(DataLoadingError error in e.Errors)
                Console.WriteLine(
                    $"The following error occurs in {error.DataSourceName}: {error.Error}");
            throw new Exception();
        }
        static void Exporter_DashboardItemDataLoadingError(object sender,
            DashboardItemDataLoadingErrorEventArgs e) {
            foreach(DashboardItemDataLoadingError error in e.Errors)
                Console.WriteLine(
                    $"The following error occurs in {error.DashboardItemName}: {error.Error}");
            throw new Exception();
        }
    }
    public class TeamInformation {
        public string TeamName { get; set; }
        public string TeamMail { get; set; }
        public List<string> TeamRegions { get; set; }
    }
}
