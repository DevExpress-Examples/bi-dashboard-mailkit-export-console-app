Imports DevExpress.DashboardCommon
Imports MailKit.Net.Smtp
Imports MailKit.Security
Imports MimeKit
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace ConsoleMailExport
	Friend Class Program

		Private Const smtpHost As String = "ENTER YOUR HOST NAME"
		Private Const smtpPort As Integer = 25
		Private Const userName As String = "ENTER YOUR USER NAME"
		Private Const password As String = "ENTER YOUR PASSWORD"
		Shared Sub Main(ByVal args() As String)
			'Get information about teams
			Dim teams As List(Of TeamInformation) = GetTeamsInformation()
			'Send Dashboard to each team
			For Each team In teams
				SendDashboardToTeamAsync(team)
			Next team
		End Sub
		Private Shared Function GetTeamsInformation() As List(Of TeamInformation)
			Dim teams As New List(Of TeamInformation)()
			teams.Add(New TeamInformation() With {
				.TeamName = "Management Team",
				.TeamMail = "management@somewhere.com",
				.TeamRegions = New List(Of String) From {"Eastern", "Northern", "Southern", "Western"}
			})
			teams.Add(New TeamInformation() With {
				.TeamName = "Eastern Team",
				.TeamMail = "eastern@somewhere.com",
				.TeamRegions = New List(Of String) From {"Eastern"}
			})
			teams.Add(New TeamInformation() With {
				.TeamName = "Northern Team",
				.TeamMail = "northern@somewhere.com",
				.TeamRegions = New List(Of String) From {"Northern"}
			})
			teams.Add(New TeamInformation() With {
				.TeamName = "Southern Team",
				.TeamMail = "southern@somewhere.com",
				.TeamRegions = New List(Of String) From {"Southern"}
			})
			teams.Add(New TeamInformation() With {
				.TeamName = "Western Team",
				.TeamMail = "western@somewhere.com",
				.TeamRegions = New List(Of String) From {"Western"}
			})
			Return teams
		End Function
		Private Shared Async Sub SendDashboardToTeamAsync(ByVal team As TeamInformation)
			'Create Mail Message
			Using mail As MimeMessage = CreateMimeMessage(team)
				If mail Is Nothing Then
					Console.WriteLine($"Sending Dashboard for the {team.TeamName} fail.")
					Return
				End If
				'Send Mail Message
				Using client = New SmtpClient()
					Try
						client.Connect(smtpHost, smtpPort, SecureSocketOptions.Auto)
						client.Authenticate(userName, password)
						Await client.SendAsync(mail)
						Console.WriteLine($"Dashboard for the {team.TeamName} was sent.")
					Catch ex As Exception
						Console.WriteLine($"Sending Dashboard for the {team.TeamName} fail. The following error occurs: {ex.Message}")
					End Try
					client.Disconnect(True)
				End Using
			End Using
		End Sub
		Private Shared Function CreateMimeMessage(ByVal team As TeamInformation) As MimeMessage
			Try
				' Create Mail Message
				Dim message = New MimeMessage()
				message.From.Add(New MailboxAddress("Someone", "someone@somewhere.com"))
				message.To.Add(New MailboxAddress(team.TeamName, team.TeamMail))
				message.Subject = "Freight Dashboard"
				Dim builder = New BodyBuilder()
				builder.TextBody = "This is a test e-mail message sent by an application."
				'Create Filter State based on Regions
				Dim regionFilterState As New DashboardState()
				Dim itemState As New DashboardItemState("listBoxDashboardItem1")
				For Each region As String In team.TeamRegions
					itemState.MasterFilterValues.Add(New Object() { region })
				Next region
				regionFilterState.Items.Add(itemState)
				'Create Exporter
				Dim exporter As New DashboardExporter()
				AddHandler exporter.ConnectionError, AddressOf Exporter_ConnectionError
				AddHandler exporter.DashboardItemDataLoadingError, AddressOf Exporter_DashboardItemDataLoadingError
				AddHandler exporter.DataLoadingError, AddressOf Exporter_DataLoadingError
				' Export Dashboard to PDf and attach to the Mail Message.
				Using stream As New MemoryStream()
					exporter.ExportToPdf(dashboardXmlPath:= "Data/MailDashboard.xml", outputStream:= stream, dashboardSize:= New System.Drawing.Size(2000, 1000), state:= regionFilterState)
					stream.Seek(0, System.IO.SeekOrigin.Begin)
					builder.Attachments.Add("Dashboard.pdf", stream.ToArray(), New ContentType("application", "pdf"))
				End Using
				message.Body = builder.ToMessageBody()
				Return message
			Catch
				Return Nothing
			End Try
		End Function
		Private Shared Sub Exporter_ConnectionError(ByVal sender As Object, ByVal e As DashboardExporterConnectionErrorEventArgs)
			Console.WriteLine($"The following error occurs in {e.DataSourceName}: {e.Exception.Message}")
			Throw New Exception()
		End Sub
		Private Shared Sub Exporter_DataLoadingError(ByVal sender As Object, ByVal e As DataLoadingErrorEventArgs)
			For Each [error] As DataLoadingError In e.Errors
				Console.WriteLine($"The following error occurs in {[error].DataSourceName}: {[error].Error}")
			Next [error]
			Throw New Exception()
		End Sub
		Private Shared Sub Exporter_DashboardItemDataLoadingError(ByVal sender As Object, ByVal e As DashboardItemDataLoadingErrorEventArgs)
			For Each [error] As DashboardItemDataLoadingError In e.Errors
				Console.WriteLine($"The following error occurs in {[error].DashboardItemName}: {[error].Error}")
			Next [error]
			Throw New Exception()
		End Sub
	End Class
	Public Class TeamInformation
		Public Property TeamName() As String
		Public Property TeamMail() As String
		Public Property TeamRegions() As List(Of String)
	End Class
End Namespace
