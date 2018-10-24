Imports Microsoft.Office.Tools.Ribbon


Public Class HOME

    Dim objItem As Outlook.MailItem
    Dim objMsg As Outlook.MailItem
    Dim app As Outlook.Application
    Dim exp As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim Application As Outlook.Application
    Dim attachments As Outlook.Attachments
    Dim objOutlookAtt As Outlook.Attachment



    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Phishing.Click

        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

        If exp.Selection.Count Then
            Dim response = MsgBox("The selected message will be forwarded to " & PhishReporterConfig.SecurityTeamEmailAlias & vbCrLf & " and removed from your inbox.  Would you like to continue?", MsgBoxStyle.YesNo, "Report Phishing To Your Security Team")
            If response = MsgBoxResult.Yes Then
                ' Added loops for attaching/deleting selected items
                ' We need a notification for if they have clicked on any attachments or followed any links
                ' Adding this in as the phishclick variable
                ' Phishclick will then go on to the body of the e-mail to report if they have clicked any attachments or followed any links
                Dim phishclick As String
                phishclick = MsgBox("Did you click on any attachments or follow any links?", MsgBoxStyle.YesNo, "Report Phishing To Your Security Team")
                If phishclick = vbYes Then
                    phishclick = "Yes."
                ElseIf phishclick = MsgBoxResult.No Then
                    phishclick = "No."
                End If
                Dim phishEmail As Outlook.MailItem = exp.Selection(1)
                Dim reportEmail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)
                For Each phishEmail In exp.Selection
                    With phishEmail
                        reportEmail.Attachments.Add(phishEmail, Outlook.OlAttachmentType.olEmbeddeditem)
                    End With
                Next
                reportEmail.Subject = PhishReporterConfig.ReportEmailSubject & " - '" & phishEmail.Subject & "' - " & "User Clicked: " & phishclick
                reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias
                reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the Phishing Reporter Outlook plugin. Please review the attached phishing email." & vbNewLine & vbNewLine & "User Clicked/Followed Attachments/Links: " & phishclick

                If String.IsNullOrEmpty(PhishReporterConfig.RunbookURL) Then
                    reportEmail.Body = reportEmail.Body & "."
                Else
                    reportEmail.Body = reportEmail.Body & vbNewLine & "Then follow the process defined in " & PhishReporterConfig.RunbookURL
                End If

                reportEmail.Send()
                For Each phishEmail In exp.Selection
                    With phishEmail
                        phishEmail.Delete()
                    End With
                Next
            Else
            End If
        Else
            MsgBox("Please Select a message To Continue.", MsgBoxStyle.OkOnly, "Phishing Reporter - No E-Mail Message Selected")
        End If

    End Sub

End Class



