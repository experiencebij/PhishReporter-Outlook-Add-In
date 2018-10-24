Public Class PhishReporterConfig
    ' Report Configuration
    Public Shared Property SecurityTeamEmailAlias As String = "jinnes@compuvision.biz"
    Public Shared Property ReportEmailSubject As String = "[Automated] Spam/Phishing Email Report"
    ' Link to the security team's runbook for handling phishing emails. If the variable is empty or not defined, defaults to a simplified message"
    ' https://corporatewiki/path/to/runbook
    Public Shared Property RunbookURL As String = "https://portal.compuvision.biz/pages/viewpage.action?pageId=19399049" & vbNewLine & vbNewLine & "Alternatively: Please go through the following steps in order to properly validate the spam/phishing e-mail:" & vbNewLine & "Research - Use MXToolbox and Google:" & vbNewLine & "- Domain:" & vbNewLine & "- Links lead to (If Any):" & vbNewLine & "- Message ID:" & vbNewLine & "- Multiple Users Affected (If Y, send e-mail to AllStaff/appropriate distribution group):" & vbNewLine & vbNewLine & "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & "GATHER INFORMATION (Call User) - Check the contents of the e-mail. Is this obvious spam from a throwaway? If yes, skip this part. If no, or if the user indicated that they clicked/opened anything, contact the user and gather the following details:" & vbNewLine & "- Do they know the person:" & vbNewLine & "- Do they do business with this domain:" & vbNewLine & vbNewLine & "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & "ONLY COMPLETE THIS SECTION IF USER(S) CLICKED/OPENED LINKS/ATTACHMENTS)" & vbNewLine & "- Password Reset:" & vbNewLine & "- Suspicious rules in Outlook (Check rules with user): " & vbNewLine & "- Suspicious rules in OWA (Check rules with user): " & vbNewLine & "- All suspicious rules removed?: " & vbNewLine & "- Malware Scan - Results: " & vbNewLine & "- Mailbox purged of malicious items?: " & vbNewLine & vbNewLine & "---------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & "OTHER NOTES (Anything additional that may be useful can be input here:" & vbNewLine & "ONCE COMPLETE: Close the ticket out."

    ' Ribbon Group Config
    Public Shared Property RibbonGroupName As String = "Report Spam/Phishing"

    ' Button Config
    Public Shared Property ButtonName As String = "Report Spam/Phishing"
    Public Shared Property ButtonHoverDescription As String = "Report a suspicious email to the CompuVision Information Security Team."
    Public Shared Property ButtonScreenTip As String = "Report spam/phishing emails"
    Public Shared Property ButtonSuperTip As String = "Use this button to report spam/phishing emails to the CompuVision Information Security team."

End Class

