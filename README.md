# PhishReporter-Outlook-Add-In
PhishReporter Outlook Add-In in an Outlook Add-In that allows users to report phishing e-mails to a specific e-mail address for further processing/investigation

This simple, yet efficient, Outlook Add-In adds a button to your Outlook Home Ribbon that allows users to simply select/highlight a phishing email and it will forward it to the appropriate mailbox/e-mail address as an attachment for further analysis.  Once the user has verified that they want to send this Phishing email, then the Outlook Add-In removes it from their inbox and places it in their “trash” folder.

There are tons of companies (now) offering this type of Add-In, but PhishReporter Outlook Add-In is completely free and can be completely customized.

## Requirements to Build/Customize the PhishReporter Outlook Add-In:

* PhishReporter Project Files - Clone this repo
* Visual Studio 2015 (Tested and working on Community and Professional)
* Visual Studio Installer Projects Extension: https://visualstudiogallery.msdn.microsoft.com/9abe329c-9bba-44a1-be59-0fbf6151054d
* Visual Studio Office Developer Tools
* To customize the plugin, just edit the setting in the PhishRepoerterConfig.vb file in the ForwardToAbuseAddin project

