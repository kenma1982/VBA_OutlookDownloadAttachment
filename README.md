# VBA_OutlookDownloadAttachment

tested in VBA 7.1 Office16\Outlook

copy the content of attachmentdownload.txt into outlook\developer tab\visual Basic, edit the path of setting file as following and save.

'Variable Setup
Call loadsetting(aSetting, "**C:\mail\settings.txt**", FileType_list, Setting) 'load setting file

Edit settings.txt:
1. Change the path for the downloaded attachment and the file of email address list
2. download_extension, only download the files with described extension
3. rename the downloaded attachment with the datastamp (YES/NO)

Put the edited settings.txt as in the described path

example:

SaveToFolder = "C:\mail\"
AddressFile = "C:\mail\addresslist.txt"
download_extension = "pdf,jpg,xls,png"
datestamp = "YES"




Setup of outlook to run a script
https://www.slipstick.com/outlook/rules/outlooks-rules-and-alerts-run-a-script/

How to add Run a script in Rule if you cannot find it:
https://www.extendoffice.com/documents/outlook/4640-outlook-rule-run-a-script-missing.html
