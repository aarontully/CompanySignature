On Error Resume Next

'Setting up the script to work with the file system.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Connecting to Active Directory to get user’s data.
Set objSysInfo = CreateObject("ADSystemInfo")
Set UserObj = GetObject("LDAP://" & objSysInfo.UserName)
strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
SigFolder = StrAppData & "\Microsoft\Signatures\"
SigFileFull = SigFolder & "Full Signature.htm"
SigFileReply = SigFolder & "Reply Signature.htm"

'Setting placeholders for the signature.
strUserName = UserObj.sAMAccountName
strTitle = UserObj.title
strFirstName = UserObj.Firstname
strFullName = UserObj.displayname
strShortName = UserObj.personalTitle
strDepartment = UserObj.department
strCompany = UserObj.Company
strCred = UserObj.info
strPhone = UserObj.TelephoneNumber
strMobile = UserObj.Mobile
strFax = UserObj.FacsimileTelephoneNumber
strWeb = UserObj.wwwHomePage
strEmail = UserObj.mail
strOfficePhone = UserObj.TelephoneNumber

'Setting global placeholders for the signature. Those values will be identical for all users - make sure to replace them with the right values!
strCompanyLogo = """\\path\to\logo"
strCompanyAddress1 = "1 Demo St, Demoville NSW Australia"
strCompanyDetails = "Yadda yadda yadda"

'Creating HTM signature file for the user's profile, if the file with such a name is found, it will be overwritten.
Set CreateSigFile = FileSysObj.CreateTextFile (SigFileFull, True, True)

'HTML code for Full Signature
CreateSigFile.WriteLine "<!DOCTYPE HTML>"
CreateSigFile.WriteLine "<HTML><HEAD><TITLE>Email Signature</TITLE>"
CreateSigFile.WriteLine "</HEAD>"
CreateSigFile.WriteLine "<BODY style='font-size: 9pt; font-family: Verdana, sans-serif;'>"
if strShortName <> "" Then
    CreateSigFile.WriteLine "<span>" & strShortName & "</span>"
Else 
    CreateSigFile.WriteLine "<span>" & strFirstName & "</span>"
End if
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<span style='display: block; font-weight: bold; line-height: 1em;'>" & strFullName & "</span>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<span>" & strTitle & "</span>"
CreateSigFile.WriteLine "<br>"
if strDepartment <> "" Then
    CreateSigFile.WriteLine "<span>" & strDepartment & "</span>"
End if
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<span><a href='#'><img src=" & """\\path\to\logo" & """ alt=" & """" & """ width=150 height=50 border=0></a></span>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<span style='font-weight: bold;'>" & strCompany & "</span>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<span>" & strCompanyDetails & "</span>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Switchboard:     +## # #### #### </pre>"
if strPhone <> "" Then
    CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Direct Line:        " & strPhone & "</pre>"
End if
if strFax <> "" Then
    CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Fax:                  " & strFax & "</pre>"
End if
if strMobile <> "" Then
    CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Mobile:              " & strMobile & "</pre>"
End if
CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Address:           " & strCompanyAddress1 & "</pre>"
CreateSigFile.WriteLine "<br>"
CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Email:              " & strEmail & "</pre>"
CreateSigFile.WriteLine "<pre style='font-size: 9pt; font-family: Verdana'>Web Page:        " & strWeb & "</pre>"
CreateSigFile.WriteLine "</BODY>"
CreateSigFile.WriteLine "</HTML>"
CreateSigFile.Close

'Applying the signature in Outlook’s settings.
Set objWord = CreateObject("Word.Application")
Set objSignatureObjects = objWord.EmailOptions.EmailSignature
objSignatureObjects.NewMessageSignature = "Full Signature"

'Setting the signature as default for new messages.
objSignatureObjects.NewMessageSignature = "Full Signature"

objWord.Quit

'Creating HTM signature file for the user's profile, if the file with such a name is found, it will be overwritten.
Set CreateReplySigFile = FileSysObj.CreateTextFile (SigFileReply, True, True)

'HTML code for Reply Signature
CreateReplySigFile.WriteLine "<!DOCTYPE HTML>"
CreateReplySigFile.WriteLine "<HTML><HEAD><TITLE>Email Reply Signature</TITLE>"
CreateReplySigFile.WriteLine "</HEAD>"
CreateReplySigFile.WriteLine "<BODY style='font-size: 9pt; font-family: Verdana, sans-serif;'>"
if strCred <> "" Then
    CreateReplySigFile.WriteLine "<span>" & strFirstName & ", " & strCred & "</span>"
Else
    if strShortName <> "" Then
        CreateReplySigFile.WriteLine "<span>" & strShortName & "</span>"
    Else
        CreateReplySigFile.WriteLine "<span>" & strFirstName & "</span>"
    End if
End if
CreateReplySigFile.WriteLine "<br>"
CreateReplySigFile.WriteLine "<br>"
CreateReplySigFile.WriteLine "<span>" & strTitle & "</span>"
CreateReplySigFile.Close

'Applying the signature in Outlook’s settings.
Set objWord = CreateObject("Word.Application")
Set objSignatureObjects = objWord.EmailOptions.EmailSignature
objSignatureObjects.ReplyMessageSignature = "Reply Signature"

'Setting the signature as default for reply messages.
objSignatureObjects.ReplyMessageSignature = "Reply Signature"

objWord.Quit