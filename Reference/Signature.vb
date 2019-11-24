'------------------------------------------------------------------------------------------------------------------
'Option Explicit
On Error Resume Next
'==================================================
'Create Outlook signature from Word template
'==================================================

'----- Declarations -----
Const wdWord = 2
Const wdParagraph = 4
Const wdExtend = 1
Const wdCollapseEnd = 0

'--------------------------------------------------------------
'----- Modify these variables appropriately ----
'--------------------------------------------------------------
strTemplatePath = "\\share\path\to\script"
strTemplateName = "Sig_Template.docx"
strReplyTemplateName = "Sig_Reply_Template.docx"


'----- Connect to AD and get user info -----'
Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strFirstname = objUser.FirstName
strLastName = objUser.givenName
strInitials = objUser.initials
strName = objUser.FullName
strTitle = objUser.Title
strDescription = objUser.Description
strOffice = objUser.physicalDeliveryOfficeName
strCompany = objUser.company
strCred = objUser.info
strStreet = objUser.StreetAddress
strCity = objUser.l
strState = objUser.st
strPostCode = objUser.PostalCode
strPhone = objUser.TelephoneNumber
strMobile = objUser.Mobile
strFax = objUser.FacsimileTelephoneNumber
strEmail = objUser.mail
strWeb = objuser.wWWHomePage

'----- Apply any modifications to Active Directory fields -----
'Use company info page if user does not have a Linked-In account specified
 if strweb = "" Then strweb = "http://www.example.edu"

'----- Open Word template in read-only mode {..Open(filename,conversion,readonly)} -----
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open(strTemplatePath & strTemplateName,,True)
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries


'----- Replace template text placeholders with user specific info -----
SearchAndRep "[Name]", strName, objWord
SearchAndRep "[Title]", strTitle, objWord
SearchAndRep "[Company]", strCompany, objWord
SearchAndRep "[Street]", strStreet, objWord
SearchAndRep "[City]", strCity, objWord
SearchAndRep "[State]", strState, objWord
SearchAndRep "[Zip]", strPostCode, objWord
SearchAndRep "[Phone]", "Phone: " & strPhone, objWord
SearchAndRep "[email]", strEmail, objWord
SearchAndRep "[web]", strWeb, objWord

'----- Replace template text placeholders with cell number if there is one -----
if strMobile = "" Then 
	SearchAndRep "| [Mobile]", strMobile, objWord
else 
	SearchAndRep "[Mobile]", "Cell: " & strMobile, objWord
End if
'----- Replace template text placeholders with fax number if there is one -----
if strFax = "" Then 
	SearchAndRep "| [Fax]", strFax, objWord
else 
	SearchAndRep "[Fax]", "Fax: " & strFax, objWord
End if

'----- Replace template hyperlink placeholders with user specific info -----
SearchAndRepHyperlink "[email]", strWeb, objDoc
SearchAndRepHyperlink "[web]", strWeb, objDoc


'----- Set signature in Outlook -----
Set objSelection = objDoc.Range()
objSignatureEntries.Add "Signature", objSelection
objSignatureObject.NewMessageSignature = "Signature"

'see note below if a different reply signature is desired
objSignatureObject.ReplyMessageSignature = "Signature"


'----- Close signature template document -----
objDoc.Saved = TRUE
objDoc.Close
objWord.Quit

'----------------------------------------------------------------------------------------------------
'note...if a different reply signature is desired, copy above code from the 
'open template section.  This time through open 
'the reply template instead.
'-----------------------------------------------------------------------------------------------------


'----- Subrouting to search and replace template text placeholders -----
Sub SearchAndRep(searchTerm, replaceTerm, WordApp)
    WordApp.Selection.GoTo 1
    With WordApp.Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWholeWord = True
        .Text = searchTerm
        .Execute ,,,,,,,,,replaceTerm
    End With
End Sub


'----- Subrouting to search and replace template hyperlink placeholders -----
'         Note this can be picky...if it does not work re-create hyperlink in the template
Sub SearchAndRepHyperlink(searchLink, replaceLink, WordDoc)
	Set colHyperlinks = WordDoc.Hyperlinks
	For Each objHyperlink in colHyperlinks
	    If objHyperlink.Address = searchLink Then                                
        	objHyperlink.Address = replaceLink
            End If
	Next
End Sub
'---------------------------------------------------------------------------------------------------------------