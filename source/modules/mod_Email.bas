Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Email
' Level:        Framework module
' Version:      1.00
' Description:  Email related functions & subroutines
'
' Source/date:  Bonnie Campbell, May 2015
' Revisions:    BLC, 5/27/2015 - 1.00 - initial version
' =================================

' ---------------------------------
'  Functions
' ---------------------------------

' =================================
' FUNCTION:     SendEmailWithOutlook
' Description:  Sends an email silently (w/o user view) via outlook
' Parameters:   datDate - date value to be converted to fiscal year
'               blnFourDigits - flag to use 4 digits to represent the year (default True)
'               blnAddPrefix - flag to add a prefix to the result (default True)
'               strPrefix - prefix to be added to the string
' Returns:      variant for the fiscal year string or integer (e.g., "FY2010")
' Throws:       none
' References:   none
' Source/date:
'               http://www.access-programmers.co.uk/forums/showthread.php?t=229533
' Revisions:    BLC, 5/27/2015 - initial version
' =================================
Public Function SendEmailWithOutlook( _
    MessageTo As String, _
    Subject As String, _
    MessageBody As String)

    On Error GoTo Err_Handler

    ' Define app variable and get Outlook using the "New" keyword
    Dim olApp As New Outlook.Application
    Dim olMailItem As Outlook.MailItem  ' An Outlook Mail item
 
    ' Create a new email object
    Set olMailItem = olApp.CreateItem(olMailItem)

    ' Add the To/Subject/Body to the message and display the message
    With olMailItem
        .To = MessageTo
        .Subject = Subject
        .Body = MessageBody
        .Send       ' Send the message immediately
    End With

    ' Release all object variables
    Set olMailItem = Nothing
    Set olApp = Nothing

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FiscalYear[mod_Time])"
    End Select
    Resume Exit_Function
End Function

http://www.blueclaw-db.com/access_email_gmail.htm
Public Function send_email()

Set cdomsg = CreateObject("CDO.message")
With cdomsg.Configuration.Fields
.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'NTLM method
.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
.item("http://schemas.microsoft.com/cdo/configuration/smptserverport") = 587
.item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
.item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
.item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mygmail@gmail.com"
.item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "mypassword"
.Update
End With
' build email parts
With cdomsg
.To = "somebody@somedomain.com"
.From = "mygmail@gmail.com"
.Subject = "the email subject"
.TextBody = "the full message body goes here. you may want to create a variable to hold the text"
.Send
End With
    Set cdomsg = Nothing
End Function