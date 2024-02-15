Attribute VB_Name = "Module2"
'Sub RemoveAttachmentBeforeForwarding()

'    Dim myolApp As Outlook.Application

'    Dim myinspector As Outlook.Inspector

'    Dim myItem As Outlook.MailItem

'    Dim myattachments As Outlook.Attachments

'    Set myolApp = CreateObject("Outlook.Application")

'    Set myinspector = myolApp.ActiveInspector

'    If Not TypeName(myinspector) = "Nothing" Then

'        Set myItem = myinspector.CurrentItem.Forward

'        Set myattachments = myItem.Attachments

'        While myattachments.Count > 0

'               myattachments.Remove 1

'        Wend

'        myItem.Display

'        myItem.Recipients.Add "Dan Wilson"

'        myItem.Send

'    Else

'        MsgBox "There is no active inspector."

'    End If

'End Sub

 

Sub Confirmation_Letter()

Dim objMail As Outlook.MailItem

Dim objAttachments As Outlook.Attachments

Dim i As Integer

Dim n As Integer

 

Set objItem = GetCurrentItem()

Set objMail = objItem.Forward

Set objAttachments = objMail.Attachments

 

 

n = objItem.Recipients.Count

For i = 1 To n

    objMail.To = objMail.To & "; " & objItem.Recipients.Item(i).Address  '& Chr(44)  & Chr(7)

Next i

 

 

objMail.Subject = Left(objItem.Subject, 40)

objMail.HTMLBody = objItem.HTMLBody

 

'---moje rozwiazanie dot kasowania zalacznikow ------------------------------------------------------------------------------

'n = objAttachments.Count

'If n > 0 Then

'    For i = 1 To n

'        objAttachments.Item(1).Delete

'    Next i

'End If

'---ponizej rozwiazanie z biblioteki ------------------------------------------------------------------------------

 

While objAttachments.Count > 0

         objAttachments.Remove 1

Wend

 

objMail.Display

 

 

Set objItem = Nothing

Set objMail = Nothing

End Sub

 

 

Sub LGE_MA_PO_to_Pantos_WAW()

Dim objMail As Outlook.MailItem

Dim myBody As Object

 

Set objItem = GetCurrentItem()

Set objMail = objItem.Forward

 

 

objMail.To = "agata.ambroziak@pantos.com; pawel.wasiak@pantos.com; "

objMail.Subject = objMail.Subject & " - Invoice No?"

objMail.HTMLBody = "<HTML><BODY>Dear Agata,</BODY></HTML>" _

        & "Please inform which shipments Invoice No will be used for this PO?</BODY></HTML>" _

        & "<HTML><BODY>~</BODY></HTML>" _

        & "<HTML><BODY>Pozdrawiam</BODY></HTML>" _

        & "<HTML><BODY>Rafal</BODY></HTML>" _

        & objMail.HTMLBody

 

 

objMail.Display

 

 

Set objItem = Nothing

Set objMail = Nothing

End Sub

 

Sub PO_with_DDD_to_Pantos_WAW()

Dim objMail As Outlook.MailItem

Dim myBody As Object

Dim objAttachments As Outlook.Attachments

 

Set objItem = GetCurrentItem()

Set objMail = objItem.Forward

Set objAttachments = objMail.Attachments

 

objMail.To = "agata.ambroziak@pantos.com; pawel.wasiak@pantos.com;"

 

objMail.Subject = objMail.Subject & " - "

objMail.HTMLBody = "<HTML><BODY>Dear All,</BODY></HTML>" _

        & "Please find in below DDD Invoice number</BODY></HTML>" _

        & "<HTML><BODY>~</BODY></HTML>" _

        & "<HTML><BODY>Pozdrawiam</BODY></HTML>" _

        & "<HTML><BODY>Rafal</BODY></HTML>" _

        & objMail.HTMLBody

 

 

'objMail.Attachments.Add (Excel.Applicaton.ActiveWorkbook.FullName)

'ActiveWorkbook.Close False

 

objMail.Display

 

 

Set objItem = Nothing

Set objMail = Nothing

End Sub

 

Sub ForwardB()

Dim objMail As Outlook.MailItem

Set objItem = GetCurrentItem()

Set objMail = objItem.Forward

objMail.To = "(E-Mail Removed)"

objMail.Display

Set objItem = Nothing

Set objMail = Nothing

End Sub

 

Function GetCurrentItem() As Object

Dim objApp As Outlook.Application

Set objApp = Application

On Error Resume Next

Select Case TypeName(objApp.ActiveWindow)

 

Case "Explorer"

Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)

 

Case "Inspector"

Set GetCurrentItem = objApp.ActiveInspector.CurrentItem

 

Case Else

End Select

 

End Function

