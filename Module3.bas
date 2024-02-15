Attribute VB_Name = "Module3"
Sub RESEND_Mail_WITHOUT_attachments()
Dim objMail As Outlook.MailItem, objAttachments As Outlook.Attachments
Dim i As Integer, n As Integer
Set objItem = GetCurrentItem()
Set objMail = objItem.Forward
Set objAttachments = objMail.Attachments
n = objItem.Recipients.Count
For i = 1 To n
    objMail.To = objMail.To & "; " & objItem.Recipients.Item(i).Address
Next i
objMail.Subject = objItem.Subject
objMail.HTMLBody = objItem.HTMLBody
While objAttachments.Count > 0
         objAttachments.Remove 1
Wend
objMail.Display
Set objItem = Nothing
Set objMail = Nothing
End Sub
Sub RESEND_Mail_with_attachments()
Dim objMail As Outlook.MailItem, objAttachments As Outlook.Attachments
Dim i As Integer, n As Integer
Set objItem = GetCurrentItem()
Set objMail = objItem.Forward
n = objItem.Recipients.Count
For i = 1 To n
    objMail.To = objMail.To & "; " & objItem.Recipients.Item(i).Address
Next i
objMail.Subject = objItem.Subject
objMail.HTMLBody = objItem.HTMLBody
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
 
Sub HelloWorldMessage()
    Dim Msg As Outlook.MailItem
    Set Msg = Application.CreateItem(olMailItem)
    Msg.Display
    Msg.Subject = "Hello World!"
    Msg.HTMLBody = "<HTML><BODY style=font-size:11pt;font-family:Arial;color:Blue>" & _
                    "<h4>Dear All,</h4>" & _
                    "<p>In attachment file" & _
                    "<br>Thank you.</p>" & _
                    "<hr></BODY></HTML>" & Msg.HTMLBody
    Set Msg = Nothing
End Sub
Sub Mail_SomeMessageWith_Signature()
    Dim OutApp As Object, OutMail As Object
    Dim StrBody As String, SigString As String, Signature As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
   
    StrBody = "<H3><B>Dear Customer Ron de Bruin</B></H3>" & _
              "Please visit this website to download the new version.<br>" & _
              "Let me know if you have problems.<br>" & _
              "<A HREF=""http://www.rondebruin.nl/tips.htm"">Ron's Excel Page</A>" & _
              "<br><br><B>Thank you</B>"
'Change only Rafal1.htm to the name of your signature
    SigString = Environ("appdata") & "\Microsoft\Signatures\Rafal1.htm"
    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    On Error Resume Next
   ' Change the mail address and subject in the macro before you run it.
    With OutMail
            .Display
            .To = "holowiak.r@lgdisplay.com"
            .cc = ""
            .BCC = ""
            .Subject = "This is the Subject line"
            .HTMLBody = StrBody & "<br>" & Signature
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
End Function


