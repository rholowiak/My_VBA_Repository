Attribute VB_Name = "Module1"

Sub MarkAllItemsAsRead_In_BRAIN_Geloeschte_Objekte_Folder()
    Dim objOutlook As Object
    Dim objnSpace As Object
    Dim objFolder As Outlook.Folder
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objnSpace = objOutlook.GetNamespace("MAPI")

    On Error Resume Next
    Set objFolder = objnSpace.Folders("Global BRAIN Connection").Folders("Gelöschte Objekte")
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "No such folder."
        Exit Sub
    End If
    'Process mail folder
    If objFolder.DefaultItemType = olMailItem Then
        Call ProcessFolders(objFolder)
    End If
    
    Set objOutlook = Nothing
    Set objnSpace = Nothing
    Set objFolder = Nothing
End Sub

Sub ProcessFolders(ByVal objCurFolder As Outlook.Folder)
    Dim objUnreadItems As Outlook.Items
    Dim i As Integer
    Dim n As Integer
    Dim m As Integer
    Dim objItem As Object
    Dim objSubFolder As Outlook.Folder
 
    Set objUnreadItems = objCurFolder.Items.Restrict("[Unread]=True")
 
    'Mark all unread emails as read
    For i = objUnreadItems.Count To 1 Step -1
        Set objItem = objUnreadItems.Item(i)
        objItem.UnRead = False
        'objItem.Save
    Next
    
    'Process subfolders recursively
    If objCurFolder.Folders.Count > 0 Then
       For Each objSubFolder In objCurFolder.Folders
           Call ProcessFolders(objSubFolder)
       Next
    End If
    
    Set objUnreadItems = Nothing
    Set objItem = Nothing
    Set objSubFolder = Nothing
End Sub

