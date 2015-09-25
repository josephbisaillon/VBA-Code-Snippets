Function SaveAttachmentsBySelection()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String
Dim tempDate As Date

' Get the path to your folder
strFolderpath = "C:\DestinationFolder\"
On Error Resume Next

' Stops the application to let the user highlight outlook files then once the user clicks ok on the msgbox it will begin the import
MsgBox ("Please highlight the items you want on outlook and then click on to import.")

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")

' Get the collection of selected objects.
Set objSelection = objOL.ActiveExplorer.Selection

' Check each selected item for attachments. If attachments exist,
' save them to the strFolderPath folder and strip them from the item.
For Each objMsg In objSelection

' Set the attachment of the mail item
Set objAttachments = objMsg.Attachments

' Count the number of attachments
lngCount = objAttachments.count
strDeletedFiles = ""

' If an attachment is seen
If lngCount > 0 Then

' We need to use a count down loop for items
' from a collection. Otherwise, the loop counter gets
' confused and only every other item is seen.

For i = lngCount To 1 Step -1

    ' Make sure the attachment is the right file (Remove this if you want all the files)
    If objAttachments.Item(i).fileName = "FileIAmLookingFor.txt" Then
    
    ' Set the file to the file name
    strFile = objAttachments.Item(i).fileName
    
    ' Set the file name to date format for attachments with the same name
    strFile = "FileIAmLookingFor_" & Format(objMsg.ReceivedTime, "MMDDYYYY") & ".txt"
    
    ' Combine with the path to the import folder.
    strFile = strFolderpath & strFile
    
    ' Save the attachment as a file.
    objAttachments.Item(i).SaveAsFile strFile
    End If
    

Next i

End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing
End Function
