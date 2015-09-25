Function SaveAttachmentsByEnteredDate()
Dim objOL As Outlook.Application
Dim oNS As Outlook.NameSpace
Dim oFolder As Outlook.Folder
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objItems As Outlook.items
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String
Dim strDeletedFiles As String
Dim strFind As String
Dim startDate As Date
Dim endDate As Date
Dim quoM As String
Dim tempDate As Date

' Get the path to your folder
strFolderpath = "C:\Attachments\"
On Error Resume Next

quoM = Chr$(34) ' Quotation mark character

startDate = DateAdd("d", -7, Date) ' Subtracting 7 days as the default start date

startDate = InputBox("Start date of file to read between? Use Format MM/DD/YYYY", "Enter Start Date", startDate) ' Ask user for start date
endDate = InputBox("End date of file to read between? Use Format MM/DD/YYYY", "Enter End Date", Date) ' Ask user for end date

' Format and create the find string to be used on the folders restrict search
strFind = "[ReceivedTime] <= " & quoM & endDate & " 11:59 PM" & quoM & " AND [ReceivedTime] > " & quoM & startDate & " 0:01 AM" & quoM

' Instantiate an Outlook Application object.
Set objOL = CreateObject("Outlook.Application")
' The GetNameSpace method is functionally equivalent to the Session property.
Set oNS = objOL.GetNamespace("MAPI")
' Set the folder to the destination folder, top down to get to the right folder.
Set oFolder = oNS.Folders("4JLabs").Folders("Inner Folder 1").Folders("Inner Inner Folder 1")
' Use the bool check statement written above by the users input to restrict the files by date.
Set objItems = oFolder.items.Restrict(strFind)

' Foreach mail in folder (after restriction) - For Each objMsg In MyFolder1.Items to be used for no restriction
For Each objMsg In objItems

' Get the Attachments collection of the item.
Set objAttachments = objMsg.Attachments

' Count attachments on the email
lngCount = objAttachments.count
strDeletedFiles = ""

' If an attachment exists
If lngCount > 0 Then

' We need to use a count down loop for items
' from a collection. Otherwise, the loop counter gets
' confused and only every other item is seen.
For i = lngCount To 1 Step -1

    ' Check for file named HelloWorld.txt
    If objAttachments.Item(i).fileName = "HelloWorld.txt" Then
	
	' Set the string to the file name or use the below to set it to a custom format, 
	' especially useful if the file coming in is named the same on multiple emails it will just overwrite the previous
    ' strFile = objAttachments.Item(i).fileName
    
    ' Set the file name to a dated normal format
    strFile = "HelloWorld_" & Format(objMsg.ReceivedTime, "MMDDYYYY") & ".txt"
    
    ' Combine with the path to the Temp folder.
    strFile = strFolderpath & strFile
    
    ' Save the attachment as a file.
    objAttachments.Item(i).SaveAsFile strFile
    End If
    
    ' Delete the attachment. Use this is you are going to delete attachment from the email
    'objAttachments.Item(i).Delete
Next i

End If
Next

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objOL = Nothing
End Function
