Attribute VB_Name = "SaveAttachments"
' Author: @alyssa.t.umanos
' Date: Feb 2022
'
' Sources:
'   -   https://stackoverflow.com/questions/15531093/save-attachments-to-a-folder-and-rename-them
'   -   https://stackoverflow.com/questions/48446838/reference-a-folder-not-under-the-default-inbox/48450185
'   -   https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem.move

Option Explicit

Public Sub SaveAttachments()
    Dim objOL As Outlook.Application
    Set objOL = CreateObject("outlook.application")

    Dim objSlctFolder As Outlook.MAPIFolder
    Dim objRecordedFolder As Outlook.MAPIFolder
    Dim objMsg As Outlook.MailItem 'Object
    Dim objAttachments As Outlook.Attachments
    Dim i As Long
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderpath As String
    Dim strDeletedFiles As String
    Dim selectedOlFolder As String
    
    ' Get the path to your My Documents folder
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
    strFolderpath = strFolderpath & "\Attachments\"
    
    'On Error Resume Next
    Debug.Print (separator1)
    Debug.Print ("STARTING PROCESS...")
    selectedOlFolder = FolderPick
    
    If selectedOlFolder <> "Cancel" Then
        ' Gets parent folder of the default inbox
        Set objSlctFolder = objOL.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Parent
        ' Gets the Recorded folder from the parent folder
        Set objRecordedFolder = objSlctFolder.Folders("Recorded")
        ' Gets the selected folder from the parent folder
        Set objSlctFolder = objSlctFolder.Folders(selectedOlFolder)
        
        ' Loop through mails in folder
        For Each objMsg In objSlctFolder.Items
            ' This code only strips attachments from mail items.
            ' If objMsg.class=olMail Then
            ' Get the Attachments collection of the item.
            Set objAttachments = objMsg.Attachments
            lngCount = objAttachments.Count
            strDeletedFiles = ""

            If lngCount > 0 Then
                Debug.Print (separator2)
                Debug.Print ("Working on email...'" & objMsg & "'")
                Debug.Print (Chr(10))

                ' We need to use a count down loop for removing items
                ' from a collection. Otherwise, the loop counter gets
                ' confused and only every other item is removed.
                For i = 1 To lngCount Step -1
                    ' Save attachment before deleting from item.
                    ' Get the file name.
                    strFile = objAttachments.Item(i).FileName
                    Debug.Print strFile
                    ' Combine with the path to the Temp folder.
                    strFile = strFolderpath & strFile
                    Debug.Print ("Saving..." & strFile)

                    ' Save the attachment as a file.
                    objAttachments.Item(i).SaveAsFile strFile

                    ' Delete the attachment.
                    ' Author's note: Removed this line below to retain the attachments in the emails.
                    ' objAttachments.Item(i).Delete

                    'write the save as path to a string to add to the message
                    'check for html and use html tags in link
                    If objMsg.BodyFormat <> olFormatHTML Then
                        strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
                    Else
                        strDeletedFiles = strDeletedFiles & "<br>" & "<a href='file://" & _
                        strFile & "'>" & strFile & "</a>"
                    End If

                    'Use the MsgBox command to troubleshoot. Remove it from the final code.
                    MsgBox strDeletedFiles
                Next
                ' Adds the filename string to the message body and save it
                ' Check for HTML body
                If objMsg.BodyFormat <> olFormatHTML Then
                    objMsg.Body = vbCrLf & "The file(s) were saved to " & strDeletedFiles & vbCrLf & objMsg.Body
                Else
                    objMsg.HTMLBody = "<p>" & "The file(s) were saved to " & strDeletedFiles & "</p>" & objMsg.HTMLBody
                End If

                ' Moves the mail to Recorded folder and marks it as read
                objMsg.Save
                objMsg.UnRead = False
                objMsg.Move objRecordedFolder
                
            Else
                ' Placeholder
            End If
            
        Next
    Else
        Debug.Print (separator2)
        Debug.Print ("No folder selected. Process termindated.")
        Debug.Print (separator1)
        Exit Sub
    End If
    Debug.Print (separator2)
    Debug.Print ("PROCESS SUCCESSFUL.")
    Debug.Print (separator1)
Exit Sub:
    Set objOL = Nothing
    Set objSlctFolder = Nothing
    Set objMsg = Nothing
    Set objAttachments = Nothing
End Sub

' https://stackoverflow.com/questions/12688476/display-dialogue-to-allow-a-user-to-select-an-outlook-folder-in-vba
' Modified
Function FolderPick() As String

    Dim objNS As NameSpace
    Dim objFolder As Folder
    
    Set objNS = Application.GetNamespace("MAPI")
    Set objFolder = objNS.PickFolder
    
    If TypeName(objFolder) <> "Nothing" Then
        'rtrnMsg = vbCr & " objFolder: " & objFolder
        FolderPick = objFolder
    Else
        'rtrnMsg = vbCr & "Cancel"
        FolderPick = "Cancel"
    End If

    Set objFolder = Nothing
    Set objNS = Nothing
    
End Function

Function separator1() As String
    Dim sprtr As String
    Dim j As Integer
    sprtr = ""
    For j = 1 To 100 Step 1
        If j = 1 Then
            sprtr = Chr(10) & "="
        Else
            sprtr = sprtr & "="
        End If
    Next j
    separator1 = sprtr
End Function

Function separator2() As String
    Dim sprtr As String
    Dim j As Integer
    sprtr = ""
    For j = 1 To 100 Step 1
        If j = 1 Then
            sprtr = Chr(10) & "_"
        Else
            sprtr = sprtr & "_"
        End If
    Next j
    separator2 = sprtr
End Function
