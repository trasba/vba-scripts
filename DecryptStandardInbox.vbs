Private Sub DecryptStandardInbox()
    Set oOutlook = CreateObject("Outlook.Application")
    Set oNamespace = oOutlook.GetNamespace("MAPI")
    If MsgBox("Decrypt all emails in Folder " & oNamespace.GetDefaultFolder(6).Name & " (no Subfolders)?", vbYesNo) = 6 Then
        Debug.Print "Confirm Folder: " & oNamespace.GetDefaultFolder(6).Name&; " = Yes"
        Set objFold = oNamespace.GetDefaultFolder(6)
        Debug.Print "Processing " & objFold.Items.Count & " Emails"
        Debug.Print "Processing Folder"
        Call ProcessFolders(objFold)
    Else
        Debug.Print "Confirm Folder: " & oNamespace.GetDefaultFolder(6).Name&; " = No"
    End If
    Debug.Print "Finished"
    
End Sub
Function ProcessFolders(ByVal objCurrentFolder As Outlook.Folder)
    If objCurrentFolder.DefaultItemType = olMailItem Then
       Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
       ulFlags = 0
       ulFlags = ulFlags Or &H0
       'Debug.Print objCurrentFolder.Name
       For i = objCurrentFolder.Items.Count To 1 Step -1
           If objCurrentFolder.Items(i).Class = olMail Then
              Set objMail = objCurrentFolder.Items(i)
                If objMail.DownloadState <> olHeaderOnly Then
              'Debug.Print objMail.MessageClass
              oProp = CLng(objMail.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
                
              If oProp = 1 Then
                Debug.Print "Processing Mail with Subject: " & objMail.Subject
                'Debug.Print objMail.Parent
                objMail.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
                objMail.Save
                'Return '(end after first hit)
                End If
              End If
            End If
        Next
    End If
End Function
