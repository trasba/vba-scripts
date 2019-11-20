Sub StandardInbox()
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
    MsgBox ("Finished")
    Debug.Print "Finished"
    
End Sub
Function ProcessFolders(ByVal objCurrentFolder As Outlook.Folder)
    If objCurrentFolder.DefaultItemType = olMailItem Then
       Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
       ulFlags = 0
       ulFlags = ulFlags Or &H0
       Debug.Print objCurrentFolder.Name
       For i = objCurrentFolder.Items.Count To 1 Step -1
           'DoEvents to retain responsiveness
            DoEvents
            Set objMail = objCurrentFolder.Items(i)
            Debug.Print objMail.Subject
            If objMail.Class = "olMail" Or objMail.Class = "43" Then
              'Debug.Print "Class: " & objMail.Class
              'exclude not donloaded (only header) mails
              'SMIME check has to happen before PropertyAccessor because when the object ist "opened" it is no SMIME anymore
              If objMail.DownloadState <> olHeaderOnly And objMail.MessageClass = "IPM.Note.SMIME" Then
                  'Debug.Print "DownloadState: " & objMail.DownloadState
                  'Debug.Print objMail.MessageClass
                  oProp = CLng(objMail.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS))
                  'Debug.Print objMail.MessageClass
                  'If oProp = 1 Then
                  'If objMail.MessageClass = "IPM.Note.SMIME" Then
                    'Debug.Print "Processing Mail with Subject: " & objMail.Subject
                    'Debug.Print objMail.Parent
                    objMail.PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
                    objMail.Save
                    'Return '(end after first hit)
                  'End If
              End If
            End If
        Next
    End If
End Function

Sub GetSelectedItems()
 
 Dim myOlExp As Outlook.Explorer
 Dim myOlSel As Outlook.Selection
 Dim MsgTxt As String
 Dim x As Integer
 
 Const PR_SECURITY_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x6E010003"
 ulFlags = 0
 ulFlags = ulFlags Or &H0
       
 'MsgTxt = "You have selected items from: "
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 For x = 1 To myOlSel.Count
    'DoEvents to retain responsiveness
    DoEvents
    'MsgTxt = MsgTxt & myOlSel.Item(x).SenderName & ";"
    Debug.Print myOlSel.Item(x).MessageClass
    If myOlSel.Item(x).MessageClass = "IPM.Note.SMIME" Then
        Debug.Print "YES"
    End If
    Debug.Print myOlSel.Item(x).Class
    Debug.Print myOlSel.Item(x).DownloadState
    
    myOlSel.Item(x).PropertyAccessor.SetProperty PR_SECURITY_FLAGS, ulFlags
    myOlSel.Item(x).Save
 Next x
 'MsgBox MsgTxt
 MsgBox ("Ende")
 End Sub