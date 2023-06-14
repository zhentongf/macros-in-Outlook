Sub Test()
    Dim Response
    Response = MsgBox("Hello World", vbYesNo + vbCritical)
    Response = MsgBox("Argument not optional", vbExclamation)
    Response = MsgBox("", vbInformation)
    Response = MsgBox("", vbQuestion)
    
End Sub
Sub TestAttachment()
    Dim myinspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    Dim myAttachments As Outlook.Attachments
    Dim myAttachment As Attachment

    Set myinspector = Application.ActiveInspector
    If Not TypeName(myinspector) = "Nothing" Then
        Set myItem = myinspector.CurrentItem
        Set myAttachments = myItem.Attachments
        myItem.Display
        MsgBox (myAttachments.Count & " attachment(s)")
        
        For Each myAttachment In myAttachments
            Debug.Print ("Index: " & myAttachment.Index)
            Debug.Print ("DisplayName: " & myAttachment.DisplayName)
            Debug.Print ("FileName: " & myAttachment.FileName)
            Debug.Print ("Size: " & myAttachment.Size)
            Debug.Print ("Type: " & myAttachment.Type & vbLf)
        Next
    Else
        MsgBox ("There is no active inspector")
    End If
        
End Sub

Sub AutoForward_OneMail()
    
    Dim myinspector As Outlook.Inspector
    Dim myItem As Outlook.MailItem
    
    Set myinspector = Application.ActiveInspector
    If Not TypeName(myinspector) = "Nothing" Then
        Set myItem = myinspector.CurrentItem.Forward
        
        myItem.Display
        myItem.Recipients.Add ("CustAutoQuote@synnex.com")
        Debug.Print "Forwarded: " & myItem.Subject
        myItem.Send
'Debug.Print "Send: " & myItem.Subject                      把Debug.Print放在这一行会报错，应该是send后MailItem对象就没了
    
    Else
        MsgBox ("There is no active inspector")
    End If
    
    Set myinspector = Nothing
    Set myItem = Nothing
    
End Sub

Sub DisplayFolder()

    Dim myNameSpace As NameSpace
    Dim myFolder As Folder
    Dim myEntryID As String
    Dim myStoreID As String
    Dim myNewFolder As Outlook.Folder
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    
'    Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox)
'    myFolder.Display
    
'    Set myFolder = Application.Session.GetDefaultFolder(olFolderTasks)
    
    Set myFolder = myNameSpace.PickFolder
    myEntryID = myFolder.EntryID
    myStoreID = myFolder.StoreID
    Debug.Print ("EntryID:" & myFolder.EntryID & " StoreID:" & myFolder.StoreID)
    Set myNewFolder = Application.Session.GetFolderFromID(myEntryID, myStoreID)
    myNewFolder.Display
    
End Sub

Sub AutoForwardFolder()

    Dim myNameSpace As NameSpace
    Dim myFolder As Folder
    Dim myItem As Outlook.MailItem
    Dim myNewItem As Outlook.MailItem
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.PickFolder
    
    Debug.Print ("Items: " & myFolder.Items.Count)
    
    If myFolder.Items.Count < 50 Then
        For Each myItem In myFolder.Items
            Set myNewItem = myItem.Forward
            myNewItem.Display
            myNewItem.Recipients.Add ("CustAutoQuote@synnex.com")
            Debug.Print ("Forwarded: " & myNewItem.Subject)
            myNewItem.Send
        Next
    Else
        MsgBox ("Error: More than 50 items")
    End If
    
End Sub

Sub ReplyWithAttachmentsFolder()

    Dim myNameSpace As NameSpace
    Dim myFolder As Folder
    Dim myItem As Outlook.MailItem
    Dim myNewItem As Outlook.MailItem
    
    Dim myAttachment As Attachment
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myFolder = myNameSpace.PickFolder
    
    Debug.Print ("Items: " & myFolder.Items.Count)
    
    If myFolder.Items.Count < 20 Then
        For Each myItem In myFolder.Items
            If myItem.Attachments.Count < 100 Then
                Set myNewItem = myItem.ReplyAll
                For Each myAttachment In myItem.Attachments
                    If Left(myAttachment.FileName, 6) <> "image0" Then
'If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file.
                        myAttachment.SaveAsFile ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                        myNewItem.Attachments.Add ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                    End If
                Next
                Debug.Print ("DisplayReplyAll: " & myNewItem.Subject)
                myNewItem.Display
            Else
                Debug.Print ("Error Over 100 Attachments: " & myItem.Subject)
                MsgBox ("Error Over 100 Attachments: " & myItem.Subject)
            End If
        Next
    Else
        MsgBox ("Error: Up to 20 Items Once")
    End If
    
End Sub

Sub ReplyWithAttachments_OneMail()

    Dim myinspector As Outlook.Inspector
    Dim myNewItem As Outlook.MailItem
    Dim myAttachment As Attachment
    
    Set myinspector = Application.ActiveInspector
    If Not TypeName(myinspector) = "Nothing" Then
        If myinspector.CurrentItem.Attachments.Count < 100 Then
            Set myNewItem = myinspector.CurrentItem.ReplyAll
            For Each myAttachment In myinspector.CurrentItem.Attachments
                If Left(myAttachment.FileName, 6) <> "image0" Then
'If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file.
                    myAttachment.SaveAsFile ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                    myNewItem.Attachments.Add ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                End If
            Next
            Debug.Print ("DisplayReplyAll: " & myNewItem.Subject)
            myNewItem.Display
        Else
            Debug.Print ("Error Over 100 Attachments: " & myItem.Subject)
            MsgBox ("Error Over 100 Attachments: " & myItem.Subject)
        End If
    Else
        MsgBox ("There is no active inspector")
    End If
    
    Set myinspector = Nothing
    Set myNewItem = Nothing
    
End Sub

Sub ReplyWithAttachments_Explorer()
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myMail As Outlook.MailItem
    Dim myAttachment As Attachment
    Dim myNewMail As Outlook.MailItem
    Dim x As Long
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    If myOlSel.Count = 0 Then
        MsgBox ("Error: Selected Nothing")
    ElseIf myOlSel.Count <= 30 Then
        Debug.Print ("Items: " & myOlSel.Count)
        For x = 1 To myOlSel.Count
        If myOlSel.Item(x).Class = OlObjectClass.olMail Then
            Set myMail = myOlSel.Item(x)
            If myMail.Attachments.Count < 50 Then
                Set myNewMail = myMail.ReplyAll
                For Each myAttachment In myMail.Attachments
                If Left(myAttachment.FileName, 6) <> "image0" Then
'If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file.
                    myAttachment.SaveAsFile ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                    myNewMail.Attachments.Add ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                End If
                Next
                Debug.Print ("DisplayReplyAll: " & myNewMail.Subject)
                myNewMail.Display
            Else
                Debug.Print ("Error Over 50 Attachments: " & myMail.Subject)
                MsgBox ("Error Over 50 Attachments: " & myMail.Subject)
            End If
        Else
            Debug.Print ("Error: Not a Mail, Index (" & x & ")")
        End If
        Next x
    Else
        MsgBox ("Error: Please select not more than 30 Items; Selected: " & myOlSel.Count)
    End If
End Sub

Sub AutoForward_Explorer()
    Dim myOlExp As Explorer, myOlSel As Selection, myNewMail As MailItem
    Dim x As Long
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    If myOlSel.Count = 0 Then
        MsgBox ("Error: Selected Nothing")
    ElseIf myOlSel.Count <= 50 Then
        If MsgBox("Confirm Forward it to CustAutoQuote(机器人NiFi) ?", vbYesNo + vbInformation) = vbYes Then
        Debug.Print ("Items: " & myOlSel.Count)
        For x = 1 To myOlSel.Count
        If myOlSel.Item(x).Class = OlObjectClass.olMail Then
            Set myNewMail = myOlSel.Item(x).Forward
            myNewMail.Recipients.Add ("CustAutoQuote@synnex.com")
            '注释掉 myNewMail.Subject = myNewMail.Subject + " - Sent to NiFi by Outlook Macros"
            myNewMail.Display
            Debug.Print ("Forwarded: " & myNewMail.Subject)
            myNewMail.Send
        Else
            Debug.Print ("Error: Not a Mail, Index (" & x & ")")
        End If
        Next x
        End If
    Else
        MsgBox ("Error: Please select not more than 50 Items; Selected: " & myOlSel.Count)
    End If
End Sub


Sub ReplyWithAttachments_Explorer_Add_Recipient()
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myMail As Outlook.MailItem
    Dim myAttachment As Attachment
    Dim myNewMail As Outlook.MailItem
    Dim x As Long
    
    Set myOlExp = Application.ActiveExplorer
    Set myOlSel = myOlExp.Selection
    If myOlSel.Count = 0 Then
        MsgBox ("Error: Selected Nothing")
    ElseIf myOlSel.Count <= 30 Then
        Debug.Print ("Items: " & myOlSel.Count)
        For x = 1 To myOlSel.Count
        If myOlSel.Item(x).Class = OlObjectClass.olMail Then
            Set myMail = myOlSel.Item(x)
            If myMail.Attachments.Count < 50 Then
                Set myNewMail = myMail.Reply
                myNewMail.Recipients.Remove (1)
                myNewMail.Recipients.Add "patrickbue@synnex.com"
                myNewMail.Recipients.Add "SNXRENEEPO@synnex.com"
                myNewMail.Recipients.ResolveAll
                
                For Each myAttachment In myMail.Attachments
                If Left(myAttachment.FileName, 6) <> "image0" Then
'If a file with the same name already exists in the destination folder, it will be overwritten with this copy of the file.
                    myAttachment.SaveAsFile ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                    myNewMail.Attachments.Add ("C:\Users\ronf\Documents\AttachmentsCache\" & myAttachment.FileName)
                End If
                Next
                Debug.Print ("DisplayReplyAll: " & myNewMail.Subject)
                myNewMail.Display
            Else
                Debug.Print ("Error Over 50 Attachments: " & myMail.Subject)
                MsgBox ("Error Over 50 Attachments: " & myMail.Subject)
            End If
        Else
            Debug.Print ("Error: Not a Mail, Index (" & x & ")")
        End If
        Next x
    Else
        MsgBox ("Error: Please select not more than 30 Items; Selected: " & myOlSel.Count)
    End If
End Sub
