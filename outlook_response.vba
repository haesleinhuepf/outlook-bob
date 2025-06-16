Sub GenerateReplyWithPython()
    Dim cb As CommandBar
    Dim btn As CommandBarControl
    
    ' Create popup menu
    On Error Resume Next
    Application.ActiveExplorer.CommandBars("ResponseTypeMenu").Delete
    On Error GoTo 0
    
    Set cb = Application.ActiveExplorer.CommandBars.Add("ResponseTypeMenu", msoBarPopup, False, True)
    
    ' Add menu items
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Caption = "Positive Response"
    btn.OnAction = "HandlePositiveResponse"
    
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Caption = "Negative Response"
    btn.OnAction = "HandleNegativeResponse"
    
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Caption = "Asking for more details"
    btn.OnAction = "HandleMoreDetailsResponse"
    
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Caption = "Saying Thanks"
    btn.OnAction = "HandleThanksResponse"
    
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Caption = "Other"
    btn.OnAction = "HandleOtherResponse"
    
    ' Show popup menu at cursor position
    cb.ShowPopup
End Sub

Sub HandlePositiveResponse()
    ProcessResponse "answering_positively"
End Sub

Sub HandleNegativeResponse()
    ProcessResponse "answering_negatively"
End Sub

Sub HandleMoreDetailsResponse()
    ProcessResponse "asking_for_more_details"
End Sub

Sub HandleThanksResponse()
    ProcessResponse "saying_thanks"
End Sub

Sub HandleOtherResponse()
    Dim customType As String
    customType = InputBox("Enter custom response:", "Custom Response")
    
    If customType = "" Then
        Exit Sub
    End If
    
    ' Convert spaces and special characters to underscores
    Dim i As Integer
    For i = 1 To Len(customType)
        Dim char As String
        char = Mid(customType, i, 1)
        If Not (char Like "[A-Za-z0-9_]") Then
            customType = Replace(customType, char, "_")
        End If
    Next i
    
    ProcessResponse customType
End Sub

Private Sub ProcessResponse(responseType As String)
    Dim currentItem As MailItem
    Dim emailBody As String
    Dim fso As Object
    Dim tempFile As String
    Dim replyText As String
    Dim replyFile As String
    Dim errorFile As String
    
    replyFile = "C:\structure\code\obob\temp\draft_reply.txt"
    tempFile = "C:\structure\code\obob\temp\email_body.txt"
    errorFile = "C:\structure\code\obob\temp\python_error.txt"
    
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email."
        Exit Sub
    End If
    
    Set currentItem = Application.ActiveExplorer.Selection.Item(1)
    emailBody = currentItem.Body
    
    ' Write email body to a temporary file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts
    Set ts = fso.CreateTextFile(tempFile, True)
    ts.Write emailBody
    ts.Close
    
    ' Call the Python script with conda environment activation and response type parameter
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd /c type " & Chr(34) & tempFile & Chr(34) & " | conda activate bob-env && python C:\structure\code\obob\email_assistant.py response_type=" & responseType & " 2> " & Chr(34) & errorFile & Chr(34), 1, True
    
   
    ' Read reply from file
    If fso.FileExists(replyFile) Then
        Dim replyTextStream
        Set replyTextStream = fso.OpenTextFile(replyFile, 1)
        replyText = replyTextStream.ReadAll
        replyTextStream.Close
    Else
        MsgBox "Reply file not found."
        Exit Sub
    End If

    ' Create reply draft with original email content
    Dim replyItem As MailItem
    Set replyItem = currentItem.Reply
    
    ' Format the original email header in HTML
    Dim originalHeader As String
    originalHeader = "<hr style='border: 1px solid #e0e0e0;'/>" & vbCrLf & _
                    "<div style='color: #666666; font-family: Arial, sans-serif; font-size: 12px;'>" & _
                    "<p><strong>From:</strong> " & currentItem.SenderName & " &lt;" & currentItem.SenderEmailAddress & "&gt;</p>" & _
                    "<p><strong>Sent:</strong> " & Format(currentItem.ReceivedTime, "dddd, mmmm d, yyyy h:mm AM/PM") & "</p>" & _
                    "<p><strong>To:</strong> " & currentItem.To & "</p>" & _
                    "<p><strong>Subject:</strong> " & currentItem.Subject & "</p></div>"
    
    ' Set the HTML body with proper formatting
    replyItem.HTMLBody = "<div style='font-family: Arial, sans-serif;'>" & _
                        replyText & _
                        "</div>" & vbCrLf & vbCrLf & _
                        originalHeader & vbCrLf & _
                        "<div style='font-family: Arial, sans-serif;'>" & _
                        Replace(emailBody, vbCrLf, "<br/>") & _
                        "</div>"
    
    replyItem.Display
End Sub
