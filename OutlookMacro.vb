Sub AddGenerateResponseButton()
    Dim myInspector As Inspector
    Set myInspector = Application.ActiveInspector

    If Not myInspector Is Nothing Then
        Dim myItem As Object
        Set myItem = myInspector.currentItem
        Call GenerateResponse(myItem.Body)
    End If
End Sub

Sub GenerateResponse(emailBody As String)
    ' Path to your Python executable
    Dim pythonExe As String
    pythonExe = "python"
    
    ' Path to your Python script
    Dim scriptPath As String
    scriptPath = "generate_response.py"
    
    ' Shell command to execute
    Dim shellCommand As String
    shellCommand = """" & pythonExe & """ """ & scriptPath & """ """ & emailBody & """"
    
    ' Create a WshShell object and execute the command with no window
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' Show progress message
    Dim progressMessage As String
    progressMessage = "Generating response, please wait..."
    MsgBox progressMessage, vbInformation + vbSystemModal, "Progress"
    
    Dim exec As Object
    On Error Resume Next
    Set exec = wsh.exec(shellCommand)
    On Error GoTo 0
    
    ' Check if the command was successfully executed
    If exec Is Nothing Then
        MsgBox "Failed to execute the Python script.", vbCritical
        Exit Sub
    End If
    
    ' Wait for the command to complete
    Do While exec.Status = 0
        DoEvents
    Loop

    ' Close progress message
    On Error Resume Next
    Application.DisplayAlerts = False
    SendKeys "{ESC}", True
    On Error GoTo 0

    ' Get the response
    Dim response As String
    response = exec.StdOut.ReadAll()

    ' Check if there is an active inspector
    If Application.ActiveInspector Is Nothing Then
        MsgBox "No active email window found.", vbExclamation
        Exit Sub
    End If

    ' Check if the current item is a mail item
    Dim currentItem As Object
    Set currentItem = Application.ActiveInspector.currentItem
    If Not TypeOf currentItem Is MailItem Then
        MsgBox "The current item is not a mail item.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    ' Create a reply and set the body
    Dim replyItem As MailItem
    Set replyItem = currentItem.Reply
    replyItem.Body = response & vbCrLf & replyItem.Body  ' Append response to original message
    replyItem.Display  ' Display before sending, allowing for edits
    Exit Sub

ErrorHandler:
    MsgBox "Failed to create a reply. Error: " & Err.Description, vbCritical
End Sub

