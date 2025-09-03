Attribute VB_Name = "mdlChatGPTAPI"
Public Sub chatGPTAPI()
''Declarations''\
    Dim i As Long
    Dim rowStart As Long, rowLast As Long
    
    Dim prompt As String
    Dim openAIAPIKey As String, APIFilePath As String
    
    Dim requestDict As New Scripting.Dictionary
        
    Dim FSO As Object, APIFile As Object
''
'log'
    Call logger("chatGPTAPI", "Start")

'Disable events'
    Call eventHandler(False)
'API'
    APIFilePath = ThisWorkbook.Path & "\API"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If IsError(FSO.opentextfile(APIFilePath & "\ChatGPT_APIkey.txt", ForReading)) = True Then
        Call errorHandler(301, "Refer to 'App Information' button for more information.")
    End If
    On Error GoTo 0
    Set APIFile = FSO.opentextfile(APIFilePath & "\ChatGPT_APIkey.txt", ForReading)
    openAIAPIKey = APIFile.ReadAll
    APIFile.Close
    
'Get row start and last'
    rowLast = wshPrompt.Cells(Rows.Count, 2).End(xlUp).Row
    rowStart = wshPrompt.UsedRange.Find("Utworzone").Row
    
'Make dictionary'
    For i = rowStart + 1 To rowLast
        Application.StatusBar = "Processing requests. Progress: " & (100 * Round((i - 1) / (rowLast - 1), 3)) & "%"
        prompt = Trim(Replace(wshPrompt.Cells(i, 2).Value, vbLf, ""))
        prompt = Replace(prompt, Chr(34), "*")
        Set requestDict = makeChatGPTDict(chatGPTModel, 20, prompt, openAIAPIKey)
'send request'
        Call httpRequest("POST", "https://api.openai.com/v1/responses", requestDict)
        wshPrompt.Cells(i, 3).Value = httpResponse
    Next i
        
    Application.StatusBar = "Busy"
     
'Formating'
    wshPrompt.Columns(3).ColumnWidth = wshPrompt.Columns(1).ColumnWidth
    wshPrompt.Columns(2).ColumnWidth = wshPrompt.Columns(1).ColumnWidth
    
'log'
    Call logger("chatGPTAPI", "Finish")

'Disable events'
    Call eventHandler(True)
    
'Go to prompts'
    ActiveWindow.Zoom = 100
    MsgBox ("Done!")
    
End Sub
Private Function makeChatGPTDict(model As String, tokenLimit As Long, prompt As String, openAIAPIKey As String) As Scripting.Dictionary
''Declarations''
    Dim message As String
    
''
'Add to dictionary'
    Set makeChatGPTDict = New Scripting.Dictionary
    makeChatGPTDict.Add "model", model
    makeChatGPTDict.Add "tokenLimit", tokenLimit
    makeChatGPTDict.Add "authorization", openAIAPIKey
    
    message = "{" & _
        Chr(34) & "model" & Chr(34) & ":" & Chr(34) & model & Chr(34) & "," & _
        Chr(34) & "input" & Chr(34) & ":" & Chr(34) & prompt & Chr(34) & _
        "}"
        
    makeChatGPTDict.Add "message", message

End Function

