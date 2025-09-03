Attribute VB_Name = "mdlHTTP"
Private module
Option Explicit
Public httpResponse As String

Public Sub httpRequest(httpRequestType As String, targetURL As String, requestDict As Scripting.Dictionary)
''Declarations''
    Dim timeStart As Date
    
    Dim objHTTP As Object
''

'log'
    Call logger("httpRequest", "Start")

'Disable events'
    Call eventHandler(False)
    
'Send request'
    timeStart = Now
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.SetTimeouts 90000, 90000, 90000, 90000
    objHTTP.Open httpRequestType, targetURL, False
    objHTTP.setRequestHeader "Content-Type", "application/json"
    objHTTP.setRequestHeader "Authorization", "Bearer " & requestDict("authorization")
    
'On timeout skip request'
    On Error GoTo skipRequest
    objHTTP.send requestDict("message")
    On Error GoTo 0
    
'Retrive response'
    httpResponse = objHTTP.responsetext
    Debug.Print ("Rensponse recived on: " & Now & ". Took " & Round(Now - timeStart, 6) & " seconds.")
    httpResponse = Replace(httpResponse, vbLf, "")
    httpResponse = Replace(httpResponse, "{", "")
    httpResponse = Replace(httpResponse, "}", "")
    httpResponse = Replace(httpResponse, Chr(34), "")
    httpResponse = Trim(httpResponse)
    httpResponse = unicodeConv(httpResponse)
    
    If Left(httpResponse, 5) = "error" Then
        Debug.Print ("Issue with request:" & vbLf & httpResponse)
    ElseIf Left(httpResponse, 5) <> "error" _
            And InStr(requestDict("model"), "gpt") > 0 Then
        httpResponse = Split(httpResponse, "text:")(1)
        httpResponse = Split(httpResponse, "]")(0)
        httpResponse = Replace(httpResponse, "\n", vbLf)
'        httpResponse = Replace(httpResponse, "**", "")
        httpResponse = Replace(httpResponse, "###", "")
        httpResponse = Replace(httpResponse, "##", "")
        Debug.Print ("Succesful request!")
    Else
        Debug.Print ("Succesful request!")
    End If
    
skipRequest:
'log'
    Call logger("httpRequest", "Finish")

'Disable events'
    Call eventHandler(True)

End Sub

Private Function unicodeConv(targetText As String) As String
'Based on :https://pl.wikipedia.org/wiki/Kodowanie_polskich_znak%C3%B3w'
'Replace spceial codes with special characters'
    targetText = Replace(targetText, "\u0104", "•")
    targetText = Replace(targetText, "\u0106", "∆")
    targetText = Replace(targetText, "\u0118", " ")
    targetText = Replace(targetText, "\u0141", "£")
    targetText = Replace(targetText, "\u0143", "—")
    targetText = Replace(targetText, "\u0D3", "”")
    targetText = Replace(targetText, "\u0d3", "”")
    targetText = Replace(targetText, "\u015A", "å")
    targetText = Replace(targetText, "\u015a", "å")
    targetText = Replace(targetText, "\u0179", "è")
    targetText = Replace(targetText, "\u017B", "Ø")
    targetText = Replace(targetText, "\u017b", "Ø")
    targetText = Replace(targetText, "\u0105", "π")
    targetText = Replace(targetText, "\u0107", "Ê")
    targetText = Replace(targetText, "\u0119", "Í")
    targetText = Replace(targetText, "\u0142", "≥")
    targetText = Replace(targetText, "\u0144", "Ò")
    targetText = Replace(targetText, "\u0F3", "Û")
    targetText = Replace(targetText, "\u0f3", "Û")
    targetText = Replace(targetText, "\u00F3", "Û")
    targetText = Replace(targetText, "\u00f3", "Û")
    targetText = Replace(targetText, "\u015B", "ú")
    targetText = Replace(targetText, "\u015b", "ú")
    targetText = Replace(targetText, "\u017A", "ü")
    targetText = Replace(targetText, "\u017a", "ü")
    targetText = Replace(targetText, "\u017C", "ø")
    targetText = Replace(targetText, "\u017c", "ø")
    targetText = Replace(targetText, "\u2013", "-")
    targetText = Replace(targetText, "\u2014", "-")
    
    unicodeConv = targetText

End Function
