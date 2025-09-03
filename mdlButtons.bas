Attribute VB_Name = "mdlButtons"
Sub btnExportAsTXT_Click()

    Call exportAsTXT(wshData, 1, 1)
    
    MsgBox ("DONE!")
    
End Sub
Sub btnExportAsWord_Click()

    Call exportAsWord(wshPrompt, 1, 3)
    
    MsgBox ("Word report done!" & vbLf & "File saved at 'Report' folder.")
    
End Sub

Sub btnClearData_Click()
''Declarations''
    Dim i As Long
    Dim rowLast As Long
''

'Disable events'
    Call eventHandler(False)
    
'Set variables'
    rowLast = wshData.Cells(Rows.Count, 1).End(xlUp).Row
    
'Clear data form datasheet'
    
    For i = 2 To rowLast
        wshData.Rows(i).Clear
        If i > rowLast Then
            Exit For
        End If
    Next i
    
'Enable events'
    Call eventHandler(True)
End Sub

Sub btnClearPrompt_Click()
''Declarations''
    Dim i As Long
    Dim rowLast As Long
''

'Disable events'
    Call eventHandler(False)
    
'Set variables'
'Find last row in wshprompt'
    rowLast = wshPrompt.Cells(Rows.Count, 1).End(xlUp).Row
    If rowLast < wshPrompt.Cells(Rows.Count, 2).End(xlUp).Row Then
        rowLast = wshPrompt.Cells(Rows.Count, 2).End(xlUp).Row
    ElseIf rowLast < wshPrompt.Cells(Rows.Count, 3).End(xlUp).Row Then
        rowLast = wshPrompt.Cells(Rows.Count, 3).End(xlUp).Row
    End If
'Clear data form datasheet'
    
    For i = 2 To rowLast
        wshPrompt.Rows(i).Clear
        If i > rowLast Then
            Exit For
        End If
    Next i
    
'Enable events'
    Call eventHandler(True)
End Sub

Sub btnSendPrompt_click()

    Call chatGPTAPI
End Sub
