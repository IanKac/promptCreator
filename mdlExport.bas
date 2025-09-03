Attribute VB_Name = "mdlExport"
Option Explicit

'This modules goal is to extract data from worksheets and export them as json file'
'Required modules:
'- mdlGeneral
'- mdlErrorHandling
'Required reference
'- Microsoft Scripting Runtime

Private Function fileExist(filePath As String) As Boolean

    If Len(Dir(filePath)) = 0 Then
        fileExist = False
    Else
        fileExist = True
    End If
 
End Function

Private Function folderExist(filePath As String) As Boolean
    
    If Len(Dir(filePath, vbDirectory)) = 0 Then
        folderExist = False
    Else
        folderExist = True
    End If

End Function

Public Sub exportAsTXT(targetSht As Worksheet, Optional StartRow As Long, Optional startcol As Long)
'This sub export all of data from target worksheet. Data should start at cells(1,1) or be provided'
''Declarations''
    Dim i As Long, j As Long
    
    Dim headerRow As Long
    Dim lastcol As Long, lastRow As Long, lastRowPrompt As Long
    
    Dim txtString As String
    
    Dim colNamesColl As Collection
    
    Dim txtFilePath As String, txtFileName As String
    
    Dim FSO As Object
    Dim txtfile As Object
''

'log'
    Call logger("exportAstxt", "Start")

'Disable events'
    Call eventHandler(False)

'Set variables'
    If startcol <> 0 Then
        i = startcol
    Else
        startcol = targetSht.UsedRange.Column
        i = 1
    End If
    If StartRow <> 0 Then
        headerRow = StartRow
    Else
        headerRow = targetSht.UsedRange.Row
    End If
    
    lastRow = targetSht.Cells(Rows.Count, startcol).End(xlUp).Row
    lastcol = targetSht.Cells(headerRow, Columns.Count).End(xlToLeft).Column
    
'Gather column names'
    Set colNamesColl = New Collection
    Do While targetSht.Cells(headerRow, i).Value <> ""
        colNamesColl.Add targetSht.Cells(headerRow, i).Value
        Debug.Print ("Column name added to collection: " & targetSht.Cells(headerRow, i).Value)
        i = i + 1
    Loop

'Create txt file'
'Create txt file name'
    txtFileName = Replace(Year(Date) & Month(Date) & Day(Date), ".", "")
    txtFileName = txtFileName & "_prompts.txt"
    
'Get filepath'
    txtFilePath = ThisWorkbook.Path & "\txt\" & txtFileName
    Debug.Print ("File path: " & txtFilePath)

'Create new folder. If present already - skip'
    If folderExist(ThisWorkbook.Path & "\txt\") = False Then
        MkDir (ThisWorkbook.Path & "\txt\")
    End If
    
'Create new file. If present already - skip'
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If fileExist(txtFilePath) = False Then
        Set txtfile = FSO.createtextfile(txtFilePath, False)
    Else
        Set txtfile = FSO.opentextfile(txtFilePath, 8)
    End If
    
'Find ID column'
    j = wshData.UsedRange.Find("ID").Column
    
'Create prompts'
'Find last row in wshprompt'
    lastRowPrompt = wshPrompt.Cells(Rows.Count, 1).End(xlUp).Row
    If lastRowPrompt < wshPrompt.Cells(Rows.Count, 2).End(xlUp).Row Then
        lastRowPrompt = wshPrompt.Cells(Rows.Count, 2).End(xlUp).Row
    ElseIf lastRowPrompt < wshPrompt.Cells(Rows.Count, 3).End(xlUp).Row Then
        lastRowPrompt = wshPrompt.Cells(Rows.Count, 3).End(xlUp).Row
    End If
    
'Clear prompts worksheet'
    For i = 2 To lastRowPrompt
        wshPrompt.Cells(i, 2).Value = ""
        wshPrompt.Cells(i, 3).Value = ""
    Next i
    
'Add headers'
        wshPrompt.Cells(1, 1).Value = "Aktualny"
        wshPrompt.Cells(1, 2).Value = "Utworzone"
        wshPrompt.Cells(1, 3).Value = "Odpowiedzi"
        wshPrompt.Cells(1, 1).Interior.Color = RGB(0, 204, 255)
        wshPrompt.Cells(1, 2).Interior.Color = RGB(0, 204, 255)
        wshPrompt.Cells(1, 3).Interior.Color = RGB(0, 204, 255)
        wshPrompt.Cells(1, 1).Font.Color = RGB(0, 0, 0)
        wshPrompt.Cells(1, 2).Font.Color = RGB(0, 0, 0)
        wshPrompt.Cells(1, 3).Font.Color = RGB(0, 0, 0)
        wshPrompt.Cells(1, 1).Font.Bold = True
        wshPrompt.Cells(1, 2).Font.Bold = True
        wshPrompt.Cells(1, 3).Font.Bold = True
        
'Collect data from worksheet'
    For i = headerRow + 1 To lastRow
'Create txt string'
        txtString = txtString & vbLf & promptCreate(i, lastcol) & ";"
'Add txt string to txt file'
        txtfile.write (txtString)
'Add txt string to prompts
        wshPrompt.Cells(i, 2).Value = txtString
    Next i
    
'Close txt file
    txtfile.Close

'Enable events'
    Call eventHandler(True)
'log'
    Call logger("exportAstxt", "Finish")
    
End Sub
Private Function promptCreate(targetRow As Long, lastcol As Long) As String
''Declarations''
    Dim i As Long
    
''
    promptCreate = wshPrompt.Cells(2, 1).Value
'Loop thorugh all columns and replace data in prompt'
    For i = 1 To lastcol
        promptCreate = Replace(promptCreate, "[" & Trim(wshData.Cells(1, i).Value) & "]", wshData.Cells(targetRow, i).Value)
    Next i

'Return string'
End Function

Public Sub exportAsWord(targetSht As Worksheet, Optional StartRow As Long, Optional startcol As Long)
'This sub export all of data from target worksheet. Data should start at cells(1,1) or be provided'
''Declarations''
    Dim i As Long, j As Long
    
    Dim headerRow As Long
    Dim lastcol As Long, nameCol As Long
    Dim lastRow As Long, lastRowPrompt As Long
    
    Dim WordString As String
    
    Dim colNamesColl As Collection
    
    Dim wordFilePath As String, wordFileName As String
    
    Dim appWord As Word.Application
    
    Dim wordCurrentSection As Word.Range
    Dim wordLastPar As Long, par As Long
''

'log'
    Call logger("exportAsWord", "Start")

'Disable events'
    Call eventHandler(False)

'Set variables'
    If startcol <> 0 Then
        i = startcol
    Else
        startcol = targetSht.UsedRange.Column
        i = 1
    End If
    If StartRow <> 0 Then
        headerRow = StartRow
    Else
        headerRow = targetSht.UsedRange.Row
    End If
    
    lastRow = targetSht.Cells(Rows.Count, startcol).End(xlUp).Row
    lastcol = targetSht.Cells(headerRow, Columns.Count).End(xlToLeft).Column
    nameCol = 2
    
'Create Word file'
'Create Word file name'
    wordFileName = Replace(Year(Date) & Month(Date) & Day(Date) & Time, ".", "")
    wordFileName = Replace(wordFileName, ":", "")
    wordFileName = wordFileName & "_Report.docx"
    
'Get filepath'
    wordFilePath = ThisWorkbook.Path & "\Report\"
    Debug.Print ("File path: " & wordFilePath & wordFileName)

'Create new folder. If present already - skip'
    If folderExist(wordFilePath) = False Then
        MkDir (wordFilePath)
    End If
    
'Create new file. If present already - skip'
    Set appWord = New Word.Application
    With appWord
        .Visible = True
        .Activate
        .Documents.Add
'Collect data from worksheet'
        For i = headerRow + 1 To lastRow
'Add Word string to Word file'
'Name'
            .Selection.Paragraphs.Add
            Set wordCurrentSection = .Selection.Range.Characters.Last
            With wordCurrentSection
                .Text = (wshData.Cells(i, nameCol).Value)
                .Style = Word.WdBuiltinStyle.wdStyleHeading2
                wordLastPar = appWord.Selection.Paragraphs.Count
            End With
'Response'
            .Selection.MoveEnd unit:=wdParagraph, Count:=.Selection.Paragraphs.Count
            .Selection.Paragraphs.Add
            Set wordCurrentSection = .Selection.Range.Characters.Last
            With wordCurrentSection
                .Text = Replace(targetSht.Cells(i, 3).Value, "---", "") & vbLf
            End With
            For par = wordLastPar + 1 To .Selection.Paragraphs.Count
                .Selection.Paragraphs(par).Style = Word.WdBuiltinStyle.wdStyleNormal
            Next par
            
'Show progress'
            .StatusBar = "Inserting chatGPT responses. Progress: " & 100 * Round((i - 1) / (lastRow - 1), 2) & "%"
        Next i
        
'Bold correct text'
        Call wordBolden(appWord, "**")
'Show navigation pane'
        .ActiveWindow.DocumentMap = True
'Save document under correct name'
        .ActiveDocument.SaveAs2 wordFilePath & wordFileName
'Clear status bar'
        .StatusBar = ""
    End With
    
'Enable events'
    Call eventHandler(True)
'log'
    Call logger("exportAsWord", "Finish")
    
End Sub

Private Sub wordBolden(targetDoc As Word.Application, targetMarker As String)
''Declarations''
    Dim i As Long
    
    Dim bolding As Boolean
    
''
'log'
    Call logger("wordBolden", "Start")
    
    For i = 1 To targetDoc.Selection.Words.Count
'If word match targetMarker then procceed to bold until another match is found then switch back'
        Debug.Print ("Checking word: " & i & " for markers out of: " & targetDoc.Selection.Words.Count & ". Progress: " & 100 * Round(i / targetDoc.Selection.Words.Count, 3))
        targetDoc.StatusBar = "Checking for markers. Progress: " & 100 * Round(i / targetDoc.Selection.Words.Count, 3) & "%"
        If bolding = False And _
                InStr(targetDoc.Selection.Words(i), targetMarker) > 0 Then
            bolding = True
        ElseIf bolding = True And _
                InStr(targetDoc.Selection.Words(i), targetMarker) > 0 Then
            bolding = False
        End If
        
        If bolding = True Then
            targetDoc.Selection.Words(i).Font.Bold = True
        End If
        
    Next i
    
'Remove marker'
    With targetDoc.Selection.Find
        .ClearFormatting
        .Text = targetMarker
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
'Remove double space'
    With targetDoc.Selection.Find
        .ClearFormatting
        .Text = "  "
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
    End With
    
'Clear status bar'
    targetDoc.StatusBar = ""
    
'log'
    Call logger("wordBolden", "Finish")
    
End Sub
