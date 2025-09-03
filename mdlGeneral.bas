Attribute VB_Name = "mdlGeneral"
Private module
Option Explicit
Public appName As String
Public firstTimeOpen As Boolean
Public firstTimeOpenMsg As String
Public userRights As String
Public chatGPTModel As String

Public Sub eventHandler(turnOn As Boolean)
'Turn on or off applications events, screen updating and such'
    Call SetGlobals
    
    With Application
        If turnOn = True Then
            .ScreenUpdating = True
            .EnableEvents = True
            .DisplayAlerts = False
            .StatusBar = appName & " is Ready"
            .Cursor = xlDefault
        ElseIf turnOn = False Then
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = True
            .StatusBar = "Busy"
            .Cursor = xlWait
        End If
    End With
    
End Sub

 Public Sub logger(commingFrom As String, action As String)
'Pring message'
    Debug.Print (commingFrom & " " & action & " on " & Now())
'Save message'
    Call saveToLogFile("Log: " & commingFrom & " " & action & " on " & Now() & ". " & vbLf)
End Sub

Private Sub saveToLogFile(log As String)
'Declarations'
    Dim FSO As Object
    Dim logFilePath As String, logFileName As String
    Dim logFile As Object
''

'Create log file name'
    logFileName = Replace(Date, ".", "_")
    logFileName = "log_" & logFileName & ".txt"
    
'Get filepath'
    logFilePath = ThisWorkbook.Path & "\Logs\" & logFileName
    Debug.Print ("File path: " & logFilePath)

'Create new folder. If present already - skip'
    If folderExist(ThisWorkbook.Path & "\Logs\") = False Then
        MkDir (ThisWorkbook.Path & "\Logs\")
    End If
    
'Create new file. If present already - skip'
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If fileExist(logFilePath) = False Then
        Set logFile = FSO.createtextfile(logFilePath, False)
    Else
        Set logFile = FSO.opentextfile(logFilePath, 8)
    End If
    
'Save log provided to the file'
    logFile.write (log)
    
'Save file'
    logFile.Close
    
End Sub

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

Function ValToboolean(val As String) As Boolean

    val = Trim(UCase(val))
    
    Select Case val
        Case "TAK", "YES", "Y", "T", "TRUE"
            ValToboolean = True
        Case "NIE", "NO", "N", "FALSE"
            ValToboolean = False
        Case Else
            errorHandler (69)
    End Select

End Function
'
'Sub DisablePasswords()
'''Declarations''
'    Dim pass As String
'
'''
'
'        If verifyUser("admin") = True Then
'            pass = UCase("zamekhauru")
'        Else
'            pass = UCase(Trim(InputBox("Provide password!")))
'        End If
'
''Disable passwords'
'        If pass = UCase(Trim("zamekhauru")) Then
'            'Remove passwords from tabs'
'            wshRecord.Unprotect pass
'            wshUsers.Unprotect pass
'            wshCurUser.Unprotect pass
'            wshWorkType.Unprotect pass
'        Else
'            MsgBox ("Wrong password!")
'        End If
'
'End Sub
'
'Sub EnablePasswords()
'''Declarations''
'    Dim pass As String
'''
'
'        pass = UCase("zamekhauru")
''
''        wshLandingPage.Protect pass
''        wshRecord.Protect pass
''        wshRecord.Unprotect pass
''        wshUsers.Unprotect pass
''        wshCurUser.Unprotect pass
''        wshWorkType.Unprotect pass
'
''    End If
'
'End Sub

Public Function formatWhole()


End Function

Private Sub SetGlobals()
'User/System Related'
    appName = wshOptions.Cells(2, 2).Value
    wshOptions.Cells(3, 2).Value = ThisWorkbook.Path
    wshOptions.Cells(4, 2).Value = Environ("UserName")
    If wshOptions.Cells(5, 2).Value = "" Then
        wshOptions.Cells(5, 2).Value = "User"
    End If
    userRights = Trim(UCase(wshOptions.Cells(5, 2).Value))
    
'First time related'
    firstTimeOpen = wshOptions.Cells(6, 2).Value
    firstTimeOpenMsg = wshOptions.Cells(8, 2).Value

'ChatGPT model'
    chatGPTModel = wshOptions.Cells(9, 2).Value
    
'Hide options tab'
    If userRights <> "DEV" _
        And userRights <> "ADMIN" Then
        wshOptions.Visible = xlSheetVeryHidden
    Else
        wshOptions.Visible = xlSheetVisible
    End If
    
End Sub

Public Sub launch()

    Call eventHandler(True)
    
    If firstTimeOpen = True Then
        Call MsgBox(firstTimeOpenMsg, vbInformation, "Application Information")
        firstTimeOpen = False
        wshOptions.Cells(6, 2).Value = False
        wshOptions.Cells(7, 2).Value = Now
    End If
    
End Sub

Public Sub displayReminder()

    Call SetGlobals
    
    Call MsgBox(firstTimeOpenMsg, vbInformation, "Application Information")
    
End Sub
