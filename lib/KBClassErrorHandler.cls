VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KBClassErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class             : KBClassErrorHandler
' Description       : An error handler class for Debug
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-19
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private blnInformationFlag      As Boolean
Private blnWarningFlag          As Boolean
Private blnErrorFlag            As Boolean
Private blnDebugFlag            As Boolean
Private strSaveFilename         As String
Private intFreeFileNumber       As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Initialize
' Description       : Initialize error handler class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

intSaveFreeFile = FreeFile
Open ThisWorkbook.Path & "\" & strSaveFilename For Append As #intSaveFreeFile

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Terminate
' Description       : Terminate error handler class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : init
' Description       : Initialize error handler objects
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub init()

Const PROCEDURE_NAME        As String = "KBClassErrorHandler.init"
Const SECTION_NAME          As String = "Tool_Log"
Const ENTRY_LOG_FILENAME    As String = "LogFilename"
Const ENTRY_INFORMATION_LOG As String = "InformationLog"
Const ENTRY_ERROR_LOG       As String = "ErrorLog"
Const ENTRY_WARNING_LOG     As String = "WarningLog"
Const ENTRY_DEBUG_LOG       As String = "DebugLog"
Dim strEntryValue           As String

If (KBFile.IsExistFile(ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE) = False) Then
    
    Call MsgBox("Initialization file is not exist" & vbCrLf & vbCrLf & _
                ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE, _
                vbCritical, _
                "ExcelVBA ERROR")
    Exit Sub
    
End If

On Error GoTo ErrorHandler

10: strSaveFilename = GetSectionEntryString(SECTION_NAME, ENTRY_LOG_FILENAME, ThisWorkbook.Path & "\1" & C_TOOL_INIT_FILE)

20: strEntryValue = GetSectionEntryString(SECTION_NAME, ENTRY_INFORMATION_LOG, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)

If UCase(strEntryValue) = "YES" Then
    blnInformationFlag = True
Else
    blnInformationFlag = False
End If

30: strEntryValue = GetSectionEntryString(SECTION_NAME, ENTRY_ERROR_LOG, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)

If UCase(strEntryValue) = "YES" Then
    blnErrorFlag = True
Else
    blnErrorFlag = False
End If

40: strEntryValue = GetSectionEntryString(SECTION_NAME, ENTRY_WARNING_LOG, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)

If UCase(strEntryValue) = "YES" Then
    blnWarningFlag = True
Else
    blnWarningFlag = False
End If

50: strEntryValue = GetSectionEntryString(SECTION_NAME, ENTRY_DEBUG_LOG, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)

If UCase(strEntryValue) = "YES" Then
    blnDebugFlag = True
Else
    blnDebugFlag = False
End If

60: intSaveFreeFile = FreeFile
Open ThisWorkbook.Path & "\" & strSaveFilename For Append As #intSaveFreeFile

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : releaseLogFile
' Description       : Release error handler objects
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub releaseLogFile()

Close #intSaveFreeFile

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : writeErrorLog
' Description       : Write error log message to log file
' Parameters        : strHeadMessage - Head message of log
'                     strBodyMessage - Body message of log
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub writeErrorLog( _
    ByVal strHeadMessage As String, _
    ByVal strBodyMessage As String _
)

Const PROCEDURE_NAME        As String = "KBClassErrorHandler.writeErrorLog"
Dim strOutputMessage As String

On Error GoTo ErrorHandler

If (blnErrorFlag = True) Then

    strOutputMessage = Date & " " & Time & " [E] " & strHeadMessage & " " & strBodyMessage
    
10: Print #intSaveFreeFile, strOutputMessage
    
End If

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")
            
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : writeInformationLog
' Description       : Write information log message to log file
' Parameters        : strHeadMessage - Head message of log
'                     strBodyMessage - Body message of log
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub writeInformationLog( _
    ByVal strHeadMessage As String, _
    ByVal strBodyMessage As String _
)

Const PROCEDURE_NAME        As String = "KBClassErrorHandler.writeInformationLog"
Dim strOutputMessage As String

On Error GoTo ErrorHandler

If (blnErrorFlag = True) Then

    strOutputMessage = Date & " " & Time & " [I] " & strHeadMessage & " " & strBodyMessage
    
10: Print #intSaveFreeFile, strOutputMessage
    
End If

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")
            
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : writeWarningLog
' Description       : Write warning log message to log file
' Parameters        : strHeadMessage - Head message of log
'                     strBodyMessage - Body message of log
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub writeWarningLog( _
    ByVal strHeadMessage As String, _
    ByVal strBodyMessage As String _
)

Const PROCEDURE_NAME        As String = "KBClassErrorHandler.writeWarningLog"
Dim strOutputMessage As String

On Error GoTo ErrorHandler

If (blnErrorFlag = True) Then

    strOutputMessage = Date & " " & Time & " [W] " & strHeadMessage & " " & strBodyMessage
    
10: Print #intSaveFreeFile, strOutputMessage
    
End If

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")
            
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : writeDebugLog
' Description       : Write debug log message to log file
' Parameters        : strHeadMessage - Head message of log
'                     strBodyMessage - Body message of log
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub writeDebugLog( _
    ByVal strHeadMessage As String, _
    ByVal strBodyMessage As String _
)

Const PROCEDURE_NAME        As String = "KBClassErrorHandler.writeDebugLog"
Dim strOutputMessage As String

On Error GoTo ErrorHandler

If (blnErrorFlag = True) Then

    strOutputMessage = Date & " " & Time & " [D] " & strHeadMessage & " " & strBodyMessage
    
10: Print #intSaveFreeFile, strOutputMessage
    
End If

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")
            
End Sub
