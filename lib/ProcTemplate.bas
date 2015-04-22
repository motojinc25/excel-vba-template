Attribute VB_Name = "ProcTemplate"
Option Explicit
Option Base 1

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Procedure         : MainTemplate
' Description       : Procedure Template
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-20
' Parameters        : Nothing
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub MainTemplate()

' Declaring Constants
Const PROCEDURE_NAME            As String = "MainTemplate"
Const PROGRESS_BAR_TITLE        As String = "KBExcelVBATemplate"
Const PROGRESS_BAR_GROUP_FIRST  As String = "Group1"
Const PROGRESS_BAR_GROUP_SECOND As String = "Group2"
Const PROGRESS_BAR_GROUP_THIRD  As String = "Group3"
Const INPUT_DIALOG_FILE_TYPE    As String = "Excel files (*.xls*),*.xls*,All files (*.*),*.*"
Const INPUT_DIALOG_TITLE        As String = "Open excel files"

' Declaring Objects
Dim objDbg As New KBClassErrorHandler      ' DebugLog
Dim objBar As New KBClassProgressBar       ' ProgressBar
Dim objEnv As New KBClassRecordCollection  ' EnvironmentSheet

' Declaring Variables
Dim varInputFiles    As Variant
Dim intInputFilesCnt As Integer
Dim intIdxLoopX      As Integer

On Error GoTo ErrorHandler

' ==================================================
' Pre-processing Section
' ==================================================

' Screen updating is turned off
100 Call LockScreen

' Initialization DebugLog object
101 Call objDbg.init

' Output debug log
102 Call objDbg.writeInformationLog(PROCEDURE_NAME, "Start VBAMacro")

' Initialization ProgressBar object
103 Call objBar.initUI(PROGRESS_BAR_TITLE, RGB(255, 0, 0), RGB(0, 0, 255), RGB(0, 255, 0), PROGRESS_BAR_GROUP_FIRST, PROGRESS_BAR_GROUP_SECOND, PROGRESS_BAR_GROUP_THIRD)

' Select input files
104 intInputFilesCnt = OpenFiles(INPUT_DIALOG_FILE_TYPE, INPUT_DIALOG_TITLE, True, varInputFiles)

' Output debug log
105 Call objDbg.writeInformationLog(PROCEDURE_NAME, "Input Files count=" & intInputFilesCnt)

' Display Progress bar window
106 Call objBar.showUI(False)

' Repaint Progress bar window
107 Call objBar.repaintUI

' ==================================================
' Main processing Section
' ==================================================
' Loop for input files
For intIdxLoopX = 1 To intInputFilesCnt
    
    ' Output debug log
    Call objDbg.writeDebugLog(PROCEDURE_NAME, "[" & intIdxLoopX & "] Start LoopX - " & varInputFiles(intIdxLoopX))
    
    ' Updating Progress bar window
    Call objBar.updateUI(1, intIdxLoopX, intInputFilesCnt)

    ' OS can process other events
    DoEvents
    
    ' Output debug log
    Call objDbg.writeDebugLog(PROCEDURE_NAME, "[" & intIdxLoopX & "] End LoopX")
    
Next intIdxLoopX  ' Loop for input files

' ==================================================
' Post-processing Section
' ==================================================
POST_PROCESSING:

' Output debug log
Call objDbg.writeInformationLog(PROCEDURE_NAME, "End VBAMacro")

' Release Environment object
Call objEnv.releaseCollection

' Hide Progress bar and release object
Call objBar.hideUI
Call objBar.releaseClass

' Release DebugLog object
Call objDbg.releaseLogFile

' Screen updating is turned on
Call UnLockScreen

Exit Sub

' ==================================================
' ErrorHandle Section
' ==================================================
ErrorHandler:

' Display message box
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")

' Output debug log
Call objDbg.writeErrorLog(PROCEDURE_NAME, "LINE=" & Erl() & ",MESSAGE=" & Err.Description & "(" & Err.Number & ")")

' Go to Post-processing section
GoTo POST_PROCESSING

End Sub
