Attribute VB_Name = "KBInit"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Procedure         : LoadInitFileInformation
' Description       : Loading information of initialization file
' Author            : Jingun Jung (Webpage: www.koreabigname.com)
' Date              : 2015-04-18
' Parameters        : Nothing
' Called By         : Nothing
' Value Returned    : Nothing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadInitFileInformation()

Const PROCEDURE_NAME     As String = "LoadInitFileInformation"
Const SECTION_NAME       As String = "Tool_Infomation"
Const ENTRY_TOOL_NAME    As String = "ToolName"
Const ENTRY_TOOL_VERSION As String = "ToolVersion"

If (KBFile.IsExistFile(ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE) = False) Then
    
    Call MsgBox("Initialization file is not exist" & vbCrLf & vbCrLf & _
                ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE, _
                vbCritical, _
                "ExcelVBA ERROR")
    Exit Sub
    
End If

On Error GoTo ErrorHandler

10: gstrToolName = GetSectionEntryString(SECTION_NAME, ENTRY_TOOL_NAME, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)
20: gstrToolVersion = GetSectionEntryString(SECTION_NAME, ENTRY_TOOL_VERSION, ThisWorkbook.Path & "\" & C_TOOL_INIT_FILE)

Exit Sub

ErrorHandler:
Call MsgBox("Procedule : " & PROCEDURE_NAME & "(" & Erl() & ")" & vbCrLf & _
            "Message : " & Err.Description & "(" & Err.Number & ")", _
            vbCritical, _
            "ExcelVBA ErrorHandler")
End Sub
