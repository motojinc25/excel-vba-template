Attribute VB_Name = "KBExport"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : ExportVBAComponents
' Description       : Export all components to files
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
Public Sub ExportVBAComponents()

Dim btnExport As Boolean
Dim objWorkbookSource As Excel.Workbook
Dim strExportPath As String
Dim strFileName As String
Dim objComponent As VBIDE.VBComponent

Set objWorkbookSource = Application.Workbooks(ThisWorkbook.Name)
strExportPath = ThisWorkbook.Path & "\..\lib\"
    
For Each objComponent In objWorkbookSource.VBProject.VBComponents
        
    btnExport = True
    strFileName = objComponent.Name

    Select Case objComponent.Type
    Case vbext_ct_ClassModule
        strFileName = strFileName & ".cls"
    Case vbext_ct_MSForm
        strFileName = strFileName & ".frm"
    Case vbext_ct_StdModule
        strFileName = strFileName & ".bas"
    Case vbext_ct_Document
        btnExport = False
    End Select
       
    If btnExport Then
        
        objComponent.Export strExportPath & strFileName
        
    End If
   
Next objComponent

End Sub


