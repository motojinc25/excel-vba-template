Attribute VB_Name = "KBWorksheet"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : IsExistWorksheet
' Description       : Check If a worksheet exists
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strWorkbookName    - Workbook name
'                     strWorksheetName   - Worksheet name
' Called By         : Nothing
' Value Returned    : Boolean - True is existing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsExistWorksheet( _
    ByVal strWorkbookName As String, _
    ByVal strWorksheetName As String _
) As Boolean

Dim objWS As Worksheet
Dim blnExistResult As Boolean

blnExistResult = False

For Each objWS In Workbooks(strWorkbookName).Worksheets
  
    If (objWS.Name = strWorksheetName) Then

        blnExistResult = True
        Exit For
        
    End If
  
Next objWS

Set objWS = Nothing

IsExistWorksheet = blnExistResult

End Function
