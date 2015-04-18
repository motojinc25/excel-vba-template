Attribute VB_Name = "KBCell"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetLastColumnNumber
' Description       : Find column number of the last cell
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strWorkbookName    - Workbook name
'                     strWorksheetName   - Worksheet name
'                     lngSearchRowNumber - Searching row number
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLastColumnNumber&( _
    ByVal strWorkbookName As String, _
    ByVal strWorksheetName As String, _
    ByVal lngSearchRowNumber As Long _
)

Dim objWS As Worksheet
Dim lngLastNumber As Long

Set objWS = Workbooks(strWorkbookName).Worksheets(strWorksheetName)

lngLastNumber = objWS.Cells(lngSearchRowNumber, Columns.Count).End(xlToLeft).Column

GetLastColumnNumber = lngLastNumber

Set objWS = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetLastRowNumber
' Description       : Find row number of the last cell
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strWorkbookName       - Workbook name
'                     strWorksheetName      - Worksheet name
'                     lngSearchColumnNumber - Searching column number
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLastRowNumber&( _
    ByVal strWorkbookName As String, _
    ByVal strWorksheetName As String, _
    ByVal lngSearchColumnNumber As Long _
)

Dim objWS As Worksheet
Dim lngLastNumber As Long

Set objWS = Workbooks(strWorkbookName).Worksheets(strWorksheetName)

lngLastNumber = objWS.Cells(Rows.Count, lngSearchColumnNumber).End(xlUp).Row

GetLastRowNumber = lngLastNumber

Set objWS = Nothing

End Function
