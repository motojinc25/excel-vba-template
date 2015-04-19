VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KBClassRecordCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class             : KBClassRecordCollection
' Description       : A Record class for Array Collection
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
Private lngCollectionRow            As Long
Private intCollectionColumn         As Integer
Private objCollection()             As Record

Private Type Record
    d() As String
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Initialize
' Description       : Initialize record class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Terminate
' Description       : Terminate record class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()

Erase objCollection

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initCollection
' Description       : Initialize record object
' Parameters        : strTargetWBName      - Workbook name
'                     strTargetWSName      - Worksheet name
'                     lngTargetStartRow    - Start row number of range
'                     intTargetStartColumn - Start column number of range
'                     lngTargetEndRow      - End row number of range
'                     intTargetEndColumn   - End column number of range
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initCollection( _
    ByVal strTargetWBName As String, _
    ByVal strTargetWSName As String, _
    ByVal lngTargetStartRow As Long, _
    ByVal intTargetStartColumn As Integer, _
    ByVal lngTargetEndRow As Long, _
    ByVal intTargetEndColumn As Integer _
)

Dim i As Long
Dim j As Long
Dim varTargetArea As Variant
Dim objWS As Worksheet

lngCollectionRow = lngTargetEndRow - lngTargetStartRow + 1
intCollectionColumn = intTargetEndColumn - intTargetStartColumn + 1

Set objWS = Workbooks(strTargetWBName).Worksheets(strTargetWSName)

Set varTargetArea = objWS.Range(objWS.Cells(lngTargetStartRow, intTargetStartColumn), _
                                objWS.Cells(lngTargetEndRow, intTargetEndColumn) _
                    )

ReDim objCollection(lngCollectionRow - 1) As Record

For i = 0 To lngCollectionRow - 1

    ReDim Preserve objCollection(i).d(intCollectionColumn - 1) As String
    
    For j = 0 To intCollectionColumn - 1
        objCollection(i).d(j) = varTargetArea(i + 1, j + 1).Value
    Next j
    
Next i

Set varTargetArea = Nothing
Set objWS = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initValueOnly
' Description       : Set row and column count
' Parameters        : intSetRow       - Set row count
'                     intSetColumn    - Set column count
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initValueOnly( _
    ByVal intSetRow As Integer, _
    ByVal intSetColumn As Integer _
)

lngCollectionRow = intSetRow
intCollectionColumn = intSetColumn

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : releaseClass
' Description       : Release record object
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub releaseCollection()

Erase objCollection

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : getRow
' Description       : Get row count
' Parameters        : Nothing
' Value Returned    : Long
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getRow&()

getRow = lngCollectionRow

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : getColumn
' Description       : Get column count
' Parameters        : Nothing
' Value Returned    : Integer
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getColumn%()

getColumn = intCollectionColumn

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : debugPrint
' Description       : Print all record
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub debugPrint()

Dim i As Long
Dim j As Long

Debug.Print "Row countÅF" & getRow()
Debug.Print "Column countÅF" & getColumn()

For i = 0 To getRow() - 1
    For j = 0 To getColumn() - 1
        Debug.Print "[" & i & "," & j & "]:" & objCollection(i).d(j)
    Next j
Next i

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : getValue
' Description       : Get record value
' Parameters        : Nothing
' Value Returned    : Variant
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getValue( _
    ByVal lngRow As Long, _
    ByVal intColumn As Integer _
) As Variant

getValue = objCollection(lngRow).d(intColumn)

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : setValue
' Description       : Set record value
' Parameters        : lngRow    - Row index number
'                     intColumn - Column index number
'                     strValue  - Set value
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub setValue( _
    ByVal lngRow As Long, _
    ByVal intColumn As Integer, _
    ByVal strValue As String _
)

objCollection(lngRow).d(intColumn) = strValue

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : addCollection
' Description       : Add record object
' Parameters        : strTargetWBName      - Workbook name
'                     strTargetWSName      - Worksheet name
'                     lngTargetStartRow    - Start row number of range
'                     intTargetStartColumn - Start column number of range
'                     lngTargetEndRow      - End row number of range
'                     intTargetEndColumn   - End column number of range
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub addCollection( _
    ByVal strTargetWBName As String, _
    ByVal strTargetWSName As String, _
    ByVal lngTargetStartRow As Long, _
    ByVal intTargetStartColumn As Integer, _
    ByVal lngTargetEndRow As Long, _
    ByVal intTargetEndColumn As Integer _
)

Dim i As Long
Dim j As Long
Dim varTargetArea As Variant
Dim objWS As Worksheet

lngCollectionRow = lngCollectionRow + lngTargetEndRow - lngTargetStartRow + 1

If (intTargetEndColumn - intTargetStartColumn + 1 <> intCollectionColumn) Then
    intTargetEndColumn = intTargetStartColumn + intCollectionColumn - 1
End If

Set objWS = Workbooks(strTargetWBName).Worksheets(strTargetWSName)

Set varTargetArea = objWS.Range(objWS.Cells(lngTargetStartRow, intTargetStartColumn), _
                                objWS.Cells(lngTargetEndRow, intTargetEndColumn) _
                    )

ReDim Preserve objCollection(lngCollectionRow - 1) As Record

For i = lngCollectionRow - (lngTargetEndRow - lngTargetStartRow + 1) To lngCollectionRow - 1

    ReDim Preserve objCollection(i).d(intCollectionColumn - 1) As String
    
    For j = 0 To intCollectionColumn - 1
    
        objCollection(i).d(j) = varTargetArea(i + 1 - (lngCollectionRow - (lngTargetEndRow - lngTargetStartRow + 1)), j + 1).Value
    
    Next j
    
Next i

Set varTargetArea = Nothing
Set objWS = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : addRecord
' Description       : Add record object
' Parameters        : pvarInsertRecordValue - Add record value
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub addRecord( _
    ByRef pvarInsertRecordValue _
)

Dim i As Long
Dim intColumn As Integer

intColumn = UBound(pvarInsertRecordValue) - LBound(pvarInsertRecordValue) + 1

lngCollectionRow = lngCollectionRow + 1
ReDim Preserve objCollection(lngCollectionRow - 1) As Record
ReDim Preserve objCollection(lngCollectionRow - 1).d(intColumn - 1) As String
    
For i = 0 To intCollectionColumn - 1
    
    If ((i + 1) <= intColumn And (i + 1) <= intCollectionColumn) Then
        objCollection(lngCollectionRow - 1).d(i) = CStr(pvarInsertRecordValue(i))
    Else
        objCollection(lngCollectionRow - 1).d(i) = ""
    End If
    
Next i

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : pasteCollection
' Description       : Paste collection value to worksheet
' Parameters        : strTargetWBName      - Workbook name
'                     strTargetWSName      - Worksheet name
'                     lngTargetStartRow    - Start row number of range
'                     intTargetStartColumn - Start column number of range
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub pasteCollection( _
    ByVal strTargetWBName As String, _
    ByVal strTargetWSName As String, _
    ByVal lngTargetStartRow As Long, _
    ByVal intTargetStartColumn As Integer _
)

Dim varPasteCollection() As Variant
Dim i As Long
Dim j As Integer
Dim objWS As Worksheet

ReDim varPasteCollection(getRow() - 1, getColumn() - 1) As Variant

For i = 0 To getRow() - 1
    For j = 0 To getColumn() - 1
        varPasteCollection(i, j) = objCollection(i).d(j)
    Next j
Next i

Set objWS = Workbooks(strTargetWBName).Worksheets(strTargetWSName)

objWS.Range(objWS.Cells(lngTargetStartRow, intTargetStartColumn), _
            objWS.Cells(lngTargetStartRow + getRow() - 1, intTargetStartColumn + getColumn() - 1) _
) = varPasteCollection

Erase varPasteCollection

Set objWS = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : getSearchItemValue
' Description       : Serch item value
' Parameters        : intSearchColumn - Search column index number
'                     intReturnColumn - Return column index number
'                     strSearchString - Search column string value
' Value Returned    : Variant
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function getSearchItemValue( _
    ByVal intSearchColumn As Integer, _
    ByVal intReturnColumn As Integer, _
    ByVal strSearchString As String _
) As Variant

Dim i As Long
Dim strReturnValue As String

For i = 0 To lngCollectionRow - 1

  If getValue(i, intSearchColumn - 1) = strSearchString Then
  
    strReturnValue = getValue(i, intReturnColumn - 1)
    Exit For
    
  End If

Next i

getSearchItemValue = strReturnValue

End Function
