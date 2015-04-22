Attribute VB_Name = "KBArray"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetSizeArray
' Description       : Get size of an array
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : pvarArray - an array which is all type by reference
' Called By         : KBFile.OpenFiles
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSizeArray&( _
    ByRef pvarArray _
)

Dim lngSize As Long

If (IsArray(pvarArray) = True) Then

    lngSize = UBound(pvarArray) - LBound(pvarArray) + 1
    
Else

    lngSize = 0

End If

GetSizeArray = lngSize

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetSumValueArray
' Description       : Sum of items in a collection
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-22
' Parameters        : pvarArray - an array which is all type by reference
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSumValueArray&( _
    ByRef pvarArray _
)

Dim lngSumValue As Long
Dim varIdxItemValue As Variant

For Each varIdxItemValue In pvarArray

    If IsNumeric(varIdxItemValue) = True Then
    
        lngSumValue = lngSumValue + varIdxItemValue
        
    End If
Next

GetSumValueArray = lngSumValue

End Function

