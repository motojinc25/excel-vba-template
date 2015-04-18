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
' Called By         : Nothing
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
