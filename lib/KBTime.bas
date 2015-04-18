Attribute VB_Name = "KBTime"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : WaitTime
' Description       : Pauses a running macro until a specified time
' Author            : Jingun Jung (Webpage: www.koreabigname.com)
' Date              : 2015-04-17
' Parameters        : intWaitSecond - Pauses a running macro until a specified time
' Called By         : Nothing
' Value Returned    : Nothing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WaitTime( _
    ByVal intWaitSecond As Integer _
)

Application.Wait Now + TimeSerial(0, 0, intWaitSecond)

End Sub


