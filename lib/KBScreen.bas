Attribute VB_Name = "KBScreen"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : LockScreen
' Description       : When vba starts, screen updating is turned off.
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : Nothing
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'   Jingun Jung     2015-05-13    Updated     Issue #25
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LockScreen()

ThisWorkbook.Activate
Worksheets(C_WS_WORK).Select
Range("A1").Select

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "Starting VBA Program."
Application.Calculate = xlCalculationManual

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : UnLockScreen
' Description       : When vba ends, screen updating is turned on.
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : Nothing
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'   Jingun Jung     2015-05-13    Updated     Issue #25
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UnLockScreen()

ThisWorkbook.Activate
Worksheets(C_WS_STARTUP).Select
Range("A1").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = False
Application.Calculate = xlCalculationAutomatic

End Sub
