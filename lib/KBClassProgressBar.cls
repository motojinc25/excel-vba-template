VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KBClassProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class             : KBClassProgressBar
' Description       : A Form class for KBFormProgressBar
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
Private Const C_BAR_MAX_VALUE As Integer = 100
Private intBar1Percent        As Integer
Private intBar2Percent        As Integer
Private intBar3Percent        As Integer
Private objForm               As New KBFormProgressBar

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Initialize
' Description       : Initialize progress bar class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

Call initUI("KBFormProgressBar", _
             RGB(255, 0, 0), _
             RGB(0, 0, 255), _
             RGB(0, 255, 0), _
             "String1", _
             "String2", _
             "String3" _
)

intBar1Percent = 0
intBar2Percent = 0
intBar3Percent = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : Class_Terminate
' Description       : Terminate progress bar class
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()

Set objForm = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initUI
' Description       : Initialize progress bar object
' Parameters        : strFormCaption - Form Caption
'                     lngBarGr1Color - Color of group 1 bar
'                     lngBarGr2Color - Color of group 2 bar
'                     lngBarGr3Color - Color of group 3 bar
'                     strBarGr1Msg   - Label of group 1 bar
'                     strBarGr2Msg   - Label of group 2 bar
'                     strBarGr3Msg   - Label of gruop 3 bar
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initUI( _
    ByVal strFormCaption As String, _
    ByVal lngBarGr1Color As Long, _
    ByVal lngBarGr2Color As Long, _
    ByVal lngBarGr3Color As Long, _
    ByVal strBarGr1Msg As String, _
    ByVal strBarGr2Msg As String, _
    ByVal strBarGr3Msg As String _
)

objForm.Caption = strFormCaption
objForm.lblBarGr1_Fg.BackColor = lngBarGr1Color
objForm.lblBarGr2_Fg.BackColor = lngBarGr2Color
objForm.lblBarGr3_Fg.BackColor = lngBarGr3Color
objForm.lblBarGr1_Msg.Caption = strBarGr1Msg
objForm.lblBarGr2_Msg.Caption = strBarGr2Msg
objForm.lblBarGr3_Msg.Caption = strBarGr3Msg

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initAllValue
' Description       : Initialize all progress bar value
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initAllValue()

intBar1Percent = 0
intBar2Percent = 0
intBar3Percent = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initValueGr1
' Description       : Initialize progress bar 1 value
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initValueGr1()

intBar1Percent = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initValueGr2
' Description       : Initialize progress bar 2 value
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initValueGr2()

intBar2Percent = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : initValueGr3
' Description       : Initialize progress bar 3 value
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub initValueGr3()

intBar3Percent = 0

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : repaintUI
' Description       : Repaint all progress bar objects
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub repaintUI()

objForm.repaint
objForm.frmBarGr1.repaint
objForm.frmBarGr2.repaint
objForm.frmBarGr3.repaint

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : showUI
' Description       : When form opens, it keeps the cocus until it is closed
' Parameters        : blnModalShowFlag
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub showUI( _
    ByVal blnModalShowFlag As Boolean _
)

Call repaintUI

If (blnModalShowFlag = True) Then
    objForm.Show (vbModal)
Else
    objForm.Show (vbModeless)
End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : hideUI
' Description       : Hide progress bar form
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub hideUI()

objForm.hide

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : hideGroup1
' Description       : Hide group 1 bar on the form
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub hideGroup1()

objForm.frmBarGr1.Visible = False
objForm.lblBarGr1_Msg.Visible = False

Call repaintUI

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : showGroup1
' Description       : Show group 1 bar on the form
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub showGroup1()

objForm.frmBarGr1.Visible = True
objForm.lblBarGr1_Msg.Visible = True

Call repaintUI

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : releaseClass
' Description       : Release progress bar objects
' Parameters        : Nothing
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub releaseClass()

Set objForm = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Method            : UpdateUI
' Description       : Update progress bar objects
' Parameters        : intBarGroupNumber - Which progress bar
'                     lngCurrentValue   - Bar value
'                     lngMaxValue       - Bar max value
' Value Returned    : Nothing
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub updateUI( _
    ByVal intBarGroupNumber As Integer, _
    ByVal lngCurrentValue As Long, _
    ByVal lngMaxValue As Long _
)

Dim intBarPercentValue As Integer
Dim intBarValue As Integer

intBarPercentValue = Int(lngCurrentValue / lngMaxValue * 100)
intBarValue = Int(intBarPercentValue * (C_ProgressBarMaxValue / 100))

Select Case intBarGroupNumber
Case 1

    If (intBarPercentValue > intBar1Percent) Then
  
        intBar1Percent = intBarPercentValue
        objForm.lblBarGr1_Fg.Width = intBarValue
    
        Call objForm.frmBarGr1.repaint
    
    End If

Case 2
  
    If (intBarPercentValue > intBar2Percent) Then
  
        intBar2Percent = intBarPercentValue
        objForm.lblBarGr2_Fg.Width = intBarValue
    
        Call objForm.frmBarGr2.repaint
    
    End If

Case 3
  
    If (intBarPercentValue > intBar3Percent) Then
  
        intBar3Percent = intBarPercentValue
        objForm.lblBarGr3_Fg.Width = intBarValue
    
        Call objForm.frmBarGr3.repaint
    
    End If

End Select

End Sub
