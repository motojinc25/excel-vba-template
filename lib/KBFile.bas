Attribute VB_Name = "KBFile"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetDriveNameForSpecifiedPath
' Description       : Get a string containing the name of the drive in a supplied path
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDriveNameForSpecifiedPath$( _
    ByVal strFullPath As String _
)

Dim objFso As Object

Set objFso = CreateObject("Scripting.FilesystemObject")

GetDriveNameForSpecifiedPath = objFso.GetDriveName(strFullPath)

Set objFso = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetParentFolderNameForSpecifiedPath
' Description       : Return a string containing the name of the parent folder of the last file or folder in a supplied path
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetParentFolderNameForSpecifiedPath$( _
    ByVal strFullPath As String _
)

Dim objFso As Object

Set objFso = CreateObject("Scripting.FilesystemObject")

GetParentFolderNameForSpecifiedPath = objFso.GetParentFolderName(strFullPath)

Set objFso = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetFileNameForSpecifiedPath
' Description       : Return the name of the last file or folder of the supplied path
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetFileNameForSpecifiedPath$( _
    ByVal strFullPath As String _
)

Dim objFso As Object

Set objFso = CreateObject("Scripting.FilesystemObject")

GetFileNameForSpecifiedPath = objFso.GetFileName(strFullPath)

Set objFso = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetBaseNameForSpecifiedPath
' Description       : Get the base name of the file or folder in a supplied path
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetBaseNameForSpecifiedPath$( _
    ByVal strFullPath As String _
)

Dim objFso As Object

Set objFso = CreateObject("Scripting.FilesystemObject")

GetBaseNameForSpecifiedPath = objFso.GetBaseName(strFullPath)

Set objFso = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetExtensionNameForSpecifiedPath
' Description       : Return a string containing the extension name of the last component in a supplied path
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetExtensionNameForSpecifiedPath$( _
    ByVal strFullPath As String _
)

Dim objFso As Object

Set objFso = CreateObject("Scripting.FilesystemObject")

GetExtensionNameForSpecifiedPath = objFso.GetExtensionName(strFullPath)

Set objFso = Nothing

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : IsExistFile
' Description       : Check If a file exists
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : KBInit.LoadInitFile
' Value Returned    : Boolean - True is existing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsExistFile( _
    ByVal strFullPath _
) As Boolean

Dim objFso As Object
Dim blnExistResult As Boolean

blnExistResult = False

Set objFso = CreateObject("Scripting.FilesystemObject")

If (objFso.FileExists(strFullPath)) Then

    blnExistResult = True
    
End If

Set objFso = Nothing

IsExistFile = blnExistResult

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : IsExistFolder
' Description       : Check If a folder exists
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFullPath - a filename with path
' Called By         : KBInit.LoadInitFile
' Value Returned    : Boolean - True is existing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsExistFolder( _
    ByVal strDirName _
) As Boolean

Dim objFso As Object
Dim blnExistResult As Boolean

blnExistResult = False

Set objFso = CreateObject("Scripting.FilesystemObject")

If (objFso.FolderExists(strDirName)) Then

    blnExistResult = True

End If

Set objFso = Nothing

IsExistFolder = blnExistResult

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : OpenFiles
' Description       : Open a file or files
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strFileFilter      - Filter (Visual Basic ƒtƒ@ƒCƒ‹ (*.bas;*.txt),*.bas;*.txt)
'                     strDialogTitle     - Open dialog title
'                     blnMultiSelectFlag - Flag multi select
'                     pvarOpenFiles      - Variant type
' Called By         : Nothing
' Value Returned    : Integer
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OpenFiles%( _
    ByVal strFileFilter As String, _
    ByVal strDialogTitle As String, _
    ByVal blnMultiSelectFlag As Boolean, _
    ByRef pvarOpenFiles _
)

Dim intFileCount As Integer

intFileCount = 0

pvarOpenFiles = Application.GetOpenFilename(FileFilter:=strFileFilter, _
                                            Title:=strDialogTitle, _
                                            MultiSelect:=blnMultiSelectFlag _
)

If (IsArray(pvarOpenFiles) = True) Then
    
    intFileCount = GetSizeArray(pvarOpenFiles)
        
End If

OpenFiles = intFileCount

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : ReadTextFileToRecordCollection
' Description       : Read a text file to record collection object
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-21
' Parameters        : strReadFileFullPath - Read text file with path
'                     strReturnCodeInFile - Return Code (vbLf, vbCrLf, vbCr)
'                     pobjCollection      - Record collection object
' Called By         : Nothing
' Value Returned    : Nothing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReadTextFileToRecordCollection( _
  ByVal strReadFileFullPath As String, _
  ByVal strReturnCodeInFile As String, _
  ByRef pobjCollection As KBClassRecordCollection _
)

Dim intFileNum          As Integer
Dim bytBuffer()         As Byte
Dim strBuffer           As String
Dim strRecord           As String
Dim lngLPos             As Long
Dim intReturnCodeLen    As Integer
Dim lngReadLen          As Long
Dim lngSurplus          As Long
Dim lngFileSize         As Long
Dim strArray(1)         As String

lngFileSize = FileLen(strReadFileFullPath)
lngReadLen = 144
lngSurplus = lngFileSize Mod lngReadLen

strReturnCodeInFile = StrConv(strReturnCodeInFile, vbFromUnicode)
intReturnCodeLen = LenB(strReturnCodeInFile)

ReDim bytBuffer(1 To lngReadLen)

intFileNum = FreeFile
Open strReadFileFullPath For Binary As intFileNum

Do Until LOF(intFileNum) <= Loc(intFileNum)

    If LOF(intFileNum) - Loc(intFileNum) <= lngReadLen Then
        ReDim bytBuffer(1 To (LOF(intFileNum) - Loc(intFileNum)))
    End If
    
    Get #intFileNum, , bytBuffer
    strBuffer = strBuffer & CStr(bytBuffer)
    lngLPos = InStrB(1, strBuffer, strReturnCodeInFile)
    Do Until (lngLPos = 0)
    
        If lngLPos > 0 Then
            strRecord = strRecord & LeftB(strBuffer, lngLPos - 1)
            strBuffer = MidB(strBuffer, lngLPos + intReturnCodeLen)
        Else
            If LOF(intFileNum) <= Loc(intFileNum) Then
                strRecord = strRecord & strBuffer
            End If
        End If
        
        If lngLPos > 0 Or LOF(intFileNum) <= Loc(intFileNum) Then
            
            strArray(0) = StrConv(strRecord, vbUnicode)
            Call pobjCollection.addRecord(strArray)
            strRecord = ""
            
        End If
        
        lngLPos = InStrB(1, strBuffer, strReturnCodeInFile)
        
    Loop
    
Loop

If (LenB(strBuffer) > 0) Then
    strArray(0) = StrConv(strBuffer, vbUnicode)
    Call pobjCollection.addRecord(strArray)
End If

Erase strArray
Close intFileNum

End Sub
