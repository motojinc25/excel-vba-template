Attribute VB_Name = "KBWin32API"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetPrivateProfileString
' Description       : Declaring GetPrivateProfileStringA Win32API
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : http://msdn.microsoft.com/
' Called By         : This.GetSectionEntryString
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
)

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetPrivateProfileInt
' Description       : Declaring GetPrivateProfileIntA Win32API
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : http://msdn.microsoft.com/
' Called By         : This.GetSectionEntryInt
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetPrivateProfileInt& Lib "kernel32" Alias "GetPrivateProfileIntA" ( _
    ByVal lpAppName As String, _
    ByVal lpKeyName As String, _
    ByVal nDefault As Long, _
    ByVal lpFileName As String _
)

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : WritePrivateProfileString
' Description       : Declaring WritePrivateProfileStringA Win32API
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : http://msdn.microsoft.com/
' Called By         : This.WriteSectionEntry
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal sSectionName As String, _
    ByVal sKeyName As String, _
    ByVal sString As String, _
    ByVal sFileName As String _
)

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetSectionEntryString
' Description       : Retrieves a string from the specified section in an initialization file
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strSectionName      - A section name of initialization file
'                     strEntryName        - An entry name of initialization file
'                     strFullPathInitFile - An initialization file with full path
' Called By         : Nothing
' Value Returned    : String
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSectionEntryString$( _
    ByVal strSectionName As String, _
    ByVal strEntryName As String, _
    ByVal strFullPathInitFile As String _
)

Dim lngStringLength As Long
Dim strReturnBuffer As String
Dim intLengthBuffer As Integer
Dim strReturnString As String

strReturnBuffer = Strings.String$(256, 0)  ' 256 null characters
strReturnString = vbNullChar
intLengthBuffer = Len(strReturnBuffer)

lngStringLength = GetPrivateProfileString(strSectionName, strEntryName, "", strReturnBuffer, intLengthBuffer, strFullPathInitFile)
strReturnString = Strings.Trim(Strings.Left$(strReturnBuffer, lngStringLength))

GetSectionEntryString = strReturnString

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : GetSectionEntryInt
' Description       : Retrieves an integer associated with a key in the specified section of an initialization file
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strSectionName      - A section name of initialization file
'                     strEntryName        - An entry name of initialization file
'                     strFullPathInitFile - An initialization file with full path
' Called By         : Nothing
' Value Returned    : Long
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetSectionEntryInt&( _
    ByVal strSectionName As String, _
    ByVal strEntryName As String, _
    ByVal strFullPathInitFile As String _
)

Dim lngReturnValue As Long

lngReturnValue = GetPrivateProfileInt(strSectionName, strEntryName, 0, strFullPathInitFile)

GetSectionEntryInt = lngReturnValue

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function          : WriteSectionEntry
' Description       : Copies a string into the specified section of an initialization file
' Author            : Jingun Jung
' Licence           : Apache License 2.0
' Source            : https://github.com/koreabigname/excel-vba-template
' Date              : 2015-04-18
' Parameters        : strSectionName      - A section name of initialization file
'                     strEntryName        - An entry name of initialization file
'                     strEntryValue       - An entry value of initialization file
'                     strFullPathInitFile - An initialization file with full path
' Called By         : Nothing
' Value Returned    : Nothing
' Modification History
'
'   Author          Date          Reason      Comment
'   ------------    ----------    --------    ---------
'
''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub WriteSectionEntryString( _
    ByVal strSectionName As String, _
    ByVal strEntryName As String, _
    ByVal strEntryValue As String, _
    ByVal strFullPathInitFile As String _
)

Dim lngWriteResult As Long
Dim strReturnBuffer As String
Dim intLengthBuffer As Integer
Dim strReturnString As String

lngWriteResult = WritePrivateProfileString(strSectionName, strEntryName, strEntryValue, strFullPathInitFile)

End Sub
