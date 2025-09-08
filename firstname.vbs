#If VBA7 Then
    Private Declare PtrSafe Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal lpValue As String, _
        ByVal dwFlags As Long, _
        ByVal pdwType As LongPtr, _
        ByVal pvData As String, _
        ByRef pcbData As Long) As Long
#Else
    Private Declare Function RegGetValue Lib "advapi32.dll" Alias "RegGetValueA" ( _
        ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal lpValue As String, _
        ByVal dwFlags As Long, _
        ByVal pdwType As Long, _
        ByVal pvData As String, _
        ByRef pcbData As Long) As Long
#End If

Const HKEY_CURRENT_USER = &H80000001
Const RRF_RT_REG_SZ = &H2

Function GetOfficeUserName() As String
    Dim sKey As String, sValue As String
    Dim ret As Long, dataSize As Long

    sKey = "Software\Microsoft\Office\Common\UserInfo"
    sValue = String(255, vbNullChar)
    dataSize = Len(sValue)

    ret = RegGetValue(HKEY_CURRENT_USER, sKey, "UserName", RRF_RT_REG_SZ, 0, sValue, dataSize)
    If ret = 0 Then
        GetOfficeUserName = Left$(sValue, dataSize - 1)
    Else
        GetOfficeUserName = ""
    End If
End Function

Sub GetFirstName()
    Dim fullName As String, firstName As String, parts As Variant
    
    fullName = GetOfficeUserName()
    
    If fullName <> "" Then
        ' Split by comma
        parts = Split(fullName, ",")
        If UBound(parts) >= 1 Then
            firstName = Trim(parts(1))
            If InStr(firstName, "[") > 0 Then
                firstName = Trim(Left(firstName, InStr(firstName, "[") - 1))
            End If
        Else
            firstName = fullName
        End If
        MsgBox firstName
    Else
        MsgBox "UserName not found in registry!"
    End If
End Sub
