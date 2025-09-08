Sub GetFirstName()
    Dim fullName As String
    Dim parts As Variant
    Dim firstName As String
    
    fullName = CreateObject("wscript.shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\UserName")
    
    ' Step 1: Split by comma
    parts = Split(fullName, ",")
    
    If UBound(parts) >= 1 Then
        ' Step 2: Take second part (after comma)
        firstName = Trim(parts(1))
        
        ' Step 3: Remove everything after space + "[" if it exists
        If InStr(firstName, "[") > 0 Then
            firstName = Trim(Left(firstName, InStr(firstName, "[") - 1))
        End If
    Else
        firstName = fullName ' fallback in case no comma
    End If
    
    MsgBox firstName
End Sub
