Sub GetFirstName()
    Dim fullName As String
    Dim firstName As String
    
    fullName = CreateObject("wscript.shell").RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\UserName")
    
    ' Split by comma and take the first part
    firstName = Split(fullName, ",")(0)
    
    ' Trim in case there are extra spaces
    firstName = Trim(firstName)
    
    MsgBox firstName
End Sub
