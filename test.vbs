Option Compare Database
Option Explicit

' ========================================
' USER ENTITLEMENTS + NAVIGATION HANDLER
' ========================================

Private m_strCurrentUser As String
Private m_colUserEntitlements As Collection   ' holds current user entitlements

' ========================================
' FORM LOAD
' ========================================
Private Sub Form_Load()
    m_strCurrentUser = GetCurrentUserName()
    Call UpdateUserEntitlements
End Sub

' ========================================
' USER ENTITLEMENT MANAGEMENT
' ========================================
Private Sub UpdateUserEntitlements()
    On Error GoTo ErrorHandler
    
    Debug.Print "Refreshing entitlements for user: " & m_strCurrentUser
    Set m_colUserEntitlements = LoadUserEntitlements(m_strCurrentUser)
    Debug.Print "Entitlements loaded: " & m_colUserEntitlements.Count
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading entitlements: " & Err.Description, vbExclamation
End Sub

Private Function LoadUserEntitlements(strUser As String) As Collection
    Dim colEntitlements As New Collection
    ' TODO: Replace with actual DB query:
    ' Example: SELECT EntitlementName FROM tblUserEntitlements WHERE UserName = strUser
    
    ' Demo entitlements (remove/modify as needed)
    colEntitlements.Add "Dashboard"
    colEntitlements.Add "Report Generator"
    colEntitlements.Add "Account"
    colEntitlements.Add "Profile"
    colEntitlements.Add "Help"
    
    Set LoadUserEntitlements = colEntitlements
End Function

Private Function HasEntitlement(strEntitlement As String) As Boolean
    Dim varItem As Variant
    HasEntitlement = False
    
    If m_colUserEntitlements Is Nothing Then Exit Function
    
    For Each varItem In m_colUserEntitlements
        If StrComp(varItem, strEntitlement, vbTextCompare) = 0 Then
            HasEntitlement = True
            Exit Function
        End If
    Next varItem
End Function

Private Function GetUserRole() As String
    ' Very basic mapping – replace with real logic from DB if available
    If HasEntitlement("Account") Then
        GetUserRole = "Administrator"
    ElseIf HasEntitlement("Report Generator") Then
        GetUserRole = "Analyst"
    ElseIf HasEntitlement("Dashboard") Then
        GetUserRole = "Standard User"
    Else
        GetUserRole = "Guest"
    End If
End Function

' ========================================
' EDGE BROWSER EVENT HANDLERS
' ========================================

Private Sub EdgeBrowser0_DocumentComplete(URL As Variant)
    Dim sCmd As String
    Dim sInitials As String
    Dim sRole As String
    
    ' --- Inject JS click interception ---
    sCmd = "var script = document.createElement('script');"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    
    sCmd = "script.innerHTML = " & Chr(34) & _
        "function interceptClickEvent(e) {" & _
        "var elementId;" & _
        "var target = e.target || e.srcElement;" & _
        "elementId = e.target.getAttribute('id');" & _
        "console.log('target->' + e.target);" & _
        "console.log('target.id->' + e.target.id);" & _
        "document.getElementById('clickedElement').value = elementId;" & _
        "e.preventDefault();" & _
        "}" & Chr(34)
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    
    sCmd = "document.body.appendChild(script);"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    
    sCmd = "var hiddenInput = document.createElement('input');"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    sCmd = "hiddenInput.type = 'hidden';"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    sCmd = "hiddenInput.id = 'clickedElement';"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    sCmd = "document.body.appendChild(hiddenInput);"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    
    sCmd = "document.addEventListener('click', interceptClickEvent);"
    Me.EdgeBrowser0.ExecuteJavascript sCmd
    
    ' --- Inject Profile Info ---
    sInitials = Left(m_strCurrentUser, 2)   ' Take first 2 letters of username
    sRole = GetUserRole()
    
    Me.EdgeBrowser0.ExecuteJavascript _
        "document.getElementById('profile-avatar').innerText = '" & sInitials & "';"
    
    Me.EdgeBrowser0.ExecuteJavascript _
        "document.getElementById('profile-name').innerText = '" & m_strCurrentUser & "';"
    
    Me.EdgeBrowser0.ExecuteJavascript _
        "document.getElementById('profile-role').innerText = '" & sRole & "';"
End Sub

Private Sub EdgeBrowser0_Click()
    Dim sClickedElem As String
    sClickedElem = Me.EdgeBrowser0.RetrieveJavascriptValue("document.getElementById('clickedElement').value")
    
    If Len(sClickedElem) = 0 Then Exit Sub
    
    If Not HasEntitlement(sClickedElem) Then
        MsgBox "Access denied. You do not have entitlement for: " & sClickedElem, vbExclamation
        Exit Sub
    End If
    
    Select Case sClickedElem
        Case "Dashboard"
            Me.subform.SourceObject = "frmDashboard"
        Case "Report Generator"
            Me.subform.SourceObject = "frmReportGenerator"
        Case "Account"
            Me.subform.SourceObject = "frmAccount"
        Case "Profile"
            Me.subform.SourceObject = "frmProfile"
        Case "security"
            MsgBox "You clicked on Security"
        Case "password"
            MsgBox "You clicked on Password"
        Case "help"
            Call OpenHelpSupportEmail
        Case "signout"
            DoCmd.Quit
    End Select
End Sub

' ========================================
' HELP / EMAIL SUPPORT
' ========================================
Private Sub OpenHelpSupportEmail()
    Dim olApp As Object, olMail As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If Not olApp Is Nothing Then
        Set olMail = olApp.CreateItem(0)
        With olMail
            .To = "support@company.com"
            .CC = "it-helpdesk@company.com"
            .Subject = "CMR Pulse - Help Request"
            .Body = "User: " & m_strCurrentUser & vbCrLf & _
                    "Issue: [Please describe your issue here]" & vbCrLf & vbCrLf & _
                    "Thank you."
            .Display
        End With
    Else
        MsgBox "Unable to launch Outlook. Please contact IT support.", vbExclamation
    End If
End Sub

' ========================================
' HELPER FUNCTIONS
' ========================================
Private Function GetCurrentUserName() As String
    GetCurrentUserName = Environ("USERNAME")
End Function
