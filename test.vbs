' ========================================
' COMPLETE CMR PULSE VBA NAVIGATION SYSTEM
' WITH USER MANAGEMENT AND LOGOUT FUNCTIONALITY
' ========================================
'
' FORM SETUP REQUIREMENTS:
' - Form Name: "frmMainNavigation" (or your preferred name)
' - WebView2 Control Name: "EdgeBrowser0"
' - Subform Control Name: "subform"
' - HTML File: Updated index.html with user management section
'
' Author: Enhanced Navigation System with User Management
' Version: 3.0
' Features: Form navigation, User management, Logout, Help & Support
' ========================================

Option Compare Database
Option Explicit

' ========================================
' MODULE-LEVEL CONSTANTS
' ========================================

' Child Form Names (update these to match your actual forms)
Private Const FORM_DASHBOARD As String = "frmDashboard"
Private Const FORM_REPORT_GENERATOR As String = "frmReportGenerator"
Private Const FORM_ADMIN As String = "frmAdmin"
Private Const FORM_USER_MANAGEMENT As String = "frmUserManagement"
Private Const FORM_LOGIN As String = "frmLogin"

' Email Configuration
Private Const SUPPORT_EMAIL As String = "support@company.com"
Private Const IT_HELPDESK_EMAIL As String = "it-helpdesk@company.com"
Private Const EMAIL_SUBJECT As String = "CMR Pulse - Help Request"

' Navigation Element IDs (matching the HTML file)
Private Const NAV_DASHBOARD As String = "Dashboard"
Private Const NAV_REPORT_GENERATOR As String = "Report Generator"
Private Const NAV_ADMIN As String = "security"
Private Const NAV_USER_MANAGEMENT As String = "password"
Private Const NAV_HELP_SUPPORT As String = "help"
Private Const NAV_LOGOUT As String = "logout"

' Update this path to your HTML file location
Private Const HTML_FILE_PATH As String = "file:///C:/CMRPulse/sidebar/index.html"

' ========================================
' MODULE-LEVEL VARIABLES
' ========================================

Private m_strCurrentFormName As String
Private m_strCurrentUser As String
Private m_bWebViewReady As Boolean

' ========================================
' FORM INITIALIZATION
' ========================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CMR Pulse Navigation System Starting ==="
    
    ' Initialize variables
    Call InitializeFormVariables
    
    ' Set up WebView2 control
    Call SetupWebViewControl
    
    ' Initialize with Dashboard
    Call NavigateToFormByName(FORM_DASHBOARD)
    
    Debug.Print "Navigation system initialized successfully"
    Exit Sub
    
ErrorHandler:
    Call HandleError("Form_Load", Err.Description)
End Sub

Private Sub InitializeFormVariables()
    m_strCurrentFormName = ""
    m_strCurrentUser = GetCurrentUserName()
    m_bWebViewReady = False
    
    ' Set form caption
    Me.Caption = "CMR Pulse - " & m_strCurrentUser
    
    Debug.Print "Current User: " & m_strCurrentUser
End Sub

Private Sub SetupWebViewControl()
    On Error GoTo ErrorHandler
    
    Debug.Print "Setting up WebView2 control..."
    
    ' Navigate to HTML file
    Me.EdgeBrowser0.Navigate HTML_FILE_PATH
    
    ' Set up the scripting object for VBA-JavaScript communication
    Set Me.EdgeBrowser0.ObjectForScripting = Me
    
    Debug.Print "WebView2 control setup complete"
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error setting up WebView2: " & Err.Description
    MsgBox "Error loading navigation sidebar. Please check if the HTML file exists at: " & HTML_FILE_PATH
End Sub

' ========================================
' WEBVIEW2 EVENT HANDLERS
' ========================================

Private Sub EdgeBrowser0_DocumentComplete()
    Debug.Print "WebView2 document loaded successfully"
    
    m_bWebViewReady = True
    
    ' Update the logged-in user in the HTML sidebar
    Call UpdateLoggedInUserDisplay
    
    ' Set up click monitoring
    Call SetupClickMonitoring
End Sub

Private Sub EdgeBrowser0_NavigateComplete()
    Debug.Print "WebView2 navigation complete"
End Sub

Private Sub SetupClickMonitoring()
    On Error Resume Next
    
    ' Set up a timer to periodically check for clicks
    ' This is a simple approach - you might want to use a more sophisticated method
    Application.OnTime Now + TimeValue("00:00:01"), "CheckForClicks"
    
    On Error GoTo 0
End Sub

' ========================================
' USER MANAGEMENT FUNCTIONS
' ========================================

Private Sub UpdateLoggedInUserDisplay()
    On Error Resume Next
    
    If m_bWebViewReady Then
        ' Call the JavaScript function to update the user display
        Dim strScript As String
        strScript = "updateLoggedInUser('" & m_strCurrentUser & "');"
        
        Me.EdgeBrowser0.Document.parentWindow.execScript strScript, "JavaScript"
        Debug.Print "Updated logged-in user display to: " & m_strCurrentUser
    End If
    
    On Error GoTo 0
End Sub

Public Sub UpdateUserName(strNewUserName As String)
    ' This function can be called from other parts of your application
    ' to update the logged-in user
    
    If Len(strNewUserName) > 0 Then
        m_strCurrentUser = strNewUserName
        Call UpdateLoggedInUserDisplay
        
        ' Update form caption
        Me.Caption = "CMR Pulse - " & m_strCurrentUser
        
        Debug.Print "User name updated to: " & strNewUserName
    End If
End Sub

Private Sub ProcessLogout()
    On Error GoTo ErrorHandler
    
    Debug.Print "Processing user logout..."
    
    ' Close current child form
    Call CloseCurrentChildForm
    
    ' Reset user display in HTML
    If m_bWebViewReady Then
        Me.EdgeBrowser0.Document.parentWindow.execScript "resetUser();", "JavaScript"
    End If
    
    ' Clear session data
    Call ClearUserSession
    
    ' Show logout confirmation
    Dim intResponse As Integer
    intResponse = MsgBox("You have been logged out. Would you like to close the application?", _
                        vbYesNo + vbQuestion, "Logout Confirmation")
    
    If intResponse = vbYes Then
        ' Close application
        DoCmd.Quit acQuitSaveNone
    Else
        ' Redirect to login form (uncomment and modify as needed)
        ' DoCmd.OpenForm FORM_LOGIN
        ' DoCmd.Close acForm, Me.Name
        
        ' For now, just reset to guest user
        Call UpdateUserName("Guest User")
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("ProcessLogout", Err.Description)
End Sub

Private Sub ClearUserSession()
    m_strCurrentFormName = ""
    m_strCurrentUser = "Guest User"
    
    ' Add any additional session cleanup here
    Debug.Print "User session cleared"
End Sub

' ========================================
' CLICK DETECTION AND NAVIGATION
' 
' IMPORTANT: NO USER ENTITLEMENT CHECKS IN SIDEBAR
' - All sidebar navigation is unrestricted
' - Users can click any menu item (Admin, Users, etc.)
' - Entitlement/permission checks should be handled:
'   * Within individual forms themselves
'   * In form Load events
'   * In specific form functions
'   * NOT in the sidebar navigation system
' ========================================

Public Sub CheckForClicks()
    On Error Resume Next
    
    If Not m_bWebViewReady Then Exit Sub
    
    ' Get clicked element from JavaScript
    Dim strClickedElement As String
    strClickedElement = Me.EdgeBrowser0.Document.parentWindow.execScript("clickedElement", "JavaScript")
    
    ' Process if we have a clicked element
    If Len(strClickedElement) > 0 Then
        Debug.Print "Detected click on: " & strClickedElement
        Call ProcessNavigationRequest(strClickedElement)
        
        ' Clear the clicked element
        Me.EdgeBrowser0.Document.parentWindow.execScript "clickedElement = '';", "JavaScript"
    End If
    
    ' Schedule next check (every second)
    Application.OnTime Now + TimeValue("00:00:01"), "CheckForClicks"
    
    On Error GoTo 0
End Sub

Private Sub ProcessNavigationRequest(strElementId As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "Processing navigation request: " & strElementId
    
    ' Route to appropriate handler - NO USER ENTITLEMENT CHECKS
    ' All sidebar navigation is unrestricted
    Select Case strElementId
        Case NAV_DASHBOARD
            Call NavigateToFormByName(FORM_DASHBOARD)
            
        Case NAV_REPORT_GENERATOR
            Call NavigateToFormByName(FORM_REPORT_GENERATOR)
            
        Case NAV_ADMIN
            ' No entitlement check - direct navigation to admin
            Call NavigateToFormByName(FORM_ADMIN)
            
        Case NAV_USER_MANAGEMENT
            ' No entitlement check - direct navigation to user management
            Call NavigateToFormByName(FORM_USER_MANAGEMENT)
            
        Case NAV_HELP_SUPPORT
            Call OpenHelpSupportEmail
            
        Case NAV_LOGOUT
            Call ProcessLogout
            
        Case Else
            Debug.Print "Unknown navigation element: " & strElementId
    End Select
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("ProcessNavigationRequest", Err.Description)
End Sub

' ========================================
' FORM NAVIGATION MANAGEMENT
' ========================================

Private Sub NavigateToFormByName(strFormName As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "Navigating to form: " & strFormName & " (NO ENTITLEMENT CHECKS)"
    
    ' Close current form if different
    If m_strCurrentFormName <> strFormName Then
        Call CloseCurrentChildForm
    End If
    
    ' Validate form exists (basic check only, no user permissions)
    If Not FormExists(strFormName) Then
        MsgBox "Form '" & strFormName & "' does not exist. Please create this form first.", _
               vbExclamation, "Form Not Found"
        Debug.Print "Form does not exist: " & strFormName
        Exit Sub
    End If
    
    ' Load the form - NO USER ENTITLEMENT VALIDATION
    ' Direct navigation regardless of user permissions
    Call LoadFormInSubformControl(strFormName)
    
    ' Update tracking
    m_strCurrentFormName = strFormName
    
    ' Update form title
    Call UpdateFormTitle(GetFormDisplayName(strFormName))
    
    Debug.Print "Successfully navigated to: " & strFormName & " (unrestricted access)"
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("NavigateToFormByName", Err.Description)
End Sub

Private Sub LoadFormInSubformControl(strFormName As String)
    On Error GoTo ErrorHandler
    
    ' Set the source object for the subform control
    Me.subform.SourceObject = strFormName
    Me.subform.Visible = True
    
    ' Optimize display
    Call OptimizeSubformDisplay
    
    Debug.Print "Form loaded in subform control: " & strFormName
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("LoadFormInSubformControl", Err.Description)
End Sub

Private Sub CloseCurrentChildForm()
    On Error Resume Next
    
    If Len(m_strCurrentFormName) > 0 Then
        Debug.Print "Closing current form: " & m_strCurrentFormName
        
        Me.subform.SourceObject = ""
        Me.subform.Visible = False
        m_strCurrentFormName = ""
    End If
    
    On Error GoTo 0
End Sub

Private Sub OptimizeSubformDisplay()
    On Error Resume Next
    
    With Me.subform
        .BorderStyle = 0  ' Transparent
        .CanGrow = True
        .CanShrink = True
    End With
    
    On Error GoTo 0
End Sub

' ========================================
' HELP AND SUPPORT SYSTEM
' ========================================

Private Sub OpenHelpSupportEmail()
    On Error GoTo ErrorHandler
    
    Debug.Print "Opening help support email"
    
    Dim strEmailBody As String
    strEmailBody = BuildHelpEmailBody()
    
    Dim strMailtoLink As String
    strMailtoLink = BuildMailtoLink(SUPPORT_EMAIL, IT_HELPDESK_EMAIL, EMAIL_SUBJECT, strEmailBody)
    
    ' Open email client
    Application.FollowHyperlink strMailtoLink
    
    Debug.Print "Help email opened successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error opening email client: " & Err.Description
    
    ' Fallback: Show message with email details
    Dim strMessage As String
    strMessage = "Unable to open email client automatically." & vbCrLf & vbCrLf
    strMessage = strMessage & "Please send an email to:" & vbCrLf
    strMessage = strMessage & "To: " & SUPPORT_EMAIL & vbCrLf
    strMessage = strMessage & "CC: " & IT_HELPDESK_EMAIL & vbCrLf
    strMessage = strMessage & "Subject: " & EMAIL_SUBJECT
    
    MsgBox strMessage, vbInformation, "Help & Support"
End Sub

Private Function BuildHelpEmailBody() As String
    Dim strBody As String
    
    strBody = "Dear Support Team," & vbCrLf & vbCrLf
    strBody = strBody & "I need assistance with the CMR Pulse application." & vbCrLf & vbCrLf
    strBody = strBody & "Issue Description:" & vbCrLf
    strBody = strBody & "[Please describe your issue here]" & vbCrLf & vbCrLf
    strBody = strBody & "System Information:" & vbCrLf
    strBody = strBody & "User: " & m_strCurrentUser & vbCrLf
    strBody = strBody & "Current Form: " & GetFormDisplayName(m_strCurrentFormName) & vbCrLf
    strBody = strBody & "Date: " & Format(Date, "mm/dd/yyyy") & vbCrLf
    strBody = strBody & "Time: " & Format(Time, "hh:nn:ss AM/PM") & vbCrLf
    strBody = strBody & "Access Version: " & SysCmd(acSysCmdAccessVer) & vbCrLf & vbCrLf
    strBody = strBody & "Thank you for your assistance." & vbCrLf & vbCrLf
    strBody = strBody & "Best regards," & vbCrLf
    strBody = strBody & m_strCurrentUser
    
    BuildHelpEmailBody = strBody
End Function

Private Function BuildMailtoLink(strTo As String, strCC As String, strSubject As String, strBody As String) As String
    Dim strLink As String
    
    strLink = "mailto:" & strTo
    strLink = strLink & "?cc=" & strCC
    strLink = strLink & "&subject=" & UrlEncode(strSubject)
    strLink = strLink & "&body=" & UrlEncode(strBody)
    
    BuildMailtoLink = strLink
End Function

' ========================================
' UTILITY FUNCTIONS
' ========================================

Private Function GetCurrentUserName() As String
    On Error Resume Next
    
    ' Try to get current Windows user
    GetCurrentUserName = Environ("USERNAME")
    
    ' Fallback if no username found
    If Len(GetCurrentUserName) = 0 Then
        GetCurrentUserName = "Current User"
    End If
    
    On Error GoTo 0
End Function

Private Function FormExists(strFormName As String) As Boolean
    On Error Resume Next
    
    Dim obj As AccessObject
    
    For Each obj In CurrentProject.AllForms
        If obj.Name = strFormName Then
            FormExists = True
            Exit Function
        End If
    Next obj
    
    FormExists = False
    On Error GoTo 0
End Function

Private Function GetFormDisplayName(strFormName As String) As String
    Select Case strFormName
        Case FORM_DASHBOARD
            GetFormDisplayName = "Dashboard"
        Case FORM_REPORT_GENERATOR
            GetFormDisplayName = "Report Generator"
        Case FORM_ADMIN
            GetFormDisplayName = "Administration"
        Case FORM_USER_MANAGEMENT
            GetFormDisplayName = "User Management"
        Case Else
            GetFormDisplayName = strFormName
    End Select
End Function

Private Sub UpdateFormTitle(strCurrentSection As String)
    On Error Resume Next
    Me.Caption = "CMR Pulse - " & strCurrentSection & " (" & m_strCurrentUser & ")"
    On Error GoTo 0
End Sub

Private Function UrlEncode(strText As String) As String
    Dim strResult As String
    strResult = Replace(strText, " ", "%20")
    strResult = Replace(strResult, vbCrLf, "%0D%0A")
    strResult = Replace(strResult, vbCr, "%0D")
    strResult = Replace(strResult, vbLf, "%0A")
    strResult = Replace(strResult, "&", "%26")
    strResult = Replace(strResult, "?", "%3F")
    strResult = Replace(strResult, "=", "%3D")
    
    UrlEncode = strResult
End Function

' ========================================
' ERROR HANDLING
' ========================================

Private Sub HandleError(strProcedureName As String, strErrorDescription As String)
    Dim strErrorMessage As String
    
    strErrorMessage = "Error in " & strProcedureName & ": " & strErrorDescription
    Debug.Print strErrorMessage
    
    ' Optionally show user-friendly error message
    ' MsgBox strErrorMessage, vbCritical, "CMR Pulse Error"
End Sub

' ========================================
' PUBLIC TESTING AND DEBUGGING FUNCTIONS
' ========================================

Public Sub TestNavigationSystem()
    Debug.Print "=== Testing Navigation System ==="
    Debug.Print "Current User: " & GetCurrentUserName()
    Debug.Print "WebView Ready: " & m_bWebViewReady
    Debug.Print "Dashboard Form Exists: " & FormExists(FORM_DASHBOARD)
    Debug.Print "Report Generator Form Exists: " & FormExists(FORM_REPORT_GENERATOR)
    Debug.Print "Admin Form Exists: " & FormExists(FORM_ADMIN)
    Debug.Print "User Management Form Exists: " & FormExists(FORM_USER_MANAGEMENT)
    Debug.Print "Current Form: " & m_strCurrentFormName
    Debug.Print "=== Test Complete ==="
End Sub

Public Sub TestUserManagement()
    Debug.Print "=== Testing User Management ==="
    Call UpdateUserName("Test User")
    Debug.Print "User updated to: Test User"
    
    ' Wait a moment
    Application.Wait Now + TimeValue("0:00:02")
    
    Call UpdateUserName("John Smith")
    Debug.Print "User updated to: John Smith"
    Debug.Print "=== User Management Test Complete ==="
End Sub

Public Sub TestLogout()
    Debug.Print "Testing logout functionality..."
    Call ProcessLogout
End Sub

Public Sub ShowCurrentStatus()
    Debug.Print "=== Current Status ==="
    Debug.Print "Current Form: " & m_strCurrentFormName
    Debug.Print "Current User: " & m_strCurrentUser
    Debug.Print "WebView Ready: " & m_bWebViewReady
    Debug.Print "Form Caption: " & Me.Caption
    Debug.Print "=================="
End Sub

' ========================================
' FORM CLEANUP
' ========================================

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' Stop the timer
    Application.OnTime Now + TimeValue("00:00:01"), "CheckForClicks", , False
    
    ' Close any open child forms
    Call CloseCurrentChildForm
    
    Debug.Print "CMR Pulse Navigation System closed"
    
    On Error GoTo 0
End Sub

' ========================================
' SETUP INSTRUCTIONS AND NOTES
' ========================================

' COMPLETE SETUP CHECKLIST:
'
' 1. FORM SETUP:
'    ✓ Create main form (e.g., "frmMainNavigation")
'    ✓ Add WebView2 control named "EdgeBrowser0"
'    ✓ Add Subform control named "subform"
'    ✓ Copy this code to the form's class module
'
' IMPORTANT: USER ENTITLEMENT IMPLEMENTATION
' 
' This sidebar navigation system does NOT check user entitlements.
' All menu items are accessible to all users. Implement security in:
'
' A. Individual Form Load Events:
'    Private Sub Form_Load()
'        If Not UserHasAdminAccess() Then
'            MsgBox "Access Denied"
'            DoCmd.Close acForm, Me.Name
'        End If
'    End Sub
'
' B. Form-Level Security Functions:
'    Private Sub cmdDeleteUser_Click()
'        If Not UserCanDeleteUsers() Then
'            MsgBox "Insufficient permissions"
'            Exit Sub
'        End If
'        ' ... delete logic
'    End Sub
'
' C. Control Visibility/Enabling:
'    Private Sub Form_Load()
'        Me.cmdAdminPanel.Visible = UserIsAdmin()
'        Me.cmdUserManagement.Enabled = UserCanManageUsers()
'    End Sub
'
' 2. CHILD FORMS TO CREATE:
'    ✓ frmDashboard - Main dashboard
'    ✓ frmReportGenerator - Report generation
'    ✓ frmAdmin - Administrative functions
'    ✓ frmUserManagement - User management
'    ✓ frmLogin - Login form (optional)
'
' 3. HTML FILE SETUP:
'    ✓ Save the updated index.html file
'    ✓ Update HTML_FILE_PATH constant with correct path
'    ✓ Ensure file is accessible to Access
'
' 4. TESTING:
'    ✓ Run TestNavigationSystem() to verify setup
'    ✓ Run TestUserManagement() to test user functions
'    ✓ Run ShowCurrentStatus() to see current state
'
' 5. CUSTOMIZATION:
'    ✓ Update email addresses in constants
'    ✓ Modify form names to match your forms
'    ✓ Customize logout behavior
'    ✓ Add additional navigation items as needed
'
' 6. TROUBLESHOOTING:
'    ✓ Check Debug.Print output in Immediate window
'    ✓ Verify HTML file path is correct
'    ✓ Ensure WebView2 control is properly installed
'    ✓ Test individual functions first
'
' USAGE EXAMPLES:
'
' ' Update logged-in user from anywhere in your application:
' Forms("frmMainNavigation").UpdateUserName "John Smith"
'
' ' Test the navigation system:
' Forms("frmMainNavigation").TestNavigationSystem
'
' ' Check current status:
' Forms("frmMainNavigation").ShowCurrentStatus
'
