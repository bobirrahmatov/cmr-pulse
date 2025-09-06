' ========================================
' ENHANCED CMR PULSE NAVIGATION SYSTEM
' Professional VBA Code for MS Access Form Navigation
' ========================================
'
' Form Name: subform
' WebView2 Control: EdgeBrowser0
' Subform Control: subform (child form container)
'
' Author: Enhanced Navigation System
' Version: 2.0
' Last Updated: [Current Date]

Option Compare Database
Option Explicit

' ========================================
' MODULE-LEVEL CONSTANTS AND VARIABLES
' ========================================

' Form names constants
Private Const FORM_DASHBOARD As String = "frmDashboard"
Private Const FORM_REPORT_GENERATOR As String = "frmReportGenerator" 
Private Const FORM_ADMIN As String = "frmAdmin"
Private Const FORM_USER_MANAGEMENT As String = "frmUserManagement"

' Email configuration constants
Private Const SUPPORT_EMAIL As String = "support@company.com"
Private Const IT_HELPDESK_EMAIL As String = "it-helpdesk@company.com"
Private Const EMAIL_SUBJECT As String = "CMR Pulse - Help Request"

' Navigation element IDs (matching HTML)
Private Const NAV_DASHBOARD As String = "Dashboard"
Private Const NAV_REPORT_GENERATOR As String = "Report Generator"
Private Const NAV_ADMIN As String = "security"
Private Const NAV_USER_MANAGEMENT As String = "password"
Private Const NAV_HELP_SUPPORT As String = "help"
Private Const NAV_SIGNOUT As String = "signout"

' Current state tracking
Private m_strCurrentFormName As String
Private m_strCurrentUser As String

' ========================================
' FORM INITIALIZATION AND SETUP
' ========================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Initialize form variables
    Call InitializeFormVariables
    
    ' Set up the WebView2 control
    Call SetupWebViewControl
    
    ' Initialize with default form
    Call NavigateToFormByName(FORM_DASHBOARD)
    
    ' Log successful initialization
    Debug.Print "CMR Pulse Navigation System initialized successfully"
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("Form_Load", Err.Description)
End Sub

Private Sub InitializeFormVariables()
    ' Initialize tracking variables
    m_strCurrentFormName = ""
    m_strCurrentUser = GetCurrentUserName()
    
    ' Set form caption
    Me.Caption = "CMR Pulse - " & m_strCurrentUser
End Sub

Private Sub SetupWebViewControl()
    On Error GoTo ErrorHandler
    
    ' Navigate to the HTML sidebar file
    ' TODO: Update this path to your actual HTML file location
    Dim strHtmlFilePath As String
    strHtmlFilePath = "file:///C:/CMRPulse/sidebar/index.html"
    
    Me.EdgeBrowser0.Navigate strHtmlFilePath
    
    ' Set up click detection JavaScript
    Call InjectClickDetectionScript
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error setting up WebView control: " & Err.Description
End Sub

Private Sub InjectClickDetectionScript()
    On Error Resume Next
    
    ' JavaScript to detect clicks and store element IDs
    Dim strScript As String
    strScript = "var script = document.createElement('script');" & _
                "Me.EdgeBrowser0.ExecuteJavascript script" & _
                "script.innerHTML = " & Chr(10) & _
                """function interceptClickEvent(e) {" & Chr(10) & _
                "    var elementId;" & Chr(10) & _
                "    var target = e.target || e.srcElement;" & Chr(10) & _
                "        elementId = e.target.getAttribute('id');" & Chr(10) & _
                "        console.log('e.target=>' + e.target);" & Chr(10) & _
                "        console.log('e.target.id=>' + e.target.id);" & Chr(10) & _
                "        console.log('e.target.getAttribute(id)=>' + e.target.getAttribute('id'));" & Chr(10) & _
                "        document.getElementById('clickedElement').value = elementId;" & Chr(10) & _
                "    e.preventDefault();" & Chr(10) & _
                "}"";" & Chr(10) & _
                "Me.EdgeBrowser0.ExecuteJavascript script"
    
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    ' Create hidden input for storing clicked element
    strScript = "document.body.appendChild(script);"
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    strScript = "var hiddenInput = document.createElement('input');"
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    strScript = "hiddenInput.type = 'hidden';"
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    strScript = "hiddenInput.id = 'clickedElement';"
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    strScript = "document.body.appendChild(hiddenInput);"
    Me.EdgeBrowser0.ExecuteJavascript strScript
    
    strScript = "document.addEventListener('click', interceptClickEvent);"
    Me.EdgeBrowser0.ExecuteJavascript strScript
End Sub

' ========================================
' WEBVIEW2 EVENT HANDLERS
' ========================================

Private Sub EdgeBrowser0_DocumentComplete(URL As Variant)
    Debug.Print "WebView2 document loaded: " & URL
    
    ' Re-inject click detection script after page loads
    Call InjectClickDetectionScript
End Sub

Private Sub EdgeBrowser0_Click()
    On Error GoTo ErrorHandler
    
    ' Get the clicked element ID from JavaScript
    Dim strClickedElementId As String
    strClickedElementId = Me.EdgeBrowser0.RetrieveJavascriptValue("document.getElementById('clickedElement').value")
    
    ' Process the navigation request
    If Len(strClickedElementId) > 0 Then
        Call ProcessNavigationRequest(strClickedElementId)
    End If
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("EdgeBrowser0_Click", Err.Description)
End Sub

' ========================================
' NAVIGATION PROCESSING ENGINE
' ========================================

Private Sub ProcessNavigationRequest(strElementId As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "Processing navigation request: " & strElementId
    
    ' Route to appropriate handler based on element ID
    Select Case strElementId
        Case NAV_DASHBOARD
            Call NavigateToFormByName(FORM_DASHBOARD)
            
        Case NAV_REPORT_GENERATOR
            Call NavigateToFormByName(FORM_REPORT_GENERATOR)
            
        Case NAV_ADMIN
            Call NavigateToFormByName(FORM_ADMIN)
            
        Case NAV_USER_MANAGEMENT
            Call NavigateToFormByName(FORM_USER_MANAGEMENT)
            
        Case NAV_HELP_SUPPORT
            Call OpenHelpSupportEmail
            
        Case NAV_SIGNOUT
            Call ProcessUserSignout
            
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
    
    Debug.Print "Navigating to form: " & strFormName
    
    ' Close current form if different
    If m_strCurrentFormName <> strFormName Then
        Call CloseCurrentChildForm
    End If
    
    ' Validate form exists
    If Not FormExists(strFormName) Then
        Debug.Print "Form does not exist: " & strFormName
        Exit Sub
    End If
    
    ' Load the new form
    Call LoadFormInSubformControl(strFormName)
    
    ' Update tracking
    m_strCurrentFormName = strFormName
    
    ' Update form title
    Call UpdateFormTitle(GetFormDisplayName(strFormName))
    
    Debug.Print "Successfully navigated to: " & strFormName
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("NavigateToFormByName", Err.Description)
End Sub

Private Sub LoadFormInSubformControl(strFormName As String)
    On Error GoTo ErrorHandler
    
    ' Set the source object for the subform control
    Me.subform.SourceObject = strFormName
    Me.subform.Visible = True
    
    ' Ensure proper sizing and display
    Call OptimizeSubformDisplay
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("LoadFormInSubformControl", Err.Description)
End Sub

Private Sub CloseCurrentChildForm()
    On Error Resume Next
    
    If Len(m_strCurrentFormName) > 0 Then
        Debug.Print "Closing current form: " & m_strCurrentFormName
        
        ' Clear the subform control
        Me.subform.SourceObject = ""
        Me.subform.Visible = False
        
        ' Clear tracking variable
        m_strCurrentFormName = ""
    End If
    
    On Error GoTo 0
End Sub

Private Sub OptimizeSubformDisplay()
    On Error Resume Next
    
    ' Optimize subform display properties
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
    
    ' Build email components
    Dim strEmailBody As String
    strEmailBody = BuildHelpEmailBody()
    
    Dim strMailtoLink As String
    strMailtoLink = BuildMailtoLink(SUPPORT_EMAIL, IT_HELPDESK_EMAIL, EMAIL_SUBJECT, strEmailBody)
    
    ' Open email client
    Call OpenEmailClient(strMailtoLink)
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("OpenHelpSupportEmail", Err.Description)
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

Private Sub OpenEmailClient(strMailtoLink As String)
    On Error GoTo ErrorHandler
    
    ' Try to open with default email client
    Application.FollowHyperlink strMailtoLink
    
    Debug.Print "Help email opened successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error opening email client: " & Err.Description
    
    ' Fallback: Copy email details to clipboard
    Call CopyEmailDetailsToClipboard(strMailtoLink)
End Sub

Private Sub CopyEmailDetailsToClipboard(strMailtoLink As String)
    On Error Resume Next
    
    ' Extract email details and copy to clipboard for manual use
    Dim strClipboardText As String
    strClipboardText = "Email Details (copy to your email client):" & vbCrLf & vbCrLf
    strClipboardText = strClipboardText & "To: " & SUPPORT_EMAIL & vbCrLf
    strClipboardText = strClipboardText & "CC: " & IT_HELPDESK_EMAIL & vbCrLf
    strClipboardText = strClipboardText & "Subject: " & EMAIL_SUBJECT & vbCrLf & vbCrLf
    strClipboardText = strClipboardText & BuildHelpEmailBody()
    
    ' Copy to clipboard (requires additional clipboard handling code)
    Debug.Print "Email details prepared for clipboard"
    
    On Error GoTo 0
End Sub

' ========================================
' USER AUTHENTICATION AND SESSION
' ========================================

Private Sub ProcessUserSignout()
    On Error GoTo ErrorHandler
    
    Debug.Print "Processing user signout"
    
    ' Close current form
    Call CloseCurrentChildForm
    
    ' Clear user session data
    Call ClearUserSession
    
    ' Log the signout
    Debug.Print "User signed out: " & m_strCurrentUser
    
    ' Close application or redirect to login
    ' Option 1: Close application
    DoCmd.Quit acQuitSaveNone
    
    ' Option 2: Redirect to login form (uncomment if needed)
    ' DoCmd.OpenForm "frmLogin"
    ' DoCmd.Close acForm, Me.Name
    
    Exit Sub
    
ErrorHandler:
    Call HandleError("ProcessUserSignout", Err.Description)
End Sub

Private Sub ClearUserSession()
    ' Clear session variables and temporary data
    m_strCurrentFormName = ""
    m_strCurrentUser = ""
    
    ' Additional session cleanup can be added here
End Sub

' ========================================
' UTILITY FUNCTIONS
' ========================================

Private Function GetCurrentUserName() As String
    On Error Resume Next
    
    ' Try to get current Windows user
    GetCurrentUserName = Environ("USERNAME")
    
    ' Fallback to Access user if available
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
    ' Basic URL encoding for email body
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
' ERROR HANDLING SYSTEM
' ========================================

Private Sub HandleError(strProcedureName As String, strErrorDescription As String)
    Dim strErrorMessage As String
    
    strErrorMessage = "Error in " & strProcedureName & ": " & strErrorDescription
    Debug.Print strErrorMessage
    
    ' Log to error table if exists
    Call LogErrorToTable(strProcedureName, strErrorDescription)
End Sub

Private Sub LogErrorToTable(strProcedure As String, strError As String)
    On Error Resume Next
    
    ' Log error to table if error logging table exists
    ' This is optional - create tblErrorLog if you want error logging
    '
    ' DoCmd.RunSQL "INSERT INTO tblErrorLog (ErrorDate, ErrorTime, UserName, " & _
    '              "Procedure, ErrorDescription) VALUES ('" & Date & "', '" & Time & "', '" & _
    '              m_strCurrentUser & "', '" & strProcedure & "', '" & strError & "')"
    
    On Error GoTo 0
End Sub

' ========================================
' DEBUGGING AND TESTING FUNCTIONS
' ========================================

Public Sub TestNavigationSystem()
    ' Test function to verify navigation system
    Debug.Print "=== Testing Navigation System ==="
    Debug.Print "Current User: " & GetCurrentUserName()
    Debug.Print "Dashboard Form Exists: " & FormExists(FORM_DASHBOARD)
    Debug.Print "Report Generator Form Exists: " & FormExists(FORM_REPORT_GENERATOR)
    Debug.Print "Admin Form Exists: " & FormExists(FORM_ADMIN)
    Debug.Print "User Management Form Exists: " & FormExists(FORM_USER_MANAGEMENT)
    Debug.Print "=== Test Complete ==="
End Sub

Public Sub TestFormNavigation()
    ' Test form navigation manually
    Call NavigateToFormByName(FORM_DASHBOARD)
    Debug.Print "Navigated to Dashboard"
End Sub

Public Sub TestHelpEmail()
    ' Test help email functionality
    Call OpenHelpSupportEmail
End Sub

Public Sub ShowCurrentStatus()
    Debug.Print "=== Current Status ==="
    Debug.Print "Current Form: " & m_strCurrentFormName
    Debug.Print "Current User: " & m_strCurrentUser
    Debug.Print "Form Caption: " & Me.Caption
    Debug.Print "=================="
End Sub

' ========================================
' SETUP AND CONFIGURATION NOTES
' ========================================

' SETUP CHECKLIST:
' 
' 1. Form Setup:
'    ✓ Main form name: "subform"
'    ✓ WebView2 control name: "EdgeBrowser0"
'    ✓ Subform control name: "subform"
'
' 2. Child Forms to Create:
'    ✓ frmDashboard - Main dashboard
'    ✓ frmReportGenerator - Report generation tools
'    ✓ frmAdmin - Administrative functions
'    ✓ frmUserManagement - User management
'
' 3. HTML File:
'    ✓ Update strHtmlFilePath in SetupWebViewControl()
'    ✓ Ensure HTML file is accessible
'
' 4. Email Configuration:
'    ✓ Update SUPPORT_EMAIL constant
'    ✓ Update IT_HELPDESK_EMAIL constant
'    ✓ Customize EMAIL_SUBJECT if needed
'
' 5. Optional Error Logging:
'    ✓ Create tblErrorLog table if you want error logging
'    ✓ Uncomment LogErrorToTable code
'
' 6. Testing:
'    ✓ Run TestNavigationSystem() to verify setup
'    ✓ Run TestFormNavigation() to test navigation
'    ✓ Run TestHelpEmail() to test email functionality
