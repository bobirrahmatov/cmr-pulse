' ========================================
' MS ACCESS VBA CODE FOR SIDEBAR NAVIGATION
' HTML Sidebar + Native Access Forms
' ========================================
' 
' SCENARIO: 
' - Left side: HTML sidebar in Edge WebView2 control
' - Right side: Native MS Access form content area
' 
' Place this code in your main form's VBA module

' ========================================
' 1. MAIN NAVIGATION FUNCTION
' ========================================
Public Sub NavigateToForm(FormName As String)
    On Error GoTo ErrorHandler
    
    ' Close current child form if any
    CloseCurrentChildForm
    
    ' Navigate to the requested form
    Select Case UCase(FormName)
        Case "DASHBOARD"
            OpenFormInMainArea "frmDashboard"
            
        Case "REPORTGENERATOR"
            OpenFormInMainArea "frmReportGenerator"
            
        Case "ADMIN"
            OpenFormInMainArea "frmAdmin"
            
        Case "USERS"
            OpenFormInMainArea "frmUsers"
            
        Case Else
            MsgBox "Form '" & FormName & "' not found or not implemented.", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error navigating to form: " & Err.Description, vbCritical
End Sub

' ========================================
' 2. FORM MANAGEMENT FUNCTIONS
' ========================================

' Open form in the main content area
Private Sub OpenFormInMainArea(ChildFormName As String)
    On Error GoTo ErrorHandler
    
    ' Method 1: Using Subform Control (RECOMMENDED)
    ' Replace "SubformControl" with your actual subform control name
    Me.SubformControl.SourceObject = ChildFormName
    Me.SubformControl.Visible = True
    
    ' Store current form name for tracking
    Me.Tag = ChildFormName
    
    ' Method 2: Alternative - Open as popup positioned in main area
    ' Uncomment if you prefer this method:
    '
    ' DoCmd.OpenForm ChildFormName, acNormal
    ' 
    ' ' Position the form in the main content area
    ' Dim ChildForm As Form
    ' Set ChildForm = Forms(ChildFormName)
    ' 
    ' ' Set position and size to match your main content area
    ' ChildForm.Move Me.Left + 3000, Me.Top + 1000, 8000, 6000  ' Adjust these values
    ' ChildForm.BorderStyle = 0  ' Remove border
    ' ChildForm.RecordSelectors = False
    ' ChildForm.NavigationButtons = False
    ' 
    ' Me.Tag = ChildFormName  ' Track current form
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening form '" & ChildFormName & "': " & Err.Description, vbCritical
End Sub

' Close current child form
Private Sub CloseCurrentChildForm()
    On Error Resume Next
    
    ' Method 1: Clear subform control
    If Not IsNull(Me.Tag) And Me.Tag <> "" Then
        Me.SubformControl.SourceObject = ""
        Me.SubformControl.Visible = False
    End If
    
    ' Method 2: Close separate form if using popup method
    ' If Not IsNull(Me.Tag) And Me.Tag <> "" Then
    '     If IsFormOpen(Me.Tag) Then
    '         DoCmd.Close acForm, Me.Tag
    '     End If
    ' End If
    
    Me.Tag = ""  ' Clear tracking
    
    On Error GoTo 0
End Sub

' Check if a form is currently open
Private Function IsFormOpen(FormName As String) As Boolean
    On Error Resume Next
    IsFormOpen = (Forms(FormName).Name = FormName)
    On Error GoTo 0
End Function

' ========================================
' 3. WEBVIEW2 INTEGRATION
' ========================================

' Call this from JavaScript in your HTML sidebar
Public Sub CallFromWebView(Action As String, Parameter As String)
    On Error GoTo ErrorHandler
    
    Select Case UCase(Action)
        Case "NAVIGATE"
            NavigateToForm Parameter
        Case "LOGOUT"
            HandleLogout
        Case Else
            MsgBox "Action '" & Action & "' not recognized.", vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing web view call: " & Err.Description, vbCritical
End Sub

' Handle logout functionality
Private Sub HandleLogout()
    If MsgBox("Are you sure you want to log out?", vbYesNo + vbQuestion) = vbYes Then
        ' Close all child forms
        CloseCurrentChildForm
        
        ' Add your logout logic here
        ' Examples:
        ' - Close the application: DoCmd.Quit
        ' - Open login form: DoCmd.OpenForm "frmLogin"
        ' - Clear user session variables
        
        MsgBox "Logging out...", vbInformation
        ' DoCmd.OpenForm "frmLogin"
        ' DoCmd.Close acForm, Me.Name
    End If
End Sub

' ========================================
' 4. FORM INITIALIZATION
' ========================================

' Add this to your main form's Form_Load event
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Initialize the HTML sidebar
    ' Replace with your actual HTML file path
    Me.WebBrowserSidebar.Navigate "file:///C:/YourPath/index.html"
    
    ' Wait a moment for the web control to load
    DoEvents
    
    ' Initialize with Dashboard form
    Call NavigateToForm("Dashboard")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbCritical
End Sub

' ========================================
' 5. WEBVIEW2 EVENT HANDLERS
' ========================================

' When the HTML page is fully loaded
Private Sub WebBrowserSidebar_DocumentComplete()
    ' Optional: Add any initialization code here
    ' The sidebar is now ready for interaction
End Sub

' Handle web view navigation events
Private Sub WebBrowserSidebar_NavigationCompleted()
    ' Optional: Handle when navigation is complete
End Sub

' ========================================
' 6. ALTERNATIVE: DIRECT VBA CALLS FROM HTML
' ========================================

' If your WebView2 control supports it, you can expose VBA functions directly
' Add this to make VBA functions callable from JavaScript:

Public Sub ExposeVBAToWebView()
    ' This depends on your WebView2 control implementation
    ' Some controls allow you to expose VBA objects to JavaScript
    
    ' Example pseudocode (syntax varies by control):
    ' Me.WebBrowserSidebar.ObjectForScripting = Me
End Sub

' ========================================
' 7. FORM LAYOUT RECOMMENDATIONS
' ========================================

' RECOMMENDED FORM LAYOUT:
' 
' +------------------+------------------------+
' |                  |                        |
' |   HTML Sidebar   |    Main Content Area   |
' |   (WebView2)     |    (Subform Control)   |
' |                  |                        |
' |   - Dashboard    |   [Current Form Here]  |
' |   - Reports      |                        |
' |   - Settings     |                        |
' |   - Links        |                        |
' |   - Help         |                        |
' |   - Logout       |                        |
' |                  |                        |
' +------------------+------------------------+
' 
' CONTROL SETUP:
' 1. WebView2 Control (left side): Name = "WebBrowserSidebar"
' 2. Subform Control (right side): Name = "SubformControl"
' 3. Set subform control properties:
'    - SourceObject: (empty initially)
'    - BorderStyle: Transparent or None
'    - CanGrow: Yes
'    - CanShrink: Yes

' ========================================
' 8. EXAMPLE CHILD FORM SETUP
' ========================================

' For each child form (frmDashboard, frmReportGenerator, etc.):
' 
' Form Properties:
' - Default View: Single Form or Continuous Forms
' - Allow Additions: No (unless needed)
' - Allow Deletions: No (unless needed)
' - Allow Edits: Yes (as needed)
' - Record Selectors: No
' - Navigation Buttons: No
' - Dividing Lines: No
' - Border Style: None or Dialog
' 
' This ensures child forms integrate seamlessly into the main form

' ========================================
' 9. DEBUGGING HELPERS
' ========================================

' Add these for troubleshooting
Public Sub TestNavigation()
    ' Test navigation without HTML sidebar
    NavigateToForm "Dashboard"
End Sub

Public Sub ShowCurrentForm()
    MsgBox "Current form: " & Me.Tag, vbInformation
End Sub

Public Sub ClearCurrentForm()
    CloseCurrentChildForm
End Sub

' ========================================
' 10. SETUP CHECKLIST
' ========================================

' SETUP CHECKLIST:
' 
' □ 1. Create main form with WebView2 control (left) and Subform control (right)
' □ 2. Name WebView2 control "WebBrowserSidebar"
' □ 3. Name Subform control "SubformControl"
' □ 4. Create child forms: frmDashboard, frmReportGenerator, frmAdmin, frmUsers
' □ 5. Copy this VBA code to main form's module
' □ 6. Update HTML file path in Form_Load event
' □ 7. Update child form names if different
' □ 8. Test navigation functionality
' □ 9. Configure child form properties for seamless integration
' □ 10. Add your specific business logic to each child form

' TROUBLESHOOTING:
' - If forms don't load: Check form names in Select Case statement
' - If WebView2 doesn't show: Check HTML file path and permissions
' - If navigation doesn't work: Verify control names match VBA code
' - If forms look wrong: Adjust child form properties (borders, selectors, etc.)
