VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mint"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   3000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMessages 
      Interval        =   100
      Left            =   600
      Top             =   4320
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   10000
      Left            =   120
      Top             =   4320
   End
   Begin VB.CommandButton cmdMenu 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4845
      Width           =   2970
   End
   Begin MSComctlLib.ImageList imlNotifyList 
      Left            =   1560
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1194
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A6E
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC0
            Key             =   "Ignored"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":279A
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3074
            Key             =   "Online Users"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34C6
            Key             =   "Offline Users"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3918
            Key             =   "All users"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D6A
            Key             =   "My Users"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41BC
            Key             =   "Offline"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A96
            Key             =   "Offline Notify"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5370
            Key             =   "Offline Ignored"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwNotifyList 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlNotifyList"
      Appearance      =   1
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextSend 
         Caption         =   "&Send Message"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuContextSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextNotify 
         Caption         =   "Notify when online/offline"
      End
      Begin VB.Menu mnuContextIgnore 
         Caption         =   "&Ignore"
      End
      Begin VB.Menu mnuContextDelete 
         Caption         =   "&Delete from list"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMenuBrowseUsers 
         Caption         =   "&Browse Online Users"
      End
      Begin VB.Menu mnuMenuAddUser 
         Caption         =   "&Add User by Login"
      End
      Begin VB.Menu mnuMenuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuExpandAll 
         Caption         =   "&Expand All"
      End
      Begin VB.Menu mnuMenuCollapseAll 
         Caption         =   "&Collapse All"
      End
      Begin VB.Menu mnuMenuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuQuiet 
         Caption         =   "&Quiet Mode"
      End
      Begin VB.Menu mnuMenuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuOptionsNext 
         Caption         =   "&Read Next"
      End
      Begin VB.Menu mnuOptionsDelete 
         Caption         =   "&Delete all"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NODE_ONLINE = 1
Const NODE_OFFLINE = 2
Const NODE_ONLINE_TEXT = "Online"
Const NODE_OFFLINE_TEXT = "Offline"

Private strDomain As String
Private colPeople As New Collection
Private colConfig As New Collection
Private strHome As String
Private WithEvents sysTray As frmSysTray
Attribute sysTray.VB_VarHelpID = -1

Private Sub cmdMenu_Click()
    frmMain.PopupMenu mnuMenu
End Sub

Private Sub Form_Paint()
    'Make the form always-on-top
    SetWindowPos frmMain.hwnd, -1, frmMain.Left / Screen.TwipsPerPixelX, frmMain.Top / Screen.TwipsPerPixelY, frmMain.Width / Screen.TwipsPerPixelX, frmMain.Height / Screen.TwipsPerPixelY, &H10 Or &H40
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        frmMain.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    If frmMain.Width < 110 Then frmMain.Width = 110
    If frmMain.Height < 770 Then frmMain.Height = 770
    tvwNotifyList.Height = frmMain.Height - 765
    tvwNotifyList.Width = frmMain.Width - 105
    cmdMenu.Width = frmMain.Width - 105
    cmdMenu.Top = tvwNotifyList.Height + 30
End Sub

Private Sub mnuContextDelete_Click()
    If tvwNotifyList.SelectedItem.Parent.Key = NODE_ONLINE_TEXT Or tvwNotifyList.SelectedItem.Parent.Key = NODE_OFFLINE_TEXT Then
        If MsgBox("Are you sure you want to delete this user from your list?", vbYesNo, "Delete User") = vbYes Then
            colPeople.Remove tvwNotifyList.SelectedItem.Tag
            tvwNotifyList.Nodes.Remove tvwNotifyList.SelectedItem.Index
        End If
    End If
End Sub

Private Sub Form_Load()
    'Make invisible
    frmMain.Visible = False
    
    'Position the window
    frmMain.Left = Screen.Width - frmMain.Width
    frmMain.Height = Screen.Height - 1000
    cmdMenu.Top = frmMain.Height - 2 * cmdMenu.Height
    tvwNotifyList.Height = frmMain.Height - 800
    
    'Initialise the tree-view
    tvwNotifyList.Nodes.Add , tvwLast, NODE_ONLINE_TEXT, NODE_ONLINE_TEXT, "Online Users"
    tvwNotifyList.Nodes.Add , tvwLast, NODE_OFFLINE_TEXT, NODE_OFFLINE_TEXT, "Offline Users"
    tvwNotifyList.Nodes(NODE_OFFLINE_TEXT).EnsureVisible
    
    'Get domain and PDC
    fnGetDomainName strDomain
    fnGetPDCName "", strDomain, strPDC
    
    'Get the users home path
    strHome = getUserHome
    
    'Read in the saved usernames
    readConfig
    
    'Update the display
    addUsers
    
    'Set up the system tray
    Set sysTray = New frmSysTray
    sysTray.ToolTip = "Mint - Online"
    sysTray.AddMenuItem "&Open Mint", "open", True
    sysTray.AddMenuItem "-", "", False
    sysTray.AddMenuItem "E&xit", "exit", False
    sysTray.IconHandle = imlNotifyList.ListImages("Online").Picture

    'Update the display
    updateUsers
End Sub

Public Sub addUser(strUsername As String)
    If userExists(strPDC, strUsername) Then
        If notOnList(strUsername) Then
            addToNotify (strUsername)
            colPeople.Add strUsername, strUsername
        End If
    Else
        MsgBox "User does not exist!", vbExclamation + vbOKOnly, "Add a user"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    writeConfig
    Unload sysTray
    Set sysTray = Nothing
End Sub

Private Sub mnuContextIgnore_Click()
    'Reverse the check status
    mnuContextIgnore.Checked = Not mnuContextIgnore.Checked
    'If they are being ignored
    If mnuContextIgnore.Checked Then
        'And they are online
        If tvwNotifyList.SelectedItem.Parent.Text = "Online" Then
            'Then ignore them
            tvwNotifyList.SelectedItem.Image = "Ignored"
        Else
            'Else offline ignore them
            tvwNotifyList.SelectedItem.Image = "Offline Ignored"
        End If
        'Don't notify
        mnuContextNotify.Checked = False
        
        'Change their people entry
        colPeople.Remove tvwNotifyList.SelectedItem.Tag
        colPeople.Add tvwNotifyList.SelectedItem.Tag & "/Ignored", tvwNotifyList.SelectedItem.Tag
    Else
        'Use standard icon
        tvwNotifyList.SelectedItem.Image = tvwNotifyList.SelectedItem.Parent.Text
    
        colPeople.Remove tvwNotifyList.SelectedItem.Tag
        colPeople.Add tvwNotifyList.SelectedItem.Tag, tvwNotifyList.SelectedItem.Tag
    End If
End Sub

Private Sub mnuContextNotify_Click()
    'See above for comments
    mnuContextNotify.Checked = Not mnuContextNotify.Checked
    If mnuContextNotify.Checked Then
        If tvwNotifyList.SelectedItem.Parent.Text = "Online" Then
            tvwNotifyList.SelectedItem.Image = "Notify"
        Else
            tvwNotifyList.SelectedItem.Image = "Offline Notify"
        End If
        mnuContextIgnore.Checked = False
    
        'Change their people entry
        colPeople.Remove tvwNotifyList.SelectedItem.Tag
        colPeople.Add tvwNotifyList.SelectedItem.Tag & "/Notify", tvwNotifyList.SelectedItem.Tag
    Else
        'Use standard icon
        tvwNotifyList.SelectedItem.Image = tvwNotifyList.SelectedItem.Parent.Text
    
        colPeople.Remove tvwNotifyList.SelectedItem.Tag
        colPeople.Add tvwNotifyList.SelectedItem.Tag, tvwNotifyList.SelectedItem.Tag
    End If
End Sub

Private Sub mnuContextSend_Click()
    Dim frmMsg As New frmMessage
    
    frmMsg.NewMessage (tvwNotifyList.SelectedItem.Tag)
End Sub

Private Sub mnuMenuAddUser_Click()
    strAddUser = InputBox("Enter the username to add:", "Add a user")
    If Trim(strAddUser) <> "" Then
        addUser (strAddUser)
    End If
End Sub

Private Sub mnuMenuBrowseUsers_Click()
    Load frmOnlineUsers
End Sub

Private Sub mnuMenuCollapseAll_Click()
    Dim a As Long

    For a = 3 To tvwNotifyList.Nodes.count
        tvwNotifyList.Nodes(a).Expanded = False
    Next a
End Sub

Private Sub mnuMenuExit_Click()
    Unload frmMain
End Sub

Private Sub mnuMenuExpandAll_Click()
    Dim a As Long
    
    For a = 1 To tvwNotifyList.Nodes.count
        tvwNotifyList.Nodes(a).Expanded = True
    Next a
End Sub

Private Sub mnuMenuQuiet_Click()
    mnuMenuQuiet.Checked = Not mnuMenuQuiet.Checked
End Sub

Private Sub mnuOptionsDelete_Click()
If MsgBox("This will delete all incoming messages - are you sure?", vbYesNo + vbExclamation, "Confirm") = vbYes Then
    frmMessage.deleteAll
End If
End Sub

Private Sub mnuOptionsNext_Click()
    frmMessage.readNextMessage
End Sub

Private Sub sysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
    Select Case sKey
    Case "open"
        frmMain.Visible = True
    Case "exit"
        Unload frmMain
    End Select
End Sub

Private Sub sysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    frmMain.Visible = True
End Sub

Private Sub sysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If eButton = vbRightButton Then
        sysTray.ShowMenu
    End If
End Sub

Private Sub tmrMessages_Timer()
    On Error Resume Next
    
    Dim strMsg As String
    Dim frmMsg As frmMessage
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim strTo As String
    Dim strFrom As String
    Dim strMessage As String
    Dim a As Long
    Dim blnIgnore As Boolean
    
    strMsg = getMessage
    If strMsg <> "" Then
        strMessage = Right(strMsg, Len(strMsg) - InStr(1, strMsg, vbCrLf & vbCrLf) - 3)
            
        'If they announce their origin
        If Left(strMessage, 5) = "From:" Then
            strFrom = Mid(strMessage, 6, InStr(1, strMessage, vbCrLf) - 6)
            strMessage = Right(strMessage, Len(strMessage) - InStr(1, strMessage, vbCrLf) - 1)
        Else
            intPos1 = InStr(1, strMsg, "from") + 5
            intPos2 = InStr(1, strMsg, "to") - 1
            strFrom = Mid(strMsg, intPos1, intPos2 - intPos1)
        End If
        
        intPos1 = InStr(1, strMsg, "to") + 3
        intPos2 = InStr(1, strMsg, "on") - 1
        strTo = Mid(strMsg, intPos1, intPos2 - intPos1)
                
        For a = 1 To tvwNotifyList.Nodes.count
            If UCase(tvwNotifyList.Nodes(a).Tag) = UCase(strFrom) And tvwNotifyList.Nodes(a).Image = "Ignored" Then
                blnIgnore = True
            End If
        Next a
        If LCase(Left(strMessage, 4)) = "ping" Then
            netSend.NetSendMessage strFrom, "-=Reply to ping:" & Str(App.Major) & "." & Str(App.Minor) & "." & Str(App.Revision) & "=-"
        ElseIf blnIgnore = False And mnuMenuQuiet.Checked = False Then
            frmMessage.RecievedMessage strTo, strFrom, strMessage
        End If
    End If
End Sub

Private Sub tmrUpdate_Timer()
    updateUsers
End Sub

Private Sub tvwNotifyList_DblClick()
    Dim frmMsg As New frmMessage
    
    If tvwNotifyList.SelectedItem.Index > NODE_OFFLINE And InStr(1, LCase(tvwNotifyList.SelectedItem.FullPath), "offline") = 0 Then
        frmMsg.NewMessage tvwNotifyList.SelectedItem.Tag
    End If
End Sub

Private Function getUserHome() As String
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    Dim userInfo As USERINFO_2_API
    Dim strName As String
    Dim a As Integer
    
    'set variables
    baServerName = strPDC & Chr$(0)
    baUserName = localUserName & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 2, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
      getUserHome = "h:\"
      Exit Function
    End If

    'Turn the pointer into a variable
    CopyMem userInfo, ByVal lngptrUserInfo, Len(userInfo)
    
    getUserHome = PointerToStringW(userInfo.usri2_home_dir) & "\"
    NetAPIBufferFree lngptrUserInfo
End Function

Private Sub readConfig()
    On Error Resume Next
    
    Dim ff As Byte
    Dim strCurrentEntry As String
    Dim strIndex As String
    
    ff = FreeFile
    Open strHome & "Mint.ini" For Input As #ff
    If Err.Number <> 53 Then
        While Not EOF(ff)
            Line Input #ff, strCurrentEntry
            If Trim(strCurrentEntry) <> "" Then
                If InStr(1, strCurrentEntry, "/") <> 0 Then strIndex = Left(strCurrentEntry, InStr(1, strCurrentEntry, "/") - 1) Else strIndex = strCurrentEntry
                colPeople.Add strCurrentEntry, strIndex
            End If
        Wend
    Else
        Err.Clear
    End If
    Close #ff
    
    Open strHome & "Mint.cfg" For Input As #ff
    If Err.Number <> 53 Then
        While Not EOF(ff)
            Line Input #ff, strCurrentEntry
            If InStr(1, strCurrentEntry, "=") <> 0 Then
                colConfig.Add Right(strCurrentEntry, Len(strCurrentEntry) - InStr(1, strCurrentEntry, "=")), Left(strCurrentEntry, InStr(1, strCurrentEntry, "=") - 1)
            End If
        Wend
    Else
        Err.Clear
    End If
    Close #ff
    
    frmMain.Left = colConfig("Left")
    frmMain.Top = colConfig("Top")
    frmMain.Width = colConfig("Width")
    frmMain.Height = colConfig("Height")
End Sub

Private Sub writeConfig()
    Dim ff As Byte
    Dim strCurrentPerson As Variant
    
    ff = FreeFile
    Open strHome & "Mint.ini" For Output As #ff
    For Each strCurrentPerson In colPeople
        Print #ff, strCurrentPerson
    Next
    Close #ff

    Open strHome & "Mint.cfg" For Output As #ff
    Print #ff, "Top=" & frmMain.Top
    Print #ff, "Left=" & frmMain.Left
    Print #ff, "Width=" & frmMain.Width
    Print #ff, "Height=" & frmMain.Height
    Close #ff
End Sub

Private Sub addUsers()
    Dim strCurrentUser As Variant
    Dim lngNodeIndex As Long
    Dim strType As String 'Notify/Ignore?
    
    'Add users as appropriate
    For Each strCurrentUser In colPeople
        If InStr(1, strCurrentUser, "/") <> 0 Then
            strType = Right(strCurrentUser, Len(strCurrentUser) - InStr(1, strCurrentUser, "/"))
            addToNotify CStr(Left(strCurrentUser, Len(strCurrentUser) - Len(strType) - 1)), strType
        Else
            addToNotify CStr(strCurrentUser)
        End If
    Next
End Sub

Private Sub updateUsers()
    Dim count As Integer

    'For each online user
    For count = tvwNotifyList.Nodes.count To 3 Step -1
        'If they ARE a user
        If tvwNotifyList.Nodes(count).Tag <> "" Then
            'And they were online
            If tvwNotifyList.Nodes(count).Parent.Key = "Online" Then
                'And they are now offline
                If SessionEnum(strPDC, "", tvwNotifyList.Nodes(count).Tag) <> 0 Then
                    'Add them to the offline list
                    tvwNotifyList.Nodes.Add NODE_OFFLINE, tvwChild, tvwNotifyList.Nodes(count).Tag, getRealName(strPDC, tvwNotifyList.Nodes(count).Tag) & " (" & tvwNotifyList.Nodes(count).Tag & ")", tvwNotifyList.Nodes(count).Image
                    tvwNotifyList.Nodes(tvwNotifyList.Nodes.count).Tag = tvwNotifyList.Nodes(count).Tag
                    'And remove them from the online list
                    tvwNotifyList.Nodes.Remove count
                    
                    'If notification is on
                    If InStr(1, tvwNotifyList.Nodes(count).Image, "Notify") <> 0 Then
                        'Doevents to update the display
                        DoEvents
                        Beep
                        'Notify them! (Duh!)
                        MsgBox getRealName(strPDC, tvwNotifyList.Nodes(count).Tag) & " (" & tvwNotifyList.Nodes(count).Tag & ") is now offline!", vbOKOnly + vbInformation, "User is offline"
                    End If
                'If they are still online
                Else
                    'Ensure their icon is appropriate
                    Select Case tvwNotifyList.Nodes(count).Image
                    Case "Offline"
                        tvwNotifyList.Nodes(count).Image = "Online"
                    Case "Offline Notify"
                        tvwNotifyList.Nodes(count).Image = "Notify"
                    Case "Offline Ignored"
                        tvwNotifyList.Nodes(count).Image = "Ignored"
                    End Select
                    'And the no of computers has changed
                    If UBound(msSessionInfo) + 1 <> tvwNotifyList.Nodes(count).Children Then
                        'Reset their entry
                        addToNotify tvwNotifyList.Nodes(count).Tag
                        tvwNotifyList.Nodes.Remove count
                    End If
                End If
            'If they were offline
            ElseIf tvwNotifyList.Nodes(count).Parent.Key = "Offline" Then
                'And they are now online
                If SessionEnum(strPDC, "", tvwNotifyList.Nodes(count).Tag) = 0 Then
                    'Add them to the online list
                    addToNotify tvwNotifyList.Nodes(count).Tag, tvwNotifyList.Nodes(count).Image
                    'Remove them from the offline list
                    tvwNotifyList.Nodes.Remove count
                    
                    'If notification is on
                    If InStr(1, tvwNotifyList.Nodes(count).Image, "Notify") <> 0 Then
                        'Doevents to update the display
                        DoEvents
                        Beep
                        'Notify them! (Duh!)
                        MsgBox getRealName(strPDC, tvwNotifyList.Nodes(count).Tag) & " (" & tvwNotifyList.Nodes(count).Tag & ") is now online!", vbOKOnly + vbInformation, "User is online"
                    End If
                Else
                    'If they are still offline
                    'Ensure their icon is appropriate
                    Select Case tvwNotifyList.Nodes(count).Image
                    Case "Online"
                        tvwNotifyList.Nodes(count).Image = "Offline"
                    Case "Notify"
                        tvwNotifyList.Nodes(count).Image = "Offline Notify"
                    Case "Ignored"
                        tvwNotifyList.Nodes(count).Image = "Offline Ignored"
                    End Select
                End If
            End If
        End If
    Next count
End Sub

Public Sub addToNotify(strUsername As String, Optional strIcon As String = "")
    On Error Resume Next
    
    Dim a As Integer
    Dim intUserIndex As Integer
    
    'If they are online
    If SessionEnum(strPDC, "", strUsername) = 0 Then
        'Notify/ignore/none
        If InStr(1, strIcon, "Notify") <> 0 Then
            tvwNotifyList.Nodes.Add NODE_ONLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline Notify"
        ElseIf InStr(1, strIcon, "Ignored") <> 0 Then
            tvwNotifyList.Nodes.Add NODE_ONLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline Ignored"
        Else
            tvwNotifyList.Nodes.Add NODE_ONLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline"
        End If
        'Current index
        intUserIndex = tvwNotifyList.Nodes.count
        'Set username
        tvwNotifyList.Nodes(intUserIndex).Tag = strUsername
        'Make it visible
        tvwNotifyList.Nodes(tvwNotifyList.Nodes.count).EnsureVisible
        'For each computer they are on
        For a = 0 To UBound(msSessionInfo)
            'Add it
            tvwNotifyList.Nodes.Add intUserIndex, tvwChild, , msSessionInfo(a).CompName, "Computer"
            tvwNotifyList.Nodes(tvwNotifyList.Nodes.count).Tag = msSessionInfo(a).CompName
        Next a
    Else
        'If they are offline
        If InStr(1, strIcon, "Notify") <> 0 Then
            tvwNotifyList.Nodes.Add NODE_OFFLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline Notify"
        ElseIf InStr(1, strIcon, "Ignored") <> 0 Then
            tvwNotifyList.Nodes.Add NODE_OFFLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline Ignored"
        Else
            tvwNotifyList.Nodes.Add NODE_OFFLINE_TEXT, tvwChild, , getRealName(strPDC, strUsername) & " (" & strUsername & ")", "Offline"
        End If
        tvwNotifyList.Nodes(tvwNotifyList.Nodes.count).Tag = strUsername
        tvwNotifyList.Nodes(tvwNotifyList.Nodes.count).EnsureVisible
    End If
End Sub

Private Sub tvwNotifyList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If (tvwNotifyList.SelectedItem.Key <> NODE_ONLINE_TEXT And tvwNotifyList.SelectedItem.Key <> NODE_OFFLINE_TEXT) Then
            mnuContextIgnore.Enabled = True
            mnuContextNotify.Enabled = True
            If tvwNotifyList.SelectedItem.Parent.Text = "Online" Then
                Select Case tvwNotifyList.SelectedItem.Image
                Case "Ignored"
                    mnuContextIgnore.Checked = True
                    mnuContextNotify.Checked = False
                Case "Notify"
                    mnuContextNotify.Checked = True
                    mnuContextIgnore.Checked = False
                Case Else
                    mnuContextIgnore.Checked = False
                    mnuContextNotify.Checked = False
                End Select
                mnuContextSend.Enabled = True
            ElseIf tvwNotifyList.SelectedItem.Parent.Text = "Offline" Then
                Select Case tvwNotifyList.SelectedItem.Image
                Case "Offline Ignored"
                    mnuContextIgnore.Checked = True
                    mnuContextNotify.Checked = False
                Case "Offline Notify"
                    mnuContextNotify.Checked = True
                    mnuContextIgnore.Checked = False
                Case Else
                    mnuContextIgnore.Checked = False
                    mnuContextNotify.Checked = False
                End Select
                mnuContextSend.Enabled = False
            Else
                mnuContextIgnore.Enabled = False
                mnuContextNotify.Enabled = False
            End If
            frmMain.PopupMenu mnuContext
        End If
    End If
End Sub

Private Function notOnList(strUsername As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    
    Err.Clear
    temp = colPeople(strUsername)
    If Err.Number <> 0 Then
        notOnList = True
    Else
        notOnList = False
    End If
End Function

Private Function strreplace(strIn As String, strReplaceText As String, strOut As String) As String
    Dim a As Integer

    For a = 1 To Len(strIn) - Len(strReplaceText)
        If Mid(strIn, a, Len(strReplaceText)) = strReplaceText Then
            strreplace = strreplace + strOut
            a = a + Len(strReplaceText) - 1
        Else
            strreplace = strreplace + Mid(strIn, a, 1)
        End If
    Next a
    strreplace = strreplace + Right(strIn, Len(strReplaceText))
End Function
