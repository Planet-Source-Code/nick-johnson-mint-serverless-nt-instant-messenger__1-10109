VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOnlineUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Online Users"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmOnlineUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCollapse 
      Caption         =   "&Collapse All"
      Height          =   375
      Left            =   2398
      TabIndex        =   2
      Top             =   3300
      Width           =   2200
   End
   Begin VB.CommandButton cmdExpand 
      Caption         =   "&Expand All"
      Height          =   375
      Left            =   96
      TabIndex        =   1
      Top             =   3300
      Width           =   2200
   End
   Begin MSComctlLib.TreeView tvwOnlineUsers 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuContext 
      Caption         =   "&Context"
      Visible         =   0   'False
      Begin VB.Menu mnuContextSend 
         Caption         =   "&Send Message"
      End
      Begin VB.Menu mnuContextSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextAdd 
         Caption         =   "&Add to list"
      End
   End
End
Attribute VB_Name = "frmOnlineUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCollapse_Click()
    Dim a As Long
    
    tvwOnlineUsers.Visible = False
    For a = 1 To tvwOnlineUsers.Nodes.count
        tvwOnlineUsers.Nodes(a).Expanded = False
    Next a
    tvwOnlineUsers.Visible = True
End Sub

Private Sub cmdExpand_Click()
    Dim a As Long
    
    tvwOnlineUsers.Visible = False
    For a = 1 To tvwOnlineUsers.Nodes.count
        tvwOnlineUsers.Nodes(a).Expanded = True
    Next a
    tvwOnlineUsers.Visible = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim nodeCount As Long
    
    frmOnlineUsers.Show vbModal
    frmOnlineUsers.Visible = True
    
    tvwOnlineUsers.ImageList = frmMain.imlNotifyList
    SessionEnum strPDC, "", ""
    For a = 0 To UBound(msSessionInfo)
        If Trim(msSessionInfo(a).UserName) <> "" Then
            tvwOnlineUsers.Nodes.Add , tvwLast, msSessionInfo(a).UserName, getRealName(strPDC, msSessionInfo(a).UserName) & " (" & msSessionInfo(a).UserName & ")", "Online"
            tvwOnlineUsers.Nodes(msSessionInfo(a).UserName).Tag = msSessionInfo(a).UserName
            nodeCount = tvwOnlineUsers.Nodes.count
            tvwOnlineUsers.Nodes.Add msSessionInfo(a).UserName, tvwChild, , msSessionInfo(a).CompName, "Computer"
            If tvwOnlineUsers.Nodes.count > nodeCount Then
                tvwOnlineUsers.Nodes(tvwOnlineUsers.Nodes.count).Tag = msSessionInfo(a).CompName
            End If
            DoEvents
        End If
    Next a
End Sub

Private Sub mnuContextAdd_Click()
    frmMain.addUser tvwOnlineUsers.SelectedItem.Key
End Sub

Private Sub mnuContextSend_Click()
    Dim frmMsg As New frmMessage
    
    frmMsg.NewMessage tvwOnlineUsers.SelectedItem.Tag
End Sub

Private Sub tvwOnlineUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        frmOnlineUsers.PopupMenu mnuContext
    End If
End Sub
