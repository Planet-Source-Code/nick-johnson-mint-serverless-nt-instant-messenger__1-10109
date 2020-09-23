VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   7
      TabIndex        =   2
      Top             =   2045
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&Reply"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2045
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.CommandButton cmdOptions 
      Cancel          =   -1  'True
      Caption         =   "&Options"
      Height          =   375
      Left            =   2360
      TabIndex        =   4
      Top             =   2045
      Width           =   2325
   End
   Begin VB.TextBox txtMessage 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   550
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   4695
      Begin VB.ComboBox cmbFrom 
         Height          =   315
         Left            =   580
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Click for a list of possible senders"
         Top             =   150
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From:"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   200
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "To:"
         Height          =   255
         Left            =   2450
         TabIndex        =   6
         Top             =   200
         Width           =   255
      End
      Begin VB.Label lblTo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2775
         TabIndex        =   5
         Top             =   150
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum messageTypes
    eNewMessage = 1
    eRecievedMessage = 2
End Enum

Private msgType As messageTypes
Private originalText As String
Private colMessages As New Collection

Public Sub NewMessage(strTo As String)
    Me.Caption = "New Message"
    lblTo.Caption = strTo
    cmbFrom.AddItem localUserName
    cmbFrom.ListIndex = 0
    cmdSend.Visible = True
    msgType = eNewMessage
    cmdOptions.Enabled = False
    Me.Show
End Sub

Public Sub RecievedMessage(strTo As String, strFrom As String, strMessage As String)
    Dim msgIncoming As New cMEssage
    
    msgIncoming.msgTo = strTo
    msgIncoming.msgFrom = strFrom
    msgIncoming.Message = strMessage
    msgType = eRecievedMessage
    Me.Caption = "Recieved Message"
    colMessages.Add msgIncoming
    
    cmdOptions.Enabled = True
    
    If Me.Visible = False Then 'First incoming message
        Me.Visible = True
        readNextMessage
    End If
    
    updateButton
End Sub

Private Sub updateButton()
    'Update the options button
    If colMessages.count > 0 Then
        cmdOptions.Caption = "Options (" & colMessages.count & " message(s) waiting)"
        frmMain.mnuOptionsDelete.Enabled = True
        frmMain.mnuOptionsNext.Enabled = True
    Else
        cmdOptions.Caption = "Options"
        frmMain.mnuOptionsDelete.Enabled = False
        frmMain.mnuOptionsNext.Enabled = False
    End If
End Sub

Public Sub readNextMessage()
    Dim msgRead As cMEssage
    
    If colMessages.count > 0 Then
        Set msgRead = colMessages.Item(1)
        colMessages.Remove 1
        cmbFrom.Clear
        cmbFrom.AddItem msgRead.msgFrom
        cmbFrom.ListIndex = 0
        
        lblTo.Caption = msgRead.msgTo
        
        originalText = msgRead.Message
        txtMessage.Text = " " 'Prompt it to reset
        cmdReply.Visible = True
        
        updateButton
    End If
End Sub

Public Sub deleteAll()
    Dim a As Integer
    
    For a = colMessages.count To 1 Step -1
        colMessages.Remove a
    Next a
    Unload Me
End Sub

Private Sub cmbFrom_Click()
    If cmbFrom.ListCount <= 1 Then
        SessionEnum strPDC, "", ""
        For a = 0 To UBound(msSessionInfo)
            If msSessionInfo(a).CompName = cmbFrom.List(0) And msSessionInfo(a).UserName <> "" Then
                cmbFrom.AddItem msSessionInfo(a).UserName
            End If
        Next a
    End If
    
    cmbFrom.ToolTipText = getRealName(strPDC, cmbFrom.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOptions_Click()
    frmMessage.PopupMenu frmMain.mnuOptions
End Sub

Private Sub cmdReply_Click()
    Dim frmMsg As New frmMessage
    
    frmMsg.NewMessage cmbFrom.Text
    'Unload Me
End Sub

Private Sub cmdSend_Click()
    netSend.NetSendMessage lblTo.Caption, "From:" & Left(localUserName, Len(localUserName) - 1) & vbCrLf & txtMessage.Text
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If colMessages.count > 0 Then
        If MsgBox("This will delete all incoming messages - are you sure?", vbYesNo + vbExclamation, "Confirm") = vbYes Then
            deleteAll
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtMessage_Change()
    If msgType = eRecievedMessage Then
        txtMessage.Text = originalText
    End If
End Sub
