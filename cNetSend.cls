VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
' Class Module clsNetSend
'
' By: Perry Harmon
'     pharmon@email.msn.com
'
' Requires VB5 and NT 4.0
' Sends message from NT 4.0 to NT 4.0 or Win95 running winpopup
'
' Properties:
'   SendTo - Read/Write
'   SendFromServer - Read/Write
'   Message - Read/Write
'   Err - Read Only
'   ErrorText - Read Only
'
' Methods:
'   NetSendMessage - Boolean Returns True/False (Success/Failure)
'       optional parameters:
'           sUser ' String - Message Recipient
'           sMsg  ' String - Message Text
'
'   ClearError - Sets object.Err = 0, object.ErrorText = ""
'
' Events:
'   Error - raised on error
'       parameters:
'           lError ' Error Code
'           ErrorText ' Just That
'   Sent - raised when object.NetSendMessage executed without errors
'       parameters: None
'
'

'**** Example  ******
'
' Add to Form Declarations Section
' Private WithEvents mNetSend As clsNetSend
'
' Add to a procedure/function
' Set mNetSend = New clsNetSend
' With mNetSend
'    .Message = "Your Message Here"
'    .SendTo = "UserNameOrWorkStation"
'    .SendFromServer = "ServerOrWorkStation" 'This is optional
'    .NetSendMessage ' Send the message
' End With

' Private Sub mNetSend_Error(ByVal lError As Long, ByVal ErrorText As String)
'    Debug.Print lError & " - " & ErrorText
' End Sub

' Private Sub mNetSend_Sent()
'    debug.print "Message Sent"
' End Sub


'**** Example  ******
'
' Private mNetSend As New clsNetSend
' Dim bSuccess As Boolean
' If mNetSend.NetSendMessage("UserOrWorkstationName", "This Is A Test") Then
'   MsgBox "Message Sent"
' else
'   MsgBox "Error " & mNetSend.Err & " - " & mNetSend.ErrorText
' End If
'
' Broadcast a message to all members of Domain called DomainName
' bSuccess = mNetSend.NetSendMessage("DomainName*", "This Is A Test")


Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_BAD_NETPATH = 53&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_INVALID_NAME = 123&
Private Const NERR_Success = 0&       ' Success
Private Const NERR_NameNotFound = 2273& ' The message alias could not be found on the network.
Private Const NERR_NetworkError = 2136& ' A general network error occurred.

Private Declare Function netSend Lib "netapi32" Alias "NetMessageBufferSend" (ByVal cServerName As String, ByVal cMsgName As String, ByVal cFromName As String, ByVal cBuf As String, ByRef iBufLen As Integer) As Integer
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type





' Events
Public Event Error(ByVal lError As Long, ByVal ErrorText As String)
Public Event Sent()

' Local copies of properties
Private m_sMsgTo As String
Private m_sMsg As String
Private m_sMsgFrom As String
Private m_lNetApiStatus As Long
Private m_sErrorText As String
Private m_bIsWinNT As Boolean





'-----------------------------------------------------------
' FUNCTION: IsWindowsNT
'
' Returns true if this program is running under Windows NT
'-----------------------------------------------------------'
Function IsWindowsNT() As Boolean

    Dim lRC As Long
    Dim typOSInfo As OSVERSIONINFO
    
    typOSInfo.dwOSVersionInfoSize = Len(typOSInfo)
    lRC = GetVersionEx(typOSInfo)
    IsWindowsNT = (typOSInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
Public Sub ClearError()
' Set Error Properties
    m_lNetApiStatus = 0
    m_sErrorText = ""
End Sub

Public Property Get ErrorText() As String
    'Set ErrorText
    ErrorText = m_sErrorText
End Property

Public Property Get Err() As Long
    ' Set Error Number
    Err = m_lNetApiStatus
End Property

Private Function SetErrorText(Error As Long) As String
    Select Case Error
        Case ERROR_ACCESS_DENIED: SetErrorText = "Access Denied!"
        Case ERROR_BAD_NETPATH: SetErrorText = "Server '" & UCase$(m_sMsgFrom) & "' not Found."
        Case ERROR_INVALID_PARAMETER: SetErrorText = "Invalid parameter specified."
        Case ERROR_NOT_SUPPORTED: SetErrorText = "Network request not supported."
        Case ERROR_INVALID_NAME: SetErrorText = "Illegal character or malformed name."
        Case NERR_Success: SetErrorText = "Message sent."
        Case NERR_NameNotFound: SetErrorText = "User/Workstation '" & m_sMsgTo & "' not found."
        Case NERR_NetworkError: SetErrorText = "General network error occurred."
        Case Else: SetErrorText = "Unknown error executing command."
    End Select
End Function

Private Sub SetLastErr(ByVal lError As Long)
    m_lNetApiStatus = lError
    m_sErrorText = SetErrorText(lError)
    If m_lNetApiStatus Then RaiseEvent Error(m_lNetApiStatus, m_sErrorText)
End Sub
Public Function NetSendMessage(Optional ByVal sUser As String = "", Optional ByVal sMsg As String = "") As Boolean
    
    Dim sBuf
    Dim sMsgFrom As String
    Dim sMsgName As String
    Dim Net_Api_Status As Long
    
    If Not m_bIsWinNT Then Exit Function
    
    If Len(sUser) Then m_sMsgTo = sUser ' Else use the SendTo Property
    
    
    If m_sMsgTo = "" Then
        NetSendMessage = False
        SetLastErr ERROR_INVALID_PARAMETER
        RaiseEvent Error(ERROR_INVALID_PARAMETER, m_sErrorText)
    Else
        Screen.MousePointer = vbHourglass
        
        If Len(sMsg) Then m_sMsg = sMsg

        sBuf = StrConv(m_sMsg, vbUnicode)
        sMsgName = StrConv(m_sMsgTo, vbUnicode)
        
        If Len(m_sMsgFrom) And sUser = "" Then
            sMsgFrom = StrConv(m_sMsgFrom, vbUnicode)
        Else
            sMsgFrom = vbNullString
        End If
        Net_Api_Status = netSend(sMsgFrom, sMsgName, vbNullString, sBuf, ByVal Len(sBuf))
        SetLastErr Net_Api_Status
        NetSendMessage = Not CBool(Net_Api_Status)
        If NetSendMessage Then RaiseEvent Sent
        Screen.MousePointer = vbNormal
    End If

End Function

Public Property Let Message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sMessage = 5
    m_sMsg = vData
End Property


Public Property Get Message() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sMessage
    Message = m_sMsg
End Property


Public Property Let SendTo(ByVal vData As String)
    m_sMsgTo = vData
End Property
Public Property Get SendTo() As String
    SendTo = m_sMsgTo
End Property

Public Property Let SendFromServer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
    m_sMsgFrom = vData
End Property


Public Property Get SendFromServer() As String
'used when retrieving value of a property, on the right side of an assignment.
    SendFromServer = m_sMsgFrom
End Property



Private Sub Class_Initialize()
    m_bIsWinNT = IsWindowsNT()
    If m_bIsWinNT Then
        m_lNetApiStatus = 0
    Else
        MsgBox "The NetSend class requires Windows NT.", vbCritical + vbOKOnly, "Net Send"
    End If
End Sub


