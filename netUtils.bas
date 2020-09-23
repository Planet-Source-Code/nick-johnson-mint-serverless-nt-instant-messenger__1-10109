Attribute VB_Name = "netUtils"
'Declares
Public Declare Function NetSessionEnum Lib "netapi32.dll" (ServerName As Byte, UncClientName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long, ByVal PreMaxLen As Long, EntriesRead As Long, TotalEntries As Long, Resume_Handle As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Public Declare Function NetWkstaGetInfo100 Lib "netapi32" Alias "NetWkstaGetInfo" (ServerName As Byte, ByVal Level As Long, BufPtr As Any) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Public Declare Function NetGetDCName Lib "netapi32.dll" (ServerName As Byte, DomainName As Byte, DCNPtr As Long) As Long
Public Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long
Public Declare Function lstrcpyW Lib "kernel32.dll" (bRet As Byte, ByVal lPtr As Long) As Long
Public Declare Function NetUserEnum Lib "netapi32.dll" (ServerName As Byte, ByVal Level As Long, ByVal Filter As Long, Buffer As Long, ByVal PrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, ResumeHwnd As Long) As Long
Public Declare Function NetUserGetInfo Lib "netapi32.dll" (ServerName As Byte, UserName As Byte, ByVal Level As Long, Buffer As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'User Types
Public Type WKSTA_INFO_100
    dw_platform_id As Long
    ptr_computername As Long
    ptr_langroup As Long
    dw_ver_major As Long
    dw_ver_minor As Long
End Type

Public Type Session_Info_10
   sesi10_cname                       As Long
   sesi10_username                    As Long
   sesi10_time                        As Long
   sesi10_idle_time                   As Long
End Type

Public Type USER_INFO_10_API
  Name As Long
  Comment As Long
  UsrComment As Long
  FullName As Long
End Type

Public Type USERINFO_2_API
  usri2_name As Long
  usri2_password As Long
  usri2_password_age As Long
  usri2_priv As Long
  usri2_home_dir As Long
  usri2_comment As Long
  usri2_flags As Long
  usri2_script_path As Long
  usri2_auth_flags As Long
  usri2_full_name As Long
  usri2_usr_comment As Long
  usri2_parms As Long
  usri2_workstations As Long
  usri2_last_logon As Long
  usri2_last_logoff As Long
  usri2_acct_expires As Long
  usri2_max_storage As Long
  usri2_units_per_week As Long
  usri2_logon_hours As Long
  usri2_bad_pw_count As Long
  usri2_num_logons As Long
  usri2_logon_server As Long
  usri2_country_code As Long
  usri2_code_page As Long
End Type

Public Type UDT_Session_Info
    CompName                   As String
    UserName                   As String
    Time                       As Long
    IdleTime                   As Long
End Type

Public Type UDT_User_Info
    Name As String
    Comment As String
    UsrComment As String
    FullName As String
End Type

'Constants
Public Const WKSTA_LEVEL_100 = 100
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const NV_MYEVENT As Long = &H5000&
Private Const BM_CLICK = &HF5
Const WM_CLOSE = &H10

'Public variables
Public msSessionInfo() As UDT_Session_Info
Public msUserInfo As UDT_User_Info
Public strPDC As String
Public strAddUser As String
Public netSend As New clsNetSend

Public Function fnGetDomainName(strDomain) As String

Dim lngReturn As Long
Dim lngTemp As Long
Dim strTemp As String
Dim bDomain(99) As Byte
Dim bServer() As Byte
Dim lngBuffPtr As Long
Dim typeWorkstation As WKSTA_INFO_100

    fnGetDomainName = 0
    
    bServer = "" + vbNullChar
    
    lngReturn = NetWkstaGetInfo100( _
        bServer(0), _
        WKSTA_LEVEL_100, _
        lngBuffPtr)
        
    If lngReturn <> 0 Then
        fnGetDomainName = lngReturn
        Exit Function
    End If
        
    CopyMem typeWorkstation, _
        ByVal lngBuffPtr, _
        Len(typeWorkstation)
        
    lngTemp = typeWorkstation.ptr_langroup
    
    lngReturn = PtrToStr( _
        bDomain(0), _
        lngTemp)
        
    strTemp = Left( _
        bDomain, _
        StrLen(lngTemp))

    strDomain = strTemp
    
End Function

Public Function fnGetPDCName(strServer As String, strDomain As String, strPDCName As String) As Long

Dim lngReturn As Long
Dim lngDCNPtr As Long
Dim bDomain() As Byte
Dim bServer() As Byte
Dim bPDCName(100) As Byte

    fnGetPDCName = 0
    
    bServer = strServer & vbNullChar
    bDomain = strDomain & vbNullChar
    lngReturn = NetGetDCName( _
        bServer(0), _
        bDomain(0), _
        lngDCNPtr)
    
    If lngReturn <> 0 Then
        fnGetPDCName = lngReturn
        Exit Function
    End If
    
    lngReturn = PtrToStr(bPDCName(0), lngDCNPtr)
    lngReturn = NetAPIBufferFree(lngDCNPtr)
    strPDCName = bPDCName()
    strPDCName = Mid$(strPDCName, 1, InStr(strPDCName, Chr$(0)) - 1)
End Function

Public Function GetPrimaryDCName(ByVal DName As String) As String
    Dim DCName As String, DCNPtr As Long
    Dim DNArray() As Byte, DCNArray(100) As Byte
    Dim result As Long
    DNArray = DName & vbNullChar
    ' Lookup the Primary Domain Controller
    result = NetGetDCName(0&, DNArray(0), DCNPtr)
    
    If result <> 0 Then
      Err.Raise vbObjectError + 4000, "CNetworkInfo", result
      Exit Function
    End If
     
    lstrcpyW DCNArray(0), DCNPtr
    result = NetAPIBufferFree(DCNPtr)
    DCName = DCNArray()
     
    GetPrimaryDCName = Left(DCName, InStr(DCName, Chr(0)) - 1)
End Function

Public Function userExists(strServer As String, strUsername As String) As Boolean
    Dim userInfo As USER_INFO_10_API
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    
    'set variables
    baServerName = strServer & Chr$(0)
    baUserName = strUsername & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 10, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
        userExists = False
    Else
        userExists = True
    End If
    
    'Free the mem
    NetAPIBufferFree lngptrUserInfo
End Function

Public Function localUserName() As String
    Dim strUsername As String * 255
    Dim lngLength As Long
    Dim lngResult As Long
    
    lngLength = 255
    lngResult = GetUserName(strUsername, lngLength)
    If lngResult <> 1 Then
        MsgBox "An error occurred with localUserName() - No " & Str(lngResult), vbCritical, "Error in getUserName"
        Exit Function
    End If
    localUserName = Left(strUsername, lngLength)
End Function

Function SessionEnum(sServerName As String, sClientName As String, sUserName As String)
     
   Dim bFirstTime           As Boolean
   Dim lRtn                 As Long
   Dim ServerName()         As Byte
   Dim UncClientName()      As Byte
   Dim UserName()           As Byte
   Dim lptrBuffer           As Long
   Dim lEntriesRead         As Long
   Dim lTotalEntries        As Long
   Dim lResume              As Long
   Dim i                    As Integer
   Dim psComputerName               As String
   Dim psUserName                   As String
   Dim plActiveTime                 As Long
   Dim plIdleTime                   As Long
   Dim typSessionInfo()             As Session_Info_10
    
    lPrefmaxlen = 65535
     
    ServerName = sServerName & vbNullChar
    UncClientName = sClientName & vbNullChar
    UserName = sUserName & vbNullChar
    
Do
   lRtn = NetSessionEnum(ServerName(0), UncClientName(0), UserName(0), 10, lptrBuffer, lPrefmaxlen, lEntriesRead, lTotalEntries, lResume)
     
    If lRtn <> 0 Then
        SessionEnum = lRtn
        Exit Function
    End If

If lTotalEntries <> 0 Then


    ReDim typSessionInfo(0 To lEntriesRead - 1)
    ReDim msSessionInfo(0 To lEntriesRead - 1)
     
    CopyMem typSessionInfo(0), ByVal lptrBuffer, Len(typSessionInfo(0)) * lEntriesRead
     
    For i = 0 To lEntriesRead - 1
     
        msSessionInfo(i).CompName = PointerToStringW(typSessionInfo(i).sesi10_cname)
        msSessionInfo(i).UserName = PointerToStringW(typSessionInfo(i).sesi10_username)
        msSessionInfo(i).Time = typSessionInfo(i).sesi10_time
        msSessionInfo(i).IdleTime = typSessionInfo(i).sesi10_idle_time
    Next i
    End If
Loop Until lEntriesRead = lTotalEntries
   
    If lptrBuffer <> 0 Then
        NetAPIBufferFree lptrBuffer
    End If
     
End Function

Public Function PointerToStringW(lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
    
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMem Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

Public Function getRealName(strServer As String, strUsername As String) As String
    Dim lngReturn As Long
    Dim baServerName() As Byte
    Dim baUserName() As Byte
    Dim lngptrUserInfo As Long
    Dim userInfo As USER_INFO_10_API
    Dim strName As String
    Dim a As Integer
    
    'set variables
    baServerName = strServer & Chr$(0)
    baUserName = strUsername & Chr$(0)
    
    'get user info
    lngReturn = NetUserGetInfo(baServerName(0), baUserName(0), 10, lngptrUserInfo)

    'any errors?
    If lngReturn <> 0 Then
      getRealName = ""
      Exit Function
    End If

    'Turn the pointer into a variable
    CopyMem userInfo, ByVal lngptrUserInfo, Len(userInfo)
    
    strName = PointerToStringW(userInfo.FullName)
    NetAPIBufferFree lngptrUserInfo
    getRealName = strName
End Function

Public Function getMessage() As String
    Dim h As Long
    Dim k As Long
    Dim hOk As Long
    Dim hMsg As Long
    Dim strName As String
    Dim strClass As String
    Dim strMessage As String
    
    hMsg = FindWindow("#32770", "Messenger Service ")
    If hMsg = 0 Then Exit Function
    h = GetWindow(hMsg, GW_CHILD)
    Do
        strClass = Space$(16)
        k = GetClassName(h, ByVal strClass, 16)
        strClass = Left(strClass, k)
        Select Case strClass
        Case "Button"
            strName = Space(16)
            k = GetWindowText(h, strName, 16)
            If k > 0 Then strName = Left(strName, k)
            If strName = "OK" Then hOk = h
        Case "Static"
            strMessage = Space(65536)
            k = GetWindowText(h, strMessage, 65536)
            If k > 0 Then strMessage = Left(strMessage, k)
        End Select
        h = GetWindow(h, GW_HWNDNEXT)
    Loop While h <> 0
    SendMessage hOk, BM_CLICK, 0&, 0&
    getMessage = strMessage
End Function
