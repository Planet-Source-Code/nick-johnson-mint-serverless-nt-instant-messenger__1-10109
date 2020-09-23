VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMEssage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pstrFrom As String
Private pstrTo As String
Private pstrMessage As String

Public Property Let msgFrom(strFrom As String)
    pstrFrom = strFrom
End Property
Public Property Get msgFrom() As String
    msgFrom = pstrFrom
End Property

Public Property Let msgTo(strTo As String)
    pstrTo = strTo
End Property
Public Property Get msgTo() As String
    msgTo = pstrTo
End Property

Public Property Let Message(strMessage As String)
    pstrMessage = strMessage
End Property
Public Property Get Message() As String
    Message = pstrMessage
End Property
