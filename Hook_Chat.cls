VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hook_Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements HookInterface
Dim data1 As Long
Dim data2 As Long
Dim data3 As Long

Private Sub HookInterface_Handle(app As HookApplication, flag As Long)
    Dim nickname As String, message As String, message2 As String
    'Dim addr As Long, addr2 As Long
    With app
        nickname = .readStr(.readu4byAddr(data1), .readu4byAddr(data1 + 4))
        message = .readStr(.readu4byAddr(data2), .readu4byAddr(data2 + 4))
        message2 = .readStr(.readu4byAddr(data3), .readu4byAddr(data3 + 4))
        'addr = .readu4byAddr(data2)
        'addr2 = .readu4byAddr(data2 + 4)
    End With
    app.restore flag '立即恢复线程运行
    'MsgBox addr & "|" & addr2
    'MsgBox message
    'app.restore flag '立即恢复线程运行
    If Len(message2) > Len(message) Then
        message = message2
    End If
    Form1.writeMsg "chat", nickname & "," & message
    'MsgBox "chat," & nickname & "," & message
    'MsgBox addr & "|" & add2
    'MsgBox "昵称（" & nickname & "）：" & message
    'MsgBox "chat," & nickname & "," & message
End Sub

Private Sub HookInterface_Register(app As HookApplication)
    With app
        .writeu1 &H90
        data1 = .getDataAddr
        data2 = .getDataAddr
        data3 = .getDataAddr
        'mov qword ptr ds:[0x00007FF701DB0FFF], rdx
        .writeu1 &H48
        .writeu1 &H89
        .writeu1 &H15
        .writeu4 data1 - .l - 4
        'mov qword ptr ds:[0x00007FF701DB0FFF], r10
        .writeu1 &H4C
        .writeu1 &H89
        .writeu1 &H15
        .writeu4 data2 - .l - 4
        'mov qword ptr ds:[0x00007FF701DB0FFF], r8
        .writeu1 &H4C
        .writeu1 &H89
        .writeu1 &H5
        .writeu4 data3 - .l - 4
    End With
End Sub


