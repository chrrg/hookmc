VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hook_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements HookInterface
Dim data1 As Long

Private Sub HookInterface_Handle(app As HookApplication, flag As Long)
    With app
        Dim nickname As String
        Dim addr As Long, add2 As Long
        addr = .readu4byAddr(data1)
        nickname = .readStr(.readu4byAddr(data1), .readu4byAddr(data1 + 4))
    End With
    app.restore flag '立即恢复线程运行
    Form1.writeMsg "login", nickname
    'MsgBox "昵称（" & nickname & "）已上线！"
End Sub

Private Sub HookInterface_Register(app As HookApplication)
    With app
        data1 = .getDataAddr
        'mov qword ptr ds:[0x00007FF701DB0FFF], r9
        .writeu1 &H4C
        .writeu1 &H89
        .writeu1 &HD
        .writeu4 data1 - .l - 4
    End With
End Sub

