VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HookMC by CH"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "开始hookMC"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "功能：1.获取玩家登录 2.获取玩家聊天内容"
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oText As Object
Private Sub Command1_Click()
    Command1.Enabled = False
    Dim hook As New HookApplication
    Dim setting As New HookSetting
    'vra=66d50000
    '68450de4
    With setting '以下设置均相对于VRA基址计算！
        .CodeAddrStart = &H1700DE4 '相对于基址的代码段开始
        .CodeAddrEnd = &H1700FFF
        .DataAddrEnd = &H26A0FFF
        '.ProgramHighAddr = &H7FF7 '64位程序的高32位 高位基址
    End With
    hook.init "bedrock_server.exe", setting
    'Exit Sub
    '以下地址均相对于VRA
    hook.addHook &H11E6338, &HFFFF8893, New Hook_Login
    hook.addHook &H98C677, &HFFFE9194, New Hook_Chat
    'hook.addHook &H8F6338, &H8EEBD0, New Hook_Login
    'hook.addHook &H9C677, &H85810, New Hook_Chat
    hook.startListener
    End
End Sub

Public Sub writeMsg(types As String, Optional str As String)
    oText.Writeline types & "," & Now & "|c|h|" & str
End Sub
Private Sub Form_Load()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oText = fso.OpenTextFile("log.txt", 8, True)
    writeMsg "start"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim msg
    msg = MsgBox("是否结束程序？如果已经hook了，程序结束后，游戏触发事件时必死！！！", vbExclamation Or vbOKCancel)
    If msg = vbOK Then End
    Cancel = -1
End Sub

