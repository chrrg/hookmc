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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼhookMC"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ܣ�1.��ȡ��ҵ�¼ 2.��ȡ�����������"
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
    With setting '�������þ������VRA��ַ���㣡
        .CodeAddrStart = &H1700DE4 '����ڻ�ַ�Ĵ���ο�ʼ
        .CodeAddrEnd = &H1700FFF
        .DataAddrEnd = &H26A0FFF
        '.ProgramHighAddr = &H7FF7 '64λ����ĸ�32λ ��λ��ַ
    End With
    hook.init "bedrock_server.exe", setting
    'Exit Sub
    '���µ�ַ�������VRA
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
    msg = MsgBox("�Ƿ������������Ѿ�hook�ˣ������������Ϸ�����¼�ʱ����������", vbExclamation Or vbOKCancel)
    If msg = vbOK Then End
    Cancel = -1
End Sub

