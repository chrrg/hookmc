VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HookSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public ProgramHighAddr As Long '程序高位基址
Public CodeAddrStart As Long    '空闲代码段开始 全0表空闲
Public CodeAddrEnd As Long    '空闲代码段结束的位置
Public DataAddrEnd As Long    '空闲数据段的结束位置
