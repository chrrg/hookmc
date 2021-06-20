VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HookApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function openProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function VirtualProtectEx64 Lib "wow64ext.dll" (ByVal hProcess As Long, ByVal l As Long, ByVal w As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Currency
Private Declare Function VirtualAllocEx64 Lib "wow64ext.dll" (ByVal hProcess As Long, ByVal l As Long, ByVal w As Long, ByVal dwnSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As dword64
Private Declare Function WriteProcessMemory64 Lib "wow64ext.dll" (ByVal hProcess As Long, ByVal l As Long, ByVal w As Long, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory64 Lib "wow64ext.dll" (ByVal hProcess As Long, ByVal l As Long, ByVal w As Long, ByRef lpBuffer As Any, ByVal nSize As Long, ByRef lpNumberOfBytesRead As Long) As Long
Private Declare Function VirtualQueryEx64 Lib "wow64ext.dll" (ByVal hProcess As Long, ByVal l As Long, ByVal w As Long, ByRef lpBuffer As MEMORY_BASIC_INFORMATION64, ByVal nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private pid As Long
Private hProcess As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const PAGE_EXECUTE_READWRITE = &H40&

Public l As Long
Public h As Long
Dim flag() As Long

Dim hookAppList() As HookInterface
Dim hookCount As Long
Dim mySetting As HookSetting
Dim VRA As Long
Sub init(name As String, setting As HookSetting)
    pid = GetPsPid(name)
    If pid = 0 Then
        MsgBox "进程未找到！"
        End
    End If
    hProcess = openProcess(PROCESS_ALL_ACCESS, 0, pid)
    HookTool.hh = hProcess
    If hProcess = 0 Then MsgBox "进程打不开！": End
    Set mySetting = setting
    '
    setting.DataAddrEnd = setting.DataAddrEnd + 1
    
    VRA = getVRA(h) '程序基址vra
    'MsgBox h
    'End
End Sub

Sub addHook(hookAddr As Long, FuncAddr As Long, hookApp As HookInterface)
'hookAddr + VRA = hook的地方的地址 这里需要找到 call xxxx指令的地址减去VRA即是需要传入的值
'FuncAddr = hook这个地方对应的函数的物理相对地址 call后面的地址
'hookAddr+VRA+FuncAddr = hook这个地方对应的函数要执行的地址

    l = hookAddr + VRA
    Dim myFuncAddress As Long 'hook写入的函数地址
    myFuncAddress = VRA + mySetting.CodeAddrStart
    
    l = l + 1
    Dim verify As Long
    
    verify = readu4()
    If verify <> FuncAddr Then
        Dim msg
        If verify = myFuncAddress - l - 4 Then
            'msg = MsgBox("已经注入过了，是否重新注入？", vbOKCancel)
            'If msg = vbCancel Then End
        Else
            msg = MsgBox("内存基址不正确！是否强行继续？？？（非常不建议，请确保版本是否匹配！）" & vbCrLf & verify & "-" & (myFuncAddress - l - 4), vbOKCancel Or vbCritical)
            If msg = vbCancel Then End
        End If
    End If
    'If verify <> FuncAddr - hookAddr - 5 Then '验证
    '    Dim msg
    '    If verify = myFuncAddress - l - 4 Then
    '        msg = MsgBox("已经注入过了，是否重新注入？", vbOKCancel)
    '        If msg = vbCancel Then End
    '    Else
    '        msg = MsgBox("内存基址不正确！是否强行继续？？？", vbOKCancel Or vbCritical)
    '        If msg = vbCancel Then End
    '    End If
    'End If
    l = l - 1
    writeu1 &HE9 'call -> jmp
    
    writeu4 myFuncAddress - l - 4 'dword4字节
    l = mySetting.CodeAddrStart + VRA
    writeu1 &H90
    writeu1 &H90
    writeu1 &H90
    writeu1 &H90

    
    hookApp.Register Me
    
    hookCount = hookCount + 1
    ReDim Preserve flag(hookCount)
    ReDim Preserve hookAppList(hookCount)
    Dim myFlag As Long
    myFlag = getDataAddr
    flag(hookCount - 1) = myFlag
    Set hookAppList(hookCount - 1) = hookApp
    'mov byte ptr ds:[0x00007FF701DB0FFF], 0x1
    writeu1 &HC6
    writeu1 &H5
    writeu4 myFlag - l - 5
    writeu1 &H1
    '设置后开始死循环
    'pos1:
    'cmp byte ptr ds:[0x00007FF701DB0FFF], 0x1
    writeu1 &H80
    writeu1 &H3D
    writeu4 myFlag - l - 5
    writeu1 &H1
    'je pos1:
    writeu1 &H74
    writeu1 &HF7
    
    'call entry'调用真正的函数
    writeu1 &HE8
    writeu4 hookAddr + VRA + FuncAddr - l + 1
    
    writeu1 &HE9 'jmp nop
    writeu4 hookAddr + VRA - l + 1 '跳回调用函数的地方的下一个语句
    mySetting.CodeAddrStart = l - VRA
End Sub
Function getVRA(ByRef highAddr As Long)
    On Error Resume Next
    getVRA = getVRAs1(highAddr)
End Function
Private Function getVRAs1(ByRef highAddr As Long) As Long
    Dim buffer As MEMORY_BASIC_INFORMATION64
    Dim addr As Long, addr2 As Long
    Dim result As Long
    result = VirtualQueryEx64(hProcess, addr, addr2, buffer, Len(buffer))
    If result = 0 Then
        MsgBox "请尝试使用管理员打开！"
        End
    End If
    Do
        addr2 = addr2 + buffer.RegionSize2
        Dim temp As Long
        temp = addr
        addr = UnsignedAdd(addr, buffer.RegionSize)
        If temp < 0 And addr >= 0 Then addr2 = addr2 + 1
        'If buffer.BaseAddress2 > &H7FF7 Then
        '    MsgBox "VRA获取失败1!"
        '    End
        'End If
        If buffer.Type = &H1000000 And buffer.RegionSize = &H1000 Then 'img
            getVRAs1 = buffer.BaseAddress
            highAddr = buffer.BaseAddress2
            'MsgBox getVRAs1 & "|" & highAddr
            Exit Do
        End If
    Loop Until VirtualQueryEx64(hProcess, addr, addr2, buffer, Len(buffer)) = 0
    If getVRAs1 = 0 Then
        MsgBox "VRA获取失败！"
        End
    End If
End Function
Private Function UnsignedAdd(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4 As Long
    Dim lY4 As Long
    Dim lX8 As Long
    Dim lY8 As Long
    Dim lResult As Long
    '提取最高位'
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    '提起次高位'
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    '剩余位直接相加'
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then '次高位同时为 1，和为 &H80000000
        'lResult = lResult + 次高位的和（即&H80000000） + X的最高位 + Y的最高位
        '其中向更高位 &H100000000 的进位直接忽略了
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then '次高位只有一个 1'
        'lResult = lResult + 次高位的和（即&H40000000） + X的最高位 + Y的最高位'
        If lResult And &H40000000 Then '剩余位的和向次高位有进位
            '  lResult + 次高位的和（即&H40000000）
            '= lResult + &H80000000 - &H40000000
            '= lResult Xor &HC0000000
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else '次高位全0，和为0'
        'lResult = lResult + 次高位和（即&H00000000） + X的最高位 + Y的最高位'
        lResult = lResult Xor lX8 Xor lY8
    End If
    UnsignedAdd = lResult
End Function
Sub restore(flag As Long)
    writeu1byAddr flag, 0 '恢复运行
End Sub
Function startListener()
    While True
        DoEvents
        Sleep 1
        Dim i As Long
        For i = 0 To hookCount - 1
            If readu1byAddr(flag(i)) Then
                hookAppList(i).Handle Me, flag(i) '处理hook
                writeu1byAddr flag(i), 0 '恢复运行
            End If
        Next
    Wend
End Function
Function getDataAddr() As Long
    mySetting.DataAddrEnd = mySetting.DataAddrEnd - 16
    getDataAddr = mySetting.DataAddrEnd + VRA
End Function
Function readu1byAddr(ByVal addr As Long) As Byte
    On Error Resume Next
    readu1byAddr = readu1s1(addr, h)
End Function
Function readu1() As Byte
    On Error Resume Next
    readu1 = readu1s1(l, h)
End Function
Private Function readu1s1(l As Long, h As Long) As Byte
    Dim i2 As Long, result As Long
    result = ReadProcessMemory64(hProcess, l, h, readu1s1, 1, i2)
    If result = 0 Then
        Dim msg
        msg = MsgBox("ReadProcessMemory64错误！是否结束hook的运行？", vbCritical Or vbOKCancel)
        If msg = vbOK Then End
        'MsgBox "ReadProcessMemory64 error"
        'Exit Function 'err
    End If
End Function
Function readu4byAddr(ByVal addr As Long) As Long
    On Error Resume Next
    readu4byAddr = readu4s1(addr, h)
End Function
Function readu4() As Long
    On Error Resume Next
    readu4 = readu4s1(l, h)
End Function
Private Function readu4s1(l As Long, h As Long) As Long
    Dim i2 As Long, result As Long
    result = ReadProcessMemory64(hProcess, l, h, readu4s1, 4, i2)
    If result = 0 Then MsgBox "ReadProcessMemory64 error": Exit Function 'err
End Function
Sub writeu1(ByVal value As Byte)
    On Error Resume Next
    writeu1s1 value
End Sub
Private Sub writeu1s1(ByVal value As Byte)
    Dim result As Long
    Dim old As Long
    Dim i As Currency
    i = VirtualProtectEx64(hProcess, l, h, 1, PAGE_EXECUTE_READWRITE, old)
    If i = 0 Then MsgBox "err": Exit Sub      'err
    Dim i2 As Long
    result = WriteProcessMemory64(hProcess, l, h, value, 1, i2)
    If result = 0 Then MsgBox "err2": Exit Sub      'err
    l = l + 1
    VirtualProtectEx64 hProcess, l, h, 1, old, old
End Sub
Sub writeu4(ByVal value As Long)
    On Error Resume Next
    writeu4s1 value
End Sub
Private Sub writeu4s1(ByVal value As Long)
    Dim result As Long
    Dim old As Long
    Dim i As Currency
    i = VirtualProtectEx64(hProcess, l, h, 4, PAGE_EXECUTE_READWRITE, old)
    If i = 0 Then MsgBox "err": Exit Sub      'err
    Dim i2 As Long
    result = WriteProcessMemory64(hProcess, l, h, value, 4, i2)
    If result = 0 Then MsgBox "err2": Exit Sub      'err
    l = l + 4
    VirtualProtectEx64 hProcess, l, h, 4, old, old
End Sub


Sub writeu1byAddr(ByVal addr As Long, ByVal value As Byte)
    On Error Resume Next
    writeu1s1byAddr addr, value
End Sub
Private Sub writeu1s1byAddr(ByVal addr As Long, ByVal value As Byte)
    Dim result As Long
    Dim old As Long
    Dim i As Currency
    i = VirtualProtectEx64(hProcess, addr, h, 1, PAGE_EXECUTE_READWRITE, old)
    If i = 0 Then MsgBox "err": Exit Sub      'err
    Dim i2 As Long
    result = WriteProcessMemory64(hProcess, addr, h, value, 1, i2)
    If result = 0 Then MsgBox "err2": Exit Sub      'err
    VirtualProtectEx64 hProcess, addr, h, 1, old, old
End Sub

Function readStr(ByVal l As Long, ByVal h As Long) As String
    On Error Resume Next
    readStr = readStrs1(l, h)
End Function
Private Function readStrs1(ByVal l As Long, ByVal h As Long) As String
    Dim i2 As Long, result As Long
    Dim buffer() As Byte
    Dim size As Long
    Dim temp As Byte
    size = 0
    Do
        result = ReadProcessMemory64(hProcess, l + size, h, temp, 1, i2)
        If result = 0 Then
            Dim msg
            msg = MsgBox("ReadProcessMemory64错误！是否结束hook的运行？", vbCritical Or vbOKCancel)
            If msg = vbOK Then End
            'MsgBox "ReadProcessMemory64 error": Exit Function 'err
        End If
        If temp = 0 Then
            Exit Do
            'Dim temp2 As Long
            'result = ReadProcessMemory64(hProcess, l + size, h, temp2, 1, i2)
            'If result = 0 Then Exit Do
            'If temp2 = 0 Then Exit Do
        End If
        size = size + 1
        ReDim Preserve buffer(size)
        buffer(size - 1) = temp
    Loop Until False
    If size = 0 Then Exit Function
    readStrs1 = Trim(UTF8ToGB2312(buffer))
    If Len(readStrs1) = 0 Then Exit Function
    readStrs1 = Left(readStrs1, Len(readStrs1) - 1)
End Function

