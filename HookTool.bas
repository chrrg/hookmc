Attribute VB_Name = "HookTool"
Option Explicit



Private Declare Function CreateToolhelp32Snapshot _
Lib "kernel32" (ByVal dwFlags As Long, _
ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First _
Lib "kernel32" (ByVal hSnapShot As Long, _
lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next _
Lib "kernel32" (ByVal hSnapShot As Long, _
lppe As PROCESSENTRY32) As Long

Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Type dword64
l As Long
h As Long
End Type
Private Const TH32CS_SNAPPROCESS = &H2&
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Type MEMORY_BASIC_INFORMATION64
    BaseAddress As Long
    BaseAddress2 As Long
    AllocationBase As Long
    AllocationBase2 As Long
    AllocationProtect As Long
    alignment1 As Long
    RegionSize As Long
    RegionSize2 As Long
    State As Long
    Protect As Long
    Type As Long
    alignment2 As Long
End Type
Public hh As Long

Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
End Function

Public Function GetPsPid(sProcess As String) As Long
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess As Long
                Dim lExitCode As Long
                GetPsPid = tPE.th32ProcessID
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Function

