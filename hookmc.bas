Attribute VB_Name = "hookmc"
Option Explicit

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll" (ByVal hProcess As Long, _
ByVal laddr As Long, ByVal waddr As Long, buffer As Any, ByVal lsize As Long, _
ByVal wsize As Long, ret As Long) As Long
Private Declare Function NtWow64WriteVirtualMemory64 Lib "ntdll" (ByVal hProcess As Long, _
ByVal laddr As Long, ByVal waddr As Long, buffer As Any, ByVal lsize As Long, _
ByVal wsize As Long, ret As Long) As Long

'VirtualProtectEx64(HANDLE hProcess, DWORD64 lpAddress, SIZE_T dwSize, DWORD flNewProtect, DWORD* lpflOldProtect);


'Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwnSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
'Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwnSize As Long, ByVal dwFreeType As Long) As Long
'Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByVal lpflOldProtect As Long) As Long


Const PROCESS_TERMINATE = 1
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RELEASE = &H8000
Private Const PAGE_EXECUTE_READWRITE = &H40&
Type DWORD64
    l As Long
    h As Long
End Type
Private Const PAGE_READWRITE = &H4
Private Const STILL_ACTIVE = &H103&
Private Const INFINITE = &HFFFF

Public pid As Long
Public hProcess As Long
Public l As Long
Public h As Long






