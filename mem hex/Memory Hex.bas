Attribute VB_Name = "MemoryReader"
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowThreadProcessId Lib "User32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long


Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const TH32CS_SNAPPROCESS As Long = 2&
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF



Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Public myHandle As Long, ProcessS As String

Function InitProcess(Hwnd As Long) As Boolean
    PHandle = OpenProcess(PROCESS_ALL_ACCESS, False, Hwnd)
    If (PHandle = 0) Then
        InitProcess = False
        myHandle = 0
    Else
        InitProcess = True
        myHandle = PHandle
    End If
End Function

Public Function ReadByte(Offset As Long) As Byte
    ReadProcessMem myHandle, Offset, ReadByte, 1, 0&
End Function

Public Function WriteByte(addr As Long, dxValue As Byte) As Boolean
    WriteByte = True
    If WriteProcessMemory(myHandle, addr, dxValue, 1, 0&) = 0 Then
        WriteByte = False
    End If
End Function

Public Function HexToDec(ByVal HexStr As String) As Byte
Dim mult As Double
Dim DecNum As Double
Dim ch As String
Dim i As Integer
mult = 1
DecNum = 0
For i = Len(HexStr) To 1 Step -1
    ch = Mid(HexStr, i, 1)
    If (ch >= "0") And (ch <= "9") Then
        DecNum = DecNum + (Val(ch) * mult)
    Else
        If (ch >= "A") And (ch <= "F") Then
            DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
        Else
            If (ch >= "a") And (ch <= "f") Then
                DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
            Else
                HexToDec = 0
                Exit Function
            End If
        End If
    End If
    mult = mult * 16
Next i
HexToDec = DecNum
End Function









Function ConvertNumberToString(number As Double) As String
    On Error Resume Next
    If number < 256 Then ConvertNumberToString = Chr(number): Exit Function
    If number < 65536 Then
        ConvertNumberToString = Chr(number And 255) & Chr((number And 65280) / 256)
        Exit Function
    End If
    b4 = number And 255: number = Int(number / 256)
    b3 = number And 255: number = Int(number / 256)
    b2 = number And 255: number = Int(number / 256)
    b1 = number And 255: number = Int(number / 256)
    ConvertNumberToString = Chr(b4) & Chr(b3) & Chr(b2) & Chr(b1)
End Function















Function DoSearch(S As String) As Long
    On Error Resume Next
    Dim c As Long, addr As Long, buffer As String * 5000, ReadLen As Long, fNo As Integer

        c = 0
        FrmMain.rCombo.Clear
        For addr = 0 To 40000
            Call ReadProcessMemory(myHandle, addr * 5000, buffer, 5000, ReadLen)
            If addr Mod 400 = 0 Then
                FrmMain.Caption = "Searching...  " & Trim(Str(Int(addr / 400))) & "%"
                DoEvents
            End If
            If ReadLen > 0 Then
                Startpos = 1
                While InStr(Startpos, buffer, Trim(S)) > 0
                    p = (addr) * 5000 + InStr(Startpos, buffer, S) - 1
                    c = c + 1
                    FrmMain.rCombo.AddItem (p - 2)
                    Startpos = InStr(Startpos, buffer, Trim(S)) + 1
                Wend
            End If
        Next addr
    Close #fNo
    MsgBox c & " Results found.", vbInformation, "Results"
    FrmMain.Caption = ProcessS
End Function


