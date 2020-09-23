VERSION 5.00
Begin VB.Form OpenProcess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Process"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   Icon            =   "OpenProcess.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ListBox Processes 
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "OpenProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PIDs(1000) As Long

Private Sub RefreshProcessList()
    Dim Process As PROCESSENTRY32, Snapshot As Long
    Processes.Clear
    Process.dwSize = Len(Process)
    Snapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    ProcessFirst Snapshot, Process
    Processes.AddItem Process.szexeFile
    PIDs(Processes.ListCount - 1) = Process.th32ProcessID
    While ProcessNext(Snapshot, Process)
        Processes.AddItem Process.szexeFile
        PIDs(Processes.ListCount - 1) = Process.th32ProcessID
    Wend
End Sub

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    RefreshProcessList
End Sub

Private Sub ok_Click()
    On Error Resume Next
    If Not InitProcess(PIDs(Processes.ListIndex)) Then
        MsgBox "This process cannot be opened. Please check this is not a system process.", vbExclamation, "Error"
    Else
        FrmMain.Show 0
        ProcessS = Processes.List(Processes.ListIndex)
        FrmMain.Caption = Processes.List(Processes.ListIndex)
        Unload Me
    End If
End Sub

