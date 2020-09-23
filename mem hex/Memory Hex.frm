VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "0"
   ClientHeight    =   12720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18600
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Memory Hex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12720
   ScaleWidth      =   18600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   10680
      TabIndex        =   33
      Top             =   11280
      Width           =   2415
      Begin VB.ListBox rCombo 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Results"
         Top             =   280
         Width           =   1935
      End
   End
   Begin VB.CommandButton cPanel 
      BackColor       =   &H8000000D&
      Caption         =   "Edit Mode On"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   15000
      TabIndex        =   32
      ToolTipText     =   "Edit Mode Off"
      Top             =   12240
      Width           =   1695
   End
   Begin VB.CommandButton cPanel 
      Caption         =   "New Process"
      DownPicture     =   "Memory Hex.frx":08CA
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13200
      Picture         =   "Memory Hex.frx":0D0C
      TabIndex        =   31
      ToolTipText     =   "Open New Process"
      Top             =   12240
      Width           =   1695
   End
   Begin VB.CommandButton cPanel 
      Caption         =   "Alpha Mode"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   16800
      TabIndex        =   30
      ToolTipText     =   "Alpha Mode"
      Top             =   12240
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   11280
      Width           =   10455
      Begin VB.TextBox NoConv 
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   36
         ToolTipText     =   "umber"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox NoConv 
         Height          =   285
         Index           =   1
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   35
         ToolTipText     =   "Characters"
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox NoConv 
         Height          =   285
         Index           =   2
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   34
         ToolTipText     =   "Hex"
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton SrchC 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         TabIndex        =   26
         ToolTipText     =   "Search"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox SrSrch 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   25
         ToolTipText     =   "Search"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdsrch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         ToolTipText     =   "Search"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox SrByte 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         ToolTipText     =   "Byte"
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Hex"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   29
         ToolTipText     =   "Hex"
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Numeric"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   28
         ToolTipText     =   "Numeric"
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Caption         =   "String"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   27
         ToolTipText     =   "String"
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Number:"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "String:"
         Height          =   375
         Left            =   3720
         TabIndex        =   38
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Hex:"
         Height          =   375
         Left            =   7320
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Search:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Go to Byte:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Converstions"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      TabIndex        =   13
      Top             =   11280
      Width           =   5295
      Begin VB.TextBox chardisp 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   1
         TabIndex        =   16
         ToolTipText     =   "Character"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox hexdispl 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   15
         ToolTipText     =   "Hex"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox asciidisp 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   14
         ToolTipText     =   "Ascii"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Char:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Ascii:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Hex:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame mFrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18375
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   17640
         Picture         =   "Memory Hex.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Bottom"
         Top             =   9840
         Width           =   615
      End
      Begin VB.VScrollBar VScroll 
         Height          =   10560
         Left            =   17160
         Max             =   20000
         TabIndex        =   9
         Top             =   495
         Value           =   1
         Width           =   375
      End
      Begin VB.PictureBox HexDisp 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10550
         Left            =   1320
         ScaleHeight     =   10485
         ScaleWidth      =   15675
         TabIndex        =   7
         Top             =   480
         Width           =   15735
         Begin VB.TextBox Edit 
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   0
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "00"
            Top             =   0
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Timer sTimer 
         Interval        =   2000
         Left            =   2280
         Top             =   7815
      End
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   17640
         Picture         =   "Memory Hex.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Top"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   1
         Left            =   17640
         Picture         =   "Memory Hex.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Up Screen"
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1855
         Index           =   2
         Left            =   17640
         Picture         =   "Memory Hex.frx":1E14
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Up Line"
         Top             =   3855
         Width           =   615
      End
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1855
         Index           =   3
         Left            =   17640
         Picture         =   "Memory Hex.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Down Line"
         Top             =   5850
         Width           =   615
      End
      Begin VB.CommandButton sDir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Index           =   4
         Left            =   17640
         Picture         =   "Memory Hex.frx":2698
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Down Screen"
         Top             =   7815
         Width           =   615
      End
      Begin VB.PictureBox sValues 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   10560
         Left            =   120
         ScaleHeight     =   10560
         ScaleWidth      =   1215
         TabIndex        =   1
         Top             =   495
         Width           =   1215
      End
      Begin VB.PictureBox dValues 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   15735
         TabIndex        =   10
         Top             =   255
         Width           =   15735
      End
      Begin VB.PictureBox Value 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   1335
         TabIndex        =   11
         Top             =   255
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sData(1 To 2500) As String, cPos As Long, HexMode As Boolean, sDet As String, sTrd As Integer
Dim bCol As Boolean, sTemp As Boolean, eMode As Boolean, cPosn As Long, tOpt As Integer

Const Max = 199997501

Public Function nHexToDec(sSource As String) As Variant
    On Error Resume Next
    Dim i As Integer, iVal As Byte, rString As String
    rString = ""
    For i = 0 To Len(sSource) - 1
        rString = rString & Mid(sSource, Len(sSource) - i, 1)
    Next i
    
    
    If (Trim$(rString) = "") Then
        nHexToDec = 0
        Exit Function
    End If
    nHexToDec = 0
    For i = 1 To Len(rString)
        iVal = CByte("&h" & Mid$(rString, i, 1))
        nHexToDec = CDec(nHexToDec + iVal * 16 ^ (Len(rString) - i))
    Next i
End Function

Private Sub asciidisp_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If asciidisp > 255 Then
        hexdispl.Text = ""
        chardisp.Text = ""
    Else
        hexdispl.Text = Hex(asciidisp.Text)
        chardisp.Text = Chr(asciidisp)
    End If
End Sub

Private Sub chardisp_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    asciidisp.Text = Asc(chardisp.Text)
    hexdispl.Text = Hex(asciidisp.Text)
End Sub

Private Sub cmdsrch_Click()
    On Error Resume Next
    GotoByte SrByte.Text
End Sub

Function GotoByte(sByte As Long)
    On Error Resume Next
    Dim tVal As Long, mNo As Single, aMod As Integer
    If sByte < 1 Then sByte = 1
    If sByte > Max Then sByte = Max
    cPos = (Int(sByte / 50) * 50) + 1
    UpdateData
    tVal = HexDisp.Width / 50
    mNo = ((((sByte Mod 50 - 1) + 0.5) / 50) * HexDisp.Width) + tVal
    HexDisp_MouseDown 0, 0, mNo, 1
    Edit.Text = sData(cPosn)
    aMod = sByte Mod 50
    Value.Cls
    Value.Print cPos - 1 + aMod
End Function

Private Sub cPanel_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        OpenProcess.Show 1
    Case 1
        If eMode = True Then
            eMode = False
            Edit.BackColor = &H800000
            Edit.ForeColor = vbWhite
            Edit.Locked = True
            cPanel(1).ToolTipText = "Edit Mode On"
            cPanel(1).Caption = "Edit Mode On"
        Else
            eMode = True
            Edit.BackColor = vbYellow
            Edit.ForeColor = vbBlack
            Edit.Locked = False
            cPanel(1).ToolTipText = "Edit Mode Off"
            cPanel(1).Caption = "Edit Mode Off"
        End If
    Case 2
        Edit.Visible = False
        If HexMode = False Then
            Edit.MaxLength = 2
            HexMode = True
            cPanel(2).Caption = "Alpha Mode"
            cPanel(2).ToolTipText = "Alpha Mode"
        Else
            Edit.MaxLength = 1
            HexMode = False
            cPanel(2).Caption = "Hex Mode"
            cPanel(2).ToolTipText = "Hex Mode"
        End If
        UpdateData
    End Select
End Sub

Private Sub Edit_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim nVal As Byte
    If Edit.Locked = True Then Exit Sub
    If KeyAscii = 13 Then
        If HexMode = True Then
            If Edit.Text = "" Then
                nVal = 0
            Else
                nVal = HexToDec(Edit.Text)
            End If
        Else
            nVal = Asc(Edit.Text)
        End If
        If WriteByte(cPosn, nVal) = True Then
            Edit.BackColor = vbGreen
            UpdateData
        Else
            Edit.BackColor = vbRed
            Edit.Text = sDet
        End If
    Else
        If HexMode = True Then
            Character = Chr(KeyAscii)
            KeyAscii = Asc(UCase(Character))
            If Chr(KeyAscii) <> vbBack Then
                If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
                    DoEvents
                Else
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub

Private Sub HexDisp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim col As Integer, row As Integer, vEdit As Byte
    col = Int((X / HexDisp.Width) * 50) + 1
    row = Int((Y / HexDisp.Height) * 50) + 1
    cPosn = ((row - 1) * 50 + col) + cPos
    Value.Cls
    Value.Print cPosn - 2
    
    If eMode = True Then Edit.BackColor = vbYellow
    
    vEdit = sData((row - 1) * 50 + col)
    If HexMode = True Then
        If Len(Hex(vEdit)) = 1 Then
            Edit.Text = "0" & Hex(vEdit)
        Else
            Edit.Text = Hex(vEdit)
        End If
    Else
        If vEdit = 0 Or vEdit = 13 Or vEdit = 10 Or vEdit = 9 Then
            Edit.Text = " "
        Else
            Edit.Text = Chr(vEdit)
        End If
    End If
    sDet = Edit.Text
    Edit.Visible = True
    Edit.Left = Int((X / HexDisp.Width) * 50) * (HexDisp.Width / 50)
    If Y > 5000 Then
        Edit.Top = Int((Y / HexDisp.Height) * 50) * (HexDisp.Height / 50) - 50
    Else
        Edit.Top = Int((Y / HexDisp.Height) * 50) * (HexDisp.Height / 50)
    End If
End Sub

Private Sub HexDisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim col As Integer, row As Integer
    col = Int((X / HexDisp.Width) * 50) + 1
    row = Int((Y / HexDisp.Height) * 50) + 1
    cPosn = ((row - 1) * 50 + col) + cPos
    Value.Cls
    Value.Print cPosn - 2
    
    vEdit = sData((row - 1) * 50 + col)
    If HexMode = True Then
        If Len(Hex(vEdit)) = 1 Then
            Edit.Text = "0" & Hex(vEdit)
        Else
            Edit.Text = Hex(vEdit)
        End If
    Else
        If vEdit = 0 Or vEdit = 13 Or vEdit = 10 Or vEdit = 9 Then
            Edit.Text = " "
        Else
            Edit.Text = Chr(vEdit)
        End If
    End If
    
    Edit.Visible = True
    Edit.Left = Int((X / HexDisp.Width) * 50) * (HexDisp.Width / 50)
    If Y > 5000 Then
        Edit.Top = Int((Y / HexDisp.Height) * 50) * (HexDisp.Height / 50) - 50
    Else
        Edit.Top = Int((Y / HexDisp.Height) * 50) * (HexDisp.Height / 50)
    End If
End Sub


Private Sub hexdispl_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Character = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Character))
    
    If Chr(KeyAscii) <> vbBack Then
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
            DoEvents
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub hexdispl_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim AscStore As Integer
    AscStore = HexToDec(hexdispl)
    asciidisp.Text = AscStore
    chardisp.Text = Chr(AscStore)
End Sub

Private Sub NoConv_Change(Index As Integer)
    On Error Resume Next
    Dim counter As Integer, g As String, fVal As String
    Select Case sTrd
    Case 0:
        NoConv(1) = ConvertNumberToString(Val(NoConv(0)))
        fVal = ""
        For counter = 1 To Len(NoConv(1))
            g = Mid(NoConv(1), counter, 1)
            If Len(Hex(Asc(g))) = 1 Then
                fVal = fVal & "0" & Hex(Asc(g))
            Else
                fVal = fVal & Hex(Asc(g))
            End If
        Next counter
        NoConv(2) = fVal
    Case 1:
        fVal = ""
        For counter = 1 To Len(NoConv(1))
            g = Mid(NoConv(1), counter, 1)
            fVal = fVal & Hex(Asc(g))
        Next counter
        NoConv(2) = fVal
        NoConv(0) = nHexToDec(NoConv(2))
    Case 2:
        NoConv(0) = nHexToDec(NoConv(2))
        NoConv(1) = ConvertNumberToString(Val(NoConv(0)))
    
    End Select
End Sub

Private Sub NoConv_GotFocus(Index As Integer)
    sTrd = Index
End Sub

Private Sub NoConv_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case Index
    Case 0:
        Character = Chr(KeyAscii)
        KeyAscii = Asc(UCase(Character))
        
        If Chr(KeyAscii) <> vbBack Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Then
                DoEvents
            Else
                KeyAscii = 0
            End If
        End If
    Case 2:
        Character = Chr(KeyAscii)
        KeyAscii = Asc(UCase(Character))
        
        If Chr(KeyAscii) <> vbBack Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
                DoEvents
            Else
                KeyAscii = 0
            End If
        End If
    End Select
End Sub

Private Sub Opt_Click(Index As Integer)
    If tOpt <> Index Then
        SrSrch.Text = ""
        tOpt = Index
    End If
End Sub



Private Sub rCombo_Click()
    On Error Resume Next
    GotoByte rCombo.List(rCombo.ListIndex)
End Sub

Private Sub rCombo_Scroll()
    On Error Resume Next
    GotoByte rCombo.List(rCombo.ListIndex)
End Sub

Private Sub SrByte_Change()
    cmdsrch.Default = True
End Sub

Private Sub SrByte_GotFocus()
    On Error Resume Next
    SrByte.Text = ""
End Sub

Private Sub SrByte_LostFocus()
    cmdsrch.Default = False
End Sub

Private Sub SrchC_Click()
    On Error Resume Next
    Dim StrSearch As String, counter As Integer, tVal As String
    Select Case tOpt
    Case 0:
        If HexMode = True Then cPanel_Click 2
        MemoryReader.DoSearch SrSrch.Text
    Case 1:
        If Not IsNumeric(SrSrch) Then MsgBox "Please enter a numeric value.", vbCritical: Exit Sub
        If SrSrch = 0 Then MsgBox "Cannot search for 0. Search will find too many results.", vbCritical: Exit Sub
        If Val(txtSearch) > 4294967295# Then MsgBox "Please enter a smaller value in search field.", vbCritical: Exit Sub
        StrSearch = ConvertNumberToString(Val(SrSrch))
        If HexMode = True Then cPanel_Click 2
        NoConv(0) = SrSrch
        DoSearch StrSearch
    Case 2:
        If Len(SrSrch.Text) Mod 2 = 1 Then MsgBox "Invalid number of digits for a hex code. Please check and try again", vbCritical, "Error":                Exit Sub
        StrSearch = ""
        For counter = 1 To Len(SrSrch) Step 2
            tVal = Mid(SrSrch, counter, 2)
            StrSearch = StrSearch & (Chr(HexToDec(tVal)))
        Next counter
        If HexMode = False Then cPanel_Click 2
        DoSearch StrSearch
    End Select
End Sub

Private Sub SrSrch_Change()
    SrchC.Default = True
End Sub

Private Sub SrSrch_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If Opt(2).Value = True Then
        Character = Chr(KeyAscii)
        KeyAscii = Asc(UCase(Character))
        
        If Chr(KeyAscii) <> vbBack Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Then
                DoEvents
            Else
                KeyAscii = 0
            End If
        End If
    End If
    If Opt(1).Value = True Then
        Character = Chr(KeyAscii)
        KeyAscii = Asc(UCase(Character))
        
        If Chr(KeyAscii) <> vbBack Then
            If (KeyAscii >= 48 And KeyAscii <= 57) Then
                DoEvents
            Else
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub SrSrch_LostFocus()
    SrchC.Default = False
End Sub

Private Sub sTimer_Timer()
    On Error Resume Next
    UpdateData
End Sub

Function UpdateData()
    On Error Resume Next
    Dim Count As Integer, dPos As Long, Count2 As Integer, pData As String
    Dim tDat As String, t As Byte, Pos As Integer, Ltemp As Byte
    
    If bChange = True Then VScroll.Value = Int(cPos / 10000)
    

    
    For Count = 1 To 2500
        dPos = Count + cPos
        DoEvents
        t = MemoryReader.ReadByte(dPos)
        sData(Count) = t
    Next Count
    
    HexDisp.Cls
    
    If HexMode = True Then
    
        For Count = 1 To 50
            pData = ""
            Pos = (Count - 1) * 50
            For Count2 = 1 To 50
                Pos = Pos + 1
                tDat = Hex(sData(Pos))
                If Len(tDat) = 1 Then tDat = "0" & tDat
                pData = pData & tDat & " "
            Next Count2
            HexDisp.ForeColor = wColour
            HexDisp.Print pData
        Next Count
    
    Else
    
        For Count = 1 To 50
            pData = ""
            Pos = (Count - 1) * 50
            For Count2 = 1 To 50
                Pos = Pos + 1
                Ltemp = sData(Pos)
                If Ltemp = 0 Or Ltemp = 13 Or Ltemp = 10 Or Ltemp = 9 Then
                    tDat = " "
                Else
                    tDat = Chr(Ltemp)
                End If
                pData = pData & tDat & "  "
            Next Count2
            HexDisp.ForeColor = wColour
            HexDisp.Print pData
        Next Count
        
    End If
    
    DispVals
    
End Function

Private Sub Form_Unload(cancel As Integer)
    Dim Ans As Integer
    Ans = MsgBox("Are you sure you want to exit?", vbYesNo, "Memory Hex Editor")
    If Ans = vbYes Then
        MemoryReader.CloseHandle myHandle
        Unload Me
        End
    Else
        cancel = 1
    End If
End Sub

Function GetSet()
    On Error Resume Next
    Dim Count As Integer, Strv As String, tVal As String
    For Count = 0 To 49
        tVal = Count
        If Len(tVal) = 1 Then tVal = " " & tVal
        Strv = Strv & tVal & " "
    Next Count
    Edit.BackColor = &H800000
    Edit.ForeColor = vbWhite
    dValues.Print Strv
    cPos = 1
    HexMode = True
    bChange = False
    bCol = False
    UpdateData
    Edit.Width = (HexDisp.Width / 50) - 50
    Edit.Height = (HexDisp.Height / 50) + 25
End Function

Private Sub VScroll_Change()
    On Error Resume Next
    Dim Temp As Long
    If sTemp = False Then
        Exit Sub
    Else
        Edit.Visible = False
        Temp = VScroll.Value
        cPos = (Temp * 10000) + 1
        If cPos > Max Then cPos = Max
        Value.Cls
        Value.Print cPos - 1
        UpdateData
    End If
End Sub

Private Sub VScroll_GotFocus()
    sTemp = True
End Sub

Private Sub VScroll_LostFocus()
    sTemp = False
End Sub

Private Sub VScroll_Scroll()
    On Error Resume Next
    Dim Temp As Long
    Edit.Visible = False
    Temp = VScroll.Value
    cPos = (Temp * 10000) + 1
    If cPos > Max Then cPos = Max
    Value.Cls
    Value.Print cPos - 1
    UpdateData
End Sub

Private Sub sDir_Click(Index As Integer)
    On Error Resume Next
    Edit.Visible = False
    Select Case Index
    Case 0
        cPos = 1
    Case 1
        cPos = cPos - 2500
    Case 2
        cPos = cPos - 50
    Case 3
        cPos = cPos + 50
    Case 4
        cPos = cPos + 2500
    Case 5
        cPos = Max
    End Select
    If cPos > Max Then cPos = Max
    If cPos < 1 Then cPos = 1
    VScroll.Value = Int(cPos / 10000)
    Value.Cls
    Value.Print cPos - 1
    UpdateData
End Sub

Private Sub Form_Load()
    GetSet
End Sub

Function wColour() As ColorConstants
    If bCol = True Then
        bCol = False
        wColour = vbBlue
    Else
        bCol = True
        wColour = vbRed
    End If
End Function

Function DispVals()
    On Error Resume Next
    Dim Count As Integer, Val As Long
    sValues.Cls
    For Count = 1 To 50
        Val = ((Count - 1) * 50) + cPos - 1
        sValues.Print Val
    Next
End Function

