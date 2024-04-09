VERSION 5.00
Begin VB.Form frmSysConfig 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Configuration"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Commit"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   7200
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   7
      ItemData        =   "frmSysConfig.frx":0000
      Left            =   480
      List            =   "frmSysConfig.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6720
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   6
      ItemData        =   "frmSysConfig.frx":00D0
      Left            =   480
      List            =   "frmSysConfig.frx":0104
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5880
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   5
      ItemData        =   "frmSysConfig.frx":013E
      Left            =   480
      List            =   "frmSysConfig.frx":0172
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5040
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   4
      ItemData        =   "frmSysConfig.frx":0396
      Left            =   480
      List            =   "frmSysConfig.frx":03CA
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4200
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   3
      ItemData        =   "frmSysConfig.frx":0479
      Left            =   480
      List            =   "frmSysConfig.frx":04AD
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3360
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   2
      ItemData        =   "frmSysConfig.frx":052E
      Left            =   480
      List            =   "frmSysConfig.frx":0562
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      ItemData        =   "frmSysConfig.frx":065A
      Left            =   480
      List            =   "frmSysConfig.frx":068E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   0
      ItemData        =   "frmSysConfig.frx":0916
      Left            =   480
      List            =   "frmSysConfig.frx":094A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Caption         =   "A B C D E F G H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "Board Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Caption         =   "Board Revision:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Caption         =   "Video:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      Caption         =   "CPU:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Audio:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "SGRAM Speed:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "ROM Bank 1:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "ROM Bank 0:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmSysConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''
' WebTV IPE (In-place Edit) 4.0 '
'                               '
' By: Eric MacDonald            '
' Date: April 24, 2005          '
'                               '
' This is a patcher tool        '
' for any SuperViewer template  '
'''''''''''''''''''''''''''''''''

Option Explicit

Public indexItem As Integer

Public Sub CalcSYSCFG(theSysConfig As Long)
    Dim theSysConfigh As String
    Dim i As Integer, j As Integer
    Label9.Caption = ""
    
    theSysConfigh = Hex(theSysConfig)
    
    j = 7
    For i = Len(theSysConfigh) + 1 To 2 Step -1
        Combo1(j).ListIndex = CInt("&H" & Mid(theSysConfigh, i - 1, 1))
        j = j - 1
    Next i
    
    Call Combo1_Click(1)
End Sub


Private Sub Combo1_Click(Index As Integer)
    Dim i As Integer
    
    Label9.Caption = ""
    For i = 0 To 7
        If Combo1(i).ListIndex <> -1 Then
            Label9.Caption = Label9.Caption & " " & Hex(Combo1(i).ListIndex)
        Else
            Label9.Caption = Label9.Caption & " 0"
        End If
    Next i
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim theSysConfigh As String
    
    theSysConfigh = ""
    
    For i = 0 To 7
        If Combo1(i).ListIndex <> -1 Then
            theSysConfigh = theSysConfigh & Hex(Combo1(i).ListIndex)
        End If
    Next i

    frmMain.cmboHeaders(indexItem).Text = CLng("&H" & theSysConfigh)
    frmSysConfig.Hide

End Sub
