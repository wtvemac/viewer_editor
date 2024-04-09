VERSION 5.00
Begin VB.Form frmPEInfo 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Portable Executable Information"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5910
   ClipControls    =   0   'False
   Icon            =   "frmPEInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   3975
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Executable:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPEInfo"
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

Private Sub Command1_Click()
    Unload frmPEInfo
End Sub

Private Sub Form_Load()
    Text1.Text = frmMain.frmPEEXE
    Text2.Text = frmMain.frmPENfo
End Sub
