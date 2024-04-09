VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks goes to VirusOmega, ShadowMafia, MattMan69 and Outatyme."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   120
      Picture         =   "frmAbout.frx":013F
      Top             =   3480
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   5370
      Left            =   -120
      Picture         =   "frmAbout.frx":2599
      Top             =   0
      Width           =   5490
   End
End
Attribute VB_Name = "frmAbout"
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


