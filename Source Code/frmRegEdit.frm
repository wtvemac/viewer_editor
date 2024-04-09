VERSION 5.00
Begin VB.Form frmRegEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registry Edits"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Viewer"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Modem"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sound"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "frmRegEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PhoneEn As String
Dim SoundEn As String
Public dontEdit As Integer


Private Sub Command1_Click()
    Dim theKeys() As String
    Dim theVals() As String
    Dim i As Integer

    
    theKeys = GetAllKeys(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator")
    
    For i = 0 To UBound(theKeys)
        Call DeleteKey(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\" & theKeys(i))
    Next i
    
    Call DeleteKey(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator")
    
    Call disableRegedit
    
    MsgBox "The viewer(s) was reset to all fresh-install values.", vbInformation, "Yo"
End Sub
Private Sub SaveSetting()
    If dontEdit = 0 Then
        Call SaveSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\PrimeTime1.1", "Use Phone", Check2.Value)
        Call SaveSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\PrimeTime1.1", "Sound Enabled", Check1.Value)
    
        Call SaveSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\Viewer2.5", "Use Phone", Check2.Value)
        Call SaveSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\Viewer2.5", "Sound Enabled", Check1.Value)
    End If
End Sub

Private Sub Command2_Click()
    Call SaveSetting
    
    Unload frmRegEdit
End Sub


Private Sub Form_Load()
    
    If dontEdit = 0 Then
        Check1.Enabled = True
        Check2.Enabled = True
        
        SoundEn = GetSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\PrimeTime1.1", "Sound Enabled")
        PhoneEn = GetSettingLong(HKEY_CURRENT_USER, "Software\WebTV Networks\WinSimulator\PrimeTime1.1", "Use Phone")

        Check1.Value = SoundEn
        Check2.Value = PhoneEn
        
        Label1.Caption = ""

    Else
        Call disableRegedit
    End If
    
        
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting
End Sub

Public Sub disableRegedit()
    Check1.Enabled = False
    Check2.Enabled = False
    
    dontEdit = 1
    
    Label1.Caption = "You must run the viewer to reenable this regedit"
End Sub
