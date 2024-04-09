VERSION 5.00
Begin VB.Form frmSSID 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SSID"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3720
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "R"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmSSID.frx":0000
      Left            =   1440
      List            =   "frmSSID.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "R"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00800000&
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      ItemData        =   "frmSSID.frx":00D0
      Left            =   1440
      List            =   "frmSSID.frx":00DA
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdChgSSID 
      Caption         =   "Commit"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400000&
      Caption         =   "sprintf(password,""%d"",computefcs(serial_number));"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      Caption         =   "sprintf(username, ""wtv_%s"", serial_number);"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      Caption         =   "Serial:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      Caption         =   "Box:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      Caption         =   "Set F:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      Caption         =   "Set E:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      Caption         =   "Set D:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      Caption         =   "Set C:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      Caption         =   "Set B:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      Caption         =   "Set A:"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblTheSSID 
      BackColor       =   &H00400000&
      Caption         =   "Unknown?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSSID"
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

Dim setA As String
Dim setB As String
Dim setC As String
Dim setD As String
Dim setE As String
Dim setF As String
Public indexItem As Integer


Public Sub CalcSSID(theSSID As String)
    Dim cp(9) As Byte
    Dim serial_number As String
    Dim i As Long
    Dim theFCS As Long
    
    setA = Mid(theSSID, 1, 2)
    setB = Mid(theSSID, 3, 4)
    setC = Mid(theSSID, 7, 2)
    setD = Mid(theSSID, 9, 2)
    
    setE = Mid(theSSID, 11, 4)
    setF = Mid(theSSID, 15, 2)
    
    If setA = "81" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    End If
    
    Text2.Text = setB
    
    Text3.Text = setC
    
    Select Case setD
        Case "00":
            Combo2.ListIndex = 0
        Case "10":
            Combo2.ListIndex = 1
        Case "20":
            Combo2.ListIndex = 2
        Case "30":
            Combo2.ListIndex = 3
        Case "40":
            Combo2.ListIndex = 4
        Case "50":
            Combo2.ListIndex = 5
        Case "60":
            Combo2.ListIndex = 6
        Case "70":
            Combo2.ListIndex = 7
        Case "90":
            Combo2.ListIndex = 8
        Case "F0":
            Combo2.ListIndex = 9
        Case Else
            Combo2.ListIndex = 0
    End Select
    
    Text4.Text = setE
    
    Text5.Text = setF
    
    Text1.Text = "wtv_" & setD & setB & setC & "0"
   
   theFCS = ComputeFCS(setD & setB & setC & "0")
    
    Text6.Text = theFCS
    lblTheSSID.Caption = setA & " " & setB & " " & setC & " " & setD & " " & setE & " " & setF & " [" & ValSm(4, Hex(theFCS)) & "]"

    
End Sub



Private Sub cmdChgSSID_Click()
    frmMain.cmboHeaders(indexItem).Text = setA & setB & setC & setD & setE & setF
    frmSSID.Hide
End Sub

Private Sub Command1_Click()
    Dim isetC As Integer
    
    setC = Text3.Text
    
    isetC = CInt("&H" & setC)
    
    isetC = isetC + 1
    
    If isetC > 255 Then isetC = 0
    
    setC = Hex(isetC)
    
    If Len(setC) <> 2 Then setC = "0" & setC

    Text3.Text = LCase(setC)
End Sub

Private Sub Command2_Click()
    Dim isetC As Integer
    
    setC = Text3.Text
    
    isetC = CInt("&H" & setC)
    
    isetC = isetC - 1
    
    If isetC < 0 Then isetC = 255
    
    setC = Hex(isetC)
    
    If Len(setC) <> 2 Then setC = "0" & setC

    Text3.Text = LCase(setC)
End Sub

Private Sub Command3_Click()
    Dim isetB As Long
    Dim i As Integer
    
    setB = Text2.Text
    
    isetB = CLng(Rnd * 65535)
    
    setB = LCase(Hex(isetB))
    
    For i = Len(setB) To 3
        setB = "0" & setB
    Next i
    
    Text2.Text = setB
End Sub

Private Sub Command4_Click()
    Dim isetF As Integer
    
    setF = Text5.Text
    
    isetF = CInt(Rnd * 255)
    
    setF = LCase(Hex(isetF))
    
    If Len(setF) <> 2 Then setF = "0" & setF
    
    Text5.Text = setF
End Sub


Public Sub updateSSID()
    Call CalcSSID(setA & setB & setC & setD & setE & setF)
End Sub

Private Sub Text2_Change()
    If Len(Text2.Text) = 4 Then
    
        setB = ValSm(4, Text2.Text)
        Call updateSSID
    End If
End Sub

Private Sub Text3_Change()
    If Len(Text3.Text) = 2 Then
        setC = ValSm(2, Text3.Text)
        Call updateSSID
    End If
End Sub

Private Sub Combo1_Click()
    
    Select Case Combo1.ListIndex
        Case 0:
            setA = "01"
        Case 1:
            setA = "81"
    End Select
    
    Call updateSSID
End Sub


Private Sub Combo2_Click()
    
    Select Case Combo2.ListIndex
        Case 0:
            setD = "00"
        Case 1:
            setD = "10"
        Case 2:
            setD = "20"
        Case 3:
            setD = "30"
        Case 4:
            setD = "40"
        Case 5:
            setD = "50"
        Case 6:
            setD = "60"
        Case 7:
            setD = "70"
        Case 8:
            setD = "90"
        Case 9:
            setD = "F0"
    End Select
    
    Call updateSSID
End Sub

Private Function ValSm(tempLen As Integer, tempSt As String)
    Dim i As Integer
    
    If Len(tempSt) > tempLen Then
        tempSt = Right(tempSt, tempLen)
    End If
    
    For i = Len(tempSt) To (tempLen - 1)
        tempSt = "0" & tempSt
    Next i
    
    ValSm = tempSt
End Function

Private Sub Text4_Change()
    If Len(Text4.Text) = 4 Then
        setE = ValSm(4, Text4.Text)
        Call updateSSID
    End If
End Sub


Private Sub Text5_Change()
    If Len(Text5.Text) = 2 Then
        setF = ValSm(2, Text5.Text)
        Call updateSSID
    End If
End Sub

