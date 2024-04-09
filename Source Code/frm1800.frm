VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm1800 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSID to 1800 Username and Password Converter"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frm1800.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "One SSID"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Bulk List"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command3(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Index           =   1
         Left            =   -69720
         TabIndex        =   12
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   -69720
         TabIndex        =   11
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Convert"
         Height          =   615
         Left            =   -73800
         TabIndex        =   10
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -74400
         TabIndex        =   9
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -74400
         TabIndex        =   8
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Convert"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "I'm looking for a SSID list in a file that is new line delimited."
         Height          =   375
         Left            =   -74400
         TabIndex        =   15
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label5 
         Caption         =   "Output File:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "SSID Input File:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "SSID:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm1800"
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

'
' Convert single
'
Private Sub Command1_Click()
    Dim theSSID As String
    Dim theUN As String
    Dim thePass As String
    
    theSSID = Text1.Text
    theUN = Text2.Text
    
    If Len(theSSID) = 16 Then
        setA = Mid(theSSID, 1, 2)
        setB = Mid(theSSID, 3, 4)
        setC = Mid(theSSID, 7, 2)
        setD = Mid(theSSID, 9, 2)
    
        setE = Mid(theSSID, 11, 4)
        setF = Mid(theSSID, 15, 2)
        
        theUN = "wtv_" & setD & setB & setC & "0"
        thePass = ComputeFCS(setD & setB & setC & "0")
    Else
        If Len(theUN) = 13 Then
            setA = "00"
            setB = Mid(theUN, 7, 4)
            setC = Mid(theUN, 11, 2)
            setD = Mid(theUN, 5, 2)
            setE = "0000"
            setF = "00"
        
            thePass = ComputeFCS(setD & setB & setC & "0")
            
            If theSSID = "" Then
                theSSID = setA & setB & setC & setD & setE & setF
            End If
        
        Else
            MsgBox "I'm sorry but I can't convert anything.  The SSID must be 16 digits long or the un must be 13 characters (including the 'wtv_')", vbCritical, "You fucked up"
        End If
        
    End If
    
    
    Text1.Text = theSSID
    Text2.Text = theUN
    Text3.Text = thePass

    
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    
    Dim FILE As Integer
    Dim theSSID As String
    Dim theFile As String
    Dim theUN As String
    Dim thePass As String
    Dim count As Integer
    Dim WshShell As Object
    

    
    theFile = ""
    
    FILE = FreeFile()
    Open Text4.Text For Input As #FILE
    
    If Err <> 0 Then
        MsgBox "OpenSSIDFile(): I couldn't open the file '" & Text4.Text & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        End
    Else
        
        Do Until EOF(FILE)
            Line Input #FILE, theSSID
        
            If Len(theSSID) = 16 Then
                setB = Mid(theSSID, 3, 4)
                setC = Mid(theSSID, 7, 2)
                setD = Mid(theSSID, 9, 2)
        
                theUN = "wtv_" & setD & setB & setC & "0"
                thePass = ComputeFCS(setD & setB & setC & "0")
                
                count = count + 1
                theFile = theFile & vbNewLine & vbNewLine & theSSID & vbNewLine & "Username: " & theUN & " Pass: " & thePass
            
            End If
        Loop
            
    
    Close FILE
    
    Open Text5.Text For Append As #FILE
        
        If Err <> 0 Then
            MsgBox "writeSSID(): I couldn't open the file '" & Text5.Text & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
            End
        Else
        
            Print #FILE, theFile
            Close #FILE
            
            If MsgBox("Conversion Complete!" & vbNewLine & vbNewLine & "I converted " & count & " SSID(s)" & vbNewLine & vbNewLine & "Open file in notepad?", vbOKCancel, "Whooohooo") = vbOK Then
                Set WshShell = CreateObject("WScript.Shell")
    
                WshShell.Run "notepad " & Text5.Text
            End If
        End If

End If
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim FileName As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist
    CommonDialog1.Filter = "SSID List files (*.TXT)|*.TXT"
    
    If Index = 0 Then
        CommonDialog1.ShowOpen
    Else
        CommonDialog1.ShowSave
    End If
    
    FileName = CommonDialog1.FileName
    
    If FileName <> "" Then
        If Index = 0 Then
            Text4.Text = FileName
        Else
            Text5.Text = FileName
        End If
    End If
End Sub
