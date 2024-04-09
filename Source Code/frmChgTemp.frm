VERSION 5.00
Begin VB.Form frmChgTemp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Template"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Change To:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Current:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmChgTemp"
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

Dim filePaths(20) As String

Private Sub Command1_Click()
    Dim configFile As String
    Dim OPENFILE As String
    Dim FILE2 As Object
    Dim FSRM As Object
    
    Set FSRM = CreateObject("Scripting.FileSystemObject")
    
    OPENFILE = App.Path & "\" & frmMain.CONFIGDIR & "\" & frmMain.configFile

    Set FILE2 = FSRM.OpenTextFile(OPENFILE, 2, False)

    configFile = "####################" & vbNewLine
    configFile = configFile & "#  ERIC MACDONALD  #" & vbNewLine
    configFile = configFile & "####################" & vbNewLine & vbNewLine
    configFile = configFile & "headers = " & Chr(34) & frmMain.HEADERSP & Chr(34) & vbNewLine
    configFile = configFile & "template = " & Chr(34) & filePaths(Combo1.ListIndex) & Chr(34) & vbNewLine
    configFile = configFile & "presets = " & Chr(34) & frmMain.PRESETSP & Chr(34) & vbNewLine
    configFile = configFile & "emulator = " & Chr(34) & frmMain.EMULP & Chr(34) & vbNewLine

    FILE2.Write (configFile)
    FILE2.Close
    
    Command1.Enabled = False
    Call frmMain.cmdRestart_Click
    Command1.Enabled = True
    
    Unload frmChgTemp
    
End Sub

Private Sub Form_Load()
    Dim TEMPD As String
    Dim TEMPD2 As String
    Dim tempFile As String
    Dim fileExt As String
    Dim OPENFILE As String
    Dim FILE2 As Object
    Dim FSRM As Object
    Dim TEMPLATE As String
    Dim fileLine As String
    Dim onItem As String
    
    Set FSRM = CreateObject("Scripting.FileSystemObject")

    Label3.Caption = frmMain.TEMPLATE

    TEMPD = frmMain.CONFIGDIR & "\" & frmMain.TEMPLATEP
    
    TEMPD = Left(TEMPD, InStrRev(TEMPD, "\"))
    TEMPD2 = Left(frmMain.TEMPLATEP, InStrRev(frmMain.TEMPLATEP, "\"))
    
    tempFile = Dir(TEMPD)
    
    onItem = 0
    
    Combo1.Clear
    
    Do While tempFile <> ""   ' Start the loop.
        If tempFile <> "." And tempFile <> ".." Then
            
            If (GetAttr(TEMPD & tempFile) And vbDirectory) <> vbDirectory Then
                fileExt = Mid(tempFile, InStrRev(tempFile, ".") + 1)
                If fileExt = "tmpl" Then
    
                    OPENFILE = App.Path & "\" & TEMPD & "\" & tempFile
                    Set FILE2 = FSRM.OpenTextFile(OPENFILE, 1, False)
                    TEMPLATE = tempFile
                    Do While FILE2.AtEndOfStream <> True
                    fileLine = FILE2.ReadLine
        
                        If checkRE(fileLine, "^\-\>(.*)$") <> 0 Then
                            TEMPLATE = REMatch.SubMatches(0)
                        End If
                    Loop
                    Combo1.AddItem TEMPLATE
                    filePaths(onItem) = TEMPD2 & tempFile
                    onItem = onItem + 1
                
                    FILE2.Close
                End If
            End If
        End If
        
        tempFile = Dir
    Loop
    Combo1.ListIndex = 0
End Sub
