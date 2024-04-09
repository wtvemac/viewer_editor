VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The SuperViewer IPE 4.0"
   ClientHeight    =   6405
   ClientLeft      =   2910
   ClientTop       =   2835
   ClientWidth     =   9405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Information"
      TabPicture(0)   =   "Form1.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Edit SV"
      TabPicture(1)   =   "Form1.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Create SV"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblCrtBP"
      Tab(2).Control(1)=   "chkCrtBP"
      Tab(2).Control(2)=   "Check4"
      Tab(2).Control(3)=   "Check3"
      Tab(2).Control(4)=   "Check1"
      Tab(2).Control(5)=   "Command1"
      Tab(2).Control(6)=   "lstCrtBP"
      Tab(2).Control(7)=   "chkLEmu"
      Tab(2).ControlCount=   8
      Begin VB.CheckBox chkLEmu 
         Caption         =   "Launch Emulator"
         Height          =   255
         Left            =   -74520
         TabIndex        =   131
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other Information"
         Height          =   2055
         Left            =   4680
         TabIndex        =   16
         Top             =   480
         Width           =   4215
         Begin VB.CommandButton Command5 
            Caption         =   "Download"
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdRestart 
            Caption         =   "Reload Configuration"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   3975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Check for update"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Agreement"
         Height          =   3255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   8895
         Begin VB.TextBox txtGPrinc 
            Height          =   2895
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Base Information"
         Height          =   2175
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   4335
         Begin VB.TextBox txtInfo_acthash 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            TabIndex        =   24
            Text            =   "Unknown?"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtInfo_thehash 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            TabIndex        =   23
            Text            =   "Unknown?"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txtInfo_pathto 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            TabIndex        =   22
            Text            =   "Unknown?"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtInfo_viewervers 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            TabIndex        =   21
            Text            =   "Unknown?"
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txtInfoTemp 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1560
            TabIndex        =   19
            Text            =   "Unknown?"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label5 
            Caption         =   "Expected Viewer:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Template:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Viewer Path:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Original MD5:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Have MD5:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.ListBox lstCrtBP 
         Height          =   1425
         ItemData        =   "Form1.frx":0342
         Left            =   -70080
         List            =   "Form1.frx":0344
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create SuperViewer"
         Height          =   495
         Left            =   -73800
         TabIndex        =   6
         Top             =   3240
         Width           =   6495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Validate using checksum"
         Height          =   255
         Left            =   -74520
         TabIndex        =   5
         Top             =   780
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Make backup"
         Height          =   255
         Left            =   -74520
         TabIndex        =   4
         Top             =   1380
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Launch vewer"
         Height          =   255
         Left            =   -74520
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkCrtBP 
         Caption         =   "Use box presets"
         Height          =   255
         Left            =   -72120
         TabIndex        =   2
         Top             =   780
         Width           =   1815
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5595
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9869
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "Headers"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label8"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "imgMoreEdit"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label9"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "VScroll1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Command3"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "frmScroll"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Combo1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Command4"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Command6"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         Begin VB.CommandButton Command6 
            Caption         =   "Add"
            Height          =   255
            Left            =   8040
            TabIndex        =   129
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Commit"
            Height          =   255
            Left            =   7200
            TabIndex        =   128
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":0346
            Left            =   3000
            List            =   "Form1.frx":034D
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   480
            Width           =   4095
         End
         Begin VB.Frame frmScroll 
            BackColor       =   &H00400000&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   4095
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   8415
            Begin VB.Frame Frame4 
               BackColor       =   &H00400000&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   11640
               Left            =   0
               TabIndex        =   32
               Top             =   0
               Width           =   8415
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   30
                  Left            =   7800
                  TabIndex        =   29
                  Top             =   14520
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   30
                  Left            =   3480
                  TabIndex        =   123
                  Text            =   "Unknown?"
                  Top             =   14520
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   29
                  Left            =   7800
                  TabIndex        =   121
                  Top             =   14040
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   29
                  Left            =   3480
                  TabIndex        =   120
                  Text            =   "Unknown?"
                  Top             =   14040
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   28
                  Left            =   7800
                  TabIndex        =   118
                  Top             =   13560
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   28
                  Left            =   3480
                  TabIndex        =   117
                  Text            =   "Unknown?"
                  Top             =   13560
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   27
                  Left            =   7800
                  TabIndex        =   115
                  Top             =   13080
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   27
                  Left            =   3480
                  TabIndex        =   114
                  Text            =   "Unknown?"
                  Top             =   13080
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   26
                  Left            =   7800
                  TabIndex        =   112
                  Top             =   12600
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   26
                  Left            =   3480
                  TabIndex        =   111
                  Text            =   "Unknown?"
                  Top             =   12600
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   25
                  Left            =   7800
                  TabIndex        =   109
                  Top             =   12120
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   25
                  Left            =   3480
                  TabIndex        =   108
                  Text            =   "Unknown?"
                  Top             =   12120
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   24
                  Left            =   7800
                  TabIndex        =   106
                  Top             =   11640
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   24
                  Left            =   3480
                  TabIndex        =   105
                  Text            =   "Unknown?"
                  Top             =   11640
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   23
                  Left            =   7800
                  TabIndex        =   103
                  Top             =   11160
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   23
                  Left            =   3480
                  TabIndex        =   102
                  Text            =   "Unknown?"
                  Top             =   11160
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   22
                  Left            =   7800
                  TabIndex        =   100
                  Top             =   10680
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   22
                  Left            =   3480
                  TabIndex        =   99
                  Text            =   "Unknown?"
                  Top             =   10680
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   21
                  Left            =   7800
                  TabIndex        =   97
                  Top             =   10200
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   21
                  Left            =   3480
                  TabIndex        =   96
                  Text            =   "Unknown?"
                  Top             =   10200
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   20
                  Left            =   7800
                  TabIndex        =   94
                  Top             =   9720
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   20
                  Left            =   3480
                  TabIndex        =   93
                  Text            =   "Unknown?"
                  Top             =   9720
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   19
                  Left            =   7800
                  TabIndex        =   91
                  Top             =   9240
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   19
                  Left            =   3480
                  TabIndex        =   90
                  Text            =   "Unknown?"
                  Top             =   9240
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   18
                  Left            =   7800
                  TabIndex        =   88
                  Top             =   8760
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   18
                  Left            =   3480
                  TabIndex        =   87
                  Text            =   "Unknown?"
                  Top             =   8760
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   17
                  Left            =   7800
                  TabIndex        =   85
                  Top             =   8280
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   17
                  Left            =   3480
                  TabIndex        =   84
                  Text            =   "Unknown?"
                  Top             =   8280
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   16
                  Left            =   7800
                  TabIndex        =   82
                  Top             =   7800
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   16
                  Left            =   3480
                  TabIndex        =   81
                  Text            =   "Unknown?"
                  Top             =   7800
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   15
                  Left            =   7800
                  TabIndex        =   79
                  Top             =   7320
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   15
                  Left            =   3480
                  TabIndex        =   78
                  Text            =   "Unknown?"
                  Top             =   7320
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   14
                  Left            =   7800
                  TabIndex        =   76
                  Top             =   6840
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   14
                  Left            =   3480
                  TabIndex        =   75
                  Text            =   "Unknown?"
                  Top             =   6840
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   13
                  Left            =   7800
                  TabIndex        =   73
                  Top             =   6360
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   13
                  Left            =   3480
                  TabIndex        =   72
                  Top             =   6360
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   12
                  Left            =   7800
                  TabIndex        =   70
                  Top             =   5880
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   12
                  Left            =   3480
                  TabIndex        =   69
                  Text            =   "Unknown?"
                  Top             =   5880
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   11
                  Left            =   7800
                  TabIndex        =   67
                  Top             =   5400
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   11
                  Left            =   3480
                  TabIndex        =   66
                  Text            =   "Unknown?"
                  Top             =   5400
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   10
                  Left            =   7800
                  TabIndex        =   64
                  Top             =   4920
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   10
                  Left            =   3480
                  TabIndex        =   63
                  Text            =   "Unknown?"
                  Top             =   4920
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   9
                  Left            =   7800
                  TabIndex        =   61
                  Top             =   4440
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   9
                  Left            =   3480
                  TabIndex        =   60
                  Text            =   "Unknown?"
                  Top             =   4440
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   8
                  Left            =   7800
                  TabIndex        =   58
                  Top             =   3960
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   8
                  Left            =   3480
                  TabIndex        =   57
                  Text            =   "Unknown?"
                  Top             =   3960
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   7
                  Left            =   7800
                  TabIndex        =   55
                  Top             =   3480
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   7
                  Left            =   3480
                  TabIndex        =   54
                  Text            =   "Unknown?"
                  Top             =   3480
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   6
                  Left            =   7800
                  TabIndex        =   52
                  Top             =   3000
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   6
                  Left            =   3480
                  TabIndex        =   51
                  Text            =   "Unknown?"
                  Top             =   3000
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   5
                  Left            =   7800
                  TabIndex        =   49
                  Top             =   2520
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   5
                  Left            =   3480
                  TabIndex        =   48
                  Text            =   "Unknown?"
                  Top             =   2520
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   4
                  Left            =   7800
                  TabIndex        =   46
                  Top             =   2040
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   4
                  Left            =   3480
                  TabIndex        =   45
                  Text            =   "Unknown?"
                  Top             =   2040
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   3
                  Left            =   7800
                  TabIndex        =   43
                  Top             =   1560
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   3
                  Left            =   3480
                  TabIndex        =   42
                  Text            =   "Unknown?"
                  Top             =   1560
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   2
                  Left            =   7800
                  TabIndex        =   40
                  Top             =   1080
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   2
                  Left            =   3480
                  TabIndex        =   39
                  Text            =   "Unknown?"
                  Top             =   1080
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   1
                  Left            =   7800
                  TabIndex        =   37
                  Top             =   600
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   36
                  Text            =   "Unknown?"
                  Top             =   600
                  Width           =   3975
               End
               Begin VB.CheckBox chkDelHeaders 
                  BackColor       =   &H00400000&
                  Caption         =   "Check2"
                  Height          =   255
                  Index           =   0
                  Left            =   7800
                  TabIndex        =   34
                  Top             =   120
                  Width           =   255
               End
               Begin VB.ComboBox cmboHeaders 
                  BackColor       =   &H00800000&
                  ForeColor       =   &H00E0E0E0&
                  Height          =   315
                  Index           =   0
                  Left            =   3480
                  TabIndex        =   33
                  Text            =   "Unknown?"
                  Top             =   120
                  Width           =   3975
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   30
                  Left            =   120
                  TabIndex        =   30
                  Top             =   14520
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   29
                  Left            =   120
                  TabIndex        =   122
                  Top             =   14040
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   28
                  Left            =   120
                  TabIndex        =   119
                  Top             =   13560
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   27
                  Left            =   120
                  TabIndex        =   116
                  Top             =   13080
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   26
                  Left            =   120
                  TabIndex        =   113
                  Top             =   12600
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   25
                  Left            =   120
                  TabIndex        =   110
                  Top             =   12120
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   24
                  Left            =   120
                  TabIndex        =   107
                  Top             =   11640
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   23
                  Left            =   120
                  TabIndex        =   104
                  Top             =   11160
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   22
                  Left            =   120
                  TabIndex        =   101
                  Top             =   10680
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   21
                  Left            =   120
                  TabIndex        =   98
                  Top             =   10200
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   20
                  Left            =   120
                  TabIndex        =   95
                  Top             =   9720
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   19
                  Left            =   120
                  TabIndex        =   92
                  Top             =   9240
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   18
                  Left            =   120
                  TabIndex        =   89
                  Top             =   8760
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   17
                  Left            =   120
                  TabIndex        =   86
                  Top             =   8280
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   16
                  Left            =   120
                  TabIndex        =   83
                  Top             =   7800
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   15
                  Left            =   120
                  TabIndex        =   80
                  Top             =   7320
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   14
                  Left            =   120
                  TabIndex        =   77
                  Top             =   6840
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   13
                  Left            =   120
                  TabIndex        =   74
                  Top             =   6360
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   71
                  Top             =   5880
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   11
                  Left            =   120
                  TabIndex        =   68
                  Top             =   5400
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   10
                  Left            =   120
                  TabIndex        =   65
                  Top             =   4920
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   9
                  Left            =   120
                  TabIndex        =   62
                  Top             =   4440
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   59
                  Top             =   3960
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   56
                  Top             =   3480
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   53
                  Top             =   3000
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   50
                  Top             =   2520
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   47
                  Top             =   2040
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   44
                  Top             =   1560
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   41
                  Top             =   1080
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   38
                  Top             =   600
                  Width           =   3135
               End
               Begin VB.Label lblHeaders 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unknown?"
                  ForeColor       =   &H00E0E0E0&
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   35
                  Top             =   120
                  Width           =   3135
               End
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Delete"
            Height          =   255
            Left            =   7680
            TabIndex        =   25
            Top             =   840
            Width           =   735
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   4095
            Left            =   8520
            TabIndex        =   125
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "Presets:"
            Height          =   255
            Left            =   2280
            TabIndex        =   124
            Top             =   480
            Width           =   615
         End
         Begin VB.Image imgMoreEdit 
            Height          =   240
            Left            =   8520
            Picture         =   "Form1.frx":0359
            Top             =   5280
            Width           =   240
         End
         Begin VB.Label Label8 
            Caption         =   "Note: To add a new header right click within the header frame."
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   5280
            Width           =   8415
         End
         Begin VB.Label Label7 
            Caption         =   "Header Object Value"
            Height          =   255
            Left            =   3840
            TabIndex        =   27
            Top             =   960
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   "Header Object Name"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   3615
         End
      End
      Begin VB.Label lblCrtBP 
         Height          =   255
         Left            =   -70080
         TabIndex        =   126
         Top             =   2520
         Width           =   4095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuIPEC 
         Caption         =   "1800 User_Pass Gen"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuChgTemp 
         Caption         =   "Change Template"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuRunView 
         Caption         =   "Run Viewer"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuRunEmu 
         Caption         =   "Run Emulator"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuRegEdit 
         Caption         =   "Small Regedit"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuLanRegEdit 
         Caption         =   "Registry Editor"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuExtPE 
         Caption         =   "Extended PE Info"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuRunCalc 
         Caption         =   "Run Calculator"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAmout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuAddHead 
      Caption         =   "ADD NEW HEADER"
      Visible         =   0   'False
      Begin VB.Menu mnuAddHead1 
         Caption         =   "mnuAddHead1"
      End
      Begin VB.Menu mnuAddHead2 
         Caption         =   "mnuAddHead2"
      End
      Begin VB.Menu mnuAddHead3 
         Caption         =   "mnuAddHead3"
      End
      Begin VB.Menu mnuAddHead4 
         Caption         =   "mnuAddHead4"
      End
      Begin VB.Menu mnuAddHead5 
         Caption         =   "mnuAddHead5"
      End
      Begin VB.Menu mnuAddHead6 
         Caption         =   "mnuAddHead6"
      End
      Begin VB.Menu mnuAddHead7 
         Caption         =   "mnuAddHead7"
      End
      Begin VB.Menu mnuAddHead8 
         Caption         =   "mnuAddHead8"
      End
      Begin VB.Menu mnuAddHead9 
         Caption         =   "mnuAddHead9"
      End
      Begin VB.Menu mnuAddHead10 
         Caption         =   "mnuAddHead10"
      End
      Begin VB.Menu mnuAddHead11 
         Caption         =   "mnuAddHead11"
      End
      Begin VB.Menu mnuAddHead12 
         Caption         =   "mnuAddHead12"
      End
      Begin VB.Menu mnuAddHead13 
         Caption         =   "mnuAddHead13"
      End
      Begin VB.Menu mnuAddHead14 
         Caption         =   "mnuAddHead14"
      End
      Begin VB.Menu mnuAddHead15 
         Caption         =   "mnuAddHead15"
      End
      Begin VB.Menu mnuAddHead16 
         Caption         =   "mnuAddHead16"
      End
      Begin VB.Menu mnuAddHead17 
         Caption         =   "mnuAddHead17"
      End
      Begin VB.Menu mnuAddHead18 
         Caption         =   "mnuAddHead18"
      End
      Begin VB.Menu mnuAddHead19 
         Caption         =   "mnuAddHead19"
      End
      Begin VB.Menu mnuAddHead20 
         Caption         =   "mnuAddHead20"
      End
      Begin VB.Menu mnuAddHead21 
         Caption         =   "mnuAddHead21"
      End
      Begin VB.Menu mnuAddHead22 
         Caption         =   "mnuAddHead22"
      End
      Begin VB.Menu mnuAddHead23 
         Caption         =   "mnuAddHead23"
      End
      Begin VB.Menu mnuAddHead24 
         Caption         =   "mnuAddHead24"
      End
      Begin VB.Menu mnuAddHead25 
         Caption         =   "mnuAddHead25"
      End
      Begin VB.Menu mnuAddHead26 
         Caption         =   "mnuAddHead26"
      End
      Begin VB.Menu mnuAddHead27 
         Caption         =   "mnuAddHead27"
      End
      Begin VB.Menu mnuAddHead28 
         Caption         =   "mnuAddHead28"
      End
      Begin VB.Menu mnuAddHead29 
         Caption         =   "mnuAddHead29"
      End
      Begin VB.Menu mnuAddHead30 
         Caption         =   "mnuAddHead30"
      End
      Begin VB.Menu mnuWriteHead 
         Caption         =   "Write header (commit edit)"
      End
      Begin VB.Menu mnuDelAllHeads 
         Caption         =   "Delete All"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Dim FSRM As Object
Dim FILE, FILE2 As Object
Dim VIEWER As Long
Dim VIEWER2 As Integer
Public CONFIGDIR As String, configFile As String
Dim VERSION As Integer
Public TEMPLATE As String
Public TEMPLATEP As String
Public pathto As String, viewervers As String, thehash As String
Dim boxPresets As Dictionary
Public blockVars As Dictionary
Dim hderObjs As Dictionary
Public EMULP As String, PRESETSP As String, HEADERSP As String
Public editCodes As Dictionary
Public tempPath As String
Dim tempDict As Dictionary
Dim mnuHeadCount As Integer
Dim knownHeadCount As Integer
Dim tempInt As Integer
Dim selMoreEdit As Integer
Dim presetFN As String
Dim strData As String
Dim strBody As String
Dim contentLen As Integer
Dim emulatorPL As String
Public frmPENfo As String
Public frmPEEXE As String





Private Sub chkDelHeaders_Click(Index As Integer)
    If chkDelHeaders(Index).Value <> 0 Then
        cmboHeaders(Index).Enabled = False
        cmboHeaders(Index).BackColor = &HFF&
    Else
        cmboHeaders(Index).Enabled = True
        cmboHeaders(Index).BackColor = 8388608
    End If
End Sub

'
' Show the box preset list if checked to do so.
'
Private Sub chkCrtBP_Click()
    If chkCrtBP.Value = 1 Then
        lstCrtBP.Visible = True
    Else
        lstCrtBP.Visible = False
    End If
End Sub


Private Sub cmboHeaders_Click(Index As Integer)
    If checkRE(cmboHeaders(Index).Text, "^(.*?)\s*\((.*?)\)") <> 0 Then
        cmboHeaders(Index).Text = REMatch.SubMatches(0)
    End If

End Sub

Private Sub cmboHeaders_GotFocus(Index As Integer)
    Label8.Caption = hderObjs(lblHeaders(Index).Caption)("0")
    
    imgMoreEdit.Visible = False
    selMoreEdit = -1
    Select Case lblHeaders(Index).Caption
        Case "wtv-client-serial-number":
            selMoreEdit = Index
            imgMoreEdit.Visible = True
        Case "wtv-system-sysconfig":
            selMoreEdit = Index
            imgMoreEdit.Visible = True
    End Select


End Sub



Public Sub cmdRestart_Click()
    cmdRestart.Enabled = False
    Call Form_Load
    cmdRestart.Enabled = True
End Sub







Private Sub Command1_Click()
    Dim hashCodes
    Dim hashCodes2
    Dim mode As Integer
    Dim backup As Boolean
    
    If Check1.Value = 1 Then
        If txtInfo_thehash.Text <> txtInfo_acthash.Text Then
            MsgBox "This is not an unedited viewer (hash didn't match)", vbCritical, "Shit guy"
            Exit Sub
        End If
    End If

    If chkCrtBP.Value = 1 Then
        mode = 3
        
        For Each hashCodes In boxPresets
            If lstCrtBP.List(lstCrtBP.ListIndex) = boxPresets(hashCodes)("description") Then
                For Each hashCodes2 In boxPresets(hashCodes)("header")
                    tempDict(hashCodes2) = boxPresets(hashCodes)("header")(hashCodes2)
                Next hashCodes2
            End If
        Next hashCodes
    Else
        mode = 2
        tempDict.RemoveAll
    End If
    
    If Check3.Value = 1 Then
        backup = True
    Else
        backup = False
    End If
    
    frmMain.BackColor = &HFF
    If writeHeader(tempDict, mode, backup) = 1 Then
        MsgBox "Write Successful", vbInformation, "Whoohooo"
    End If
    frmMain.BackColor = &H8000000F
    
    If chkLEmu.Value = 1 Then
        Call mnuRunEmu_Click
    End If
    
    If Check4.Value = 1 Then
        Call mnuRunView_Click
    End If
End Sub

Private Sub mnuAmout_Click()
    frmAbout.Show
End Sub

Private Sub mnuChgTemp_Click()
    frmChgTemp.Show
End Sub

Private Sub mnuExtPE_Click()
    frmPEInfo.Show
End Sub

Private Sub mnuLanRegEdit_Click()
    
    Call SaveSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Applets\Regedit", "LastKey", "HKEY_CURRENT_USER\Software\WebTV Networks")
    Shell "regedit.exe"
End Sub

Private Sub mnuRegEdit_Click()
    frmRegEdit.Show
End Sub

Private Sub mnuRunEmu_Click()
    Dim WshShell As Object
    Dim pathStr As String
    
    pathStr = App.Path & "\" & emulatorPL
        
    Set WshShell = CreateObject("WScript.Shell")
    
    WshShell.Run "perl " & Chr(34) & pathStr & Chr(34) & " " & Chr(34) & Left(pathStr, InStrRev(pathStr, "\") - 1) & Chr(34)
End Sub

Private Sub mnuWriteHead_Click()
    Dim hashItem
    Dim objCount As Integer
    Dim makeBack As Boolean
    
    tempDict.RemoveAll
    If MsgBox("Are you sure you want to write this header to the viewer?", vbOKCancel, "Yikes! brother sikes") = vbOK Then
        For Each hashItem In lblHeaders
            If cmboHeaders(objCount).Text <> "" And chkDelHeaders(objCount).Value <> 1 And hashItem.Visible = True Then
                tempDict(hashItem.Caption) = cmboHeaders(objCount).Text
            End If
            objCount = objCount + 1
        Next hashItem
        
        If MsgBox("Make Backup?", vbYesNo, "Hold up homeslice") = vbYes Then
            makeBack = True
        Else
            makeBack = False
        End If
    
        frmMain.BackColor = &HFF
        If writeHeader(tempDict, 1, makeBack) <> 1 Then
            MsgBox "Couln't write to viewer!", vbCritical, "Something's fucked up"
        Else
            MsgBox "Write successful", vbOKOnly, "Woohooo"
        End If
        frmMain.BackColor = &H8000000F
    End If
    
End Sub

Private Function getMD5()
    On Error Resume Next
    
    Dim DestinationString() As Byte
    Dim OPENFILE As String
    Dim hashItem
    Dim headerSplit
    Dim headerSplit2
    Dim headerVal As String
    Dim headerPart
    Dim fileSiz As Long
    Dim readByes As Long
    Dim lngBytesRead As Long
    Dim DestinationStringD As String
    Dim sectionName As String
    Dim crazyMZHeader As IMAGE_DOS_HEADER
    Dim crazyPEHeader1 As IMAGE_FILE_HEADER
    Dim crazyPEHeader2 As IMAGE_OPTIONAL_HEADER
    Dim crazyPESections() As IMAGE_SECTION_HEADER
    Dim i As Integer
    Dim timeX As String
    
    Err = 0
    
    OPENFILE = App.Path & "\" & pathto
    
    frmPEEXE = OPENFILE
    
    VIEWER = CreateFile(OPENFILE, ByVal &H80000000, 0, ByVal 0&, 3, 0, ByVal 0)

    If Err <> 0 Then
        MsgBox "getMD5(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
        fileSiz = GetFileSize(VIEWER, 0)
        
        frmPENfo = "File Size: " & fileSiz & " bytes" & vbNewLine & vbNewLine
        
        ReDim DestinationString(0 To fileSiz)
        
        ReadFile VIEWER, DestinationString(0), fileSiz, readByes, ByVal 0&
    
        '------------'
        ' DOS HEADER
        SetFilePointer VIEWER, ByVal 0, 0, 0
        ReadFile VIEWER, crazyMZHeader, ByVal Len(crazyMZHeader), lngBytesRead, ByVal 0&
        
        If crazyMZHeader.Magic <> 23117 Then
            MsgBox "getMD5(): The viewer that I'm trying to read is not a valid Win32 PE.  Please correct this.", vbCritical, "Something's fucked up"
            Unload frmMain
        End If
        
        ' PE Header
        SetFilePointer VIEWER, ByVal crazyMZHeader.lfanew + 4, 0, 0
        ReadFile VIEWER, crazyPEHeader1, ByVal Len(crazyPEHeader1), lngBytesRead, ByVal 0&
        ReadFile VIEWER, crazyPEHeader2, ByVal Len(crazyPEHeader2), lngBytesRead, ByVal 0&
    
        ' --Section header--
        ReDim crazyPESections(crazyPEHeader1.NumberOfSections - 1) As IMAGE_SECTION_HEADER
        ReadFile VIEWER, crazyPESections(0), ByVal Len(crazyPESections(0)) * crazyPEHeader1.NumberOfSections, lngBytesRead, ByVal 0&
        
        
        timeX = FormatELTime(crazyPEHeader1.TimeDateStamp)
        timeX = Left(timeX, Len(timeX) - 1)
        frmPENfo = frmPENfo & "PE Header Offset: 0x" & Hex(crazyMZHeader.lfanew) & vbNewLine _
                            & "Date Compiled: " & timeX _
                            & vbNewLine & "Linker version: " & crazyPEHeader2.MajorLinkVer & "." & crazyPEHeader2.MinorLinkVer & vbNewLine _
                            & "Memory Base: 0x" & Hex(crazyPEHeader2.ImageBase) & vbNewLine _
                            & "Code Base: 0x" & Hex(crazyPEHeader2.CodeBase) & vbNewLine _
                            & "EXE Size Aligned: " & crazyPEHeader2.ImageSize & " bytes" & vbNewLine _
                            & "Header Size: " & crazyPEHeader2.HeaderSize & " bytes" & vbNewLine _
                            & "Section Alignment: 0x" & Hex(crazyPEHeader2.SectionAlignment) & vbNewLine _
                            & "File Alignment: 0x" & Hex(crazyPEHeader2.FileAlignment) & vbNewLine _
                            & "Subsystem: " & findSub(crazyPEHeader2.Subsystem) & vbNewLine _
                            & "Machine Code is For: " & findMac(crazyPEHeader1.Machine) & vbNewLine _

        For i = 0 To UBound(crazyPESections)
            sectionName = StrConv(crazyPESections(i).sectionName, vbUnicode)
            sectionName = Left(sectionName, InStr(sectionName, Chr(0)) - 1)
            
            frmPENfo = frmPENfo & vbNewLine & "SECTION '" & sectionName & "' (" & getSecDesc(sectionName) & ")" & vbNewLine & _
                                   "File Offset: 0x" & Hex(crazyPESections(i).PData) & vbNewLine & _
                                   "Size: 0x" & Hex(crazyPESections(i).SizeOfData) & vbNewLine & _
                                   "Virtual Offset: 0x" & Hex(crazyPESections(i).VirtualAddress + crazyPEHeader2.ImageBase) & vbNewLine & _
                                   "Absolute Size: 0x" & Hex(crazyPESections(i).Address) & vbNewLine

        Next i
        '------------'
        
        
        CloseHandle VIEWER
    
    For Each hashItem In blockVars
        If blockVars(hashItem).Exists("headers") <> 0 Then
            headerVal = Mid(StrConv(DestinationString, vbUnicode), (CLng("&H" & blockVars(hashItem)("block-offset")) + 1), CLng("&H" & blockVars(hashItem)("block-size")))
            headerSplit = Split(headerVal, vbNewLine, -1, vbBinaryCompare)
            
            For Each headerPart In headerSplit
                headerSplit2 = Split(headerPart, ": ", -1, vbBinaryCompare)
                Call addKnownHeader(headerSplit2(0), headerSplit2(1))
            Next headerPart
            
        End If
    Next hashItem
    
    DestinationStringD = StrConv(DestinationString, vbUnicode)
    DestinationStringD = Left(DestinationStringD, Len(DestinationStringD) - 1)
    getMD5 = LCase(MD5(DestinationStringD, True))
    frmPENfo = frmPENfo & vbNewLine & "MD5 Checksum: " & UCase(getMD5) & vbNewLine
        
    End If
    
End Function


Private Function findSub(subs As Integer)
    '#define IMAGE_SUBSYSTEM_UNKNOWN              0   // Unknown subsystem.
    '#define IMAGE_SUBSYSTEM_NATIVE               1   // Image doesn't require a subsystem.
    '#define IMAGE_SUBSYSTEM_WINDOWS_GUI          2   // Image runs in the Windows GUI subsystem.
    '#define IMAGE_SUBSYSTEM_WINDOWS_CUI          3   // Image runs in the Windows character subsystem.
    '#define IMAGE_SUBSYSTEM_OS2_CUI              5   // image runs in the OS/2 character subsystem.
    '#define IMAGE_SUBSYSTEM_POSIX_CUI            7   // image runs in the Posix character subsystem.
    '#define IMAGE_SUBSYSTEM_NATIVE_WINDOWS       8   // image is a native Win9x driver.
    '#define IMAGE_SUBSYSTEM_WINDOWS_CE_GUI       9   // Image runs in the Windows CE subsystem.
    '#define IMAGE_SUBSYSTEM_EFI_APPLICATION      10
    '#define IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER  11
    '#define IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER   12
    '#define IMAGE_SUBSYSTEM_EFI_ROM              13
    '#define IMAGE_SUBSYSTEM_XBOX                 14
    
    Select Case subs
        Case 1:
            findSub = "Any"
        Case 2:
            findSub = "Win32 GUI"
        Case 3:
            findSub = "Win32 Console"
        Case 5:
            findSub = "OS/2"
        Case 7:
            findSub = "POSIX"
        Case 8:
            findSub = "Windows 9x Driver"
        Case 9:
            findSub = "Windows CE"
        Case 10:
            findSub = "EFI Application"
        Case 11:
            findSub = "EFI Boot Service Driver"
        Case 12:
            findSub = "EFI Runtime Driver"
        Case 13:
            findSub = "EFI ROM"
        Case 14:
            findSub = "XBOX"
        Case Else
            findSub = "Unkown"
    End Select
    
End Function

Private Function findMac(mac As Integer)
    '#define IMAGE_FILE_MACHINE_UNKNOWN           0
    '#define IMAGE_FILE_MACHINE_I386              0x014c  // Intel 386.
    '#define IMAGE_FILE_MACHINE_R3000             0x0162  // MIPS little-endian, 0x160 big-endian
    '#define IMAGE_FILE_MACHINE_R4000             0x0166  // MIPS little-endian
    '#define IMAGE_FILE_MACHINE_R10000            0x0168  // MIPS little-endian
    '#define IMAGE_FILE_MACHINE_WCEMIPSV2         0x0169  // MIPS little-endian WCE v2
    '#define IMAGE_FILE_MACHINE_ALPHA             0x0184  // Alpha_AXP
    '#define IMAGE_FILE_MACHINE_POWERPC           0x01F0  // IBM PowerPC Little-Endian
    '#define IMAGE_FILE_MACHINE_SH3               0x01a2  // SH3 little-endian
    '#define IMAGE_FILE_MACHINE_SH3E              0x01a4  // SH3E little-endian
    '#define IMAGE_FILE_MACHINE_SH4               0x01a6  // SH4 little-endian
    '#define IMAGE_FILE_MACHINE_ARM               0x01c0  // ARM Little-Endian
    '#define IMAGE_FILE_MACHINE_THUMB             0x01c2
    '#define IMAGE_FILE_MACHINE_IA64              0x0200  // Intel 64
    '#define IMAGE_FILE_MACHINE_MIPS16            0x0266  // MIPS
    '#define IMAGE_FILE_MACHINE_MIPSFPU           0x0366  // MIPS
    '#define IMAGE_FILE_MACHINE_MIPSFPU16         0x0466  // MIPS
    '#define IMAGE_FILE_MACHINE_ALPHA64           0x0284  // ALPHA64
    '#define IMAGE_FILE_MACHINE_AXP64             IMAGE_FILE_MACHINE_ALPHA64
    
    Select Case mac
        Case &H14C:
            findMac = "Intel 386"
        Case &H160:
            findMac = "MIPS big-endian"
        Case &H162:
            findMac = "MIPS little-endian"
        Case &H166:
            findMac = "MIPS little-endian"
        Case &H168:
            findMac = "MIPS little-endian"
        Case &H169:
            findMac = "MIPS little-endian Windows CE"
        Case &H184:
            findMac = "Alpha AXP"
        Case &H1F0:
            findMac = "IBM PowerPC Little-endian"
        Case &H1A2:
            findMac = "SH3E little-endian"
        Case &H1A4:
            findMac = "SH3E little-endian"
        Case &H1A6:
            findMac = "SH4 little-endian"
        Case &H1C0:
            findMac = "ARM little-endian"
        Case &H200:
            findMac = "Intel 64-bit"
        Case &H266:
            findMac = "MIPS"
        Case &H366:
            findMac = "MIPS"
        Case &H466:
            findMac = "MIPS"
        Case &H284:
            findMac = "Alpha 64-bit"
        Case Else
            findMac = "Unkown"
    End Select

End Function

Private Function getSecDesc(sectionName As String)

    Select Case sectionName
        Case ".text":
            getSecDesc = "Code: Assembly"
        Case ".data":
            getSecDesc = "Constants, Strings, Structures"
        Case ".rdata":
            getSecDesc = "Import/export table, Debug information"
        Case ".rsrc":
            getSecDesc = "Resources: file menus, icons etc..."
        Case Else
            getSecDesc = "Unkown"
    End Select
End Function

Private Sub addKnownHeader(headerName, HeaderValue)
    Dim dd
    
    If knownHeadCount > 30 Then
        MsgBox "You can't add more than 30 header objects", vbCritical, "You fucked up"
    Else
    
        cmboHeaders(knownHeadCount).Visible = True
        lblHeaders(knownHeadCount).Visible = True
        chkDelHeaders(knownHeadCount).Visible = True
    
        lblHeaders(knownHeadCount).Caption = headerName
        cmboHeaders(knownHeadCount).Clear
        cmboHeaders(knownHeadCount).Text = HeaderValue
        
        For Each dd In hderObjs(lblHeaders(knownHeadCount).Caption)("1")
            cmboHeaders(knownHeadCount).AddItem (dd)
        Next dd
        
        knownHeadCount = knownHeadCount + 1
    
        Call calcHeadFramSiz
        Call buildMenu
    End If
End Sub




Private Sub Command2_Click()
    On Error Resume Next
    
    strData = ""
    contentLen = -1
    strBody = ""
    
    Command2.Enabled = False
    
    Winsock1.Connect "testdrive04.mine.nu", 80
    
End Sub

Private Sub doUpdate(updateStuff As String)
    Dim tempHash
    Dim objCount As Integer
    
    objCount = 0
    
    Command5.Visible = False
    
    For Each tempHash In Split(updateStuff, "=-=")
        Select Case objCount
            Case 0:
                If tempHash > VERSION Then
                    Text2.ForeColor = &HFF0000
                Else
                    Text2.ForeColor = &HFF&
                    Text2.Text = "No new version available"
                    Exit Sub
                End If
            Case 1:
                Text2.Text = "New version available (" & tempHash & "): " & vbNewLine & vbNewLine
            Case 2:
                Text2.Text = Text2.Text & tempHash
            Case 3:
                Command5.Visible = True
                Command5.Tag = tempHash
        End Select
        
        objCount = objCount + 1
    Next tempHash
End Sub

Private Sub Command5_Click()
    Dim WshShell As Object
    
    Set WshShell = CreateObject("WScript.Shell")
    
    WshShell.Run Command5.Tag
End Sub

Private Sub Command6_Click()
    Dim desc As String
    Dim addFile As String
    Dim objCount As Integer
    Dim hashItem
    Dim psID As String
    Dim OPENFILE As String
    Dim RandNum As String
    
    addFile = vbNewLine & vbNewLine & "# Generated preset from 4.0" & vbNewLine
    
    desc = InputBox("Please type a description of the current header set", "Plop")
    
    If desc <> "" Then
        psID = "PS" & MD5(desc, True)
        
        addFile = addFile & "define " & psID & ": {" & vbNewLine
        addFile = addFile & "description: " & Chr(34) & desc & Chr(34) & vbNewLine
        Set boxPresets(psID) = New Dictionary
        Set boxPresets(psID)("header") = New Dictionary
        boxPresets(psID)("description") = desc
    
        For Each hashItem In chkDelHeaders
        
            If hashItem.Value = 0 And hashItem.Visible = True Then
                addFile = addFile & lblHeaders(objCount).Caption & ": " & cmboHeaders(objCount).Text & vbNewLine
                boxPresets(psID)("header")(lblHeaders(objCount).Caption) = cmboHeaders(objCount).Text
            End If
            objCount = objCount + 1
        Next hashItem

        addFile = addFile & "}"
    
        lstCrtBP.Clear
        Combo1.Clear
        For Each hashItem In boxPresets
            If boxPresets(hashItem).Exists("description") <> 0 Then
                lstCrtBP.AddItem (boxPresets(hashItem)("description"))
                Combo1.AddItem (boxPresets(hashItem)("description"))
            End If
        Next hashItem
    
        OPENFILE = App.Path & "\" & CONFIGDIR & "\" & presetFN
        
        Err = 0
        Open OPENFILE For Append As #VIEWER
        
        If Err <> 0 Then
            MsgBox "addPreset(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
            Unload frmMain
        Else
        
            Print #VIEWER, addFile
            Close #VIEWER
        End If
        
        MsgBox "Preset added", vbOKOnly, "Whoohooo"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuIPEC_Click()
    frm1800.Show
End Sub

Private Sub mnuRunCalc_Click()
    Dim WshShell As Object
    
    Set WshShell = CreateObject("WScript.Shell")
    
    WshShell.Run "calc"
End Sub

Private Sub mnuRunView_Click()
    Dim WshShell As Object
    
    Set WshShell = CreateObject("WScript.Shell")
    
    WshShell.Run pathto
    
    frmRegEdit.dontEdit = 0
End Sub



Private Sub evalDownload()
    Call doUpdate(strData)
    Command2.Enabled = True
End Sub


Private Sub Winsock1_Connect()
    On Error Resume Next
    
    Winsock1.SendData "GET /40update.txt HTTP/1.1" & vbCrLf & _
                      "Host: testdrive04.mine.nu" & vbCrLf & _
                      "User-Agent: SuperViewer 4.0 Update" & vbCrLf & vbCrLf
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim recData As String
    Dim hashBullShit
    Dim header As String
    Dim headerVal As String
    
    Winsock1.GetData recData, vbString
    strData = strData & recData
    
    If InStr(1, strData, vbCrLf & vbCrLf) Then
        If contentLen = -1 Then
            For Each hashBullShit In Split(Left(strData, InStr(strData, vbCrLf & vbCrLf)), vbCrLf)
                header = Left(hashBullShit, InStr(hashBullShit, ":") - 1)
                headerVal = Mid(hashBullShit, InStr(hashBullShit, ":") + 1)
                If LCase(header) = "content-length" Then
                    contentLen = headerVal
                End If
            Next hashBullShit
        
            strBody = Mid(strData, InStr(strData, vbCrLf & vbCrLf) + 4)
            
            If Len(strBody) = contentLen Or Len(strBody) > contentLen Then
                strData = strBody
                Winsock1.Close
                Call evalDownload
            End If
        Else
            
            strBody = strBody & recData
            
            If Len(strBody) = contentLen Or Len(strBody) > contentLen Then
                strData = strBody
                Winsock1.Close
                Call evalDownload
            End If

        End If
        
    End If
        
End Sub

Private Sub Command3_Click()
    Dim hashItem
    Dim objCount As Integer
    
    objCount = 0
    
    For Each hashItem In chkDelHeaders
        
        If hashItem.Value = 0 And hashItem.Visible = True Then
            tempDict(lblHeaders(objCount).Caption) = cmboHeaders(objCount).Text
        End If
        objCount = objCount + 1
    Next hashItem
    
    Call delAllHeads
    
    For Each hashItem In tempDict.Keys
        Call addKnownHeader(hashItem, tempDict(hashItem))
    Next hashItem
    
    tempDict.RemoveAll

End Sub

Private Sub Command4_Click()
    Dim boxTtoUse As String
    Dim headVal As String
    Dim hashValue
    
    boxTtoUse = boxPresets.Keys(Combo1.ListIndex)
    
    Call delAllHeads
    
    For Each hashValue In boxPresets(boxTtoUse)("header")
        headVal = boxPresets(boxTtoUse)("header")(hashValue)
        If headVal = "=GET(?)" Then
            Call addKnownHeader(hashValue, "PUT SOMETHING IN HERE!!")
        Else
            Call addKnownHeader(hashValue, headVal)
        End If
    Next hashValue
End Sub

Private Sub imgMoreEdit_Click()
    If selMoreEdit <> -1 Then
        Select Case lblHeaders(selMoreEdit).Caption
            Case "wtv-client-serial-number":
            
                If Len(cmboHeaders(selMoreEdit).Text) = 16 Then
                    Call frmSSID.CalcSSID(cmboHeaders(selMoreEdit).Text)
                    frmSSID.indexItem = selMoreEdit
                    frmSysConfig.Hide
                    frmSSID.Show
                Else
                    MsgBox "SSID Must be 16 digits in length", vbCritical, "You fucked up"
                End If
            Case "wtv-system-sysconfig":
                Call frmSysConfig.CalcSYSCFG(cmboHeaders(selMoreEdit).Text)
                frmSysConfig.indexItem = selMoreEdit
                frmSSID.Hide
                frmSysConfig.Show
        End Select

    End If
End Sub

Private Sub lblHeaders_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuAddHead
    End If
End Sub

Private Sub mnuAddHead1_Click()
    Call addKnownHeader(mnuAddHead1.Caption, "")
End Sub
Private Sub mnuAddHead2_Click()
    Call addKnownHeader(mnuAddHead2.Caption, "")
End Sub
Private Sub mnuAddHead3_Click()
    Call addKnownHeader(mnuAddHead3.Caption, "")
End Sub
Private Sub mnuAddHead4_Click()
    Call addKnownHeader(mnuAddHead4.Caption, "")
End Sub
Private Sub mnuAddHead5_Click()
    Call addKnownHeader(mnuAddHead5.Caption, "")
End Sub
Private Sub mnuAddHead6_Click()
    Call addKnownHeader(mnuAddHead6.Caption, "")
End Sub
Private Sub mnuAddHead7_Click()
    Call addKnownHeader(mnuAddHead7.Caption, "")
End Sub
Private Sub mnuAddHead8_Click()
    Call addKnownHeader(mnuAddHead8.Caption, "")
End Sub
Private Sub mnuAddHead9_Click()
    Call addKnownHeader(mnuAddHead9.Caption, "")
End Sub
Private Sub mnuAddHead10_Click()
    Call addKnownHeader(mnuAddHead10.Caption, "")
End Sub
Private Sub mnuAddHead11_Click()
    Call addKnownHeader(mnuAddHead11.Caption, "")
End Sub
Private Sub mnuAddHead12_Click()
    Call addKnownHeader(mnuAddHead12.Caption, "")
End Sub
Private Sub mnuAddHead13_Click()
    Call addKnownHeader(mnuAddHead13.Caption, "")
End Sub
Private Sub mnuAddHead14_Click()
    Call addKnownHeader(mnuAddHead14.Caption, "")
End Sub
Private Sub mnuAddHead15_Click()
    Call addKnownHeader(mnuAddHead15.Caption, "")
End Sub
Private Sub mnuAddHead16_Click()
    Call addKnownHeader(mnuAddHead16.Caption, "")
End Sub
Private Sub mnuAddHead17_Click()
    Call addKnownHeader(mnuAddHead17.Caption, "")
End Sub
Private Sub mnuAddHead18_Click()
    Call addKnownHeader(mnuAddHead18.Caption, "")
End Sub
Private Sub mnuAddHead19_Click()
    Call addKnownHeader(mnuAddHead19.Caption, "")
End Sub
Private Sub mnuAddHead20_Click()
    Call addKnownHeader(mnuAddHead20.Caption, "")
End Sub
Private Sub mnuAddHead21_Click()
    Call addKnownHeader(mnuAddHead21.Caption, "")
End Sub
Private Sub mnuAddHead22_Click()
    Call addKnownHeader(mnuAddHead22.Caption, "")
End Sub
Private Sub mnuAddHead23_Click()
    Call addKnownHeader(mnuAddHead23.Caption, "")
End Sub
Private Sub mnuAddHead24_Click()
    Call addKnownHeader(mnuAddHead24.Caption, "")
End Sub
Private Sub mnuAddHead25_Click()
    Call addKnownHeader(mnuAddHead25.Caption, "")
End Sub
Private Sub mnuAddHead26_Click()
    Call addKnownHeader(mnuAddHead26.Caption, "")
End Sub
Private Sub mnuAddHead27_Click()
    Call addKnownHeader(mnuAddHead27.Caption, "")
End Sub
Private Sub mnuAddHead28_Click()
    Call addKnownHeader(mnuAddHead28.Caption, "")
End Sub
Private Sub mnuAddHead29_Click()
    Call addKnownHeader(mnuAddHead29.Caption, "")
End Sub
Private Sub mnuAddHead30_Click()
    Call addKnownHeader(mnuAddHead30.Caption, "")
End Sub

Private Sub calcHeadFramSiz()
    
    Frame4.Height = (knownHeadCount * 480) + 120
    
    If Frame4.Height < frmScroll.Height Then
        Frame4.Height = frmScroll.Height
        VScroll1.Enabled = False
        Frame4.Top = 0
    Else
        VScroll1.Enabled = True
        VScroll1.Max = Abs(Frame4.Height - frmScroll.Height)
        VScroll1.LargeChange = VScroll1.Max / 0.5
        VScroll1.SmallChange = VScroll1.Max / 10
    End If
    
End Sub


Private Sub Form_Load()
    Dim SizeOfFile As Double
    Dim hashItem
    Dim hashItem2
    Dim objCount As Integer
    Dim tempStr As String
        
    VERSION = "400"
    CONFIGDIR = "..\Config"
    configFile = "Config.ini"
    knownHeadCount = 0
    
    If frmRegEdit.dontEdit <> 1 Then
        frmRegEdit.dontEdit = 0
        
    End If
    
    Set FSRM = CreateObject("Scripting.FileSystemObject")
        
    Set tempDict = New Dictionary
    tempDict.RemoveAll
    Set hderObjs = New Dictionary
    hderObjs.RemoveAll
    Set blockVars = New Dictionary
    blockVars.RemoveAll
    Set boxPresets = New Dictionary
    boxPresets.RemoveAll
    Set editCodes = New Dictionary
    editCodes.RemoveAll
    
    tempInt = 0
    
    
    imgMoreEdit.Visible = False

    For Each hashItem In lblHeaders
        hashItem.Visible = False
    Next hashItem
    
    For Each hashItem In cmboHeaders
        hashItem.Visible = False
    Next hashItem
    
    For Each hashItem In chkDelHeaders
        hashItem.Visible = False
    Next hashItem
    
    Call showAgeement
    Call loadConfig
    
    txtInfoTemp.Text = TEMPLATE
    txtInfo_pathto.Text = pathto
    txtInfo_viewervers.Text = viewervers
    txtInfo_thehash.Text = thehash
    
    txtInfo_acthash.Text = getMD5()
    
    If txtInfo_acthash.Text <> txtInfo_thehash.Text Then
        txtInfo_acthash.ForeColor = &HFF&
    Else
        txtInfo_acthash.ForeColor = &H80000008
    End If
    
    lstCrtBP.Clear
    Combo1.Clear
    For Each hashItem In boxPresets
        If boxPresets(hashItem).Exists("description") <> 0 Then
            lstCrtBP.AddItem (boxPresets(hashItem)("description"))
            Combo1.AddItem (boxPresets(hashItem)("description"))
        End If
    Next hashItem

    Call buildMenu
    
    Call calcHeadFramSiz
End Sub

Private Sub buildMenu()
    Dim hashItem
    Dim hashItem2
    Dim hashItem3
    Dim flagNo As Integer
    
    mnuHeadCount = 0
    
    mnuAddHead1.Visible = False
    mnuAddHead2.Visible = False
    mnuAddHead3.Visible = False
    mnuAddHead4.Visible = False
    mnuAddHead5.Visible = False
    mnuAddHead6.Visible = False
    mnuAddHead7.Visible = False
    mnuAddHead8.Visible = False
    mnuAddHead9.Visible = False
    mnuAddHead10.Visible = False
    mnuAddHead11.Visible = False
    mnuAddHead12.Visible = False
    mnuAddHead13.Visible = False
    mnuAddHead14.Visible = False
    mnuAddHead15.Visible = False
    mnuAddHead16.Visible = False
    mnuAddHead17.Visible = False
    mnuAddHead18.Visible = False
    mnuAddHead19.Visible = False
    mnuAddHead20.Visible = False
    mnuAddHead21.Visible = False
    mnuAddHead22.Visible = False
    mnuAddHead23.Visible = False
    mnuAddHead24.Visible = False
    mnuAddHead25.Visible = False
    mnuAddHead26.Visible = False
    mnuAddHead27.Visible = False
    mnuAddHead28.Visible = False
    mnuAddHead29.Visible = False
    mnuAddHead30.Visible = False
    
    For Each hashItem In blockVars
        If blockVars(hashItem).Exists("headers") <> 0 Then
            For Each hashItem2 In blockVars(hashItem)("headers")
                flagNo = 0
                For Each hashItem3 In lblHeaders
                    If hashItem3.Caption = hashItem2 Then
                        flagNo = 1
                    End If
                Next hashItem3
                If flagNo <> 1 Then
                    Call addHeadMenu(hashItem2)
                End If
            Next hashItem2
        End If
    Next hashItem

End Sub

Private Sub addHeadMenu(menuCaption)
    
    mnuHeadCount = mnuHeadCount + 1
    
    Select Case mnuHeadCount
        Case 1:
            mnuAddHead1.Caption = menuCaption
            mnuAddHead1.Visible = True
        Case 2:
            mnuAddHead2.Caption = menuCaption
            mnuAddHead2.Visible = True
        Case 3:
            mnuAddHead3.Caption = menuCaption
            mnuAddHead3.Visible = True
        Case 4:
            mnuAddHead4.Caption = menuCaption
            mnuAddHead4.Visible = True
        Case 5:
            mnuAddHead5.Caption = menuCaption
            mnuAddHead5.Visible = True
        Case 6:
            mnuAddHead6.Caption = menuCaption
            mnuAddHead6.Visible = True
        Case 7:
            mnuAddHead7.Caption = menuCaption
            mnuAddHead7.Visible = True
        Case 8:
            mnuAddHead8.Caption = menuCaption
            mnuAddHead8.Visible = True
        Case 9:
            mnuAddHead9.Caption = menuCaption
            mnuAddHead9.Visible = True
        Case 10:
            mnuAddHead10.Caption = menuCaption
            mnuAddHead10.Visible = True
        Case 11:
            mnuAddHead11.Caption = menuCaption
            mnuAddHead11.Visible = True
        Case 12:
            mnuAddHead12.Caption = menuCaption
            mnuAddHead12.Visible = True
        Case 13:
            mnuAddHead13.Caption = menuCaption
            mnuAddHead13.Visible = True
        Case 14:
            mnuAddHead14.Caption = menuCaption
            mnuAddHead14.Visible = True
        Case 15:
            mnuAddHead15.Caption = menuCaption
            mnuAddHead15.Visible = True
        Case 16:
            mnuAddHead16.Caption = menuCaption
            mnuAddHead16.Visible = True
        Case 17:
            mnuAddHead17.Caption = menuCaption
            mnuAddHead17.Visible = True
        Case 18:
            mnuAddHead18.Caption = menuCaption
            mnuAddHead18.Visible = True
        Case 19:
            mnuAddHead19.Caption = menuCaption
            mnuAddHead19.Visible = True
        Case 20:
            mnuAddHead20.Caption = menuCaption
            mnuAddHead20.Visible = True
        Case 21:
            mnuAddHead21.Caption = menuCaption
            mnuAddHead21.Visible = True
        Case 22:
            mnuAddHead22.Caption = menuCaption
            mnuAddHead22.Visible = True
        Case 23:
            mnuAddHead23.Caption = menuCaption
            mnuAddHead23.Visible = True
        Case 24:
            mnuAddHead24.Caption = menuCaption
            mnuAddHead24.Visible = True
        Case 25:
            mnuAddHead25.Caption = menuCaption
            mnuAddHead25.Visible = True
        Case 26:
            mnuAddHead26.Caption = menuCaption
            mnuAddHead26.Visible = True
        Case 27:
            mnuAddHead27.Caption = menuCaption
            mnuAddHead27.Visible = True
        Case 28:
            mnuAddHead28.Caption = menuCaption
            mnuAddHead28.Visible = True
        Case 29:
            mnuAddHead29.Caption = menuCaption
            mnuAddHead29.Visible = True
        Case 30:
            mnuAddHead30.Caption = menuCaption
            mnuAddHead30.Visible = True
    End Select
End Sub


Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuAddHead
    End If
End Sub

Private Sub delAllHeads()
    Dim hashItem
    
    knownHeadCount = 0
    For Each hashItem In lblHeaders
        hashItem.Visible = False
        hashItem.Caption = "Unkown?"
    Next hashItem
    
    For Each hashItem In cmboHeaders
        hashItem.Visible = False
        hashItem.Clear
    Next hashItem
    
    For Each hashItem In chkDelHeaders
        hashItem.Visible = False
        hashItem.Value = 0
    Next hashItem
        
    Call calcHeadFramSiz
    Call buildMenu
End Sub

Private Sub mnuDelAllHeads_Click()
    
    If (MsgBox("Are you sure you want to delete all of the headers?", vbOKCancel, "Yikes brother sikes!")) = vbOK Then
        Call delAllHeads
    End If
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub


Private Sub showAgeement()
    On Error Resume Next
    
    Dim OPENFILE As String
    
    Err = 0
    
    OPENFILE = App.Path & "\..\Guiding Principles.txt"
    
    Set FILE = FSRM.OpenTextFile(OPENFILE, 1, False)
    

    If Err <> 0 Then
        MsgBox "showAgeement(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
    
        txtGPrinc.Text = FILE.ReadAll
        FILE.Close
    End If
    
End Sub



Private Sub readTemp(FileName As String)
    On Error Resume Next
    
    Dim fileLine As String
    Dim fileMatch As Match
    Dim variab As String, subvar As String, flatval As String
    Dim OPENFILE As String
    Dim notlevel1 As Integer
    Dim inlevel(20) As String
    Dim defining As String
    Dim i As Integer
    Dim temp As String
    Dim OPENFILE2 As String
    Dim DestinationString As String

    Err = 0

    OPENFILE = App.Path & "\" & CONFIGDIR & "\" & FileName
    
    Set FILE2 = FSRM.OpenTextFile(OPENFILE, 1, False)
    If Err <> 0 Then
        MsgBox "readTemp(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
    
        Do While FILE2.AtEndOfStream <> True
            fileLine = FILE2.ReadLine
        
            If checkRE(fileLine, "^\-\>(.*)$") <> 0 Then
                TEMPLATE = REMatch.SubMatches(0)
            Else
                If checkRE(fileLine, "^\s*\}") <> 0 Then
                    notlevel1 = notlevel1 - 1
                    If notlevel1 < 0 Then notlevel1 = 0
                Else
                    If checkRE(fileLine, "^(\S*)\s*(\S*)\:\s*(.*)$") <> 0 Then
                        variab = REMatch.SubMatches(0)
                        subvar = REMatch.SubMatches(1)
                        flatval = REMatch.SubMatches(2)
                    
                        If notlevel1 <> 0 Then
                            
                            If inlevel(notlevel1 - 1) = "define" And notlevel1 = 1 Then
                                Select Case variab
                                    Case "block-offset":
    
                                        If checkRE(flatval, "^[A-Fa-f0-9]*$") <> 0 Then
                                            blockVars(defining)("block-offset") = flatval
                                         End If
    
                                    Case "block-size":
    
                                        If checkRE(flatval, "^\d*$") <> 0 Then
                                            blockVars(defining)("block-size") = flatval
                                        End If
                                    
                                    Case "description":
                                    
                                        If checkRE(flatval, "^(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                                            blockVars(defining)("description") = REMatch.SubMatches(1)
                                        End If
                                    
                                    
                                    Case "multiple":
                                    
                                        blockVars(defining)("multiple") = flatval
                                    
                                    Case "views":
                                    
                                        blockVars(defining)("views") = flatval
                                    
                                    Case "noblanks":
                                    
                                        blockVars(defining)("noblanks") = flatval
                                    
                                    
                                    Case "isspecial":
                                    
                                        blockVars(defining)("isspecial") = flatval
                                    
                                    Case "write-end":
                                    
                                        blockVars(defining)("write-end") = flatval
                                    
                                    Case "headers":
                                        notlevel1 = notlevel1 + 1
                                        inlevel(notlevel1 - 1) = "headers"
                                        Set blockVars(defining)("headers") = New Dictionary
                                End Select

                            Else
                                If inlevel(notlevel1 - 1) = "headers" And notlevel1 = 2 Then
                                    

                                    blockVars(defining)("headers")(variab) = ""
                                
                                End If
                            End If
                        
                        Else
                            Select Case variab
                                Case "path":
                                    If checkRE(flatval, "^(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                                        pathto = REMatch.SubMatches(1)
                                    End If
                        
                                Case "vers":
                                    If checkRE(flatval, "^(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                                        viewervers = REMatch.SubMatches(1)
                                    End If
                        
                                Case "unedited-hash":
                                    thehash = flatval
                                
                                Case "string-code":
                                    If checkRE(subvar, "^[A-Fa-f0-9]*$") <> 0 And checkRE(flatval, "^(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                                        editCodes(subvar) = REMatch.SubMatches(1) & Chr(0)
                                    End If
                                Case "section-code":
                                    Err = 0
    
                                    OPENFILE2 = App.Path & "\" & CONFIGDIR & "\" & flatval
                                    VIEWER = FreeFile
                                    Open OPENFILE2 For Binary Access Read As #VIEWER
    
                                    If Err <> 0 Then
                                        MsgBox "readTemplate(): I couldn't open the file '" & OPENFILE2 & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
                                    Else
                                        DestinationString = Space(LOF(VIEWER))
                                        Get #VIEWER, , DestinationString
                                        Close #VIEWER
                                        editCodes("sec_" & subvar) = DestinationString
                                    End If
                                Case "file-code":
                                    If checkRE(subvar, "^[A-Fa-f0-9]*$") <> 0 Then
                                        Err = 0
    
                                        OPENFILE2 = App.Path & "\" & CONFIGDIR & "\" & flatval
                                        VIEWER = FreeFile
                                        Open OPENFILE2 For Binary Access Read As #VIEWER
    
                                        If Err <> 0 Then
                                            MsgBox "readTemplate(): I couldn't open the file '" & OPENFILE2 & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
                                        Else
                                            DestinationString = Space(LOF(VIEWER))
                                            Get #VIEWER, , DestinationString
                                            Close #VIEWER
                                            editCodes(subvar) = DestinationString
                                        End If
                                    End If
                                Case "hex-code":
                                    If checkRE(subvar, "^[A-Fa-f0-9]*$") <> 0 Then
                                        temp = ""
                                        For i = 0 To Len(flatval) Step 2
                                            temp = temp & Chr(CInt("&H" & Mid(flatval, (i - 1), 2)))
                                        Next i
                                        editCodes(subvar) = temp
                                    End If
                            End Select
                        End If
                    Else
                        If (variab = "define") And (checkRE(flatval, "^\{") <> 0) Then
                            notlevel1 = 1
                            inlevel(notlevel1 - 1) = variab
                            defining = subvar
                            Set blockVars(defining) = New Dictionary

                        End If
                    End If
                End If
        
            End If
        Loop
    
        FILE2.Close
    End If

End Sub

Private Sub addHObjs(FileName As String)
    On Error Resume Next
    
    Dim fileLine As String
    Dim fileMatch As Match
    Dim OPENFILE As String
    
    Err = 0

    OPENFILE = App.Path & "\" & CONFIGDIR & "\" & FileName
    
    Set FILE2 = FSRM.OpenTextFile(OPENFILE, 1, False)
    
    If Err <> 0 Then
        MsgBox "addHObjs(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
    
        Do While FILE2.AtEndOfStream <> True
        fileLine = FILE2.ReadLine
        
        If checkRE(fileLine, "^(\S*)\:\s*\u0022(.*?)\u0022\=?(.*)$") <> 0 Then
            Set hderObjs(REMatch.SubMatches(0)) = New Dictionary
            hderObjs(REMatch.SubMatches(0))("0") = REMatch.SubMatches(1)
            
            If REMatch.SubMatches(4) <> "" Then
                hderObjs(REMatch.SubMatches(0))("1") = Split(REMatch.SubMatches(2), ",")
            End If
        End If
        
        Loop
    
        FILE2.Close
    End If
End Sub

Private Sub readPresets()
    On Error Resume Next
    Dim fileLine As String
    Dim fileMatch As Match
    Dim OPENFILE As String
    Dim notlevel1 As Integer
    Dim inlevel(20) As String
    Dim defining As String
    Dim variab, subvar, flatval As String
    
    notlevel1 = 0
    
    Err = 0
    
    OPENFILE = App.Path & "\" & CONFIGDIR & "\" & presetFN
    
    Set FILE2 = FSRM.OpenTextFile(OPENFILE, 1, False)
    

    If Err <> 0 Then
        MsgBox "readPresets(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
    
        Do While FILE2.AtEndOfStream <> True
            fileLine = FILE2.ReadLine
        
            If checkRE(fileLine, "^\s*\}") <> 0 Then
                
                notlevel1 = notlevel1 - 1
                If notlevel1 < 0 Then notlevel1 = 0
                
            Else
            
                If checkRE(fileLine, "^(\S*)\s*(\S*)\:\s*(.*)$") <> 0 Then
                    variab = REMatch.SubMatches(0)
                    subvar = REMatch.SubMatches(1)
                    flatval = REMatch.SubMatches(2)
                    
                    If notlevel1 <> 0 Then
                        If inlevel(notlevel1 - 1) = "define" And notlevel1 = 1 Then
    
                            If variab = "description" And checkRE(flatval, "^(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                                
                                boxPresets(defining)("description") = REMatch.SubMatches(1)
                            Else
                                 boxPresets(defining)("header")(variab) = flatval
                                 
                                 
                            End If

                        End If
                    Else
                        If (variab = "define") And (checkRE(flatval, "^\{") <> 0) Then
                            notlevel1 = 1
                            inlevel(notlevel1 - 1) = variab
                            defining = subvar
                            Set boxPresets(defining) = New Dictionary
                            Set boxPresets(defining)("header") = New Dictionary
                        End If
                    
                    End If
            
                Else
        
                End If
        
            End If
        Loop
        
        
        FILE2.Close
    End If

End Sub

Private Sub loadConfig()
    On Error Resume Next
    
    Dim configLine As String
    Dim variable, Value As String
    Dim OPENFILE As String
    
    Err = 0
    
    OPENFILE = App.Path & "\" & CONFIGDIR & "\" & configFile
    
    Set FILE = FSRM.OpenTextFile(OPENFILE, 1, False)

    If Err <> 0 Then
        MsgBox "loadConfig(): I couldn't open the file '" & OPENFILE & "' I had an Error: " & Err, vbCritical, "Something's fucked up"
        Unload frmMain
    Else
    
        Do While FILE.AtEndOfStream <> True
            configLine = FILE.ReadLine
        

            If checkRE(configLine, "^(\S*)\s*\=\s*(\u0022|\'|)(.*?)(\u0022|\'|)$") <> 0 Then
                variable = REMatch.SubMatches(0)
                Value = REMatch.SubMatches(2)
                Select Case variable
                    Case "template":
                        TEMPLATEP = Value
                        TEMPLATE = Value
                        Call readTemp(Value)
                    Case "headers":
                        HEADERSP = Value
                        Call addHObjs(Value)
                    Case "presets":
                        PRESETSP = Value
                        presetFN = Value
                        Call readPresets
                    Case "emulator":
                        EMULP = Value
                        emulatorPL = Value
                End Select
            
            End If
        
        Loop
    
        FILE.Close
    End If

End Sub

Private Sub VScroll1_Change()
    Call VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()

    If VScroll1.Value < (Frame4.Height - 4095) Then
        Frame4.Top = -(VScroll1.Value)
    End If
End Sub
