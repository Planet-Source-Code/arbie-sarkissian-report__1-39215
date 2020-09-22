VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Option 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Option.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sst_Option 
      Height          =   3015
      Left            =   143
      TabIndex        =   3
      Top             =   120
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Backup"
      TabPicture(0)   =   "frm_Option.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Backup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Logon"
      TabPicture(1)   =   "frm_Option.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tbx_Password"
      Tab(1).Control(1)=   "tbx_UserName"
      Tab(1).Control(2)=   "cbx_ShowLogon"
      Tab(1).Control(3)=   "lbl_Text(4)"
      Tab(1).Control(4)=   "lbl_Text(3)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "General"
      TabPicture(2)   =   "frm_Option.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl_Text(5)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fra_Startup"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fra_Startup 
         Caption         =   " Startup "
         Height          =   855
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   4575
         Begin VB.CheckBox cbx_ShowShellNotifyIcon 
            Caption         =   "Show Shell Notify Icon On Startup"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.TextBox tbx_Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72735
         PasswordChar    =   "#"
         TabIndex        =   16
         Top             =   1890
         Width           =   1695
      End
      Begin VB.TextBox tbx_UserName 
         Height          =   285
         Left            =   -72735
         TabIndex        =   15
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox cbx_ShowLogon 
         Caption         =   "Show Logon Dialog On Startup"
         Height          =   255
         Left            =   -73695
         TabIndex        =   10
         Top             =   1050
         Width           =   2535
      End
      Begin VB.Frame fra_Backup 
         Caption         =   " Backup "
         Height          =   2415
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   4695
         Begin VB.ComboBox cbo_Backup 
            Height          =   315
            ItemData        =   "frm_Option.frx":091E
            Left            =   1080
            List            =   "frm_Option.frx":0940
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox tbx_AutoBackupFileName 
            Height          =   285
            Left            =   2400
            TabIndex        =   11
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox cbx_AutoBackup 
            Caption         =   "Auto Backup"
            Height          =   495
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox tbx_MaxAutoBackupFileNo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2625
            TabIndex        =   5
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lbl_Text 
            AutoSize        =   -1  'True
            Caption         =   "Backup:"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   885
            Width           =   570
         End
         Begin VB.Label lbl_Text 
            AutoSize        =   -1  'True
            Caption         =   "Auto Backup File(s) Name:"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   8
            Top             =   1350
            Width           =   1890
         End
         Begin VB.Label lbl_Text 
            AutoSize        =   -1  'True
            Caption         =   "Max Auto Backup File(s) No:"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   7
            Top             =   1830
            Width           =   2025
         End
      End
      Begin VB.Label lbl_Text 
         Caption         =   $"frm_Option.frx":09D1
         Height          =   795
         Index           =   5
         Left            =   -74640
         TabIndex        =   19
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label lbl_Text 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   4
         Left            =   -73695
         TabIndex        =   14
         Top             =   1905
         Width           =   750
      End
      Begin VB.Label lbl_Text 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   195
         Index           =   3
         Left            =   -73695
         TabIndex        =   13
         Top             =   1470
         Width           =   840
      End
   End
   Begin VB.CommandButton btn_Default 
      Caption         =   "&Default"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton btn_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton btn_OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -15
      X2              =   5485
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   -15
      X2              =   5485
      Y1              =   3255
      Y2              =   3255
   End
End
Attribute VB_Name = "frm_Option"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Cancel_Click()
    Unload frm_Option
End Sub

Private Sub btn_Default_Click()
    With frm_Option
        .cbx_AutoBackup.Value = 1
        .cbo_Backup.ListIndex = 5
        .tbx_AutoBackupFileName = "Backup"
        .tbx_MaxAutoBackupFileNo = "3"
    End With
End Sub

Private Sub btn_OK_Click()
    If frm_Option.cbx_ShowLogon.Value = 1 Then
        If (frm_Option.tbx_UserName.Text = "") Or (frm_Option.tbx_Password.Text = "") Then
            MsgBox "You should enter UserName & Password", vbInformation
            frm_Option.tbx_UserName.SetFocus
            Exit Sub
        End If
    End If

    Call WriteOption
    frm_Option.Hide
End Sub

Private Sub cbx_ShowLogon_Click()
    If frm_Option.cbx_ShowLogon.Value = 0 Then
        frm_Option.tbx_UserName.Enabled = False
        frm_Option.tbx_Password.Enabled = False
    Else
        frm_Option.tbx_UserName.Enabled = True
        frm_Option.tbx_Password.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call LoadOption
End Sub
