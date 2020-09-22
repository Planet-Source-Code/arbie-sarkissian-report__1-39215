VERSION 5.00
Begin VB.Form frm_Logon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Logon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Logon 
      Caption         =   "&Logon"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton btn_Exit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox tbx_Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1470
      PasswordChar    =   "#"
      TabIndex        =   3
      Top             =   735
      Width           =   1575
   End
   Begin VB.TextBox tbx_UserName 
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   255
      Width           =   1575
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   30
      X2              =   3630
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   30
      X2              =   3630
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Index           =   1
      Left            =   510
      TabIndex        =   1
      Top             =   780
      Width           =   750
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   0
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "frm_Logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v_sUserName, v_sPassword As String

Private Sub btn_Exit_Click()
    Unload frm_Logon
    End
End Sub

Private Sub btn_Logon_Click()
    If (frm_Logon.tbx_UserName.Text = v_sUserName) And (frm_Logon.tbx_Password.Text = v_sPassword) Then
        Unload frm_Logon
        frm_Main.Show
    Else
        MsgBox "The UserName or Password is not correct, try again", vbCritical
        frm_Logon.tbx_UserName.Text = ""
        frm_Logon.tbx_Password.Text = ""
        frm_Logon.tbx_UserName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim v_sString As String

    v_sString = Space(255)
    GetPrivateProfileString "Logon", "User Name", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    v_sUserName = TrimString(v_sString)

    v_sString = Space(255)
    GetPrivateProfileString "Logon", "Password", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    v_sPassword = TrimString(v_sString)
End Sub
