VERSION 5.00
Begin VB.Form frm_Input 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Input.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Clear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   5505
      TabIndex        =   26
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton btn_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4545
      TabIndex        =   25
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton btn_Add 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3585
      TabIndex        =   24
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   3998
      TabIndex        =   23
      Top             =   2655
      Width           =   2175
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   1478
      TabIndex        =   22
      Top             =   2655
      Width           =   1575
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   3998
      TabIndex        =   21
      Top             =   2175
      Width           =   2175
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1358
      TabIndex        =   20
      Top             =   2175
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4838
      TabIndex        =   19
      Top             =   1695
      Width           =   1335
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1718
      TabIndex        =   18
      Top             =   1695
      Width           =   1455
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   4958
      TabIndex        =   17
      Top             =   1215
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1838
      TabIndex        =   16
      Top             =   1215
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4958
      TabIndex        =   15
      Top             =   735
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1838
      TabIndex        =   14
      Top             =   735
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4238
      TabIndex        =   13
      Top             =   255
      Width           =   1935
   End
   Begin VB.TextBox tbx_Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1245
      TabIndex        =   0
      Top             =   255
      Width           =   2055
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   90
      X2              =   6450
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   90
      X2              =   6450
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Shipper:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   405
      TabIndex        =   12
      Top             =   255
      Width           =   585
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Liner:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3638
      TabIndex        =   11
      Top             =   255
      Width           =   405
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Release:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3398
      TabIndex        =   10
      Top             =   735
      Width           =   1245
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Request:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   398
      TabIndex        =   9
      Top             =   735
      Width           =   1275
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Return:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   398
      TabIndex        =   8
      Top             =   1215
      Width           =   1170
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Departure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3398
      TabIndex        =   7
      Top             =   1215
      Width           =   1410
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Master B/L No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   398
      TabIndex        =   6
      Top             =   1695
      Width           =   1065
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "House B/L No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   3518
      TabIndex        =   5
      Top             =   1695
      Width           =   1020
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "B/L Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   398
      TabIndex        =   4
      Top             =   2175
      Width           =   675
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Buying Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   2918
      TabIndex        =   3
      Top             =   2175
      Width           =   930
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Selling Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   398
      TabIndex        =   2
      Top             =   2655
      Width           =   900
   End
   Begin VB.Label lbl_Text 
      AutoSize        =   -1  'True
      Caption         =   "Profit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   3398
      TabIndex        =   1
      Top             =   2655
      Width           =   450
   End
End
Attribute VB_Name = "frm_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Add_Click()
    Dim v_iLoop As Integer
    Dim v_iTemp As Integer
    Dim v_bAllowAdd As Boolean
    
    On Error GoTo Error
    
    For v_iLoop = 0 To 11
        If frm_Input.tbx_Text(v_iLoop) <> "" Then v_bAllowAdd = True
        If frm_Input.tbx_Text(v_iLoop) = "" Then frm_Input.tbx_Text(v_iLoop) = " "
    Next v_iLoop
    
    If v_bAllowAdd = True Then
    
    pQuery = "SELECT * FROM Report"
    pDatabase.Open pQuery, pActiveConnection, adOpenDynamic, adLockPessimistic
    
    If Not pDatabase.BOF Then
        pDatabase.MoveLast
        v_iTemp = Val(pDatabase.Fields(0).Value)
    Else
        v_iTemp = 0
    End If
    
    pDatabase.AddNew
    pDatabase!Refrence = Right(Str(v_iTemp + 1), Len(Str(v_iTemp + 1)) - 1)
    pDatabase!Consignee = frm_Input.tbx_Text(0).Text
    pDatabase!Liner = frm_Input.tbx_Text(1).Text
    pDatabase!DateOfRequest = frm_Input.tbx_Text(2).Text
    pDatabase!DateOfRelease = frm_Input.tbx_Text(3).Text
    pDatabase!DateOfReturn = frm_Input.tbx_Text(4).Text
    pDatabase!DateOfDeparture = frm_Input.tbx_Text(5).Text
    pDatabase!MasterBLNo = frm_Input.tbx_Text(6).Text
    pDatabase!HouseBLNo = frm_Input.tbx_Text(7).Text
    pDatabase!BLDate = frm_Input.tbx_Text(8).Text
    pDatabase!BuyingRate = frm_Input.tbx_Text(9).Text
    pDatabase!SellingRate = frm_Input.tbx_Text(10).Text
    pDatabase!Profit = frm_Input.tbx_Text(11).Text
    pDatabase!Result = "Unchecked"
    pDatabase.Update
    
    Call btn_Clear_Click
    Call ShowReport(pDatabase)
    pDatabase.Close
    frm_Input.tbx_Text(0).SetFocus
    
    Else
        For v_iLoop = 0 To 11
            frm_Input.tbx_Text(v_iLoop) = ""
        Next v_iLoop
    End If
    Exit Sub
    
Error:
    MsgBox Err.Source, vbCritical
End Sub

Private Sub btn_Cancel_Click()
    frm_Input.Hide
End Sub

Private Sub btn_Clear_Click()
    Dim v_iLoop As Integer
    
    For v_iLoop = 0 To 11
        frm_Input.tbx_Text(v_iLoop).Text = ""
    Next v_iLoop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim v_lMsg As Single

    v_lMsg = X / Screen.TwipsPerPixelX
    Select Case v_lMsg
        Case WM_LBUTTONUP
        Case WM_RBUTTONUP
            PopupMenu frm_Main.pdm_File
        Case WM_MOUSEMOVE
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONDBLCLK
            frm_Main.WindowState = 2
            frm_Main.Show
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONDBLCLK
        Case Else
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call btn_Clear_Click
End Sub
