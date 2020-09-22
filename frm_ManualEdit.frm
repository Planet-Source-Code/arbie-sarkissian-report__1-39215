VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_ManualEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manual Edit"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ManualEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider sldr_Slide 
      Height          =   630
      Left            =   2625
      TabIndex        =   28
      Top             =   2145
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1111
      _Version        =   393216
      TickStyle       =   2
   End
   Begin VB.CommandButton btn_Previous 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   2250
      Width           =   855
   End
   Begin VB.CommandButton btn_Next 
      Caption         =   "&Next"
      Height          =   375
      Left            =   1200
      TabIndex        =   26
      Top             =   2250
      Width           =   855
   End
   Begin VB.CommandButton btn_Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   2250
      Width           =   855
   End
   Begin VB.CommandButton btn_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8160
      TabIndex        =   24
      Top             =   2250
      Width           =   855
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   7185
      TabIndex        =   11
      Top             =   1635
      Width           =   1815
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   4665
      TabIndex        =   10
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   1305
      TabIndex        =   9
      Top             =   1635
      Width           =   2055
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   7545
      TabIndex        =   8
      Top             =   1155
      Width           =   1455
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   4665
      TabIndex        =   7
      Top             =   1155
      Width           =   1575
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   1545
      TabIndex        =   6
      Top             =   1155
      Width           =   1455
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   7665
      TabIndex        =   5
      Top             =   675
      Width           =   1335
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   4545
      TabIndex        =   4
      Top             =   675
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   1665
      TabIndex        =   3
      Top             =   675
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   7785
      TabIndex        =   2
      Top             =   195
      Width           =   1215
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   4065
      TabIndex        =   1
      Top             =   195
      Width           =   1935
   End
   Begin VB.TextBox tbx_Text 
      DataSource      =   "dat_Data"
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
      Left            =   1065
      TabIndex        =   0
      Top             =   195
      Width           =   2055
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   -105
      X2              =   9295
      Y1              =   2115
      Y2              =   2115
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -105
      X2              =   9295
      Y1              =   2100
      Y2              =   2100
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
      Left            =   225
      TabIndex        =   23
      Top             =   195
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
      Left            =   3465
      TabIndex        =   22
      Top             =   195
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
      Left            =   225
      TabIndex        =   21
      Top             =   675
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
      Left            =   6345
      TabIndex        =   20
      Top             =   195
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
      Left            =   3225
      TabIndex        =   19
      Top             =   675
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
      Left            =   6105
      TabIndex        =   18
      Top             =   675
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
      Left            =   225
      TabIndex        =   17
      Top             =   1155
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
      Left            =   3345
      TabIndex        =   16
      Top             =   1155
      Width           =   1080
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
      Left            =   6585
      TabIndex        =   15
      Top             =   1155
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
      Left            =   225
      TabIndex        =   14
      Top             =   1635
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
      Left            =   3585
      TabIndex        =   13
      Top             =   1635
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
      Left            =   6585
      TabIndex        =   12
      Top             =   1635
      Width           =   450
   End
End
Attribute VB_Name = "frm_ManualEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gIndex As Integer

Private Sub btn_Cancel_Click()
    Unload frm_ManualEdit
End Sub

Private Sub btn_Next_Click()
    gIndex = gIndex + 1
    frm_ManualEdit.tbx_Text(0).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(1).Text
    frm_ManualEdit.tbx_Text(1).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(2).Text
    frm_ManualEdit.tbx_Text(2).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(3).Text
    frm_ManualEdit.tbx_Text(3).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(4).Text
    frm_ManualEdit.tbx_Text(4).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(5).Text
    frm_ManualEdit.tbx_Text(5).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(6).Text
    frm_ManualEdit.tbx_Text(6).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(7).Text
    frm_ManualEdit.tbx_Text(7).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(8).Text
    frm_ManualEdit.tbx_Text(8).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(9).Text
    frm_ManualEdit.tbx_Text(9).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(10).Text
    frm_ManualEdit.tbx_Text(10).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(11).Text
    frm_ManualEdit.tbx_Text(11).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(12).Text
    
    If gIndex = frm_Main.lvw_List.ListItems.Count Then
        frm_ManualEdit.btn_Next.Enabled = False
    End If
    
    frm_ManualEdit.btn_Previous.Enabled = True
    frm_ManualEdit.sldr_Slide.Value = frm_ManualEdit.sldr_Slide.Value + 1
End Sub

Private Sub btn_Previous_Click()
    gIndex = gIndex - 1
    frm_ManualEdit.tbx_Text(0).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(1).Text
    frm_ManualEdit.tbx_Text(1).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(2).Text
    frm_ManualEdit.tbx_Text(2).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(3).Text
    frm_ManualEdit.tbx_Text(3).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(4).Text
    frm_ManualEdit.tbx_Text(4).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(5).Text
    frm_ManualEdit.tbx_Text(5).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(6).Text
    frm_ManualEdit.tbx_Text(6).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(7).Text
    frm_ManualEdit.tbx_Text(7).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(8).Text
    frm_ManualEdit.tbx_Text(8).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(9).Text
    frm_ManualEdit.tbx_Text(9).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(10).Text
    frm_ManualEdit.tbx_Text(10).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(11).Text
    frm_ManualEdit.tbx_Text(11).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(12).Text

    If gIndex = 1 Then
        frm_ManualEdit.btn_Next.Enabled = True
        frm_ManualEdit.btn_Previous.Enabled = False
    End If

    frm_ManualEdit.sldr_Slide.Value = frm_ManualEdit.sldr_Slide.Value - 1
End Sub

Private Sub btn_Update_Click()
    On Error GoTo Error

    pDatabase.Open pQuery, pActiveConnection, adOpenDynamic, adLockPessimistic
    
    pDatabase.Move (gIndex - 1)
    pDatabase!Consignee = frm_ManualEdit.tbx_Text(0).Text
    pDatabase!Liner = frm_ManualEdit.tbx_Text(1).Text
    pDatabase!DateOfRequest = frm_ManualEdit.tbx_Text(2).Text
    pDatabase!DateOfRelease = frm_ManualEdit.tbx_Text(3).Text
    pDatabase!DateOfReturn = frm_ManualEdit.tbx_Text(4).Text
    pDatabase!DateOfDeparture = frm_ManualEdit.tbx_Text(5).Text
    pDatabase!MasterBLNo = frm_ManualEdit.tbx_Text(6).Text
    pDatabase!HouseBLNo = frm_ManualEdit.tbx_Text(7).Text
    pDatabase!BLDate = frm_ManualEdit.tbx_Text(8).Text
    pDatabase!BuyingRate = frm_ManualEdit.tbx_Text(9).Text
    pDatabase!SellingRate = frm_ManualEdit.tbx_Text(10).Text
    pDatabase!Profit = frm_ManualEdit.tbx_Text(11).Text
    pDatabase.Update
    
    pDatabase.Close
    Exit Sub
    
Error:
    MsgBox Err.Source, vbCritical
End Sub

Private Sub Form_Load()
    If frm_Main.lvw_List.ListItems.Count > 0 Then
    
    gIndex = frm_Main.lvw_List.SelectedItem.Index
    frm_ManualEdit.tbx_Text(0).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(1).Text
    frm_ManualEdit.tbx_Text(1).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(2).Text
    frm_ManualEdit.tbx_Text(2).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(3).Text
    frm_ManualEdit.tbx_Text(3).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(4).Text
    frm_ManualEdit.tbx_Text(4).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(5).Text
    frm_ManualEdit.tbx_Text(5).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(6).Text
    frm_ManualEdit.tbx_Text(6).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(7).Text
    frm_ManualEdit.tbx_Text(7).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(8).Text
    frm_ManualEdit.tbx_Text(8).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(9).Text
    frm_ManualEdit.tbx_Text(9).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(10).Text
    frm_ManualEdit.tbx_Text(10).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(11).Text
    frm_ManualEdit.tbx_Text(11).Text = frm_Main.lvw_List.ListItems(gIndex).ListSubItems(12).Text
    
    frm_ManualEdit.sldr_Slide.Min = 1
    frm_ManualEdit.sldr_Slide.Max = frm_Main.lvw_List.ListItems.Count
    frm_ManualEdit.sldr_Slide.Value = gIndex
    frm_ManualEdit.btn_Update.Enabled = True
    frm_ManualEdit.btn_Next.Enabled = True
    
    Else
        gIndex = 0
        frm_ManualEdit.sldr_Slide.Max = 1
        frm_ManualEdit.btn_Update.Enabled = False
        frm_ManualEdit.btn_Next.Enabled = False
    End If
End Sub

Private Sub sldr_Slide_Change()
    Dim v_iIndex As Integer
    
    If frm_Main.lvw_List.ListItems.Count > 0 Then
    
    v_iIndex = frm_ManualEdit.sldr_Slide.Value
    frm_ManualEdit.tbx_Text(0).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(1).Text
    frm_ManualEdit.tbx_Text(1).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(2).Text
    frm_ManualEdit.tbx_Text(2).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(3).Text
    frm_ManualEdit.tbx_Text(3).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(4).Text
    frm_ManualEdit.tbx_Text(4).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(5).Text
    frm_ManualEdit.tbx_Text(5).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(6).Text
    frm_ManualEdit.tbx_Text(6).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(7).Text
    frm_ManualEdit.tbx_Text(7).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(8).Text
    frm_ManualEdit.tbx_Text(8).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(9).Text
    frm_ManualEdit.tbx_Text(9).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(10).Text
    frm_ManualEdit.tbx_Text(10).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(11).Text
    frm_ManualEdit.tbx_Text(11).Text = frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems(12).Text
    
    If v_iIndex = 1 Then
        frm_ManualEdit.btn_Next.Enabled = True
        frm_ManualEdit.btn_Previous.Enabled = False
    End If
    
    If v_iIndex = frm_Main.lvw_List.ListItems.Count Then
        frm_ManualEdit.btn_Next.Enabled = False
        frm_ManualEdit.btn_Previous.Enabled = True
    End If
    
    gIndex = v_iIndex
    
    End If
End Sub
