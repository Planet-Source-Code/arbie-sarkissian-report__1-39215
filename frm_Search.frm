VERSION 5.00
Begin VB.Form frm_Search 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_Clear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   2115
      Width           =   855
   End
   Begin VB.CommandButton btn_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   2123
      Width           =   855
   End
   Begin VB.CommandButton btn_Search 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   24
      Top             =   2123
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
      Index           =   0
      Left            =   1185
      TabIndex        =   0
      Top             =   203
      Width           =   2055
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
      Left            =   4180
      TabIndex        =   1
      Top             =   203
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
      Index           =   2
      Left            =   7900
      TabIndex        =   2
      Top             =   203
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
      Left            =   1780
      TabIndex        =   3
      Top             =   683
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
      Left            =   4660
      TabIndex        =   4
      Top             =   683
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
      Index           =   5
      Left            =   7780
      TabIndex        =   5
      Top             =   683
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
      Left            =   1660
      TabIndex        =   6
      Top             =   1163
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
      Index           =   7
      Left            =   4780
      TabIndex        =   7
      Top             =   1163
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
      Index           =   8
      Left            =   7660
      TabIndex        =   8
      Top             =   1163
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
      Index           =   9
      Left            =   1420
      TabIndex        =   9
      Top             =   1643
      Width           =   2055
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
      Left            =   4780
      TabIndex        =   10
      Top             =   1643
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
      Index           =   11
      Left            =   7300
      TabIndex        =   12
      Top             =   1643
      Width           =   1815
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   10
      X2              =   9410
      Y1              =   2003
      Y2              =   2003
   End
   Begin VB.Line lin_Separator 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   10
      X2              =   9410
      Y1              =   2018
      Y2              =   2018
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
      Left            =   6700
      TabIndex        =   23
      Top             =   1643
      Width           =   450
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
      Left            =   3700
      TabIndex        =   22
      Top             =   1643
      Width           =   900
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
      Left            =   340
      TabIndex        =   21
      Top             =   1643
      Width           =   930
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
      Left            =   6700
      TabIndex        =   20
      Top             =   1163
      Width           =   675
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
      Left            =   3460
      TabIndex        =   19
      Top             =   1163
      Width           =   1080
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
      Left            =   340
      TabIndex        =   18
      Top             =   1163
      Width           =   1065
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
      Left            =   6220
      TabIndex        =   17
      Top             =   683
      Width           =   1410
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
      Left            =   3340
      TabIndex        =   16
      Top             =   683
      Width           =   1170
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
      Left            =   6460
      TabIndex        =   15
      Top             =   203
      Width           =   1275
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
      Left            =   340
      TabIndex        =   14
      Top             =   683
      Width           =   1245
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
      Left            =   3580
      TabIndex        =   13
      Top             =   203
      Width           =   405
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
      Left            =   345
      TabIndex        =   11
      Top             =   210
      Width           =   585
   End
End
Attribute VB_Name = "frm_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Cancel_Click()
    Unload frm_Search
End Sub

Private Sub btn_Clear_Click()
    Dim v_iLoop As Integer
    
    For v_iLoop = 0 To 11
        frm_Search.tbx_Text(v_iLoop).Text = ""
    Next v_iLoop
End Sub

Private Sub btn_Search_Click()
    Dim v_sQuery As String
    
    v_sQuery = "SELECT * FROM Report WHERE "
    
    If frm_Search.tbx_Text(0).Text <> "" Then
        v_sQuery = v_sQuery & "Consignee LIKE '" & frm_Search.tbx_Text(0).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(1).Text <> "" Then
        v_sQuery = v_sQuery & "Liner LIKE '" & frm_Search.tbx_Text(1).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(2).Text <> "" Then
        v_sQuery = v_sQuery & "DateOfRequest LIKE '" & frm_Search.tbx_Text(2).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(3).Text <> "" Then
        v_sQuery = v_sQuery & "DateOfRelease LIKE '" & frm_Search.tbx_Text(3).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(4).Text <> "" Then
        v_sQuery = v_sQuery & "DateOfReturn LIKE '" & frm_Search.tbx_Text(4).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(5).Text <> "" Then
        v_sQuery = v_sQuery & "DateOfDeparture LIKE '" & frm_Search.tbx_Text(5).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(6).Text <> "" Then
        v_sQuery = v_sQuery & "MasterBLNo LIKE '" & frm_Search.tbx_Text(6).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(7).Text <> "" Then
        v_sQuery = v_sQuery & "HouseBLNo LIKE '" & frm_Search.tbx_Text(7).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(8).Text <> "" Then
        v_sQuery = v_sQuery & "BLDate LIKE '" & frm_Search.tbx_Text(8).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(9).Text <> "" Then
        v_sQuery = v_sQuery & "BuyingRate LIKE '" & frm_Search.tbx_Text(9).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(10).Text <> "" Then
        v_sQuery = v_sQuery & "SellingRate LIKE '" & frm_Search.tbx_Text(10).Text & "' AND "
    End If
    
    If frm_Search.tbx_Text(11).Text <> "" Then
        v_sQuery = v_sQuery & "Profit LIKE '" & frm_Search.tbx_Text(11).Text & "' AND "
    End If
    
    v_sQuery = Left(v_sQuery, Len(v_sQuery) - 5)
    pQuery = v_sQuery
    pDatabase.Open v_sQuery, pActiveConnection
    Call ShowReport(pDatabase)
    frm_Search.Hide
End Sub

