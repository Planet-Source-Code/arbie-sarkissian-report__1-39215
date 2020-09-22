VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_PrintPreview 
   Caption         =   "Print Preview"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frm_PrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbr_Preview 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "iml_Preview"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.CommandButton btn_Close 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton btn_Page 
         Caption         =   "Page"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   5
         Top             =   30
         Width           =   855
      End
      Begin VB.TextBox tbx_Page 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   30
         Width           =   855
      End
   End
   Begin VB.HScrollBar hsc_Scroll 
      Height          =   255
      LargeChange     =   50
      Left            =   0
      SmallChange     =   30
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.VScrollBar vsc_Scroll 
      Height          =   1215
      LargeChange     =   50
      Left            =   3000
      SmallChange     =   30
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox pic_Viewport 
      BackColor       =   &H00808080&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      Begin VB.PictureBox pic_Preview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.PictureBox pic_Shadow 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         ScaleHeight     =   495
         ScaleWidth      =   1215
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList iml_Preview 
      Left            =   3240
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_PrintPreview.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu ppm_Page 
      Caption         =   "Page"
      Visible         =   0   'False
      Begin VB.Menu ppi_PageNo 
         Caption         =   "1"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frm_PrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Close_Click()
    Unload frm_PrintPreview
End Sub

Private Sub btn_Page_Click()
    PopupMenu ppm_Page, , frm_PrintPreview.btn_Page.Left, frm_PrintPreview.btn_Page.Top + frm_PrintPreview.btn_Page.Height
End Sub

Private Sub Form_Load()
    Dim v_iRecordsPerPage As Integer

    v_iRecordsPerPage = 50
    With frm_PrintPreview
        .pic_Preview.Width = Printer.Width
        .pic_Preview.Height = Printer.Height
        .pic_Shadow.Width = .pic_Preview.Width
        .pic_Shadow.Height = .pic_Preview.Height
        
        frm_PrintPreview.tbx_Page.Text = "1"
            
        v_iPrintPageCount = Int(frm_Main.lvw_List.ListItems.Count / v_iRecordsPerPage) + 1
        For v_iLoop = 2 To v_iPrintPageCount
            Load frm_PrintPreview.ppi_PageNo(v_iLoop)
            frm_PrintPreview.ppi_PageNo(v_iLoop).Caption = v_iLoop
            frm_PrintPreview.ppi_PageNo(v_iLoop).Visible = True
        Next v_iLoop
        
        Call MakePrintPreview
    End With
End Sub

Private Sub Form_Resize()
    With frm_PrintPreview
        If .ScaleHeight - .hsc_Scroll.Height - .tbr_Preview.Height > 0 Then
        
        .pic_Viewport.Width = .ScaleWidth - .vsc_Scroll.Width - 15
        .pic_Viewport.Height = .ScaleHeight - .hsc_Scroll.Height - 15
        .pic_Preview.Width = Printer.Width
        .pic_Preview.Height = Printer.Height
        .pic_Shadow.Width = .pic_Preview.Width
        .pic_Shadow.Height = .pic_Preview.Height
        
        .pic_Preview.ScaleLeft = Printer.ScaleLeft
        .pic_Preview.ScaleTop = Printer.ScaleTop
        .pic_Preview.ScaleWidth = Printer.ScaleWidth
        .pic_Preview.ScaleHeight = Printer.ScaleHeight
        .pic_Shadow.ScaleWidth = .pic_Preview.ScaleWidth
        .pic_Shadow.ScaleHeight = .pic_Preview.ScaleHeight
        
        .pic_Preview.Left = (.pic_Viewport.ScaleWidth / 2) - (.pic_Preview.Width / 2)
        If .pic_Preview.Left <= 0 Then
            .pic_Preview.Left = 100
        End If
        .pic_Shadow.Left = .pic_Preview.Left + 60
        
        .vsc_Scroll.Left = .ScaleWidth - .vsc_Scroll.Width
        .vsc_Scroll.Height = .ScaleHeight - .hsc_Scroll.Height - .tbr_Preview.Height
        .vsc_Scroll.Min = 120
        .vsc_Scroll.Max = .ScaleHeight - .hsc_Scroll.Height - .pic_Preview.Height - 510
        If .vsc_Scroll.Max < 0 Then
            .vsc_Scroll.Visible = True
        Else
            .vsc_Scroll.Visible = False
        End If
     
        .hsc_Scroll.Top = .ScaleHeight - .hsc_Scroll.Height
        .hsc_Scroll.Width = .ScaleWidth - .vsc_Scroll.Width
        .hsc_Scroll.Min = 100
        .hsc_Scroll.Max = .ScaleWidth - .vsc_Scroll.Width - .pic_Preview.Width - 240
        If .hsc_Scroll.Max < 0 Then
            .hsc_Scroll.Visible = True
        Else
            .hsc_Scroll.Visible = False
        End If
        
        End If
    End With
End Sub

Private Sub hsc_Scroll_Change()
    frm_PrintPreview.pic_Preview.Left = frm_PrintPreview.hsc_Scroll.Value
    frm_PrintPreview.pic_Shadow.Left = frm_PrintPreview.pic_Preview.Left + 60
End Sub

Private Sub hsc_Scroll_Scroll()
    frm_PrintPreview.pic_Preview.Left = frm_PrintPreview.hsc_Scroll.Value
    frm_PrintPreview.pic_Shadow.Left = frm_PrintPreview.pic_Preview.Left + 60
End Sub

Private Sub ppi_PageNo_Click(Index As Integer)
    frm_PrintPreview.tbx_Page.Text = Index
    Call MakePrintPreview(Index)
End Sub

Private Sub tbr_Preview_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim v_iLoop As Integer
    Dim v_iRecordsPerPage As Integer
    
    v_iRecordsPerPage = 50
    Select Case Button.Index
    Case 1:
        For v_iLoop = 1 To Int(frm_Main.lvw_List.ListItems.Count / v_iRecordsPerPage) + 1
            Call PrintPage(v_iLoop)
            Printer.EndDoc
        Next v_iLoop
    End Select
End Sub

Private Sub vsc_Scroll_Change()
    frm_PrintPreview.pic_Preview.Top = frm_PrintPreview.vsc_Scroll.Value
    frm_PrintPreview.pic_Shadow.Top = frm_PrintPreview.pic_Preview.Top + 60
End Sub

Private Sub vsc_Scroll_Scroll()
    frm_PrintPreview.pic_Preview.Top = frm_PrintPreview.vsc_Scroll.Value
    frm_PrintPreview.pic_Shadow.Top = frm_PrintPreview.pic_Preview.Top + 60
End Sub
