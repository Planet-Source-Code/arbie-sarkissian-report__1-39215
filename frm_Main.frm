VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Kabrian Report"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList iml_Menu 
      Left            =   1080
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":0FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":10E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":133E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":17F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1A52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1C7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbr_Tools 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "iml_Images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Input"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList iml_Images 
      Left            =   480
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":248A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":32B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":39C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":40DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":47EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg_Printer 
      Left            =   0
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbr_Status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10345
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/20/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "9:34 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvw_List 
      Height          =   6375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rf"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Shipper"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Liner"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Of Request"
         Object.Width           =   2258
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Of Release"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Of Return"
         Object.Width           =   2170
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date Of Departure"
         Object.Width           =   2222
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Master B/L No."
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "House B/L No."
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "B/L Date"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Buying Rate"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Selling Rate"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Profit"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Result"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu pdm_File 
      Caption         =   "&File"
      Begin VB.Menu pdi_Input 
         Caption         =   "&Input..."
         Shortcut        =   ^I
      End
      Begin VB.Menu pdi_Search 
         Caption         =   "&Search..."
         Shortcut        =   ^S
      End
      Begin VB.Menu Separator01 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_ShowAll 
         Caption         =   "Show &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu Separator02 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_PrintPreview 
         Caption         =   "Print Pre&view"
         Shortcut        =   ^V
      End
      Begin VB.Menu pdi_PrintSetup 
         Caption         =   "Print Setup"
      End
      Begin VB.Menu pdi_Print 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu Separator03 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu pdm_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu pdi_ManualEdit 
         Caption         =   "&Manual Edit..."
         Shortcut        =   ^M
      End
      Begin VB.Menu pdi_AutomaticEdit 
         Caption         =   "A&utomatic Edit..."
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu pdm_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu pdi_Baclup 
         Caption         =   "&Backup..."
         Shortcut        =   ^B
      End
      Begin VB.Menu pdi_BackupsHistory 
         Caption         =   "Backup(s) History"
         Begin VB.Menu pdii_Backup 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu separator04 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Restore 
         Caption         =   "&Restore..."
         Shortcut        =   ^R
      End
      Begin VB.Menu pdi_ReplaceRestoredDatabase 
         Caption         =   "Replace Restored Database"
         Enabled         =   0   'False
      End
      Begin VB.Menu Separator05 
         Caption         =   "-"
      End
      Begin VB.Menu pdi_Option 
         Caption         =   "&Option..."
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu pdm_help 
      Caption         =   "&Help"
      Begin VB.Menu pdi_About 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim v_rsData As New Recordset
    
    pQuery = "SELECT * FROM Report"
    v_rsData.Open pQuery, pActiveConnection
    
    frm_Main.sbr_Status.Panels(1).Text = "Welcome to Kabrian Report Program Version 2.0 Copyright(c) 2002 by Arbie Sarkissian"
    Call ShowReport(v_rsData)
End Sub

Private Sub Form_Resize()
    If frm_Main.WindowState <> 1 Then
        frm_Main.lvw_List.Width = frm_Main.Width - 105
        If frm_Main.Height - frm_Main.tbr_Tools.Height - frm_Main.sbr_Status.Height - 650 > 0 Then
            frm_Main.lvw_List.Height = frm_Main.Height - frm_Main.tbr_Tools.Height - frm_Main.sbr_Status.Height - 650
        End If
    Else
        frm_Main.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
    Unload frm_Main
    Unload frm_Input
    Unload frm_Search
    Unload frm_Option
    End
End Sub

Private Sub lvw_List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim v_rsData As New Recordset
    
    If pQueryTemp <> "" Then
        pQuery = pQueryTemp
    End If
    
    Select Case ColumnHeader.Index
    Case 1:
        If Right(pQuery, 18) = " ORDER BY Refrence" Then
            Exit Sub
        Else
            pQuery = pQuery & " ORDER BY Refrence"
            pQueryTemp = Left(pQuery, Len(pQuery) - 18)
        End If
    Case 2:
        If Right(pQuery, 19) = " ORDER BY Consignee" Then
        Else
            pQuery = pQuery & " ORDER BY Consignee"
            pQueryTemp = Left(pQuery, Len(pQuery) - 19)
        End If
    Case 3:
        If Right(pQuery, 15) = " ORDER BY Liner" Then
        Else
            pQuery = pQuery & " ORDER BY Liner"
            pQueryTemp = Left(pQuery, Len(pQuery) - 15)
        End If
    Case 4:
        If Right(pQuery, 23) = " ORDER BY DateOfRequest" Then
        Else
            pQuery = pQuery & " ORDER BY DateOfRequest"
            pQueryTemp = Left(pQuery, Len(pQuery) - 23)
        End If
    Case 5:
        If Right(pQuery, 23) = " ORDER BY DateOfRelease" Then
        Else
            pQuery = pQuery & " ORDER BY DateOfRelease"
            pQueryTemp = Left(pQuery, Len(pQuery) - 23)
        End If
    Case 6:
        If Right(pQuery, 22) = " ORDER BY DateOfReturn" Then
        Else
            pQuery = pQuery & " ORDER BY DateOfReturn"
            pQueryTemp = Left(pQuery, Len(pQuery) - 22)
        End If
    Case 7:
        If Right(pQuery, 25) = " ORDER BY DateOfDeparture" Then
        Else
            pQuery = pQuery & " ORDER BY DateOfDeparture"
            pQueryTemp = Left(pQuery, Len(pQuery) - 25)
        End If
    Case 8:
        If Right(pQuery, 20) = " ORDER BY MasterBLNo" Then
        Else
            pQuery = pQuery & " ORDER BY MasterBLNo"
            pQueryTemp = Left(pQuery, Len(pQuery) - 20)
        End If
    Case 9:
        If Right(pQuery, 19) = " ORDER BY HouseBLNo" Then
        Else
            pQuery = pQuery & " ORDER BY HouseBLNo"
            pQueryTemp = Left(pQuery, Len(pQuery) - 19)
        End If
    Case 10:
        If Right(pQuery, 16) = " ORDER BY BLDate" Then
        Else
            pQuery = pQuery & " ORDER BY BLDate"
            pQueryTemp = Left(pQuery, Len(pQuery) - 16)
        End If
    Case 11:
        If Right(pQuery, 20) = " ORDER BY BuyingRate" Then
        Else
            pQuery = pQuery & " ORDER BY BuyingRate"
            pQueryTemp = Left(pQuery, Len(pQuery) - 20)
        End If
    Case 12:
        If Right(pQuery, 21) = " ORDER BY SellingRate" Then
        Else
            pQuery = pQuery & " ORDER BY SellingRate"
            pQueryTemp = Left(pQuery, Len(pQuery) - 21)
        End If
    Case 13:
        If Right(pQuery, 16) = " ORDER BY Profit" Then
        Else
            pQuery = pQuery & " ORDER BY Profit"
            pQueryTemp = Left(pQuery, Len(pQuery) - 16)
        End If
    Case 14:
        If Right(pQuery, 16) = " ORDER BY Result" Then
        Else
            pQuery = pQuery & " ORDER BY Result"
            pQueryTemp = Left(pQuery, Len(pQuery) - 16)
        End If
    End Select
    
    v_rsData.Open pQuery, pActiveConnection
    Call ShowReport(v_rsData)
    'v_rsData.Close
End Sub

Private Sub lvw_List_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim v_rsData As New Recordset
    
    v_rsData.Open "Select * From Report", pActiveConnection, adOpenDynamic, adLockPessimistic
    v_rsData.Move (Val(Item.Text) - 1)
    If Item.Checked = True Then
        v_rsData!Result = "Checked"
        frm_Main.lvw_List.ListItems(Item.Index).ListSubItems(13).Text = "Checked"
    Else
        v_rsData!Result = "Unchecked"
        frm_Main.lvw_List.ListItems(Val(Item.Text)).ListSubItems(13).Text = "Unchecked"
    End If
    v_rsData.Update
    'v_rsData.Close
End Sub

Private Sub pdi_About_Click()
    frm_About.Show 1
End Sub

Private Sub pdi_AutomaticEdit_Click()
    frm_AutomaticEdit.Show 1
End Sub

Private Sub pdi_Baclup_Click()
    frm_Main.cdlg_Printer.DialogTitle = "Backup"
    frm_Main.cdlg_Printer.Filter = "Microsoft Access Database (*.mdb)|*.mdb"
    frm_Main.cdlg_Printer.ShowSave
    
    If frm_Main.cdlg_Printer.FileName <> "" Then
        FileCopy App.Path & "\Database.mdb", frm_Main.cdlg_Printer.FileName
    End If
End Sub

Private Sub pdi_Exit_Click()
    Unload frm_Main
    Unload frm_Input
    Unload frm_Search
    Unload frm_Option
    End
End Sub

Private Sub pdi_Input_Click()
    frm_Input.Show 1
End Sub

Private Sub pdi_ManualEdit_Click()
    frm_ManualEdit.Show 1
End Sub

Private Sub pdi_Option_Click()
    frm_Option.Show 1
End Sub

Private Sub pdi_Print_Click()
    Dim v_iLoop As Integer
    Dim v_iRecordsPerPage As Integer
    
    v_iRecordsPerPage = 50
    For v_iLoop = 1 To Int(frm_Main.lvw_List.ListItems.Count / v_iRecordsPerPage) + 1
        Call PrintPage(v_iLoop)
        Printer.EndDoc
    Next v_iLoop
End Sub

Private Sub pdi_PrintPreview_Click()
    frm_PrintPreview.Show 1
End Sub

Private Sub pdi_PrintSetup_Click()
    frm_Main.cdlg_Printer.ShowPrinter
End Sub

Private Sub pdi_ReplaceRestoredDatabase_Click()
    FileCopy frm_Main.cdlg_Printer.FileName, App.Path & "\Database.mdb"
End Sub

Private Sub pdi_Restore_Click()
    frm_Main.cdlg_Printer.DialogTitle = "Restore"
    frm_Main.cdlg_Printer.Filter = "Microsoft Access Database (*.mdb)|*.mdb"
    frm_Main.cdlg_Printer.ShowOpen
    
    If frm_Main.cdlg_Printer.FileName <> "" Then
        pActiveConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & frm_Main.cdlg_Printer.FileName
        pDatabasePath = frm_Main.cdlg_Printer.FileName
        frm_Main.pdi_ReplaceRestoredDatabase.Enabled = True
        Call pdi_ShowAll_Click
    End If
End Sub

Private Sub pdi_Search_Click()
    frm_Search.Show 1
End Sub

Private Sub pdi_ShowAll_Click()
    Dim v_rsDatabase As New Recordset
    
    v_rsDatabase.Open "SELECT * FROM Report", pActiveConnection
    
    Call ShowReport(v_rsDatabase)
    'v_rsDatabase.Close
End Sub

Private Sub pdii_Backup_Click(Index As Integer)
    pActiveConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & frm_Main.pdii_Backup(Index).Caption
    pDatabasePath = frm_Main.pdii_Backup(Index).Caption
    frm_Main.pdi_ReplaceRestoredDatabase.Enabled = True
    Call pdi_ShowAll_Click
End Sub

Private Sub tbr_Tools_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
             Call pdi_Print_Click
        Case 2:
            Call pdi_PrintPreview_Click
        Case 4:
            Call pdi_Input_Click
        Case 5:
            Call pdi_Search_Click
        Case 6:
            Call pdi_ShowAll_Click
        Case 7:
            Call pdi_ManualEdit_Click
        Case 9:
            Call pdi_Exit_Click
    End Select
End Sub
