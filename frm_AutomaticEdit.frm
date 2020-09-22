VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_AutomaticEdit 
   Caption         =   "Automatic Edit"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   Icon            =   "frm_AutomaticEdit.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dat_Data 
      Caption         =   "Data"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   1740
   End
   Begin MSDBGrid.DBGrid dbg_List 
      Bindings        =   "frm_AutomaticEdit.frx":08CA
      Height          =   5295
      Left            =   113
      OleObjectBlob   =   "frm_AutomaticEdit.frx":08E1
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frm_AutomaticEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frm_AutomaticEdit.dat_Data.DatabaseName = pDatabasePath
    frm_AutomaticEdit.dat_Data.RecordSource = "Report"
End Sub

Private Sub Form_Resize()
    frm_AutomaticEdit.dbg_List.Width = frm_AutomaticEdit.ScaleWidth - 240
    If frm_AutomaticEdit.Height - 600 > 0 Then
        frm_AutomaticEdit.dbg_List.Height = frm_AutomaticEdit.Height - 600
    End If
    frm_AutomaticEdit.dbg_List.Left = 120
    frm_AutomaticEdit.dbg_List.Top = 120
End Sub
