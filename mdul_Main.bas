Attribute VB_Name = "mdul_Main"
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const MF_BITMAP = &H4&

Type NOTIFYICONDATA
    cbSize           As Long
    hwnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type

Type TBackup
    AutoBackup As Boolean
    BackupRate As Integer
    BackupFileName As String
    BackupFileNo As Integer
    BackupNo As Integer
    StartYear As Integer
    StartMonth As Integer
    StartDay As Integer
    BackupYear As Integer
    BackupMonth As Integer
    BackupDay As Integer
End Type

Public nfIconData As NOTIFYICONDATA
Public pActiveConnection As String
Public pDatabase As New Recordset
Public pQuery As String
Public pQueryTemp As String
Public pDatabasePath As String
Public pBackup As TBackup
Public pShowLogon As Boolean
Public pUserName As String
Public pPassword As String
Public pShowShellNotifyIcon As Boolean

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
    
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Sub Main()
    pActiveConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Database.mdb"
    pDatabasePath = App.Path & "\Database.mdb"
    
    If InStr(Command, FileName) > 0 Then
    
        Select Case UCase(Left(Command, 2))
            Case "/C"
                Call LoadOption
                Call CheckForBackup
                End
        End Select
    
    End If
    
    With nfIconData
        .hwnd = frm_Input.hwnd
        .uID = frm_Input.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm_Input.Icon.Handle
        .szTip = "Kabrian Report" & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
        
    Call AddIconToMenu
    Call LoadOption
    Call CheckForBackup
    If pShowLogon = True Then
        frm_Logon.Show
    Else
        frm_Main.Show
    End If
    
    If pShowShellNotifyIcon = True Then
        frm_Main.Hide
    End If
End Sub

Public Sub AddIconToMenu()
    Dim v_lMenuHnd    As Long
    Dim v_lSubMenuHnd As Long
    Dim v_lMenuCnt    As Long
    Dim v_lSubMenuCnt As Long
    Dim v_lSubMenuID  As Long

    v_lMenuHnd = GetMenu(frm_Main.hwnd)
    v_lMenuCnt = GetMenuItemCount(lMenuHnd)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 0)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(1).Picture, frm_Main.iml_Menu.ListImages(1).Picture)
        
    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 1)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(2).Picture, frm_Main.iml_Menu.ListImages(2).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 3)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(3).Picture, frm_Main.iml_Menu.ListImages(3).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 5)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(4).Picture, frm_Main.iml_Menu.ListImages(4).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 7)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(5).Picture, frm_Main.iml_Menu.ListImages(5).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 0)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 9)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(6).Picture, frm_Main.iml_Menu.ListImages(6).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 1)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 0)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(7).Picture, frm_Main.iml_Menu.ListImages(7).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 1)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 1)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(8).Picture, frm_Main.iml_Menu.ListImages(8).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 2)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 0)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BYCOMMAND, frm_Main.iml_Menu.ListImages(9).Picture, frm_Main.iml_Menu.ListImages(9).Picture)

    v_lSubMenuHnd = GetSubMenu(v_lMenuHnd, 2)
    v_lSubMenuID = GetMenuItemID(v_lSubMenuHnd, 3)
    Call SetMenuItemBitmaps(v_lMenuHnd, v_lSubMenuID, MF_BITMAP, frm_Main.iml_Menu.ListImages(10).Picture, frm_Main.iml_Menu.ListImages(10).Picture)
End Sub

Public Sub ShowReport(m_Recordset As Recordset)
    Dim v_iIndex As Integer
    
    frm_Main.lvw_List.ListItems.Clear
    
    If Not m_Recordset.BOF Then
    
    m_Recordset.MoveFirst
    While Not m_Recordset.EOF
       v_iIndex = v_iIndex + 1
       frm_Main.lvw_List.ListItems.Add , , m_Recordset.Fields!Refrence
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 1, , m_Recordset.Fields!Consignee
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 2, , m_Recordset.Fields!Liner
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 3, , m_Recordset.Fields!DateOfRequest
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 4, , m_Recordset.Fields!DateOfRelease
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 5, , m_Recordset.Fields!DateOfReturn
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 6, , m_Recordset.Fields!DateOfDeparture
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 7, , m_Recordset.Fields!MasterBLNo
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 8, , m_Recordset.Fields!HouseBLNo
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 9, , m_Recordset.Fields!BLDate
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 10, , m_Recordset.Fields!BuyingRate
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 11, , m_Recordset.Fields!SellingRate
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 12, , m_Recordset.Fields!Profit
       frm_Main.lvw_List.ListItems(v_iIndex).ListSubItems.Add 13, , m_Recordset.Fields!Result
       If m_Recordset.Fields!Result = "Checked" Then
          frm_Main.lvw_List.ListItems(v_iIndex).Checked = True
       Else
          frm_Main.lvw_List.ListItems(v_iIndex).Checked = False
       End If
       m_Recordset.MoveNext
    Wend
    End If
    
    m_Recordset.Close
End Sub

Public Sub MakePrintPreview(Optional m_PageNo As Integer)
    Dim v_iLoop As Integer
    Dim v_iRecordsPerPage As Integer
    Dim v_iPrintPageCount As Integer
    Dim v_iIndex As Integer
    Dim v_iRecordLoop As Integer
    
    On Error GoTo Error
    
    v_iRecordsPerPage = 50
    With frm_PrintPreview
        .pic_Preview.Font = "Tahoma"
        .pic_Preview.ForeColor = &H0&
        
        .pic_Preview.Cls
        .pic_Preview.Font.Bold = True
        .pic_Preview.CurrentX = 700
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Rf"
        .pic_Preview.CurrentX = 1100
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Shipper"
        .pic_Preview.CurrentX = 2400
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Liner"
        .pic_Preview.CurrentX = 3400
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Date Of Req."
        .pic_Preview.CurrentX = 4600
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Date Of Rel."
        .pic_Preview.CurrentX = 5800
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Date Of Ret."
        .pic_Preview.CurrentX = 7000
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Date Of Dep."
        .pic_Preview.CurrentX = 8200
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Master B/L No."
        .pic_Preview.CurrentX = 9600
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "House B/L No."
        .pic_Preview.CurrentX = 10900
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "B/L Date"
        .pic_Preview.CurrentX = 11800
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Buying Rate"
        .pic_Preview.CurrentX = 13000
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Selling Rate"
        .pic_Preview.CurrentX = 14200
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Profit"
        .pic_Preview.CurrentX = 15200
        .pic_Preview.CurrentY = 800
        .pic_Preview.Print "Result"
        
        .pic_Preview.Line (700, 1000)-(.pic_Preview.Width - 700, 1000)
        
        .pic_Preview.Font.Bold = False
    
        If m_PageNo = 0 Then
            v_iIndex = 0
        Else
            v_iIndex = (m_PageNo - 1) * v_iRecordsPerPage
        End If
        
        v_iRecordLoop = v_iRecordsPerPage + v_iIndex
        If v_iRecordLoop >= frm_Main.lvw_List.ListItems.Count Then
            v_iRecordLoop = frm_Main.lvw_List.ListItems.Count
        End If
                
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 700
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).Text
        Next v_iLoop
        
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 1100
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(1).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 2400
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(2).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 3400
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(3).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 4600
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(4).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 5800
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(5).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 7000
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(6).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 8200
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(7).Text
        Next v_iLoop
        
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 9600
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(8).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 10900
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(9).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 11800
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(10).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 13000
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(11).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 14200
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(12).Text
        Next v_iLoop
    
        .pic_Preview.CurrentY = 1100
        For v_iLoop = v_iIndex + 1 To v_iRecordLoop
            .pic_Preview.CurrentX = 15200
            .pic_Preview.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(13).Text
        Next v_iLoop
    End With
    Exit Sub
    
Error:
    MsgBox Err.Source, vbCritical
End Sub

Public Sub PrintPage(m_PageNo As Integer)
    Dim v_iLoop As Integer
    Dim v_iRecordsPerPage As Integer
    Dim v_iPrintPageCount As Integer
    Dim v_iIndex As Integer
    Dim v_iRecordLoop As Integer
    
    On Error GoTo Error
    
    v_iRecordsPerPage = 50
    
    Printer.Font = "Tahoma"
    Printer.ForeColor = &H0&
        
    Printer.Font.Bold = True
    Printer.CurrentX = 700
    Printer.CurrentY = 800
    Printer.Print "Rf"
    Printer.CurrentX = 1100
    Printer.CurrentY = 800
    Printer.Print "Shipper"
    Printer.CurrentX = 2400
    Printer.CurrentY = 800
    Printer.Print "Liner"
    Printer.CurrentX = 3400
    Printer.CurrentY = 800
    Printer.Print "Date Of Req."
    Printer.CurrentX = 4600
    Printer.CurrentY = 800
    Printer.Print "Date Of Rel."
    Printer.CurrentX = 5800
    Printer.CurrentY = 800
    Printer.Print "Date Of Ret."
    Printer.CurrentX = 7000
    Printer.CurrentY = 800
    Printer.Print "Date Of Dep."
    Printer.CurrentX = 8200
    Printer.CurrentY = 800
    Printer.Print "Master B/L No."
    Printer.CurrentX = 9600
    Printer.CurrentY = 800
    Printer.Print "House B/L No."
    Printer.CurrentX = 10900
    Printer.CurrentY = 800
    Printer.Print "B/L Date"
    Printer.CurrentX = 11800
    Printer.CurrentY = 800
    Printer.Print "Buying Rate"
    Printer.CurrentX = 13000
    Printer.CurrentY = 800
    Printer.Print "Selling Rate"
    Printer.CurrentX = 14200
    Printer.CurrentY = 800
    Printer.Print "Profit"
    Printer.CurrentX = 15200
    Printer.CurrentY = 800
    Printer.Print "Result"
        
    Printer.Line (700, 1000)-(Printer.Width - 700, 1000)
        
    Printer.Font.Bold = False
    
    v_iIndex = (m_PageNo - 1) * v_iRecordsPerPage
        
    v_iRecordLoop = v_iRecordsPerPage + v_iIndex
    If v_iRecordLoop >= frm_Main.lvw_List.ListItems.Count Then
        v_iRecordLoop = frm_Main.lvw_List.ListItems.Count
    End If
                
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 700
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).Text
    Next v_iLoop
        
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 1100
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(1).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 2400
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(2).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 3400
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(3).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 4600
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(4).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 5800
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(5).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 7000
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(6).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 8200
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(7).Text
    Next v_iLoop
        
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 9600
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(8).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 10900
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(9).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 11800
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(10).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 13000
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(11).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 14200
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(12).Text
    Next v_iLoop
    
    Printer.CurrentY = 1100
    For v_iLoop = v_iIndex + 1 To v_iRecordLoop
        Printer.CurrentX = 15200
        Printer.Print frm_Main.lvw_List.ListItems(v_iLoop).ListSubItems(13).Text
    Next v_iLoop
    Exit Sub
    
Error:
    MsgBox Err.Source, vbCritical
End Sub

Public Function TrimString(Str As String) As String
    Str = RTrim$(Str)
    Str = Left(Str, Len(Str) - 1)
    TrimString = Str
End Function

Public Sub WriteOption()
    If frm_Option.cbx_AutoBackup.Value = 0 Then
        WritePrivateProfileString "Option", "Auto Backup", "False", App.Path & "\Settings.ini"
    Else
        WritePrivateProfileString "Option", "Auto Backup", "True", App.Path & "\Settings.ini"
    End If
    
    WritePrivateProfileString "Option", "Backup Rate", Str(frm_Option.cbo_Backup.ListIndex), App.Path & "\Settings.ini"
    WritePrivateProfileString "Option", "Auto Backup File(s) Name", frm_Option.tbx_AutoBackupFileName, App.Path & "\Settings.ini"
    WritePrivateProfileString "Option", "Max Auto Backup File(s) No", frm_Option.tbx_MaxAutoBackupFileNo, App.Path & "\Settings.ini"
    WritePrivateProfileString "Option", "Start Date", Format(Now, "mm/dd/yyyy"), App.Path & "\Settings.ini"

    If frm_Option.cbx_ShowLogon.Value = 0 Then
        WritePrivateProfileString "Option", "Show Logon", "False", App.Path & "\Settings.ini"
    Else
        WritePrivateProfileString "Option", "Show Logon", "True", App.Path & "\Settings.ini"
    End If
    
    WritePrivateProfileString "Logon", "User Name", frm_Option.tbx_UserName, App.Path & "\Settings.ini"
    WritePrivateProfileString "Logon", "Password", frm_Option.tbx_Password, App.Path & "\Settings.ini"

    If frm_Option.cbx_ShowShellNotifyIcon.Value = 1 Then
        WritePrivateProfileString "Option", "Show Shell Notify Icon", "True", App.Path & "\Settings.ini"
    Else
        WritePrivateProfileString "Option", "Show Shell Notify Icon", "False", App.Path & "\Settings.ini"
    End If
End Sub

Public Sub LoadOption()
    Dim v_sString As String
    
    v_sString = Space(255)
    GetPrivateProfileString "Option", "Auto Backup", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    If TrimString(v_sString) = "True" Then
        frm_Option.cbx_AutoBackup.Value = 1
        pBackup.AutoBackup = True
    Else
        frm_Option.cbx_AutoBackup.Value = 0
        pBackup.AutoBackup = False
    End If
    
    v_sString = Space(255)
    GetPrivateProfileString "Option", "Backup Rate", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pBackup.BackupRate = Val(TrimString(v_sString))
    frm_Option.cbo_Backup.ListIndex = pBackup.BackupRate

    v_sString = Space(255)
    GetPrivateProfileString "Option", "Auto Backup File(s) Name", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pBackup.BackupFileName = TrimString(v_sString)
    frm_Option.tbx_AutoBackupFileName.Text = pBackup.BackupFileName

    v_sString = Space(255)
    GetPrivateProfileString "Option", "Max Auto Backup File(s) No", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pBackup.BackupFileNo = Val(TrimString(v_sString))
    frm_Option.tbx_MaxAutoBackupFileNo.Text = pBackup.BackupFileNo

    v_sString = Space(255)
    GetPrivateProfileString "Backup", "Backup No", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pBackup.BackupNo = Val(TrimString(v_sString))

    v_sString = Space(255)
    GetPrivateProfileString "Option", "Start Date", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    v_sString = TrimString(v_sString)
    pBackup.StartYear = Right(v_sString, 4)
    pBackup.StartMonth = Left(v_sString, 2)
    pBackup.StartDay = Mid(v_sString, 4, 2)
    pBackup.BackupYear = pBackup.StartYear
    pBackup.BackupMonth = pBackup.StartMonth
    pBackup.BackupDay = pBackup.StartDay

    v_sString = Space(255)
    GetPrivateProfileString "Option", "Show Logon", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    If TrimString(v_sString) = "True" Then
        pShowLogon = True
        frm_Option.cbx_ShowLogon.Value = 1
        frm_Option.tbx_UserName.Enabled = True
        frm_Option.tbx_Password.Enabled = True
    Else
        pShowLogon = False
        frm_Option.cbx_ShowLogon.Value = 0
        frm_Option.tbx_UserName.Enabled = False
        frm_Option.tbx_Password.Enabled = False
    End If
    
    v_sString = Space(255)
    GetPrivateProfileString "Logon", "User Name", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pUserName = TrimString(v_sString)
    frm_Option.tbx_UserName.Text = pUserName

    v_sString = Space(255)
    GetPrivateProfileString "Logon", "Password", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    pPassword = TrimString(v_sString)
    frm_Option.tbx_Password.Text = pPassword

    v_sString = Space(255)
    GetPrivateProfileString "Option", "Show Shell Notify Icon", "", v_sString, Len(v_sString), App.Path & "\Settings.ini"
    If TrimString(v_sString) = "True" Then
        frm_Option.cbx_ShowShellNotifyIcon.Value = 1
        pShowShellNotifyIcon = True
    Else
        frm_Option.cbx_ShowShellNotifyIcon.Value = 0
        pShowShellNotifyIcon = False
    End If
End Sub

Public Sub CheckForBackup()
    Dim v_iYear, v_iMonth, v_iDay As Integer
    Dim v_iHour, v_iMinute, v_iSecond As Integer
    Dim v_sTemp As String
    Dim v_iLoop As Integer
    
    v_sTemp = Format(Now, "hh:mm:ss")
    v_iHour = Val(Left(v_sTemp, 2))
    v_iMinute = Val(Mid(v_sTemp, 4, 2))
    v_iSecond = Val(Mid(v_sTemp, 7, 2))
    
    v_sTemp = Format(Now, "mm:dd:yyyy")
    v_iYear = Val(Right(v_sTemp, 4))
    v_iMonth = Val(Left(v_sTemp, 2))
    v_iDay = Val(Mid(v_sTemp, 4, 2))
    
    Call SetBackupDateTime
    If (pBackup.BackupYear <= v_iYear) And (pBackup.StartMonth <= v_iMonth) _
    And (pBackup.BackupDay <= v_iDay) And (pBackup.AutoBackup = True) Then
        If pBackup.BackupNo = pBackup.BackupFileNo Then
            pBackup.BackupNo = 0
        End If
        pBackup.BackupNo = pBackup.BackupNo + 1
        WritePrivateProfileString "Backup", "Backup No", Str(pBackup.BackupNo), App.Path & "\Settings.ini"
        FileCopy App.Path & "\Database.mdb", App.Path & "\" & pBackup.BackupFileName & Format(pBackup.BackupNo, "00") & ".mdb"
        WritePrivateProfileString "Option", "Start Date", Format(Now, "mm/dd/yyyy"), App.Path & "\Settings.ini"
    End If
    
    If pBackup.BackupNo = 0 Then
        frm_Main.pdi_BackupsHistory.Visible = False
    End If
    
    For v_iLoop = 1 To pBackup.BackupFileNo - 1
        Load frm_Main.pdii_Backup(v_iLoop)
        frm_Main.pdii_Backup(v_iLoop).Visible = False
    Next v_iLoop

    If pBackup.BackupNo <> 0 Then
        For v_iLoop = 0 To pBackup.BackupNo - 1
            frm_Main.pdii_Backup(v_iLoop).Caption = App.Path & "\" & pBackup.BackupFileName & Format(pBackup.BackupNo, "00") & ".mdb"
        Next v_iLoop
    End If
End Sub

Public Sub SetBackupDateTime()
    Dim v_iDayTemp, v_iMonthTemp As Integer

    Select Case pBackup.BackupRate
        Case 0:
            v_iDayTemp = 1
        Case 1:
            v_iDayTemp = 2
        Case 2:
            v_iDayTemp = 3
        Case 3:
            v_iDayTemp = 5
        Case 4:
            v_iDayTemp = 10
    End Select
    
    If pBackup.StartMonth = 2 Then
        pBackup.BackupDay = pBackup.StartDay + v_iDayTemp
        If pBackup.BackupDay > 28 Then pBackup.BackupDay = pBackup.BackupDay - 28
    ElseIf (pBackup.StartMonth = 4) Or (pBackup.StartMonth = 6) _
        Or (pBackup.StartMonth = 9) Or (pBackup.StartMonth = 11) Then
        pBackup.BackupDay = pBackup.StartDay + v_iDayTemp
    If pBackup.StartDay > 30 Then pBackup.BackupDay = pBackup.BackupDay - 30
    Else
        pBackup.BackupDay = pBackup.StartDay + v_iDayTemp
        If pBackup.BackupDay > 31 Then pBackup.BackupDay = pBackup.BackupDay - 31
    End If
    
    Select Case pBackup.BackupRate
        Case 5:
            v_iMonthTemp = 1
        Case 6:
            v_iMonthTemp = 2
        Case 7:
            v_iMonthTemp = 3
        Case 8:
            v_iMonthTemp = 6
    End Select
    
    pBackup.BackupMonth = pBackup.StartMonth + v_iMonthTemp
    If pBackup.BackupMonth > 12 Then
        pBackup.BackupMonth = pBackup.BackupMonth - 12
    End If
    
    If pBackup.BackupRate = 9 Then
        pBackup.BackupYear = pBackup.StartYear + 1
    End If
End Sub
