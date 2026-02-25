Attribute VB_Name = "ConfigFile"
Option Explicit
'Dim ret As Integer      ' Scratch integer.
'Dim Temp As String      ' Scratch string.

'------------------------------------------------------------------
'------------------------------------------------------------------
'------------------------------------------------------------------
'------------------------------------------------------------------
Function GetFileLn(fname As String, section, key)
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
   
    worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), fname)
    If worked = 0 Then
        GetFileLn = ""
    Else
        GetFileLn = Left(retVal, InStr(retVal, Chr(0)) - 1)
    End If
    
End Function

'------------------------------------------------------------------
'   Write Configuration File
'------------------------------------------------------------------
Public Sub write_Config_File(fname As String)
Dim worked As Integer
'
'    worked = WritePrivateProfileString("COMDEF", "COMPORT", CommPort, fname)
'    worked = WritePrivateProfileString("COMDEF", "BAUD", Str$(BaudRate), fname)
'    worked = WritePrivateProfileString("STRINGSREF", "VER", StrVersion, fname)
'    worked = WritePrivateProfileString("SRECFILE", "SFILE", progFileName, fname)
''    worked = WritePrivateProfileString("COMDEF", "COMPORT", Str$(CommPort), fname)
''    worked = WritePrivateProfileString("STRINGSREF", "VER", Str$(StrVersion), fname)
''    worked = WritePrivateProfileString("SRECFILE", "SFILE", Str$(progFileName), fname)
'
End Sub

'------------------------------------------------------------------
'   Read Configuration File
'------------------------------------------------------------------
Public Sub read_Config_File(fname As String)
    
'    CommPort = GetFileLn(fname, "COMDEF", "COMPORT")
'    BaudRate = Val(GetFileLn(fname, "COMDEF", "BAUD"))
'    StrVersion = GetFileLn(fname, "STRINGSREF", "VER")
'    progFileName = GetFileLn(fname, "SRECFILE", "SFILE")
        
End Sub

''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub mnuNewFile_Click()
'
'    Dim replace
'    Dim filenm As String
'    Dim StrDefault, StrTitle, StrPrompt
'    Dim msg As String
'
'    NewFileFlag = False
'    Call loadfrmStimulation
'
'    If NewFileFlag Then
'
'        On Error Resume Next
'        'init_Electrodes_Connection
'        'CmDialog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
'        CmDialog.Flags = cdlOFNOverwritePrompt Or cdlOFNExplorer
'
'        CmDialog.CancelError = True
'
'        ' Get the filename from the user.
'        CmDialog.DialogTitle = "HC11 memory Download: Configuration File"
'        CmDialog.Filter = "Configuration Files (*.INI)|*.ini|All Files (*.*)|*.*"
'
'        CmDialog.filename = ""
'        CmDialog.ShowSave
'
'        If Err = cdlCancel Then Exit Sub
'        Temp = CmDialog.filename
'
'        Call write_File(Temp)
'
'    End If
'
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub mnuOpenFile_Click()
'    Dim replace
'
'    On Error Resume Next
'    'init_Electrodes_Connection
'    CmDialog.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer
'    CmDialog.CancelError = True
'
'    ' Get the filename from the user.
'    CmDialog.DialogTitle = "HC11 Download: Open an existing Configuration File"
'    CmDialog.Filter = "Configuration Files (*.INI)|*.ini|All Files (*.*)|*.*"
'
'    CmDialog.filename = ""
'    CmDialog.ShowOpen
'    If Err = cdlCancel Then Exit Sub
'    Temp = CmDialog.filename
'
'    Call read_File(Temp)
'
'    'NewFileFlag = False
'    'Call loadfrmStimulation
'
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'' Toggles the state of the port (open or closed).
'Private Sub mnuPortOpen_Click()
'    On Error Resume Next
'    Dim OpenFlag
'
'    On Error Resume Next
'    Call open_port(MSC)
'    If (MSComm1.PortOpen) Then
'        imgConnected.ZOrder
'        sbrStatus.Panels("Settings").Text = "Settings: " & MSComm1.Settings
'        StartTiming
'    Else
'        imgNotConnected.ZOrder
'        sbrStatus.Panels("Settings").Text = "Settings: "
'        StopTiming
'End If
'    mnuPortOpen.Checked = MSComm1.PortOpen
'
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub mnuProperties_Click()
'  ' Show the CommPort properties form
'    frmProperties.Show vbModal
'End Sub
'
'Private Sub loadfrmStimulation()
'    ' Show the CommPort properties form
'    ' ActiveValidate = True
'    FrmFile.Show vbModal
'End Sub
'
'Private Sub mnuSave_Click()
'    Call write_File(Temp)
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'
'Private Sub tbrToolbar_ButtonClick(ByVal Button As ComctlLib.Button)
'    Select Case Button.key
'        Case "tbrNewFile"
'            Call mnuNewFile_Click
'        Case "tbrOpenFile"
'            Call mnuOpenFile_Click
'        Case "tbrSaveFile"
'            Call mnuSave_Click
'        Case "tbrCommPort"
'            Call mnuPortOpen_Click
'        Case "tbrComProperties"
'            Call mnuProperties_Click
'    End Select
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub SCITimer1_Timer()
'    'imgConnected.Visible = Not imgConnected.Visible
'    'imgNotConnected.Visible = Not imgNotConnected.Visible
'    Call putch(MSC, "a")
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub Timer2_Timer()
'    sbrStatus.Panels("ConnectTime").Text = Format(Time, "hh:nn") & " "
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub Form_Unload(Cancel As Integer)
'    Dim Counter As Long
'
'    If MSC.comID.PortOpen Then
'       ' Wait 10 seconds for data to be transmitted.
'       Counter = Timer + 10
'       Do While MSC.comID.OutBufferCount
'          ret = DoEvents()
'          If Timer > Counter Then
'             Select Case MsgBox("Data cannot be sent", 34)
'                ' Cancel.
'                Case 3
'                   Cancel = True
'                   Exit Sub
'                ' Retry.
'                Case 4
'                   Counter = Timer + 10
'                ' Ignore.
'                Case 5
'                   Exit Do
'             End Select
'          End If
'       Loop
'
'       MSC.comID.PortOpen = False
'    End If
'    Unload Me
'End Sub
'
''------------------------------------------------------------------
''------------------------------------------------------------------
'
'Private Sub Form_Load()
'
'    Dim CommPort As String
'    Dim Handshaking As String
'    Dim Settings As String
'
'    On Error Resume Next
'
'    '  ........     Set comport and sci timer
'    Set MSC.comID = MSComm1
'    Set MSC.ptrT = SCITimer1
'
'    ' Set Title
'    App.Title = "BioCELL application"
'
'    ' Center Form
'    frmBioCell.Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
'
'    ' Load Registry Settings
'    Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
'    If Settings <> "" Then
'        MSComm1.Settings = Settings
'        If Err Then
'            MsgBox Error$, vbExclamation
'            Exit Sub
'        End If
'    End If
'
'    CommPort = GetSetting(App.Title, "Properties", "CommPort", "") ' frmTerminal.MSComm1.CommPort
'    If CommPort <> "" Then MSComm1.CommPort = CommPort
'
'    Handshaking = GetSetting(App.Title, "Properties", "Handshaking", "") 'frmTerminal.MSComm1.Handshaking
'    If Handshaking <> "" Then
'        MSComm1.Handshaking = Handshaking
'        If Err Then
'            MsgBox Error$, vbExclamation
'            Exit Sub
'        End If
'    End If
'
'    Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
'    On Error GoTo 0
'
'    ' Set up status indicator light
'    imgNotConnected.ZOrder
'
'    ' Set up Status Panel Width
'    sbrStatus.Panels("Status").Width = 2 * ScaleWidth / 6
'    sbrStatus.Panels("Settings").Width = 3 * ScaleWidth / 6
'    sbrStatus.Panels("ConnectTime").Width = ScaleWidth / 6
'
'    Timer2.Enabled = True
'    sbrStatus.Panels("ConnectTime").Text = Format(Time, "hh:nn") & " "
'
'    ' Position the status indicator light
'    imgConnected.Left = ScaleWidth - imgConnected.Width * 3.1
'    imgNotConnected.Left = ScaleWidth - imgNotConnected.Width * 3.1
'
'    ' .................
'    MuxName(0) = "MUX0"
'    MuxName(1) = "MUX1"
'    MuxName(2) = "MUX2"
'    MuxName(3) = "MUX3"
'
'    ChanName(0) = "CH0"
'    ChanName(1) = "CH1"
'    ChanName(2) = "CH2"
'    ChanName(3) = "CH3"
'    ChanName(4) = "CH4"
'    ChanName(5) = "CH5"
'    ChanName(6) = "CH6"
'    ChanName(7) = "CH7"
'    ChanName(8) = "S_+"
'    ChanName(9) = "S_-"
'    ChanName(10) = "GND"
'    ChanName(11) = "RES"
'    ' .................
'
'
'End Sub
''------------------------------------------------------------------
''------------------------------------------------------------------

