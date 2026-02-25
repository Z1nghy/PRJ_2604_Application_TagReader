Attribute VB_Name = "RdTagDrv"
Option Explicit
Dim ret As Integer      ' Scratch integer.
Dim Temp As String      ' Scratch string.
Dim hSimFile As Integer ' Handle of open simulation file.
Dim hProgPromFile As Integer ' Handle of open PROM Program file.


'------------------------------------------------------------------
'------------------------------------------------------------------
'Public Sub comAutoDetect(sciID As SciDef)
Public Function comAutoDetect() As Integer
Dim ncom As Integer
Dim strVer As String
Dim ret As Integer
    
    On Error Resume Next
    Err.Clear
    frmWinInfo.LstDoc.Clear
    
    Screen.MousePointer = vbHourglass       ' Hourglass (wait).
    comAutoDetect = 55
    
    For ncom = 1 To 16
        ret = DoEvents()
        MSC.comID.CommPort = ncom
        BaudRate = 19200
        CommPortSettings = Trim$(Str$(BaudRate)) & ",N,8,1"
        frmWinInfo.LstDoc.AddItem "Com" & ncom   ' Send to win_INFO
        ' Set Differents parameters.
        MSC.comID.InputLen = 1
        ' ....... Set Up the receive buffer
        MSC.comID.InBufferSize = RXBUFFERSIZE
        MSC.comID.InBufferCount = 0
        MSC.comID.RThreshold = 1
        ' ....... Set Up the transmit buffer
        MSC.comID.OutBufferSize = TXBUFFERSIZE
        MSC.comID.OutBufferCount = 0

        MSC.comID.DTREnable = True
        MSC.comID.RTSEnable = True
        MSC.comID.EOFEnable = True

        MSC.comID.Settings = CommPortSettings
        MSC.rxIndex = 0
        MSC.rxUserIndex = 0
        MSC.rxBufLength = 0

        MSC.comID.PortOpen = True
        If Err Then
            'MsgBox "Err = " + Str(Err.Number) + "   " + Error$, vbExclamation
            MSC.comID.PortOpen = False
            MSC.ptrT.Enabled = False
            frmWinInfo.LstDoc.RemoveItem ncom - 1
            frmWinInfo.LstDoc.AddItem "Com" & ncom & " ... KO"  ' Send to win_INFO
            
            Select Case Err
            Case 8002   ' Invalid port number
                Err.Clear
'                Screen.MousePointer = vbDefault ' (Default) Shape determined by the object.
'                Exit Function
'                GoTo XNewPort
            Case 8005   ' Port already open
                Err.Clear
                GoTo XNewPort
            Case 8012   ' Device is not open
                Err.Clear
'                Screen.MousePointer = vbDefault ' (Default) Shape determined by the object.
'                Exit Function
'                GoTo XNewPort
            Case Else
                Err.Clear
'                GoTo XNewPort
           End Select

        End If
        MSC.ptrT.Enabled = True

        Call delayT(0.1)
        Call cleanRxBuffer(MSC)
        strVer = GetVersion()
        If (strVer = "") Then
            Call close_port(MSC)
            frmWinInfo.LstDoc.RemoveItem ncom - 1
            frmWinInfo.LstDoc.AddItem "Com" & ncom & " ... KO"  ' Send to win_INFO
        Else
            If (Left(strVer, 7) = "0,RF125") Then
                frmWinInfo.LstDoc.RemoveItem ncom - 1
                frmWinInfo.LstDoc.AddItem "Com" & ncom & " ... ok"  ' Send to win_INFO
                Call addItem_INFO("Rx: " & strVer)
                comAutoDetect = 0
                Screen.MousePointer = vbDefault     ' (Default) Shape determined by the object.
                Exit Function
            Else
                Call close_port(MSC)
                frmWinInfo.LstDoc.RemoveItem ncom - 1
                frmWinInfo.LstDoc.AddItem "Com" & ncom & " ... KO"  ' Send to win_INFO
                GoTo XNewPort
            End If
        End If

XNewPort:
    Next ncom
    Screen.MousePointer = vbDefault             ' (Default) Shape determined by the object.
              
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Function ctrlTagReaderInit() As Integer
Dim retV As Integer
Dim strVer As String

    retV = 0
    If (ReaderInitOk = False) Then
        retV = 55
        strVer = GetVersion()
        If (strVer = "") Then
            MsgBox "Tag Reader not found !!!", vbCritical, "Tag Read"
            ReaderInitOk = False
            retV = 55
            'Exit Function
        Else
            Call addItem_INFO("Tx: " & strVer)
            ReaderInitOk = True
            retV = 0
        End If
    End If
    ctrlTagReaderInit = retV
    
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Sub ctrlKeyPress(txtCtrl As Variant, KeyAscii As Integer, lenData As Integer)
Dim charP As String
Dim txt As String
Dim s As Long

    With txtCtrl
        If (Len(.Text) < lenData) Then
            Select Case lenData
                Case 2
                    .Text = "00"
                Case 8
                    .Text = "00000000"
                Case Else
            End Select
        End If
        charP = Chr(KeyAscii)
        KeyAscii = Asc(UCase(charP))
        If (((KeyAscii >= &H30) And (KeyAscii <= &H39)) _
            Or ((KeyAscii >= &H41) And (KeyAscii <= &H46))) Then
            txt = .Text
            s = .SelStart + 1
            If (s <= lenData) Then
                Mid(txt, s, 1) = Chr(KeyAscii)
            Else
                Beep
                KeyAscii = 0    ' Suprimme la touche
                Exit Sub
            End If
            .Text = txt
            .SelStart = s
            KeyAscii = 0
        Else
            Beep
            KeyAscii = 0    ' Suprimme la touche
        End If
    End With
    
End Sub

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Sub delayT(aa As Single)
Dim tt As Single
Dim ret As Integer
    tt = Timer
    Do
        If (tt > Timer) Then tt = Timer
        ret = DoEvents()
    Loop Until (Timer - tt) > aa

End Sub


'------------------------------------------------------------------
'   Converts A to 10, B to 11, ..., F to 15
'------------------------------------------------------------------
Public Function hex2int(xvalue As Variant) As Integer
Dim value As Integer
Dim svalue As Variant

    value = Val(xvalue)
    
    If (xvalue = "A") Then value = 10
    If (xvalue = "B") Then value = 11
    If (xvalue = "C") Then value = 12
    If (xvalue = "D") Then value = 13
    If (xvalue = "E") Then value = 14
    If (xvalue = "F") Then value = 15
    hex2int = value
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function ConvertValue(xvalue As Variant) As Integer
Dim Nibble As Variant
Dim msb, lsb As Integer

    Select Case Len(xvalue)
    Case 0:
        msb = 0
        lsb = 0
    Case 1:
        msb = 0
        lsb = hex2int(xvalue)
    Case 2:
        Nibble = Left(xvalue, 1)
        msb = hex2int(Nibble) * 16
    
        Nibble = Right(xvalue, 1)
        lsb = hex2int(Nibble)
    End Select
    
    ConvertValue = msb + lsb

End Function

'------------------------------------------------------------------
'   GetVersion :
'------------------------------------------------------------------
Public Function GetVersion() As Variant
Dim ValStr As String
Dim tx_ok As Boolean

    tx_ok = False
    ValStr = "V?" + Chr(vbKeyReturn)
    'Call addItem_INFO("Tx: " & ValStr)
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        GetVersion = gets(MSC, 2)             ' Get Target response
    Else
        GetVersion = ""
    End If
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Sub addItem_INFO(ValStr As String)
Dim tmp As Variant
Dim tt As Integer
    
    tmp = ValStr
    tt = InStr(1, ValStr, Chr(vbKeyReturn))
    If (tt) Then
        tmp = Mid(ValStr, 1, Len(ValStr) - 1)
    End If
    frmWinInfo.LstDoc.AddItem tmp  ' Send to win_INFO
    frmWinInfo.LstDoc.TopIndex = frmWinInfo.LstDoc.NewIndex
'    frmDocum(win_INFO).LstDoc(win_INFO - 1).AddItem tmp  ' Send to win_INFO
'    frmDocum(win_INFO).LstDoc(win_INFO - 1).TopIndex = _
'        frmDocum(win_INFO).LstDoc(win_INFO - 1).NewIndex

End Sub

Public Sub ErrorActionSet(rV As Integer, txt As String)
    Select Case rV
        Case 3
            TagTransactionError = txt + ": Error in argument(Err=" + Str(rV) + ")"
        Case 4
            TagTransactionError = txt + ": Exec Error (Err=" + Str(rV) + ")"
        Case 5
            TagTransactionError = txt + ": TimeOut Error, no tag detect(Err=" + Str(rV) + ")"
        Case 6
            TagTransactionError = txt + ": R/W Transaction Error(Err=" + Str(rV) + ")"
        Case Else
            TagTransactionError = txt + ": Transaction/no Reader"
    End Select
End Sub


'------------------------------------------------------------------
'   SetTypeTag
'       TypeOfTag   :  in Hexadecimal
'       TypeOfTag   =  0    Temic e5530
'       TypeOfTag   =  1    Temic e5550
'       TypeOfTag   = 10    Micro Electronic Marin H400x / H410x
'       TypeOfTag   = 12    Micro Electronic Marin V4050
'------------------------------------------------------------------
Public Function SetTypeTag(TypeTag As Integer) As Boolean
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "TT," + Hex(TypeTag) + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    SetTypeTag = tx_ok
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        If (xvalue = 0) Then
            SetTypeTag = True
        Else
            SetTypeTag = False
        End If
    Else
        SetTypeTag = False
    End If
    
End Function

'------------------------------------------------------------------
'   SetMagnetic(OnOff)
'       OnOff   =  1    Magnetic ON
'       OnOff   =  0    Magnetic OFF
'------------------------------------------------------------------
Public Function SetMagnetic(OnOff As Integer) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    If (OnOff = 0) Then
        ValStr = "H0" + Chr(vbKeyReturn)
    Else
        ValStr = "H1" + Chr(vbKeyReturn)
    End If
    Call addItem_INFO("Tx: " & ValStr)
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    SetMagnetic = tx_ok
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetMagnetic = xvalue
    Else
        SetMagnetic = 6
        Call addItem_INFO("Rx: " & "6")
    End If
    
End Function

'------------------------------------------------------------------
'   SetPassWord:    PW,tagPassWord<CR>
'       tagPassWord in Hexadecimal : 00000000H .. FFFFFFFFH
'
'       return:    0<CR>
'------------------------------------------------------------------
Public Function SetPassWord(tagPassWord As Variant) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "PW," + tagPassWord + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    'SetPassWord = tx_ok
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetPassWord = xvalue
    Else
        SetPassWord = 6
        Call addItem_INFO("Rx: " & "6")
    End If
    
End Function

'------------------------------------------------------------------
'   SetLogin:    LG,tagPassWord<CR>
'       tagPassWord in Hexadecimal : 00000000H .. FFFFFFFFH
'
'       return:
'           0<CR>   ==> OK, case V4050
'           4<CR>   ==> Wrong PassWord
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function SetLogin(tagPWord As Variant) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String
    
    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "LG," + tagPWord + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    SetLogin = tx_ok
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetLogin = xvalue
    Else
        SetLogin = 6
        Call addItem_INFO("Rx: " & "6")
    End If
    
End Function

'------------------------------------------------------------------
'   writePassWord:    only for V4050
'       WP,CurrentTagPassword,NewTagPossword<CR>
'       CurrentPassword : 00000000H .. FFFFFFFFH, the current password actualy
'                           stored in the current tag.
'       NewPassword     : 00000000H .. FFFFFFFFH, the new value that will replace
'                           the CurrentPassword if authorized.
'
'       return:
'           0<CR>   ==> OK, case V4050
'           4<CR>   ==> Wrong PassWord
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function writePassWord(CurrentPWRD As Variant, NewPWRD As Variant) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "WP," + CurrentPWRD + NewPWRD + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    writePassWord = tx_ok
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
    Else
        writePassWord = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   SetReset:    RS<CR>
'
'       return:
'           0<CR>   ==> OK
'           5<CR>   ==> End of Timeout, or no tag detect
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function SetReset() As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "RS" + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetReset = xvalue
    Else
        SetReset = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   SetTimeOut:    TO,timeout<CR>
'
'       return:
'           0<CR>   ==> OK
'           5<CR>   ==> End of Timeout, or no tag detect
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function SetTimeOut(t_OUT As Integer) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String

    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "TO," + CStr(Hex(t_OUT)) + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetTimeOut = xvalue
    Else
        SetTimeOut = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   SetTryCount:    TC,tcnt<CR>
'
'       return:
'           0<CR>   ==> OK
'           5<CR>   ==> End of Timeout, or no tag detect
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function SetTryCount(tcnt As Integer) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String
    
    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "TC," + CStr(Hex(tcnt)) + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetTryCount = xvalue
    Else
        SetTryCount = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   SetFormatTag:    FM<CR>
'   Format TK5550, V4050
'       return:
'           0<CR>   ==> OK
'           5<CR>   ==> End of Timeout, or no tag detect
'           6<CR>   ==> Tag detected but transaction error
'------------------------------------------------------------------
Public Function SetFormatTag() As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String
    
    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "FM" + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        SetFormatTag = xvalue
    Else
        SetFormatTag = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   standardReadTag:
'       RD<CR>
' Standard read : TT5530, TT5550, TT5551, TT5560, V4050/P4150 ...
' Get Serial Number: Hitag1 & Hitag2 ...
'   return:
'       0,DataStream<CR>    ==> Reading Ok, data followed
'       3<CR>               ==> Error in argument
'       5<CR>               ==> TimeOut error, no tag detect
'       6<CR>               ==> Reading error, error during transaction
'------------------------------------------------------------------
Public Function standardReadTag() As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim RdVariant As Variant
Dim tmp As Long


    xvalue = 0
    tx_ok = False
    RdVariant = ""
    tmp = 0
    Select Case TypeOfTag
        Case TT_TK5530, TT_TK5550, TT_TK5552, TT_H400X, TT_V4050
            ValStr = "RD" + Chr(vbKeyReturn)
        Case TT_HITAG1, TT_HITAG2
            ValStr = "GS" + Chr(vbKeyReturn)
        Case Else

    End Select
    
    Call cleanRxBuffer(MSC)
    dataStream = ""
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    tx_ok = puts(MSC, ValStr)
    tx_ok = True
    If (tx_ok) Then
        RdVariant = gets(MSC, 2)            ' Get Target response
        'Call addItem_INFO("Rx: " & RdVariant)      ' Send to win_INFO
        tmp = Len(RdVariant)
        If (tmp <> 0) Then
            xvalue = Val(Mid(RdVariant, 1, 1))
            If (xvalue = 0) And tmp > 3 Then
                dataStream = Mid(RdVariant, 3, Len(RdVariant) - 3)
                Call addItem_INFO("Rx: " & RdVariant)      ' Send to win_INFO
            Else
                dataStream = ""
                Call addItem_INFO("Rx: " & Str(xvalue))      ' Send to win_INFO
            End If
            standardReadTag = xvalue
        Else
            standardReadTag = 6
            Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
            'Call addItem_INFO("Rx: " & RdVariant)      ' Send to win_INFO
        End If
    Else
        standardReadTag = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   readTag:
'       RD,StartAddress,WordCount<CR>
'
'   return:
'       0,DataStream<CR>    ==> Reading Ok, data followed
'       3<CR>               ==> Error in argument
'       5<CR>               ==> TimeOut error, no tag detect
'       6<CR>               ==> Reading error, error during transaction
'------------------------------------------------------------------
Public Function readTag(StartAddress As Integer, WordCount As Integer) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim RdVariant As Variant
Dim tmp As Long

    dataStream = ""
    xvalue = 0
    tx_ok = False
    RdVariant = ""
    tmp = 0
    ValStr = "RD," + Hex(StartAddress) + "," + Hex(WordCount) + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    Call cleanRxBuffer(MSC)
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        RdVariant = gets(MSC, 2)            ' Get Target response
        tmp = Len(RdVariant)
        If (tmp <> 0) Then
            xvalue = Val(Mid(RdVariant, 1, 1))
            If (xvalue = 0) And (tmp > 3) Then
                dataStream = Mid(RdVariant, 3, Len(RdVariant) - 3)
                Call addItem_INFO("Rx: " & RdVariant)      ' Send to win_INFO
            Else
                dataStream = ""
                Call addItem_INFO("Rx: " & Str(xvalue))      ' Send to win_INFO
            End If
            readTag = xvalue
        Else
            readTag = 6
            Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
            'Call addItem_INFO("Rx: " & RdVariant)      ' Send to win_INFO
        End If
    Else
        readTag = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function

'------------------------------------------------------------------
'   writeTag:
'       WR,dataAddress,DataStream,LocBit<CR>
'       StartAddress    in Hexadecimal 00H .. FFH. Default Value = 0H
'       DataStream      in Hexadecimal datastream
'       LockBit         in Hexadecimal.
'   return:
'       0<CR>    ==> Writing Ok
'       5<CR>    ==> TimeOut error, no tag detect
'       6<CR>    ==> Writing error, error during transaction
'------------------------------------------------------------------
Public Function writeTag(dataAddress As Integer, dataStream As Variant, _
                LockBit As Integer) As Integer
Dim xvalue As Integer
Dim ValStr As String
Dim tx_ok As Boolean
Dim stVal As String
    
    LockBit = 0                 ' Force LockBit to 0 OSSENI. !!!!
'    ValStr = "WR," + Hex(dataAddress) + "," + dataStream + Chr(LockBit) + Chr(vbKeyReturn)
    stVal = ""
    xvalue = 0
    tx_ok = False
    ValStr = "WR," + Hex(dataAddress) + "," + dataStream + Chr(vbKeyReturn)
    Call addItem_INFO("Tx: " & ValStr)      ' Send to win_INFO
    tx_ok = puts(MSC, ValStr)
    If (tx_ok) Then
        stVal = gets(MSC, 2)                ' Get Target response
        xvalue = Val(stVal)
        Call addItem_INFO("Rx: " & stVal)
        writeTag = xvalue
    Else
        writeTag = 6
        Call addItem_INFO("Rx: " & "6")      ' Send to win_INFO
    End If
    
End Function





''------------------------------------------------------------------
'' fname = File to download to target or to verify.
''------------------------------------------------------------------
'Public Sub checkSrecFile(fname As String)
'Dim cnt_i As Integer
'Dim cnt_j As Integer
'Dim ErrFile
'
'    On Error Resume Next
'    ' Open the S-Record file.
'    hSimFile = FreeFile
'    RecordCount = 0
'    Open fname For Input As #hSimFile
'    If Err Then
'        MsgBox Error$, vbExclamation
'        Close #hSimFile
'        hSimFile = 0
'        ok_SFile = False
'        Exit Sub
'    Else
'        '
'        While (Not EOF(hSimFile))
'            Input #hSimFile, record
'            'If Left(record, 2) = "S0" Then GoTo NoS1Records
'            If Left(record, 2) = "S0" Then
'                ' Header record found
'            End If
'            If (Left(record, 2) = "S1") Then
'                RecordCount = RecordCount + 1
'            End If
'            If (Left(record, 2) = "S9") Then
'                If (RecordCount = 0) Then
'                    ok_SFile = False
'                    ' There are no S1 in this file
'                    ErrFile = MsgBox("There are no S1 in file " + fname, vbOKOnly)
'                Else
'                    ' GoTo SendFile: it's ok to send file
'                    ok_SFile = True
'                End If
'
'            Else
'            End If
'            ' This file contains one or more none S19 records.
'            ' Debug.Print record
'        Wend
'        '
'        Close #hSimFile
'        ' Debug.Print "RecordCount = " & RecordCount
'    End If
'    '
'End Sub


