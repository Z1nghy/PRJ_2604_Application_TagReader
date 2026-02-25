Attribute VB_Name = "RS232"
Option Explicit
'------------------------------------------------------------------
'---- Sélectionne tout le contenu du contrôle
'------------------------------------------------------------------
Public Sub SetSel(ctl As Control)
    ctl.SelStart = 0
    ctl.SelLength = Len(ctl.Text)
End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Sub open_port(sciID As SciDef)

    On Error Resume Next

    With sciID
        If (Not .comID.PortOpen) Then
            ' The InputLen property, which determines how
            ' many bytes of data are read each time Input is used
            ' to retrieve data from the input buffer.
            ' Setting .InputLen to 0 specifies that
            ' the entire contents of the buffer should be read.
            .comID.InputLen = 1
        
            ' ....... Set Up the receive buffer
            .comID.InBufferSize = RXBUFFERSIZE
            .comID.InBufferCount = 0
            .comID.RThreshold = 1
            ' ....... Set Up the transmit buffer
            .comID.OutBufferSize = TXBUFFERSIZE
            .comID.OutBufferCount = 0
            
            .comID.DTREnable = True
            .comID.RTSEnable = True
            .comID.EOFEnable = True

            .comID.Settings = CommPortSettings
            '.Settings = Trim$(frmProperties.cboSpeed.Text) _
            '            & "," & Left$(frmProperties.cboParity.Text, 1) _
            '            & "," & Trim$(frmProperties.cboDataBits.Text) _
            '            & "," & Trim$(frmProperties.cboStopBits.Text)
                        
            .rxIndex = 0
            .rxUserIndex = 0
            .rxBufLength = 0

            .comID.PortOpen = True

            If Err Then
                'MsgBox Error$, vbExclamation
                MsgBox "Err = " + Str(Err.Number) + "   " + Error$, vbExclamation
                .comID.PortOpen = False
                .ptrT.Enabled = False
                Exit Sub
            End If
            ' Set Comm Port Timer to form timer
            'Set ptrTimer = ptrT
            'ptrTimer.Enabled = True
            .ptrT.Enabled = True
        Else
            .comID.PortOpen = Not .comID.PortOpen
            If Err Then
                MsgBox Error$, vbExclamation
                .comID.PortOpen = False
            End If
            .ptrT.Enabled = False
        End If
    End With
            
End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Sub close_port(sciID As SciDef)  'comID As MSComm, RxState As SciBuff, ptrT As Timer)

    On Error Resume Next
    With sciID
        If .comID.PortOpen Then .comID.PortOpen = False
        .comID.InBufferCount = 0
        .comID.OutBufferCount = 0
        .rxIndex = 0
        .rxUserIndex = 0
        .rxBufLength = 0
        .ptrT.Enabled = False
    End With
    
End Sub

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function putch_timeout(sciID As SciDef, ch$, timeout As Integer) As Boolean
    Dim ttox As Long, ret As Integer
    
    With sciID
        If (.comID.PortOpen) Then
            'ttox = .ptrT + timeout
            ' Transmit the char.
            .comID.Output = ch$
            If Err Then
                putch_timeout = False
                Exit Function
            End If
      
            ' Wait for all the data to be sent.
            ttox = Timer + timeout
            Do
                ret = DoEvents()
                If timeout Then
                    'If .ptrT > ttox Then
                    If Timer > ttox Then
                        putch_timeout = False
                        Exit Function
                    End If
                End If
            Loop Until .comID.OutBufferCount = 0
            putch_timeout = True
        Else
            putch_timeout = False
        End If
    End With
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function putch(sciID As SciDef, ch$) As Boolean
    putch = putch_timeout(sciID, ch$, 1)
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function rsputs(sciID As SciDef, sstr As String) As Boolean
    Dim ll As Integer, ret As Integer
    Dim cnt As Integer
    Dim stt As Boolean
    
    ll = Len(sstr)
    cnt = 1
    Do
        ret = DoEvents()
        stt = putch(sciID, Mid(sstr, cnt, 1))
        If (stt) Then
            ll = ll - 1
            cnt = cnt + 1
        End If
    Loop Until (ll = 0) Or (Not stt)
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function puts(sciID As SciDef, strO As String) As Boolean
    Dim ret As Integer
    
    On Error Resume Next
    If (sciID.comID.PortOpen) Then
        ' Send the string
        sciID.comID.Output = strO
        If Err Then GoTo err_puts
         ' Wait for all the data to be sent.
         Do
            ret = DoEvents()
         Loop Until sciID.comID.OutBufferCount = 0
    End If
    puts = True
    GoTo fin_puts
err_puts:
    puts = False
fin_puts:

End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function getch(sciID As SciDef) As Integer

    If sciID.comID.PortOpen Then
        If (sciID.rxBufLength <> 0) Then
            getch = sciID.rxbuffer(sciID.rxUserIndex)
            If (sciID.rxUserIndex >= RXBUFFERSIZE - 1) Then
                sciID.rxUserIndex = 0
            Else
                sciID.rxUserIndex = sciID.rxUserIndex + 1
            End If
            sciID.rxBufLength = sciID.rxBufLength - 1
        Else
            getch = -1
        End If
    Else
        getch = -1
    End If
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function getchar(sciID As SciDef) As Integer
Dim ccr As Integer
Dim ret As Integer

    If sciID.comID.PortOpen Then
        ' Wait for data to come back to the serial port.
        endTimeOut = Timer + 10          ' timeout de 10 sec.
        'endTimeOut = sciID.ptrT + 2      ' timeout de 2 sec.
        Do
            ret = DoEvents()
            ccr = getch(sciID)
            If (ccr <> -1) Then
                getchar = ccr
                Exit Do
            End If
        Loop Until (Timer >= endTimeOut)
    End If
    
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Function gets(sciID As SciDef, tTime As Integer) As String
Dim ret As Integer, ccr As Integer
    
    On Error Resume Next
    gets = ""
    If sciID.comID.PortOpen Then
        ' Wait for data to come back to the serial port.
        endTimeOut = Timer + tTime          ' timeout de 2 sec.
        'endTimeOut = sciID.ptrT + 2      ' timeout de 2 sec.
        Do
            ret = DoEvents()
            ccr = getch(sciID)
            If (ccr <> -1) Then
                gets = gets & Chr(ccr)
            'Else
            '    Exit Do
            End If
'Debug.Print ccr & "     Char : " & Chr(ccr)
        Loop Until (ccr = 13) Or (Timer >= endTimeOut)
    End If

End Function

'------------------------------------------------------------------
' state = True  ==> Turn DTR on
' state = False ==> Turn DTR off
' Toggle the DTREnabled property.
'------------------------------------------------------------------
Public Sub set_DTR(sciID As SciDef, state As Boolean)
    On Error Resume Next
    If sciID.comID.PortOpen Then
        sciID.comID.DTREnable = state
    End If
End Sub

'------------------------------------------------------------------
' Display the value of the CDHolding property.
'------------------------------------------------------------------
Public Function rd_CD(sciID As SciDef) As Boolean
    On Error Resume Next
    If sciID.comID.PortOpen Then
        If sciID.comID.CDHolding Then
            rd_CD = True
        Else
            rd_CD = False
        End If
    End If
End Function

'------------------------------------------------------------------
' Display the value of the CTSHolding property.
'------------------------------------------------------------------
Public Function rd_CTS(sciID As SciDef) As Boolean
    On Error Resume Next
    If sciID.comID.PortOpen Then
        If sciID.comID.CTSHolding Then
            rd_CTS = True
        Else
            rd_CTS = False
        End If
    End If
End Function

'------------------------------------------------------------------
' Display the value of the DSRHolding property.
'------------------------------------------------------------------
Public Function rd_DSR(sciID As SciDef) As Boolean
    On Error Resume Next
    If sciID.comID.PortOpen Then
        If sciID.comID.DSRHolding Then
            rd_DSR = True
        Else
            rd_DSR = False
        End If
    End If
End Function

'------------------------------------------------------------------
'------------------------------------------------------------------
Public Sub cleanRxBuffer(sciID As SciDef)
Dim bidon As Integer, i As Integer
Dim ret As Integer
    On Error Resume Next
    i = 0
    If sciID.comID.PortOpen Then
        Do
            ret = DoEvents()
            bidon = getch(sciID)
            If bidon = -1 Then i = i + 1
        Loop Until (bidon <> -1) Or (i = 500)
    End If
End Sub
