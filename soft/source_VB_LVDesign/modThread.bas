Attribute VB_Name = "modThread"
Option Explicit

' ... Fonctions de l'API pour les threads
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const INFINITE = &HFFFF      '  Infinite timeout
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'' ... Fonctions de l'API pour les sections critiques
Public Type CRITICAL_SECTION
    dummy As Long
End Type
Public Declare Sub InitializeCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub DeleteCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)

' ... Collection des documents
Private ColDocs As New Collection
' Section critique pour la synchronisation de l'accès à la collection
Public csCol As CRITICAL_SECTION


'' User-defined type to store information about child forms
'Type FormState
'    Deleted As Integer
''    Dirty As Integer
''    Color As Long
'End Type
'
'Public FState(7)  As FormState           ' Array of user-defined types
'Public frmDocum(7) As New frmDocument    ' Array of child form objects
Public frmDocum(NB_MAX_WINDOW) As New frmDocument    ' Array of child form objects


'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

'
'Public Function AnyPadsLeft() As Integer
'    Dim i As Integer        ' Counter variable
'
'    ' Cycle through the document array.
'    ' Return true if there is at least one open document.
'    For i = 1 To UBound(frmDocum)
'        If Not FState(i).Deleted Then
'            AnyPadsLeft = True
'            Exit Function
'        End If
'    Next
'End Function

'--------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------
'Private Sub listDoc_Visible(win_tTag As Integer)
'Dim lst_n As Integer
'
'    For lst_n = win_INFO To NB_MAX_WINDOW
'        frmDocum(win_tTag).LstDoc(lst_n - 1).Visible = False
'    Next
'    frmDocum(win_tTag).LstDoc(win_tTag - 1).Visible = True
'
'End Sub
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Public Sub PositionCmd(win_tTag As Integer)
    With frmDocum(win_tTag)
        Select Case win_tTag
            Case win_TT_V4050
                .CmdLogin.Height = PHeight
                .CmdLogin.Left = PLeftLogin
                .CmdLogin.Top = PTop
                .CmdLogin.Width = PWidth

                .CmdRead.Height = PHeight
                .CmdRead.Left = PLeftRead
                .CmdRead.Top = PTop
                .CmdRead.Width = PWidth

                .CmdAutoRd.Height = PHeight
                .CmdAutoRd.Left = PLeftAutoRd
                .CmdAutoRd.Top = PTop
                .CmdAutoRd.Width = PWidth

                .CmdWrite.Height = PHeight
                .CmdWrite.Left = PLeftWrite
                .CmdWrite.Top = PTop
                .CmdWrite.Width = PWidth

                .CmdFormat.Height = PHeight
                .CmdFormat.Left = PLeftFormat
                .CmdFormat.Top = PTop
                .CmdFormat.Width = PWidth
                
                .CmdChPwd.Height = PHeight
                .CmdChPwd.Left = PLeftChPwd
                .CmdChPwd.Top = PTop
                .CmdChPwd.Width = PWidth

                .CmdReset.Height = PHeight
                .CmdReset.Left = PLeftReset
                .CmdReset.Top = PTop
                .CmdReset.Width = PWidth
                
            Case win_TT_H400X, win_TT_TK5530
'                .CmdLogin.Height = PHeight
'                .CmdLogin.Left = PLeftLogin
'                .CmdLogin.Top = PTop
'                .CmdLogin.Width = PWidth

                .CmdRead.Height = PHeight
                .CmdRead.Left = PLeftLogin
                .CmdRead.Top = PTop
                .CmdRead.Width = PWidth

                .CmdAutoRd.Height = PHeight
                .CmdAutoRd.Left = PLeftRead
                .CmdAutoRd.Top = PTop
                .CmdAutoRd.Width = PWidth

'                .CmdWrite.Height = PHeight
'                .CmdWrite.Left = PLeftWrite
'                .CmdWrite.Top = PTop
'                .CmdWrite.Width = PWidth
'
'                .CmdFormat.Height = PHeight
'                .CmdFormat.Left = PLeftFormat
'                .CmdFormat.Top = PTop
'                .CmdFormat.Width = PWidth
'
'                .CmdChPwd.Height = PHeight
'                .CmdChPwd.Left = PLeftChPwd
'                .CmdChPwd.Top = PTop
'                .CmdChPwd.Width = PWidth
'
'                .CmdReset.Height = PHeight
'                .CmdReset.Left = PLeftReset
'                .CmdReset.Top = PTop
'                .CmdReset.Width = PWidth
                            
            Case win_TT_TK5550, win_TT_TK5552
                        
'                .CmdLogin.Height = PHeight
'                .CmdLogin.Left = PLeftLogin
'                .CmdLogin.Top = PTop
'                .CmdLogin.Width = PWidth

                .CmdRead.Height = PHeight
                .CmdRead.Left = PLeftLogin
                .CmdRead.Top = PTop
                .CmdRead.Width = PWidth

                .CmdAutoRd.Height = PHeight
                .CmdAutoRd.Left = PLeftRead
                .CmdAutoRd.Top = PTop
                .CmdAutoRd.Width = PWidth

                .CmdWrite.Height = PHeight
                .CmdWrite.Left = PLeftAutoRd
                .CmdWrite.Top = PTop
                .CmdWrite.Width = PWidth

                .CmdFormat.Height = PHeight
                .CmdFormat.Left = PLeftWrite
                .CmdFormat.Top = PTop
                .CmdFormat.Width = PWidth
                
'                .CmdChPwd.Height = PHeight
'                .CmdChPwd.Left = PLeftChPwd
'                .CmdChPwd.Top = PTop
'                .CmdChPwd.Width = PWidth

                .CmdReset.Height = PHeight
                .CmdReset.Left = PLeftFormat
                .CmdReset.Top = PTop
                .CmdReset.Width = PWidth
                
            Case Else

        End Select
    End With

                
End Sub

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Sub init_frmDoc_Cmd(win_tTag As Integer)

    Call PositionCmd(win_tTag)
    With frmDocum(win_tTag)
        Select Case win_tTag
            Case win_INFO
                ' ... frmDoc Size
                .Width = MDIfrmRDT.Width / 2
                .Height = MDIfrmRDT.Height * 0.8
                ' ... ------ ----
                
                .PctBox.Visible = False
'                .PctBox.AutoSize = True
                
                .LstDoc(win_tTag - 1).Visible = True
                .LstDoc(win_tTag - 1).Width = .Width * 0.99
                .LstDoc(win_tTag - 1).Move 10, 0
            
                .LblHex.Visible = False
                .LblPassword.Visible = False
                .LblWriteFrom.Visible = False
                .LblWriteTo.Visible = False
                .txtPassWord.Visible = False
                .txtStartAdd.Visible = False
                .txtEndAdd.Visible = False
                
                .docframe.Visible = False
                
            Case win_TT_V4050
                ' ... frmDoc Size
                '.Width = MDIfrmRDT.Width * 0.6
                .Width = MDIfrmRDT.Width * 0.68
                .Height = MDIfrmRDT.Height * 0.8
                ' ... ------ ----
                
                .PctBox.Visible = True
'                .PctBox.AutoSize = True

                .LstDoc(win_tTag - 1).Visible = False
                '.LstDoc(win_tTag - 1).Width = .Width * 0.99
                '.LstDoc(win_tTag - 1).Move 10, 50
                               
                .docframe.Visible = True
                .docframe.Caption = "Tag Data"
                .docframe.Width = .Width * 0.99
                .docframe.Move 10, 100
                               
                .CmdLogin.Visible = True
                .CmdRead.Visible = True
                .CmdAutoRd.Visible = True
                .CmdWrite.Visible = True
                .CmdFormat.Visible = True
                
                .CmdChPwd.Visible = True
                .CmdChPwd.Enabled = False
                .CmdReset.Visible = True
    
                .LblHex.Visible = True
                .LblPassword.Visible = True
                .LblWriteFrom.Visible = True
                .LblWriteTo.Visible = True
                .txtPassWord.Visible = True
                .txtStartAdd.Visible = True
                .txtEndAdd.Visible = True
    
                writeStartAdd = 3
                writeEndAdd = 31
                .txtPassWord.Text = "00000000"
                .txtStartAdd.Text = Hex(writeStartAdd)
                .txtEndAdd.Text = Hex(writeEndAdd)
            
            Case win_TT_H400X
                ' ... frmDoc Size
                .Width = MDIfrmRDT.Width / 2
                .Height = MDIfrmRDT.Height / 4
                ' ... ------ ----
                .PctBox.Visible = True
                .PctBox.AutoSize = True
                
                .LstDoc(win_tTag - 1).Visible = False
                '.LstDoc(win_tTag - 1).Width = .Width * 0.99
                '.LstDoc(win_tTag - 1).Move 10, 50
                                
                .docframe.Visible = True
                .docframe.Caption = "Tag Data"
                .docframe.Width = .Width * 0.99
                .docframe.Move 10, 100
                                
                .CmdLogin.Visible = False
                .CmdRead.Visible = True
                .CmdAutoRd.Visible = True
                .CmdWrite.Visible = False
                .CmdFormat.Visible = False
        
                .CmdChPwd.Visible = False
                .CmdReset.Visible = False
                
                .LblHex.Visible = False
                .LblPassword.Visible = False
                .LblWriteFrom.Visible = False
                .LblWriteTo.Visible = False
                .txtPassWord.Visible = False
                .txtStartAdd.Visible = False
                .txtEndAdd.Visible = False
    
            Case win_TT_TK5530
                ' ... frmDoc Size
                .Width = MDIfrmRDT.Width / 2
                .Height = MDIfrmRDT.Height / 4
                ' ... ------ ----
                .PctBox.Visible = True
                .PctBox.AutoSize = True
                
                .LstDoc(win_tTag - 1).Visible = False
                '.LstDoc(win_tTag - 1).Width = .Width * 0.99
                '.LstDoc(win_tTag - 1).Move 10, 50
                
                .docframe.Visible = True
                .docframe.Caption = "Tag Data"
                .docframe.Width = .Width * 0.99
                .docframe.Move 10, 100
                                
                .CmdLogin.Visible = False
                .CmdRead.Visible = True
                .CmdAutoRd.Visible = True
                .CmdWrite.Visible = False
                .CmdFormat.Visible = False
        
                .CmdChPwd.Visible = False
                .CmdReset.Visible = False
                
                .LblHex.Visible = False
                .LblPassword.Visible = False
                .LblWriteFrom.Visible = False
                .LblWriteTo.Visible = False
                .txtPassWord.Visible = False
                .txtStartAdd.Visible = False
                .txtEndAdd.Visible = False
    
            Case win_TT_TK5550, win_TT_TK5552
                ' ... frmDoc Size
                .Width = MDIfrmRDT.Width * 0.6
                .Height = MDIfrmRDT.Height * 0.8
                ' ... ------ ----
                .PctBox.Visible = True
                .PctBox.AutoSize = True
                
                .LstDoc(win_tTag - 1).Visible = False
                '.LstDoc(win_tTag - 1).Width = .Width * 0.99
                ''.LstDoc(win_tTag - 1).Move 10, 50
                '.LstDoc(win_tTag - 1).Move 10, 100
                
                .docframe.Visible = True
                .docframe.Caption = "Tag Data"
                .docframe.Width = .Width * 0.99
                .docframe.Move 10, 100
                
                .CmdLogin.Visible = False
                .CmdRead.Visible = True
                .CmdAutoRd.Visible = True
                .CmdWrite.Visible = True
                .CmdFormat.Visible = True
        
                .CmdChPwd.Visible = False
                .CmdReset.Visible = True
                
                .LblHex.Visible = True
                .LblPassword.Visible = False
                .LblWriteFrom.Visible = True
                .LblWriteTo.Visible = True
                .txtPassWord.Visible = False
                .txtStartAdd.Visible = True
                .txtEndAdd.Visible = True
                
                writeStartAdd = 1
                writeEndAdd = 7
                .txtPassWord.Text = 0
                .txtStartAdd.Text = Hex(writeStartAdd)
                .txtEndAdd.Text = Hex(writeEndAdd)
    
'            Case win_TT_TK5552
'                ' ... frmDoc Size
'                .Width = MDIfrmRDT.Width * 0.6
'                .Height = MDIfrmRDT.Height * 0.8
'                ' ... ------ ----
'                .PctBox.Visible = True
'                .PctBox.AutoSize = True
'
'                .LstDoc(win_tTag - 1).Visible = True
'                .LstDoc(win_tTag - 1).Width = .Width * 0.99
''                .LstDoc(win_tTag - 1).Move 10, 50
'                .LstDoc(win_tTag - 1).Move 10, 100
'
'                .CmdLogin.Visible = False
'                .CmdRead.Visible = True
'                .CmdAutoRd.Visible = True
'                .CmdWrite.Visible = True
'                .CmdFormat.Visible = True
'
'                .CmdChPwd.Visible = False
'                .CmdReset.Visible = True
'
'                .LblHex.Visible = True
'                .LblPassword.Visible = False
'                .LblWriteFrom.Visible = True
'                .LblWriteTo.Visible = True
'                .txtPassWord.Visible = False
'                .txtStartAdd.Visible = True
'                .txtEndAdd.Visible = True
'
'                writeStartAdd = 1
'                writeEndAdd = 7
'                .txtPassWord.Text = 0
'                .txtStartAdd.Text = Hex(writeStartAdd)
'                .txtEndAdd.Text = Hex(writeEndAdd)
'
            Case Else

        End Select
    End With
End Sub

'----------------------------------------------------------------------------------
'---- Création d'un nouveau document
'----------------------------------------------------------------------------------
Public Sub LoadNewDoc(nDoc_win As Integer)

    ' ... nDoc_win = Numéro du document
    Set frmDocum(nDoc_win) = New frmDocument
    ' ... l'Ajouta à la collection
    ColDocs.Add frmDocum(nDoc_win), CStr(nDoc_win)
    
    ' ... Set Window
    frmDocum(nDoc_win).Visible = True
    frmDocum(nDoc_win).WindowState = WIN_Normal
    
    frmDocum(nDoc_win).Init nDoc_win

End Sub

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Sub frmWinState(win_typeTag As Integer)
Dim win_n As Integer

    If (win_TT(win_typeTag)) Then
        frmDocum(win_typeTag).WindowState = WIN_Minimized
        frmDocum(win_typeTag).Visible = False
        win_TT(win_typeTag) = False
        ' ... Termine le document
        Call UnloadDoc(win_typeTag)
    End If

End Sub

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public Sub frmWinVisible(win_typeTag As Integer)
Dim win_TT As Integer

    For win_TT = win_TT_V4050 To NB_MAX_WINDOW
        frmDocum(win_TT).Visible = False
    Next
    frmDocum(win_typeTag).Visible = True

End Sub

'----------------------------------------------------------------------------------
'---- Fin d'un document
'----------------------------------------------------------------------------------
Public Sub UnloadDoc(ByVal Num As Integer)
    ' ... Retire le document de la collection
    ColDocs.Remove CStr(Num)
End Sub

'----------------------------------------------------------------------------------
'---- La fonction de thread
'----------------------------------------------------------------------------------
Public Function MonThread(ByVal Num As Long) As Long
    ' ... Retrouve le document courant
'    Dim frm As frmDocument
'    Set frm = ColDocs(CStr(Num))
    Set frmDocum(Num) = ColDocs(CStr(Num))
    ' ... Boucle
    Do While frmDocum(Num).fThreadRunning
        'frm.Dessin
    Loop
    ' ... Termine
    MonThread = 0
End Function

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

