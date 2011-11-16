Attribute VB_Name = "modMessage"
Option Explicit

' ---------- �W�����W���[�� ----------
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

'�T�u�N���X���֐�
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const GWL_WNDPROC = (-4) '�E�C���h�E�v���V�[�W��

Private Const WM_ACTIVATE = &H6
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_SETCURSOR = &H20
Private Const WM_KEYDOWN = &H100
Private Const WM_SYSCOMMAND = &H112
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_CTLCOLORSCROLLBAR = &H137
Private Const WM_MOUSEWHEEL = &H20A

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

Private Const MM_MCINOTIFY = &H3B9

Private Const MCI_NOTIFY_SUCCESSFUL = 1
Private Const MCI_NOTIFY_SUPERSEDED = 2
Private Const MCI_NOTIFY_ABORTED = 4
Private Const MCI_NOTIFY_FAILURE = 8

Private Const WA_ACTIVE = 1
Private Const WA_CLICKACTIVE = 2
Private Const WA_INACTIVE = 0

Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8

'�f�t�H���g�̃E�C���h�E�v���V�[�W��
Public OldWindowhWnd As Long


'---------------------------------------------------------------------------
' �֐����F SubClass
' �@�\ �F �T�u�N���X�����J�n����
' ���� �F (in) hWnd �c �Ώۃt�H�[���̃E�C���h�E�n���h��
' �Ԃ�l �F �Ȃ�
'---------------------------------------------------------------------------
Public Sub SubClass(ByVal hwnd As Long)


    OldWindowhWnd = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)


End Sub


'---------------------------------------------------------------------------
' �֐����F UnSubClass
' �@�\ �F �T�u�N���X�����I������
' ���� �F (in) hWnd �c �Ώۃt�H�[���̃E�C���h�E�n���h��
' �Ԃ�l �F �Ȃ�
'---------------------------------------------------------------------------
Public Sub UnSubClass(ByVal hwnd As Long)


    Dim ret As Long


    If OldWindowhWnd <> 0 Then
    
        '���̃v���V�[�W���A�h���X�ɐݒ肷��
        ret = SetWindowLong(hwnd, GWL_WNDPROC, OldWindowhWnd)


        OldWindowhWnd = 0&
    
    End If


End Sub

'---------------------------------------------------------------
' �֐����F strNullCut
' �@�\ �F ������� vbNullChar �܂ł��擾����
' ���� �F (in) srcStr �c �Ώە�����
' �Ԃ�l �F�ҏW���ꂽ������
'---------------------------------------------------------------
Public Function strNullCut(ByVal srcStr As String) As String


    Dim NullCharPos As Integer


    NullCharPos = InStr(srcStr, Chr$(0))


    If NullCharPos = 0 Then
    
        strNullCut = srcStr
        
        Exit Function
    
    End If


    strNullCut = Left$(srcStr, NullCharPos - 1)


End Function


'���́A��M���鑤�̃R�[�h�B������擾���@�͎擾����������ւ̃|�C���^��� NULL �܂ł̒������擾���A���̒������o�C�g�P�ʂŃR�s�[���Ă��΂悢�B



'-------------------------------------------------------------------------
' �֐����F WindowProc
' �@�\ �F �E�C���h�E���b�Z�[�W���t�b�N����
' ���� �F (in) hWnd �c �Ώۃt�H�[���̃E�C���h�E�n���h��
'�@�@�@�@�@(in) uMsg �c �E�C���h�E���b�Z�[�W
'�@�@�@�@�@(in) wParam �c �ǉ����P
'�@�@�@�@�@(in) lParam �c �ǉ����Q
' �Ԃ�l �F �Ȃ�
' ���l �F ���ɂȂ�
'---------------------------------------------------------------------------
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


    'Dim udtCDP As COPYDATASTRUCT
    'Dim SentText As String '�����Ă���������
    'Dim SentTextLen As Long '�����Ă���������̐�
    
    Dim lngRet  As Long
    
    
    If frmMain.hwnd = GetActiveWindow() Then
    
        Select Case uMsg
        
            Case WM_ACTIVATEAPP
            
                If wParam Then
                
                    If frmMain.mnuOptionsActiveIgnore.Checked Then g_blnIgnoreInput = True
                
                End If
            
            Case WM_SETCURSOR
            
                If wParam <> frmMain.picMain.hwnd Then
                
                    g_Obj(UBound(g_Obj)).intCh = 0
                    
                    frmMain.staMain.Panels("Position").Text = "Position:"
                    
                    'Call frmMain.picMain.Cls
                    Call modEasterEgg.DrawEffect
                    
                    Select Case wParam
                    
                        Case frmMain.lstWAV.hwnd
                        
                            Call frmMain.lstWAV.SetFocus
                        
                        Case frmMain.lstBMP.hwnd
                        
                            Call frmMain.lstBMP.SetFocus
                        
                        Case frmMain.lstBGA.hwnd
                        
                            Call frmMain.lstBGA.SetFocus
                        
                        Case frmMain.lstMeasureLen.hwnd
                        
                            Call frmMain.lstMeasureLen.SetFocus
                    
                    End Select
                
                Else
                
                    'Call frmMain.vsbMain.SetFocus
                    Call frmMain.picMain.SetFocus
                
                End If
            
            Case WM_MOUSEWHEEL
            
                If HWORD(wParam) > 0 Then
                
                    lngRet = SB_LINEUP
                
                Else
                
                    lngRet = SB_LINEDOWN
                
                End If
                
                Call WindowProc(frmMain.hwnd, WM_VSCROLL, lngRet, frmMain.vsbMain.hwnd)
                Call WindowProc(frmMain.hwnd, WM_VSCROLL, SB_ENDSCROLL, frmMain.vsbMain.hwnd)
            
            Case MM_MCINOTIFY
            
                If wParam = MCI_NOTIFY_SUCCESSFUL Then
                
                    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
                
                End If
            
            Case WM_CTLCOLORSCROLLBAR '�X�N���[���o�[�ςȐF�΍�
            
                Exit Function
        
        End Select
        
        'Debug.Print uMsg, wParam, lParam
    
    End If
    
    WindowProc = CallWindowProc(OldWindowhWnd, hwnd, uMsg, wParam, lParam)

End Function

Public Function HWORD(ByVal LongValue As Long) As Integer

    '�������l�����ʃ��[�h���擾����
    HWORD = (LongValue And &HFFFF0000) \ &H10000

End Function

Public Function LWORD(ByVal LongValue As Long) As Integer

    '�������l���牺�ʃ��[�h���擾����
    If (LongValue And &HFFFF&) > &H7FFF Then
        
        LWORD = (LongValue And &HFFFF&) - &H10000
    
    Else
        
        LWORD = LongValue And &HFFFF&
    
    End If

End Function
