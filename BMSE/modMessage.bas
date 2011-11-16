Attribute VB_Name = "modMessage"
Option Explicit

' ---------- 標準モジュール ----------
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

'サブクラス化関数
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const GWL_WNDPROC = (-4) 'ウインドウプロシージャ

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

'デフォルトのウインドウプロシージャ
Public OldWindowhWnd As Long


'---------------------------------------------------------------------------
' 関数名： SubClass
' 機能 ： サブクラス化を開始する
' 引数 ： (in) hWnd … 対象フォームのウインドウハンドル
' 返り値 ： なし
'---------------------------------------------------------------------------
Public Sub SubClass(ByVal hwnd As Long)


    OldWindowhWnd = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)


End Sub


'---------------------------------------------------------------------------
' 関数名： UnSubClass
' 機能 ： サブクラス化を終了する
' 引数 ： (in) hWnd … 対象フォームのウインドウハンドル
' 返り値 ： なし
'---------------------------------------------------------------------------
Public Sub UnSubClass(ByVal hwnd As Long)


    Dim ret As Long


    If OldWindowhWnd <> 0 Then
    
        '元のプロシージャアドレスに設定する
        ret = SetWindowLong(hwnd, GWL_WNDPROC, OldWindowhWnd)


        OldWindowhWnd = 0&
    
    End If


End Sub

'---------------------------------------------------------------
' 関数名： strNullCut
' 機能 ： 文字列を vbNullChar までを取得する
' 引数 ： (in) srcStr … 対象文字列
' 返り値 ：編集された文字列
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


'次は、受信する側のコード。文字列取得方法は取得した文字列へのポインタより NULL までの長さを取得し、その長さ分バイト単位でコピーしてやればよい。



'-------------------------------------------------------------------------
' 関数名： WindowProc
' 機能 ： ウインドウメッセージをフックする
' 引数 ： (in) hWnd … 対象フォームのウインドウハンドル
'　　　　　(in) uMsg … ウインドウメッセージ
'　　　　　(in) wParam … 追加情報１
'　　　　　(in) lParam … 追加情報２
' 返り値 ： なし
' 備考 ： 特になし
'---------------------------------------------------------------------------
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


    'Dim udtCDP As COPYDATASTRUCT
    'Dim SentText As String '送られてきた文字列
    'Dim SentTextLen As Long '送られてきた文字列の数
    
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
            
            Case WM_CTLCOLORSCROLLBAR 'スクロールバー変な色対策
            
                Exit Function
        
        End Select
        
        'Debug.Print uMsg, wParam, lParam
    
    End If
    
    WindowProc = CallWindowProc(OldWindowhWnd, hwnd, uMsg, wParam, lParam)

End Function

Public Function HWORD(ByVal LongValue As Long) As Integer

    '長整数値から上位ワードを取得する
    HWORD = (LongValue And &HFFFF0000) \ &H10000

End Function

Public Function LWORD(ByVal LongValue As Long) As Integer

    '長整数値から下位ワードを取得する
    If (LongValue And &HFFFF&) > &H7FFF Then
        
        LWORD = (LongValue And &HFFFF&) - &H10000
    
    Else
        
        LWORD = LongValue And &HFFFF&
    
    End If

End Function
