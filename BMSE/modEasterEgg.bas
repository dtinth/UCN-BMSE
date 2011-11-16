Attribute VB_Name = "modEasterEgg"
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_WORDBREAK = &H10

Public Enum EASTEREGG
    OFF
    TIPS
    RASTER
    SNOW
    SIROMARU
    NOISE
    STORM
    SIROMARU2
    SIROMARU3
    DISP_LOG
    STAFFROLL
    STAFFROLL2
    G20
    BLUESCREEN
    Max
End Enum

Private Type m_udtSnow
    X   As Single
    Y   As Single
    dX  As Single
    dY  As Single
    counter As Long
    Angle   As Integer
End Type

Private m_objSnow() As m_udtSnow

'Private m_sngRaster(511)    As Single
'Private m_lngRasterHeight   As Long
Private m_lngCounter        As Long

Private m_strStaffRoll()    As String

Public Sub InitEffect()

    'modInput.LoadBMSEnd にエイプリルフール用コードあり
    'If strGet_ini("EasterEgg", "Snow", False, "bmse.ini") = True Or (Month(Now) = 12 And Day(Now) = 25) Then
    If strGet_ini("EasterEgg", "Snow", False, "bmse.ini") Then
    
        g_disp.intEffect = SNOW
        
        Call modEasterEgg.InitSnow
    
    ElseIf strGet_ini("EasterEgg", "siromaru", False, "bmse.ini") Then
    
        g_disp.intEffect = SIROMARU
        
        Call modEasterEgg.InitSnow
    
    'ElseIf Month(Now) = 12 And (Day(Now) = 24 Or Day(Now) = 25) And strGet_ini("EasterEgg", "Snow", True, "bmse.ini") = True Then
    ElseIf Month(Now) = 12 And Day(Now) = 25 And strGet_ini("EasterEgg", "Snow", True, "bmse.ini") = True Then
        
        g_disp.intEffect = SNOW
        
        Call modEasterEgg.InitSnow
        
        g_strAppTitle = g_strAppTitle & " (Xmas mode: Only once!)"
        frmMain.Caption = g_strAppTitle
        
        Call modMain.lngSet_ini("EasterEgg", "Snow", False)
    
    End If

End Sub

Public Sub LoadEffect()

    If Month(Now) = 2 And Day(Now) = 15 Then 'シロマルの誕生日
    
        If g_BMS.strArtist = "siromaru" Then
        
            If modMain.strGet_ini("EasterEgg", "siromaru", False, "bmse.ini") = False And modMain.strGet_ini("EasterEgg", "siromaru", True, "bmse.ini") = True Then
                
                g_disp.intEffect = SIROMARU
                
                Call modEasterEgg.InitSnow
                
                'Call modMain.lngSet_ini("EasterEgg", "siromaru", False)
            
            End If
        
        End If
    
    End If

End Sub

Public Sub EndEffect()

    Dim strRet  As String
    
    If frmMain.tmrEffect.Enabled Then
    
        'If Month(Now) = 4 And Day(Now) = 1 And g_disp.intEffect = RASTER Then 'April Fool
        If Month(Now) = 2 And Day(Now) = 15 And g_disp.intEffect = SIROMARU Then 'シロマルの誕生日
        
            'If modMain.strGet_ini("EasterEgg", "RasterScroll", False, "bmse.ini") = False And modMain.strGet_ini("EasterEgg", "RasterScroll", True, "bmse.ini") = True Then
            If modMain.strGet_ini("EasterEgg", "siromaru", False, "bmse.ini") = False And modMain.strGet_ini("EasterEgg", "siromaru", True, "bmse.ini") = True Then
            
                'strRet = "エイプリル・フールだというのにわざわざ BMSE を使ってくれてありがとう!"
                'strRet = strRet & vbCrLf & "Thank you for using BMSE on April Fool's Day!"
                'strRet = strRet & vbCrLf & "びっくりしたかな?"
                'strRet = strRet & vbCrLf & "Were you surprised?"
                'strRet = strRet & vbCrLf
                'strRet = strRet & vbCrLf & "さて、この演出は今回限りだから安心してほしい。"
                'strRet = strRet & vbCrLf & "So, be relieved that this effect is only once."
                'strRet = strRet & vbCrLf & "それじゃ、、、"
                'strRet = strRet & vbCrLf & "Well..."
                'strRet = strRet & vbCrLf & """また後でな (ニヤリ)"""
                'strRet = strRet & vbCrLf & """BE SEEING YOU! (GRIN)"""
                
                strRet = "A Happy New Year!"
                strRet = strRet & vbCrLf & "本年もよろしくお願いします。"
                strRet = strRet & vbCrLf
                strRet = strRet & vbCrLf & "*Please note that this effect will appear only once."
                strRet = strRet & vbCrLf & "*大変恐縮ですが、このサービスは1回限りとさせて頂きます。"
                
                Call MsgBox(strRet, vbInformation, g_strAppTitle)
                
                'Call modMain.lngSet_ini("EasterEgg", "RasterScroll", False)
                Call modMain.lngSet_ini("EasterEgg", "siromaru", False)
            
            End If
        
        End If
    
    End If

End Sub

Public Sub DrawEffect()

    Select Case g_disp.intEffect
    
        Case SNOW, SIROMARU
        
            Call DrawSnow
        
        Case SIROMARU2
        
            Call DrawSiromaru2
        
        Case STAFFROLL, STAFFROLL2
        
            Call DrawStaffRoll
        
        Case DISP_LOG
        
            Call DrawLog
        
        Case BLUESCREEN
        
            Call DrawBlueScreen
        
        Case Else
        
            Call frmMain.picMain.Cls
    
    End Select

End Sub

Public Sub KeyCheck(ByVal KeyCode As Integer, ByVal Shift As Integer)

    Static buf As String * 16
    
    Select Case KeyCode
    
        Case vbKeyF9 'ほわいる
        
            If Len(frmMain.txtGenre.Text) = 0 And Len(frmMain.txtArtist.Text) = 0 Then
            
                frmMain.txtGenre.Text = "Unstable Pitch Song"
                frmMain.txtArtist.Text = "while"
            
            End If
        
        Case vbKeyF10 'シロディウス
        
            Call ShellExecute(0, vbNullString, "sirodius.exe", vbNullString, vbNullString, SW_SHOWNORMAL)
        
        Case vbKeyA To vbKeyZ, vbKey0 To vbKey9 'バッファに保存
        
            buf = Right$(buf, 15) & Chr$(KeyCode)
        
        Case vbKeyNumpad0 To vbKeyNumpad9 'バッファに保存
        
            buf = Right$(buf, 15) & KeyCode - vbKeyNumpad0
        
        Case vbKeySpace 'バッファに保存
        
            buf = Right$(buf, 15) & " "
        
        Case vbKeyReturn 'イースターエッグ発動
        
            If Right$(buf, 3) = "OFF" Then 'OFF
            
                frmMain.tmrEffect.Enabled = False
                g_disp.intEffect = OFF
                
                Call DrawEffect
            
            ElseIf Right$(buf, 4) = "TIPS" Then 'TIPS
            
                With frmWindowTips
                
                    .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
                    .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
                    
                    Call .Show(vbModal, frmMain)
                
                End With
            
            ElseIf Right$(buf, 4) = "SNOW" Then 'SNOW
            
                If g_disp.intEffect = SNOW Then
                
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = OFF
                
                Else
                
                    g_disp.intEffect = SNOW
                    
                    Call InitSnow
                
                End If
                
                Call DrawEffect
            
            ElseIf Right$(buf, 8) = "SIROMARU" Or Right$(buf, 9) = "SIROMARU1" Then 'SIROMARU
            
                If g_disp.intEffect = SIROMARU Then
                
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = OFF
                
                Else
                
                    g_disp.intEffect = SIROMARU
                    
                    Call InitSnow
                
                End If
                
                Call DrawEffect
            
            ElseIf Right$(buf, 9) = "SIROMARU2" Then 'SIROMARU2
            
                If g_disp.intEffect = SIROMARU2 Then
                
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = OFF
                
                Else
                
                    g_disp.intEffect = SIROMARU2
                    
                    Call InitSiromaru2
                
                End If
                
                Call DrawEffect
            
            ElseIf Right$(buf, 3) = "LOG" Then 'LOG
            
                frmMain.tmrEffect.Enabled = False
                
                If g_disp.intEffect = DISP_LOG Then
                
                    g_disp.intEffect = OFF
                
                Else
                
                    g_disp.intEffect = DISP_LOG
                
                End If
                
                Call DrawEffect
                Call modDraw.Redraw
            
            ElseIf Right$(buf, 9) = "STAFFROLL" Then 'STAFFROLL, STAFFROLL2
            
                If g_disp.intEffect = STAFFROLL Or g_disp.intEffect = STAFFROLL2 Then
                
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = OFF
                
                Else
                
                    frmMain.tmrEffect.Interval = 100
                    
                    g_disp.intEffect = STAFFROLL
                    
                    Call InitStaffRoll
                
                End If
                
                Call DrawEffect
            
            ElseIf Right$(buf, 10) = "STAFFROLL2" Then 'STAFFROLL, STAFFROLL2
            
                If g_disp.intEffect = STAFFROLL Or g_disp.intEffect = STAFFROLL2 Then
                
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = OFF
                
                Else
                
                    frmMain.tmrEffect.Interval = 10
                    
                    g_disp.intEffect = STAFFROLL
                    
                    Call InitStaffRoll
                
                End If
                
                Call DrawEffect
            
            ElseIf Right$(buf, 10) = "BLUESCREEN" Or Right$(buf, 4) = "BSOD" Then 'BLUESCREEN OF DEATH
            
                If g_disp.intEffect = BLUESCREEN Then
                
                    g_disp.intEffect = OFF
                
                Else
                
                    'frmMain.tmrEffect.Interval = 1
                    'frmMain.tmrEffect.Enabled = True
                    frmMain.tmrEffect.Enabled = False
                    g_disp.intEffect = BLUESCREEN
                
                End If
                
                Call DrawEffect
            
            End If
        
            buf = ""
    
    End Select

End Sub

Public Sub InitSnow()

    Dim i       As Long
    Dim lngRet  As Long
    Dim sngRet  As Single
    
    If g_disp.intEffect = OFF Then Exit Sub
    
    ReDim m_objSnow((Screen.Width \ Screen.TwipsPerPixelX) * 0.5 - 1)
    
    If g_disp.intEffect <> SNOW Then ReDim m_objSnow((Screen.Width \ Screen.TwipsPerPixelX) \ 8 - 1)
    
    lngRet = Screen.Height \ Screen.TwipsPerPixelY
    sngRet = ((Screen.Width \ Screen.TwipsPerPixelX) / UBound(m_objSnow))
    
    Call Randomize
    
    For i = 0 To UBound(m_objSnow)
    
        With m_objSnow(i)
        
            .counter = (Rnd * 1024) \ 4
            
            .X = sngRet * i
            .Y = Rnd * lngRet + 1 - lngRet
            
            If g_disp.intEffect = SNOW Then
            
                .dY = Rnd * 2 + 1
            
            Else
            
                .dY = Rnd * 4 + 4
            
            End If
            
            .dX = .X + g_sngSin(.counter And 255) * 5 * .dY
        
        End With
    
    Next i
    
    If g_disp.intEffect = SNOW Then Call QuickSortA(0, UBound(m_objSnow))
    
    frmMain.tmrEffect.Enabled = True
    frmMain.tmrEffect.Interval = 100

End Sub

Public Sub FallingSnow()

    Dim i       As Long
    Dim lngRet  As Long
    
    For i = 0 To UBound(m_objSnow)
    
        With m_objSnow(i)
        
            .counter = .counter + 4
            
            If g_disp.intEffect = SNOW Then
            
                .Y = .Y + .dY
                .X = .X + g_sngSin(.counter * 2 And 255) * .dY / 2
            
            Else
            
                lngRet = (.counter \ 4) And 7
                
                If lngRet = 0 Then
                
                    .Angle = Rnd * 128
                
                ElseIf lngRet > 1 Then
                
                    .X = .X + g_sngSin(.Angle + 64) * .dY
                    .Y = .Y + g_sngSin(.Angle) * .dY
                    
                    '.x = .x + g_sngSin((.Counter \ 32 And 7) * 32 + 16) * .dY
                    '.y = .y + g_sngSin(((.Counter \ 32 And 7) * 32 + 16 + 64) And 127) * .dY
                    '.X = .X + g_sngSin((.Counter) And 255) * .dY
                    '.Y = .Y + g_sngSin((.Counter + 64) And 127) * .dY
                
                End If
            
            End If
        
        End With
    
    Next i
    
    If g_disp.intEffect <> SNOW Then Call QuickSortY(0, UBound(m_objSnow))

End Sub

Public Sub DrawSnow()

    Dim i       As Long
    Dim X       As Long
    Dim Y       As Long
    Dim Width   As Long
    Dim Height  As Long
    Dim Size    As Long
    'Dim srcX    As Long
    Dim srcY    As Long
    Dim intRet  As Integer
    'Dim lngRet  As Long
    
    'lngRet = timeGetTime()
    
    Width = (Screen.Width \ Screen.TwipsPerPixelX)
    Height = (Screen.Height \ Screen.TwipsPerPixelY)
    
    For i = 0 To UBound(m_objSnow)
    
        With m_objSnow(i)
        
            X = (.X - frmMain.hsbMain.value * g_disp.Width) Mod Width
            Y = (.Y + frmMain.vsbMain.value * g_disp.Height) Mod Height
            
            'If Y < frmMain.picMain.ScaleHeight And X < frmMain.picMain.ScaleWidth Then
            
                Select Case .dY
                
                    Case Is < 3
                    
                        Size = Int(3 + (.dY - 1) * 3)  '3-8
                        
                        Call Ellipse(frmMain.picMain.hdc, X, Y, X + Size, Y + Size)
                        
                        If Y + Size > Height Then
                        
                            intRet = Y - Height
                            
                            Call Ellipse(frmMain.picMain.hdc, X, intRet, X + Size, intRet + Size)
                        
                        End If
                        
                        If X + Size > Width Then
                        
                            intRet = X - Width
                            
                            Call Ellipse(frmMain.picMain.hdc, intRet, Y, intRet + Size, Y + Size)
                        
                        End If
                    
                    Case Else
                    
                        srcY = ((.counter \ 4) And 7)
                        
                        If srcY > 1 Then
                        
                            Y = Y - g_sngSin((srcY - 1) * 128 \ 6 And 127) * 4 * .dY
                        
                        End If
                        
                        srcY = srcY * 32
                        
                        'Call Ellipse(frmMain.picmain.hdc, X - 16, .y - 16, X + 16, .y + 16)
                        Call BitBlt(frmMain.picMain.hdc, X, Y, 32, 32, frmMain.picSiromaru.hdc, 32, srcY, SRCAND)
                        Call BitBlt(frmMain.picMain.hdc, X, Y, 32, 32, frmMain.picSiromaru.hdc, 0, srcY, SRCPAINT)
                        
                        If Y + 32 > Height Then
                        
                            intRet = Y + 32 - Height
                            
                            Call BitBlt(frmMain.picMain.hdc, X, 0, 32, intRet, frmMain.picSiromaru.hdc, 32, srcY + 32 - intRet, SRCAND)
                            Call BitBlt(frmMain.picMain.hdc, X, 0, 32, intRet, frmMain.picSiromaru.hdc, 0, srcY + 32 - intRet, SRCPAINT)
                        
                        End If
                        
                        If X + 32 > Width Then
                        
                            intRet = X + 32 - Width
                            
                            Call BitBlt(frmMain.picMain.hdc, 0, Y, intRet, 32, frmMain.picSiromaru.hdc, 64 - intRet, srcY, SRCAND)
                            Call BitBlt(frmMain.picMain.hdc, 0, Y, intRet, 32, frmMain.picSiromaru.hdc, 32 - intRet, srcY, SRCPAINT)
                        
                        End If
                
                End Select
            
            'End If
        
        End With
    
    Next i
    
    'frmMain.cboDirectInput.Text = timeGetTime() - lngRet
    
    Exit Sub

End Sub

Private Sub InitSiromaru2()

    frmMain.tmrEffect.Enabled = True
    frmMain.tmrEffect.Interval = 100
    
    m_lngCounter = 0
    
    ReDim m_objSnow(0)
    m_objSnow(0).X = 1
    m_objSnow(0).dX = 0

End Sub

Public Sub ZoomSiromaru2()

    m_lngCounter = m_lngCounter + 1
    
    If (m_lngCounter And 7) > 1 And (m_lngCounter And 7) < 7 Then
    
        If m_objSnow(0).X < frmMain.picMain.ScaleWidth * 2 Then
        
            m_objSnow(0).dX = m_objSnow(0).dX + 0.1
            m_objSnow(0).X = m_objSnow(0).X + m_objSnow(0).dX
        
        End If
    
    End If

End Sub

Public Sub DrawSiromaru2()

    Dim X As Long
    Dim Y As Long
    Dim srcY As Integer
    
    srcY = m_lngCounter And 7
    
    If srcY > 1 Then
    
        Y = Y - g_sngSin((srcY - 1) * 128 \ 6 And 127) * m_objSnow(0).X
    
    End If
    
    srcY = srcY * 32
    
    With frmMain.picMain
    
        X = (.ScaleWidth - m_objSnow(0).X) \ 2
        Y = Y + (.ScaleHeight - m_objSnow(0).X) \ 2
        
        Call StretchBlt(.hdc, X, Y, m_objSnow(0).X, m_objSnow(0).X, frmMain.picSiromaru.hdc, 32, srcY, 32, 32, SRCAND)
        Call StretchBlt(.hdc, X, Y, m_objSnow(0).X, m_objSnow(0).X, frmMain.picSiromaru.hdc, 0, srcY, 32, 32, SRCPAINT)
    
    End With

End Sub

Public Sub DrawLog()

    '1.3.6 にて削除

End Sub

Private Sub DrawLogText(ByVal X As Long, ByVal Y As Long, ByVal Text As String, Optional ByVal Color As Long = 16777215)

    Dim intRet  As Integer
    
    With frmMain.picMain
    
        intRet = LenB(StrConv(Text, vbFromUnicode))
        
        Call SetTextColor(.hdc, 0) 'RGB(0, 0, 0)
        
        Call TextOut(.hdc, X, Y - 1, Text, intRet)
        Call TextOut(.hdc, X + 1, Y, Text, intRet)
        Call TextOut(.hdc, X, Y + 1, Text, intRet)
        Call TextOut(.hdc, X - 1, Y, Text, intRet)
        
        Call SetTextColor(.hdc, Color)
        
        Call TextOut(.hdc, X, Y, Text, intRet)
    
    End With

End Sub

Public Sub InitStaffRoll()

    If g_disp.intEffect = OFF Then Exit Sub
    
    frmMain.tmrEffect.Enabled = True
    
    m_lngCounter = 0
    
    ReDim m_strStaffRoll(0)
    
    Call AddStaffRoll("BMx Sequence Editor", 1)
    Call AddStaffRoll("Staff Credit", 5)
    
    Call AddStaffRoll("-Program-", 1)
    'Call AddStaffRoll("tokonats", 3)
    Call AddStaffRoll("Hayana", 0)
    Call AddStaffRoll("(aka tokonats)", 3)
    
    Call AddStaffRoll("-Program Icon, Toolbar Icon, BMSE Image-", 1)
    Call AddStaffRoll("AOiRO_Manbow", 3)
    
    Call AddStaffRoll("-Technical Adviser-", 1)
    Call AddStaffRoll("aska sakurano", 3)
    
    Call AddStaffRoll("-Language File Support-", 1)
    Call AddStaffRoll("Aruhito", 0)
    Call AddStaffRoll("sfmddrex", 0)
    Call AddStaffRoll("MW", 3)
    
    Call AddStaffRoll("-Tips Writing-", 1)
    Call AddStaffRoll("sfmddrex", 0)
    Call AddStaffRoll("Aruhito", 3)
    
    Call AddStaffRoll("-siromaru Animation-", 1)
    Call AddStaffRoll("tutidama", 0)
    Call AddStaffRoll("●▼●", 3)
    
    Call AddStaffRoll("-Easter Egg Adviser-", 1)
    Call AddStaffRoll("shammy", 0)
    'Call AddStaffRoll("Clock", 0)
    Call AddStaffRoll("sfmddrex", 0)
    Call AddStaffRoll("Lai", 0)
    Call AddStaffRoll("Yamajet", 0)
    Call AddStaffRoll("AOiRO_Manbow", 3)
    
    Call AddStaffRoll("-Programming Assistant-", 1)
    Call AddStaffRoll("Coca-Cola Classic", 3)
    
    Call AddStaffRoll("-Special Thanks-", 1)
    Call AddStaffRoll("tix", 0)
    Call AddStaffRoll("J.T.", 1)
    'Call AddStaffRoll("Shunsuke Kudo a.k.a. OBONO", 3)
    Call AddStaffRoll("Shunsuke Kudo", 0)
    Call AddStaffRoll("(aka OBONO)", 3)
    
    Call AddStaffRoll("-Special ""NO"" Thanks-", 1)
    Call AddStaffRoll("FontSize Property", 0)
    Call AddStaffRoll("FontBold Property", 0)
    Call AddStaffRoll("FontItalic Property", 0)
    Call AddStaffRoll("FontName Property", 0)
    Call AddStaffRoll("FontStrikethru Property", 0)
    Call AddStaffRoll("FontUnderline Property", 1)
    Call AddStaffRoll("TabStrip Control", 0)
    Call AddStaffRoll("SSTab Control", 1)
    Call AddStaffRoll("PitcureBox.MouseDown", 0)
    Call AddStaffRoll("PitcureBox.MouseMove", 0)
    Call AddStaffRoll("PictureBox.MouseUp", 1)
    'Call AddStaffRoll("Microsoft Visual Basic 6.0", 3)
    Call AddStaffRoll("Microsoft Visual Basic 6.0", 0)
    Call AddStaffRoll("(Oh No, I Love Her!)", 3)
    
    Call AddStaffRoll("-Debugger-", 1)
    Call AddStaffRoll("All BMSE Users:)", 5)
    
    'Call AddStaffRoll("Copyright(C) tokonats/UCN-Soft 2004.", 0)
    Call AddStaffRoll("Copyright(C) Hayana/UCN-Soft 2004-2006.", 0)
    Call AddStaffRoll("http://ucn.tokonats.net/", 0)
    Call AddStaffRoll("ucn@tokonats.net", 0)
    
    'ReDim Preserve m_strStaffRoll(1)

End Sub

Private Sub AddStaffRoll(ByVal Text As String, Optional ByVal Break As Integer)

    Dim lngRet  As Long
    
    If Break < 0 Then Break = 0
    
    lngRet = UBound(m_strStaffRoll) + 1
    
    ReDim Preserve m_strStaffRoll(lngRet + Break)
    
    m_strStaffRoll(lngRet) = Text

End Sub

Public Sub StaffRollScroll()

    m_lngCounter = m_lngCounter + 100 \ frmMain.tmrEffect.Interval

End Sub

Public Sub DrawStaffRoll()

    Dim i       As Long
    Dim X       As Long
    Dim Y       As Long
    Dim Color   As Long
    Dim intRet  As Long
    Dim lngRet  As Long
    Dim retSize As Size
    
    With frmMain.picMain
    
        Call SetTextColor(.hdc, RGB(255, 255, 255))
        .Font.Size = 12
        
        lngRet = .ScaleHeight - m_lngCounter
        
        For i = 0 To UBound(m_strStaffRoll)
        
            If Len(m_strStaffRoll(i)) Then
            
                intRet = LenB(StrConv(m_strStaffRoll(i), vbFromUnicode))
                
                Call GetTextExtentPoint32(.hdc, m_strStaffRoll(i), intRet, retSize)
                
                X = (frmMain.picMain.ScaleWidth - retSize.Width) \ 2
                Y = lngRet - retSize.Height \ 2
                
                If (Y < .ScaleHeight And Y + retSize.Height > 0) Or g_disp.intEffect = STAFFROLL2 Then
                
                    If g_disp.intEffect = STAFFROLL Then
                    
                        If .ScaleHeight < 128 Then
                        
                            Color = 255
                        
                        ElseIf lngRet < 64 Then
                        
                            Color = lngRet * 4
                            
                            If Color < 0 Then Color = 0
                        
                        ElseIf lngRet > .ScaleHeight - 64 Then
                        
                            Color = 255 - (lngRet - (.ScaleHeight - 64)) * 4
                            
                            If Color < 0 Then Color = 0
                        
                        Else
                        
                            Color = 255
                        
                        End If
                    
                    Else
                    
                        Select Case m_lngCounter
                        
                            Case Is > 95
                            
                                frmMain.tmrEffect.Enabled = False
                                g_disp.intEffect = OFF
                                
                                Exit Sub
                            
                            Case Is < 32: Color = m_lngCounter * 8 '0-31
                            Case Is > 63: Color = (95 - m_lngCounter) * 8 '64-95
                            Case Else: Color = 255 '32-63
                        
                        End Select
                        
                        Y = (.ScaleHeight - retSize.Height * UBound(m_strStaffRoll)) \ 2 + retSize.Height * i
                    
                    End If
                    
                    If m_strStaffRoll(i) <> "●▼●" Then
                    
                        Call DrawLogText(X, Y, m_strStaffRoll(i), RGB(Color, Color, Color))
                    
                    End If
                
                End If
                
                If m_strStaffRoll(i) = "●▼●" Then
                
                    Dim srcY    As Integer
                    
                    X = (frmMain.picMain.ScaleWidth - 32) \ 2
                    Y = lngRet '.ScaleHeight - lngRet
                    
                    srcY = (m_lngCounter And 7)
                    
                    If srcY > 1 Then
                    
                        Y = Y - g_sngSin((srcY - 1) * 128 \ 6 And 127) * 4 * 8
                    
                    End If
                    
                    srcY = srcY * 32
                    
                    Call BitBlt(frmMain.picMain.hdc, X, Y, 32, 32, frmMain.picSiromaru.hdc, 32, srcY, SRCAND)
                    Call BitBlt(frmMain.picMain.hdc, X, Y, 32, 32, frmMain.picSiromaru.hdc, 0, srcY, SRCPAINT)
                
                End If
                
                lngRet = lngRet + retSize.Height + 2
            
            Else
            
                lngRet = lngRet + 12
            
            End If
        
        Next i
        
        If lngRet < 0 Then
        
            If g_disp.intEffect = STAFFROLL Then
            
                g_disp.intEffect = STAFFROLL2
                
                ReDim m_strStaffRoll(0)
                
                'm_lngCounter = Rnd * 4
                
                'Select Case m_lngCounter
                    'Case 0: m_strStaffRoll(0) = """HAVE YOU FORGOTTEN SOMETHING?"""
                    'Case 1: m_strStaffRoll(0) = "I'M PERFECT! ARE YOU?"
                    'Case 2: m_strStaffRoll(0) = "The Matrix has you..."
                    'Case 3
                        'm_strStaffRoll(0) = "WAS ITS PHANTASM"
                        'Call AddStaffRoll("THE LAST ATTACKING")
                        'Call AddStaffRoll("OR ITS LAST MOMENTS", 1)
                        'Call AddStaffRoll("AND WAS THIS FOR REAL")
                        'Call AddStaffRoll("OR WAS I DREAMING", 1)
                        'Call AddStaffRoll("NOBODY KNOWS YET....")
                'End Select
                
                m_lngCounter = 0
                m_strStaffRoll(0) = """HAVE YOU FORGOTTEN SOMETHING?"""
            
            Else
            
                frmMain.tmrEffect.Enabled = False
                
                Erase m_strStaffRoll()
                
                g_disp.intEffect = OFF
            
            End If
        
        End If
    
    End With

End Sub

Public Sub DrawBlueScreen()

    Dim hBrushNew   As Long
    Dim hBrushOld   As Long
    Dim temp As RECT
    
    With frmMain.picMain
    
        hBrushNew = CreateSolidBrush(vbBlue)
        hBrushOld = SelectObject(.hdc, hBrushNew)
        
        Call Rectangle(.hdc, 0, 0, .ScaleWidth, .ScaleHeight)
        
        hBrushNew = SelectObject(.hdc, hBrushOld)
        Call DeleteObject(hBrushNew)
        
        Call SetTextColor(.hdc, 16777215)
        
        temp.Left = 8
        temp.Right = .ScaleWidth - 8
        temp.Top = 8
        temp.Bottom = .ScaleHeight
        
        .Font.Size = 9
        
        Call DrawText(.hdc, _
            "A problem has been detected and BMSE has been shut down to prevent damage to your mind." & vbCrLf & vbCrLf & _
            "The problem seems to be caused by the following file: BMSE.EXE" & vbCrLf & vbCrLf & _
            "EASTER_EGG_BLUE_SCREEN_OF_DEATH" & vbCrLf & vbCrLf & _
            "If this is the first time you've seen this stop error screen, restart your BMSE. If this screen appears again, follow these steps:" & vbCrLf & vbCrLf & _
            "1) Bury me from your computer." & vbCrLf & _
            "2) Access UCN-Soft BBS, and write your shout of spirit." & vbCrLf & _
            "       ex) ""BMSE is the worst software in the world!!!!!!!!!!!!!!111111""" & vbCrLf & _
            "3) Sing ""asdf song"":" & vbCrLf & _
            "       This is the sound of the asdf song." & vbCrLf & _
            "       asdf fdsa" & vbCrLf & _
            "       asdffdsa ye-ye" & vbCrLf & _
            "       (clap clap clap)" & vbCrLf & _
            "4) Throw your computer from window." & vbCrLf & vbCrLf & _
            "If you are satiated with joke:" & vbCrLf & vbCrLf & _
            "Launch BMSE and type your key ""OFF"", then press return key." & vbCrLf & vbCrLf & _
            "Meaningless information:" & vbCrLf & vbCrLf & _
            "*** STOP: 0x88710572 (0xASDFFDSA,0x00004126,0xD0SUK01,0x○0▽0○)" & vbCrLf & vbCrLf & vbCrLf & _
            "***  BMSE.EXE - Public Sub DrawBlueScreen() at modEasterEgg.bas, DateStamp 2006-12-26", _
            -1, temp, DT_WORDBREAK)
    
    End With

End Sub

Private Sub QuickSortY(ByVal lngLeft As Long, ByVal lngRight As Long)

    Dim i   As Long
    Dim j   As Long

    If lngLeft >= lngRight Then Exit Sub
    
    i = lngLeft + 1
    j = lngRight
    
    Do While i <= j
    
        Do While i <= j
        
            If m_objSnow(i).Y > m_objSnow(lngLeft).Y Then
                Exit Do
            End If
            
            i = i + 1
        Loop
        
        Do While i <= j
        
            If m_objSnow(j).Y < m_objSnow(lngLeft).Y Then
                Exit Do
            End If
            
            j = j - 1
        
        Loop
        
        If i >= j Then Exit Do
        
        Call SwapObj(m_objSnow(j), m_objSnow(i))
        
        i = i + 1
        j = j - 1
    
    Loop
    
    Call SwapObj(m_objSnow(j), m_objSnow(lngLeft))
    Call QuickSortY(lngLeft, j - 1)
    Call QuickSortY(j + 1, lngRight)

End Sub

Private Sub QuickSortA(ByVal lngLeft As Long, ByVal lngRight As Long)

    Dim i   As Long
    Dim j   As Long

    If lngLeft >= lngRight Then Exit Sub
    
    i = lngLeft + 1
    j = lngRight
    
    Do While i <= j
    
        Do While i <= j
        
            If m_objSnow(i).dY > m_objSnow(lngLeft).dY Then
                Exit Do
            End If
            
            i = i + 1
        Loop
        
        Do While i <= j
        
            If m_objSnow(j).dY < m_objSnow(lngLeft).dY Then
                Exit Do
            End If
            
            j = j - 1
        
        Loop
        
        If i >= j Then Exit Do
        
        Call SwapObj(m_objSnow(j), m_objSnow(i))
        
        i = i + 1
        j = j - 1
    
    Loop
    
    Call SwapObj(m_objSnow(j), m_objSnow(lngLeft))
    Call QuickSortA(lngLeft, j - 1)
    Call QuickSortA(j + 1, lngRight)

End Sub

Private Sub SwapObj(ByRef Obj1 As m_udtSnow, ByRef Obj2 As m_udtSnow)

    Dim dummyObj    As m_udtSnow
    
    With dummyObj
    
        .Angle = Obj1.Angle
        .counter = Obj1.counter
        .dX = Obj1.dX
        .dY = Obj1.dY
        .X = Obj1.X
        .Y = Obj1.Y
    
    End With
    
    With Obj1
    
        .Angle = Obj2.Angle
        .counter = Obj2.counter
        .dX = Obj2.dX
        .dY = Obj2.dY
        .X = Obj2.X
        .Y = Obj2.Y
    
    End With
    
    With Obj2
    
        .Angle = dummyObj.Angle
        .counter = dummyObj.counter
        .dX = dummyObj.dX
        .dY = dummyObj.dY
        .X = dummyObj.X
        .Y = dummyObj.Y
    
    End With

End Sub


