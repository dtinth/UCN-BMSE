Attribute VB_Name = "modDraw"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long

'Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'CreatePen 関連
Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6

'CreateHatchBrush 関連
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||

'CreateBrushIndirect 関連
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8

Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

'BitBlt 関連の定数
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046

'GetTextExtentPoint32 関連
Public Type Size
    Width   As Long
    Height  As Long
End Type

Public Const OBJ_DIFF = -1 'オブジェのずれ

'# Ch早見表 #
' 1 BGM
' 2 小節長
' 3 BPM Hex
' 4 BGA
' 6 Poor
' 7 Layer
' 8 BPM Dec
' 9 STOP
'11 1P 1Key
'12 1P 2Key
'13 1P 3Key
'14 1P 4Key
'15 1P 5Key
'18 1P 6Key
'19 1P 7Key
'16 1P SC
'21 2P 1Key
'22 2P 2Key
'23 2P 3Key
'24 2P 4Key
'25 2P 5Key
'28 2P 6Key
'29 2P 7Key
'26 2P SC
'31-49 不可視オブジェ
'51-69 ロングノート

Public g_lngPenColor(75)    As Long
Public g_lngBrushColor(36)  As Long
Public g_lngSystemColor(6)  As Long

Private m_hPen(75)          As Long
Private m_hBrush(37)        As Long

Private m_retObj()  As g_udtObj

Public Enum COLOR_NUM
    MEASURE_NUM
    MEASURE_LINE
    GRID_MAIN
    GRID_SUB
    VERTICAL_MAIN
    VERTICAL_SUB
    INFO
    
    Max
End Enum

Public Enum PEN_NUM
    BGM_LIGHT
    BGM_SHADOW
    BPM_LIGHT
    BPM_SHADOW
    BGA_LIGHT
    BGA_SHADOW
    KEY01_LIGHT
    KEY02_LIGHT
    KEY03_LIGHT
    KEY04_LIGHT
    
    KEY05_LIGHT
    KEY06_LIGHT
    KEY07_LIGHT
    KEY08_LIGHT
    KEY11_LIGHT
    KEY12_LIGHT
    KEY13_LIGHT
    KEY14_LIGHT
    KEY15_LIGHT
    KEY16_LIGHT
    
    KEY17_LIGHT
    KEY18_LIGHT
    KEY01_SHADOW
    KEY02_SHADOW
    KEY03_SHADOW
    KEY04_SHADOW
    KEY05_SHADOW
    KEY06_SHADOW
    KEY07_SHADOW
    KEY08_SHADOW
    
    KEY11_SHADOW
    KEY12_SHADOW
    KEY13_SHADOW
    KEY14_SHADOW
    KEY15_SHADOW
    KEY16_SHADOW
    KEY17_SHADOW
    KEY18_SHADOW
    INV_KEY01_LIGHT
    INV_KEY02_LIGHT
    
    INV_KEY03_LIGHT
    INV_KEY04_LIGHT
    INV_KEY05_LIGHT
    INV_KEY06_LIGHT
    INV_KEY07_LIGHT
    INV_KEY08_LIGHT
    INV_KEY11_LIGHT
    INV_KEY12_LIGHT
    INV_KEY13_LIGHT
    INV_KEY14_LIGHT
    
    INV_KEY15_LIGHT
    INV_KEY16_LIGHT
    INV_KEY17_LIGHT
    INV_KEY18_LIGHT
    INV_KEY01_SHADOW
    INV_KEY02_SHADOW
    INV_KEY03_SHADOW
    INV_KEY04_SHADOW
    INV_KEY05_SHADOW
    INV_KEY06_SHADOW
    
    INV_KEY07_SHADOW
    INV_KEY08_SHADOW
    INV_KEY11_SHADOW
    INV_KEY12_SHADOW
    INV_KEY13_SHADOW
    INV_KEY14_SHADOW
    INV_KEY15_SHADOW
    INV_KEY16_SHADOW
    INV_KEY17_SHADOW
    INV_KEY18_SHADOW
    
    LONGNOTE_LIGHT
    LONGNOTE_SHADOW
    SELECT_OBJ_LIGHT
    SELECT_OBJ_SHADOW
    EDIT_FRAME
    DELETE_FRAME
    
    Max
End Enum

Public Enum BRUSH_NUM
    BGM
    BPM
    BGA
    KEY01
    KEY02
    KEY03
    KEY04
    KEY05
    KEY06
    KEY07
    
    KEY08
    KEY11
    KEY12
    KEY13
    KEY14
    KEY15
    KEY16
    KEY17
    KEY18
    INV_KEY01
    
    INV_KEY02
    INV_KEY03
    INV_KEY04
    INV_KEY05
    INV_KEY06
    INV_KEY07
    INV_KEY08
    INV_KEY11
    INV_KEY12
    INV_KEY13
    
    INV_KEY14
    INV_KEY15
    INV_KEY16
    INV_KEY17
    INV_KEY18
    LONGNOTE
    SELECT_OBJ
    DELETE_FRAME
    EDIT_FRAME
    
    Max
End Enum

Public Enum GRID
    NUM_BLANK_1
    
    NUM_BPM
    NUM_STOP
    
    NUM_BLANK_2
    
    NUM_FOOTPEDAL
    
    NUM_1P_SC_L
    NUM_1P_1KEY
    NUM_1P_2KEY
    NUM_1P_3KEY
    NUM_1P_4KEY
    NUM_1P_5KEY
    NUM_1P_6KEY
    NUM_1P_7KEY
    NUM_1P_SC_R
    
    NUM_BLANK_3
    
    NUM_2P_SC_L
    NUM_2P_1KEY
    NUM_2P_2KEY
    NUM_2P_3KEY
    NUM_2P_4KEY
    NUM_2P_5KEY
    NUM_2P_6KEY
    NUM_2P_7KEY
    NUM_2P_SC_R
    
    NUM_BLANK_4
    
    NUM_BGA
    NUM_LAYER
    NUM_POOR
    
    NUM_BLANK_5

    NUM_BGM
End Enum

Public Const OBJ_WIDTH = 28
Public Const OBJ_HEIGHT = 9

Public Const GRID_WIDTH = OBJ_WIDTH
Public Const GRID_HALF_WIDTH = GRID_WIDTH \ 2
Public Const GRID_HALF_EDGE_WIDTH = (GRID_WIDTH * 3) \ 4
Public Const SPACE_WIDTH = 4
Public Const FRAME_WIDTH = GRID_WIDTH \ 2
Public Const LEFT_SPACE = FRAME_WIDTH + SPACE_WIDTH
Public Const RIGHT_SPACE = FRAME_WIDTH + SPACE_WIDTH * 2

Public g_sngSin(256 + 64)   As Single

Public Sub InitVerticalLine()
On Error GoTo Err:

    Dim i       As Long
    Dim lngRet  As Long
    
    With frmMain
    
        If .cboDispFrame.ListIndex Then
        
            For i = GRID.NUM_1P_1KEY To GRID.NUM_1P_7KEY
            
                g_VGrid(i).intWidth = GRID_WIDTH
            
            Next i
            
            For i = GRID.NUM_2P_1KEY To GRID.NUM_2P_7KEY
            
                g_VGrid(i).intWidth = GRID_WIDTH
            
            Next i
        
        Else
        
            g_VGrid(GRID.NUM_1P_1KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            For i = GRID.NUM_1P_2KEY To GRID.NUM_1P_6KEY
            
                g_VGrid(i).intWidth = GRID_HALF_WIDTH
            
            Next i
            
            If frmMain.cboDispKey.ListIndex Then
            
                g_VGrid(GRID.NUM_1P_7KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            Else
            
                g_VGrid(GRID.NUM_1P_5KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            End If
        
            g_VGrid(GRID.NUM_2P_1KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            For i = GRID.NUM_2P_2KEY To GRID.NUM_2P_6KEY
            
                g_VGrid(i).intWidth = GRID_HALF_WIDTH
            
            Next i
            
            If frmMain.cboDispKey.ListIndex Then
            
                g_VGrid(GRID.NUM_2P_7KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            Else
            
                g_VGrid(GRID.NUM_2P_5KEY).intWidth = GRID_HALF_EDGE_WIDTH
            
            End If
        
        End If
        
        Select Case .cboPlayer.ListIndex
        
            Case 0, 1, 2 '1P/2P/DP
            
                g_VGrid(GRID.NUM_FOOTPEDAL).blnVisible = False
                g_VGrid(GRID.NUM_2P_SC_L - 1).blnVisible = True
                
                If .cboDispKey.ListIndex = 0 Then
                
                    g_VGrid(GRID.NUM_1P_6KEY).blnVisible = False
                    g_VGrid(GRID.NUM_1P_7KEY).blnVisible = False
                
                Else
                
                    g_VGrid(GRID.NUM_1P_6KEY).blnVisible = True
                    g_VGrid(GRID.NUM_1P_7KEY).blnVisible = True
                
                End If
                
                If .cboDispSC1P.ListIndex = 0 Then
                
                    g_VGrid(GRID.NUM_1P_SC_L).blnVisible = True
                    g_VGrid(GRID.NUM_1P_SC_R).blnVisible = False
                
                Else
                
                    g_VGrid(GRID.NUM_1P_SC_L).blnVisible = False
                    g_VGrid(GRID.NUM_1P_SC_R).blnVisible = True
                
                End If
                
                If .cboPlayer.ListIndex <> 0 Then
                
                    For i = GRID.NUM_2P_SC_L To GRID.NUM_2P_SC_R + 1
                    
                        g_VGrid(i).blnVisible = True
                    
                    Next i
                    
                    If .cboDispKey.ListIndex = 0 Then
                    
                        g_VGrid(GRID.NUM_2P_6KEY).blnVisible = False
                        g_VGrid(GRID.NUM_2P_7KEY).blnVisible = False
                    
                    Else
                    
                        g_VGrid(GRID.NUM_2P_6KEY).blnVisible = True
                        g_VGrid(GRID.NUM_2P_7KEY).blnVisible = True
                    
                    End If
                    
                    If .cboDispSC2P.ListIndex = 0 Then
                    
                        g_VGrid(GRID.NUM_2P_SC_L).blnVisible = True
                        g_VGrid(GRID.NUM_2P_SC_R).blnVisible = False
                    
                    Else
                    
                        g_VGrid(GRID.NUM_2P_SC_L).blnVisible = False
                        g_VGrid(GRID.NUM_2P_SC_R).blnVisible = True
                    
                    End If
                
                Else
                
                    For i = GRID.NUM_2P_SC_L To GRID.NUM_2P_SC_R + 1
                    
                        g_VGrid(i).blnVisible = False
                    
                    Next i
                
                End If
                
            Case 3 'PMS
            
                g_VGrid(GRID.NUM_FOOTPEDAL).blnVisible = False
                g_VGrid(GRID.NUM_1P_SC_L).blnVisible = False
                g_VGrid(GRID.NUM_1P_6KEY).blnVisible = False
                g_VGrid(GRID.NUM_1P_7KEY).blnVisible = False
                g_VGrid(GRID.NUM_1P_SC_R).blnVisible = False
                g_VGrid(GRID.NUM_2P_SC_L - 1).blnVisible = False
                g_VGrid(GRID.NUM_2P_SC_L).blnVisible = False
                g_VGrid(GRID.NUM_2P_1KEY).blnVisible = False
                g_VGrid(GRID.NUM_2P_SC_R + 1).blnVisible = True
                
                For i = GRID.NUM_2P_2KEY To GRID.NUM_2P_5KEY
                
                    g_VGrid(i).blnVisible = True
                
                Next i
                
                For i = GRID.NUM_2P_6KEY To GRID.NUM_2P_SC_R
                
                    g_VGrid(i).blnVisible = False
                
                Next i
                
                If .cboDispFrame.ListIndex = 0 Then
                
                    g_VGrid(GRID.NUM_1P_5KEY).intWidth = GRID_HALF_WIDTH
                    g_VGrid(GRID.NUM_2P_5KEY).intWidth = GRID_HALF_EDGE_WIDTH
                
                End If
            
            Case 4 'Oct
            
                g_VGrid(GRID.NUM_FOOTPEDAL).blnVisible = True
                g_VGrid(GRID.NUM_1P_SC_L).blnVisible = True
                g_VGrid(GRID.NUM_1P_6KEY).blnVisible = True
                g_VGrid(GRID.NUM_1P_7KEY).blnVisible = True
                g_VGrid(GRID.NUM_2P_SC_L - 1).blnVisible = False
                g_VGrid(GRID.NUM_2P_1KEY).blnVisible = False
                g_VGrid(GRID.NUM_2P_SC_R).blnVisible = True
                g_VGrid(GRID.NUM_2P_SC_R + 1).blnVisible = True
                
                For i = GRID.NUM_2P_2KEY To GRID.NUM_2P_7KEY
                
                    g_VGrid(i).blnVisible = True
                
                Next i
                
                If .cboDispFrame.ListIndex = 0 Then
                
                    g_VGrid(GRID.NUM_1P_5KEY).intWidth = GRID_HALF_WIDTH
                    g_VGrid(GRID.NUM_1P_7KEY).intWidth = GRID_HALF_WIDTH
                    g_VGrid(GRID.NUM_2P_5KEY).intWidth = GRID_HALF_WIDTH
                    g_VGrid(GRID.NUM_2P_7KEY).intWidth = GRID_HALF_EDGE_WIDTH
                
                End If
        
        End Select
    
    End With
    
    lngRet = 0
    
    For i = 0 To 999
    
        g_Measure(i).lngY = lngRet
        lngRet = lngRet + g_Measure(i).intLen
    
    Next i
    
    g_disp.lngMaxY = g_Measure(999).lngY + g_Measure(999).intLen
    
    Call Redraw
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "InitVerticalLine")
End Sub

Public Sub Redraw()
On Error GoTo Err:

    Dim i           As Long
    Dim lngRet      As Long
    'Dim lngTimer    As Long
    
    'lngTimer = timeGetTime()
    
    'If frmMain.Visible = False Or frmMain.Enabled = False Then Exit Sub
    If frmMain.Visible = False Then Exit Sub
    
    For i = 0 To g_disp.intMaxMeasure
    
        lngRet = lngRet + g_Measure(i).intLen
    
    Next i
    
    'frmMain.vsbMain.Min = lngRet \ 96
    frmMain.vsbMain.Min = lngRet \ g_disp.intResolution
    
    frmMain.picMain.AutoRedraw = True
    
    With g_disp
    
        '.Width = frmMain.hsbDispWidth.Value / 100
        '.Height = frmMain.hsbDispHeight.Value / 100
        .Width = frmMain.cboDispWidth.ItemData(frmMain.cboDispWidth.ListIndex) / 100
        .Height = frmMain.cboDispHeight.ItemData(frmMain.cboDispHeight.ListIndex) / 100
        .intStartMeasure = 999
        .intEndMeasure = 999
        .lngStartPos = .Y - OBJ_HEIGHT
        .lngEndPos = .Y + frmMain.picMain.ScaleHeight / .Height
    
    End With
    
    'lngRet = 16
    lngRet = FRAME_WIDTH
    
    For i = 0 To UBound(g_intVGridNum)
    
        g_intVGridNum(i) = 0
    
    Next i
    
    For i = 0 To UBound(g_VGrid)
    
        With g_VGrid(i)
        
            If .blnVisible Then
            
                Select Case .intCh
                
                    Case 11 To 29
                    
                        g_intVGridNum(.intCh) = i
                        g_intVGridNum(.intCh + 20) = i
                        g_intVGridNum(.intCh + 40) = i
                    
                    Case Is > 0
                    
                        g_intVGridNum(.intCh) = i
                
                End Select
                
                .lngLeft = lngRet
                
                Select Case .intCh
                
                    Case 15
                    
                        If frmMain.cboDispKey.ListIndex = 1 Or frmMain.cboPlayer.ListIndex > 2 Then
                        
                            .lngObjLeft = .lngLeft + (.intWidth - GRID_WIDTH) \ 2
                        
                        Else
                        
                            .lngObjLeft = .lngLeft + .intWidth - GRID_WIDTH
                        
                        End If
                    
                    Case 25
                    
                        If frmMain.cboPlayer.ListIndex = 4 Then
                        
                            .lngObjLeft = .lngLeft + (.intWidth - GRID_WIDTH) \ 2
                        
                        ElseIf frmMain.cboDispKey.ListIndex = 0 Or frmMain.cboPlayer.ListIndex = 3 Then
                        
                            .lngObjLeft = .lngLeft + .intWidth - GRID_WIDTH
                        
                        Else
                        
                            .lngObjLeft = .lngLeft + (.intWidth - GRID_WIDTH) \ 2
                        
                        End If
                    
                    Case 19
                    
                        If frmMain.cboPlayer.ListIndex > 2 Then
                        
                            .lngObjLeft = .lngLeft + (.intWidth - GRID_WIDTH) \ 2
                        
                        Else
                        
                            .lngObjLeft = .lngLeft + .intWidth - GRID_WIDTH
                        
                        End If
                    
                    Case 29
                    
                        .lngObjLeft = .lngLeft + .intWidth - GRID_WIDTH
                    
                    Case 12 To 18, 22 To 28
                    
                        .lngObjLeft = .lngLeft + (.intWidth - GRID_WIDTH) \ 2
                    
                    Case Else
                    
                        .lngObjLeft = lngRet
                
                End Select
                
                'If (lngRet + .intWidth) * g_disp.Width >= g_disp.X And (g_disp.X + frmMain.picMain.ScaleWidth) / g_disp.Width >= lngRet Then
                If .lngLeft + .intWidth >= g_disp.X And frmMain.picMain.ScaleWidth + (g_disp.X - .lngLeft) * g_disp.Width >= 0 Then
                
                    .blnDraw = True
                
                Else
                    
                    .blnDraw = False
                
                End If
                
                lngRet = lngRet + .intWidth
            
            Else
            
                .blnDraw = False
            
            End If
        
        End With
    
    Next i
    
    g_disp.lngMaxX = lngRet
    
    lngRet = 0
    
    For i = 0 To 999
    
        lngRet = lngRet + g_Measure(i).intLen
        
        If lngRet > g_disp.Y Then
        
            g_disp.intStartMeasure = i
            
            Exit For
        
        End If
    
    Next i
    
    For i = g_disp.intStartMeasure + 1 To 999
    
        lngRet = lngRet + g_Measure(i).intLen
        
        If (lngRet - g_disp.Y) * g_disp.Height >= frmMain.picMain.ScaleHeight Then
        
            g_disp.intEndMeasure = i
            
            Exit For
        
        End If
    
    Next i
    
    lngRet = 0
    
    Call frmMain.picMain.Cls
    
    Call DrawGridBG '背景色
    
    Call DrawMeasureNum '小節番号
    
    Call DrawVerticalGrayLine '縦線(灰色)
    
    Call DrawHorizonalLine '横線(灰色)
    
    Call DrawVerticalWhiteLine '縦線(白)
    
    Call DrawMeasureLine '横線(白)
    
    Call InitPen
        
    With frmMain.picMain.Font
    
        .Size = 8
        .Italic = False
    
    End With
    
    ReDim m_retObj(0)
    
    For i = 0 To UBound(g_Obj) - 1 'オブジェ
    
        With g_Obj(i)
        
            If .intCh > 0 And .intCh < 133 Then
            
                If g_VGrid(g_intVGridNum(.intCh)).blnDraw Then
                
                    If g_disp.lngStartPos <= g_Measure(.intMeasure).lngY + .lngPosition And g_disp.lngEndPos >= g_Measure(.intMeasure).lngY + .lngPosition Then
                    
                        Call DrawObj(g_Obj(i))
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next i
    
    For i = 0 To UBound(m_retObj) - 1
    
        With m_retObj(i)
        
            If g_disp.lngStartPos <= g_Measure(.intMeasure).lngY + .lngPosition And g_disp.lngEndPos >= g_Measure(.intMeasure).lngY + .lngPosition And g_VGrid(g_intVGridNum(.intCh)).blnDraw = True And .intCh <> 0 Then
            
                Call DrawObj(m_retObj(i))
            
            End If
        
        End With
    
    Next i
    
    Erase m_retObj
    
    Call DeletePen
    
    Call DrawGridInfo 'グリッド情報
    
    With frmMain.picMain
    
        If (g_disp.lngMaxX + 16) * g_disp.Width - .ScaleWidth < 0 Then
        
            frmMain.hsbMain.Max = 0
        
        Else
        
            'frmMain.hsbMain.Max = (g_disp.lngMaxX + 16) * g_disp.Width - .ScaleWidth
            'frmMain.hsbMain.Max = (g_disp.lngMaxX + 16) - .ScaleWidth / g_disp.Width
            frmMain.hsbMain.Max = (g_disp.lngMaxX + FRAME_WIDTH) - .ScaleWidth / g_disp.Width
        
        End If
        
        .AutoRedraw = False
    
    End With
    
    If g_disp.intEffect Then Call modEasterEgg.DrawEffect
    
    'Debug.Print timeGetTime() - lngTimer
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "Redraw")
End Sub

Private Sub DrawGridBG()

    Dim i           As Long
    Dim hPenNew     As Long
    Dim hPenOld     As Long
    Dim hBrushNew   As Long
    Dim hBrushOld   As Long
    
    If frmMain.mnuOptionsLaneBG.Checked Then
    
        For i = 0 To UBound(g_VGrid) '背景色
        
            With g_VGrid(i)
            
                If .blnDraw Then
                
                    If .intCh Then
                    
                        hPenNew = CreatePen(PS_SOLID, 1, .lngBackColor)
                        hPenOld = SelectObject(frmMain.picMain.hdc, hPenNew)
                        hBrushNew = CreateSolidBrush(.lngBackColor)
                        hBrushOld = SelectObject(frmMain.picMain.hdc, hBrushNew)
                        
                        'Call Rectangle(frmMain.picMain.hdc, .lngLeft * g_disp.Width - g_disp.X, 0, (.lngLeft + .intWidth + 1) * g_disp.Width - g_disp.X, frmMain.picMain.ScaleHeight)
                        Call Rectangle(frmMain.picMain.hdc, (.lngLeft - g_disp.X) * g_disp.Width, 0, (.lngLeft + .intWidth + 1 - g_disp.X) * g_disp.Width, frmMain.picMain.ScaleHeight)
                        
                        hPenNew = SelectObject(frmMain.picMain.hdc, hPenOld)
                        Call DeleteObject(hPenNew)
                        hBrushNew = SelectObject(frmMain.picMain.hdc, hBrushOld)
                        Call DeleteObject(hBrushNew)
                    
                    End If
                
                End If
            
            End With
        
        Next i
    
    End If

End Sub

Private Sub DrawMeasureNum()

    Dim i       As Long
    Dim strRet  As String * 4
    Dim retSize As Size
    
    With frmMain.picMain
    
        .Font.Size = 72
        .Font.Italic = True
        
        For i = g_disp.intStartMeasure To g_disp.intEndMeasure '#小節番号
        
            strRet = "#" & Format$(i, "000")
            
            Call GetTextExtentPoint32(.hdc, strRet, 4, retSize)
            
            Call SetTextColor(.hdc, g_lngSystemColor(COLOR_NUM.MEASURE_NUM))  'RGB(64, 64, 64)
            Call TextOut(.hdc, (.ScaleWidth - retSize.Width) \ 2, .ScaleHeight - retSize.Height - (g_Measure(i).lngY - g_disp.Y) * g_disp.Height, strRet, 4)
            
        Next i
    
    End With
    
End Sub

Private Sub DrawVerticalGrayLine()

    Dim i       As Long
    Dim hNew    As Long
    Dim hOld    As Long
    
    hNew = CreatePen(PS_SOLID, 1, g_lngSystemColor(COLOR_NUM.VERTICAL_SUB))  'RGB(128, 128, 128)
    hOld = SelectObject(frmMain.picMain.hdc, hNew)
    
    For i = 0 To UBound(g_VGrid) '縦線(灰色)
    
        With g_VGrid(i)
        
            If .blnDraw Then
            
                If .intCh Then
                
                    Call PrintLine(.lngLeft + .intWidth, g_disp.Y, 0, frmMain.picMain.ScaleHeight)
                
                End If
            
            End If
        
        End With
    
    Next i
    
    hNew = SelectObject(frmMain.picMain.hdc, hOld)
    Call DeleteObject(hNew)

End Sub

Private Sub DrawHorizonalLine()

    Dim i       As Long
    Dim j       As Long
    Dim intRet  As Integer
    Dim hNew    As Long
    Dim hOld    As Long
    
    hNew = CreatePen(PS_SOLID, 1, g_lngSystemColor(COLOR_NUM.GRID_MAIN)) 'RGB(96, 96, 96)
    hOld = SelectObject(frmMain.picMain.hdc, hNew)
    
    For i = g_disp.intStartMeasure To g_disp.intEndMeasure '横線(灰色)
    
        With frmMain.cboDispGridSub
        
            If .ItemData(.ListIndex) Then
            
                intRet = 192 \ .ItemData(.ListIndex)
                
                For j = 0 To g_Measure(i).intLen Step intRet
                
                    Call PrintLine(LEFT_SPACE, g_Measure(i).lngY + j, g_disp.lngMaxX - RIGHT_SPACE, 0)
                
                Next j
            
            End If
        
        End With
    
    Next i
    
    
    hNew = SelectObject(frmMain.picMain.hdc, hOld)
    Call DeleteObject(hNew)
    
    hNew = CreatePen(PS_SOLID, 1, g_lngSystemColor(COLOR_NUM.GRID_SUB))  'RGB(192, 192, 192))
    hOld = SelectObject(frmMain.picMain.hdc, hNew)
    
    For i = g_disp.intStartMeasure To g_disp.intEndMeasure '横線(灰色・補助)
    
        With frmMain.cboDispGridMain
        
            If .ItemData(.ListIndex) Then
            
                intRet = 192 \ .ItemData(.ListIndex)
                
                For j = intRet To g_Measure(i).intLen Step intRet
                
                    'Call PrintLine(16, g_Measure(i).lngY + j, g_disp.lngMaxX - 16, 0)
                    Call PrintLine(FRAME_WIDTH, g_Measure(i).lngY + j, g_disp.lngMaxX - FRAME_WIDTH, 0)
                
                Next j
            
            End If
        
        End With
    
    Next i
    
    hNew = SelectObject(frmMain.picMain.hdc, hOld)
    Call DeleteObject(hNew)

End Sub

Private Sub DrawVerticalWhiteLine()

    Dim i       As Long
    Dim hNew    As Long
    Dim hOld    As Long
    
    hNew = CreatePen(PS_SOLID, 1, g_lngSystemColor(COLOR_NUM.VERTICAL_MAIN))
    hOld = SelectObject(frmMain.picMain.hdc, hNew)
    
    For i = 0 To UBound(g_VGrid) '縦線(白)
    
        With g_VGrid(i)
        
            If .blnDraw = True Then
            
                If .intCh = 0 Then
                
                    Call PrintLine(.lngLeft, g_disp.Y, 0, frmMain.picMain.ScaleHeight)
                    Call PrintLine(.lngLeft + .intWidth, g_disp.Y, 0, frmMain.picMain.ScaleHeight)
                
                End If
            
            End If
        
        End With
    
    Next i
    
    hNew = SelectObject(frmMain.picMain.hdc, hOld)
    Call DeleteObject(hNew)

End Sub

Private Sub DrawMeasureLine()

    Dim i   As Long
    Dim hNew    As Long
    Dim hOld    As Long
    
    hNew = CreatePen(hNew, 1, g_lngSystemColor(COLOR_NUM.MEASURE_LINE))
    hOld = SelectObject(frmMain.picMain.hdc, hNew)
    
    For i = g_disp.intStartMeasure To g_disp.intEndMeasure '横線(白)
    
        'Call PrintLine(16, g_Measure(i).lngY, g_disp.lngMaxX - 16, 0)
        Call PrintLine(FRAME_WIDTH, g_Measure(i).lngY, g_disp.lngMaxX - FRAME_WIDTH, 0)
    
    Next i
    
    'If g_disp.intEndMeasure = 999 Then Call PrintLine(16, g_Measure(999).lngY + g_Measure(999).intLen, g_disp.lngMaxX - 16, 0)
    If g_disp.intEndMeasure = 999 Then Call PrintLine(FRAME_WIDTH, g_Measure(999).lngY + g_Measure(999).intLen, g_disp.lngMaxX - FRAME_WIDTH, 0)
    
    hNew = SelectObject(frmMain.picMain.hdc, hOld)
    Call DeleteObject(hNew)

End Sub

Private Sub DrawGridInfo()

    Dim i       As Long
    Dim j       As Long
    Dim X       As Long
    Dim Width   As Long
    Dim intRet  As Integer
    Dim lngRet  As Long
    Dim strRet  As String
    Dim retSize As Size
    
    With frmMain.picMain
    
        .Font.Size = 9
    
    End With
        
    For i = 0 To UBound(g_VGrid) '文字
    
        With g_VGrid(i)
        
            If .blnDraw Then
            
                If .intCh Then
            
                    If frmMain.mnuOptionsVertical.Checked Then
                    
                        'lngRet = (.lngLeft + (.intWidth \ 2)) * g_disp.Width - g_disp.X
                        lngRet = (.lngLeft + (.intWidth \ 2) - g_disp.X) * g_disp.Width
                        
                        For j = 0 To Len(.strText) - 1
                        
                            strRet = Mid$(.strText, j + 1, 1)
                            intRet = LenB(StrConv(strRet, vbFromUnicode))
                            Call GetTextExtentPoint32(frmMain.picMain.hdc, strRet, intRet, retSize)
                            
                            X = lngRet - retSize.Width \ 2
                            
                            '無理やり縁取り
                            Call SetTextColor(frmMain.picMain.hdc, 0)
                            'Call TextOut(frmMain.picMain.hdc, X - 1, 0 + 11 * j, strRet, intRet)
                            Call TextOut(frmMain.picMain.hdc, X, 0 + 11 * j, strRet, intRet)
                            'Call TextOut(frmMain.picMain.hdc, X + 1, 0 + 11 * j, strRet, intRet)
                            Call TextOut(frmMain.picMain.hdc, X - 1, 1 + 11 * j, strRet, intRet)
                            Call TextOut(frmMain.picMain.hdc, X + 1, 1 + 11 * j, strRet, intRet)
                            'Call TextOut(frmMain.picMain.hdc, X - 1, 2 + 11 * j, strRet, intRet)
                            Call TextOut(frmMain.picMain.hdc, X, 2 + 11 * j, strRet, intRet)
                            'Call TextOut(frmMain.picMain.hdc, X + 1, 2 + 11 * j, strRet, intRet)
                            Call SetTextColor(frmMain.picMain.hdc, g_lngSystemColor(COLOR_NUM.INFO))
                            Call TextOut(frmMain.picMain.hdc, X, 1 + 11 * j, strRet, intRet)
                        
                        Next j
                    
                    Else
                    
                        intRet = LenB(StrConv(.strText, vbFromUnicode))
                        Call GetTextExtentPoint32(frmMain.picMain.hdc, .strText, intRet, retSize)
                        
                        'X = (.lngLeft + .intWidth \ 2) * g_disp.Width - (retSize.Width) \ 2 - g_disp.X + 1
                        X = (.lngLeft + .intWidth \ 2 - g_disp.X) * g_disp.Width - (retSize.Width) \ 2 + 1
                        
                        '無理やり縁取り
                        Call SetTextColor(frmMain.picMain.hdc, 0)
                        'Call TextOut(frmMain.picMain.hdc, X - 1, 0, .strText, intRet)
                        Call TextOut(frmMain.picMain.hdc, X, 0, .strText, intRet)
                        'Call TextOut(frmMain.picMain.hdc, X + 1, 0, .strText, intRet)
                        Call TextOut(frmMain.picMain.hdc, X - 1, 1, .strText, intRet)
                        Call TextOut(frmMain.picMain.hdc, X + 1, 1, .strText, intRet)
                        'Call TextOut(frmMain.picMain.hdc, X - 1, 2, .strText, intRet)
                        Call TextOut(frmMain.picMain.hdc, X, 2, .strText, intRet)
                        'Call TextOut(frmMain.picMain.hdc, X + 1, 2, .strText, intRet)
                        Call SetTextColor(frmMain.picMain.hdc, g_lngSystemColor(COLOR_NUM.INFO))
                        Call TextOut(frmMain.picMain.hdc, X, 1, .strText, intRet)
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next i

End Sub

Private Sub PrintLine(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long)

    Width = Width * g_disp.Width
    'X = X * g_disp.Width
    
    If X - g_disp.X < 0 Then
    
        'If Width Then Width = Width + (X - g_disp.X)
        If Width Then Width = Width + (X - g_disp.X) * g_disp.Width
        
        X = 0
    
    Else
    
        'X = X - g_disp.X
        X = (X - g_disp.X) * g_disp.Width
    
    End If
    
    If Y + g_disp.Y < 0 Then
    
        If Height Then Height = Height + (Y - g_disp.Y)
        
        Y = 0
    
    Else
    
        Y = (Y - g_disp.Y) * g_disp.Height
    
    End If
    
    Call MoveToEx(frmMain.picMain.hdc, X, frmMain.picMain.ScaleHeight - 1 - Y, 0)
    Call LineTo(frmMain.picMain.hdc, X + Width, frmMain.picMain.ScaleHeight - 1 - Y - Height)

End Sub

Public Sub DrawObj(ByRef retObj As g_udtObj) '(ByVal lngNum As Long)
On Error GoTo Err:

    Dim intRet          As Integer
    Dim Text            As String
    Dim strArray()      As String
    Dim X               As Long
    Dim Y               As Long
    Dim Width           As Integer
    Dim retSize         As Size
    Dim intLightNum     As Integer
    Dim intShadowNum    As Integer
    Dim intBrushNum     As Integer
    Dim hOldBrush       As Long
    Dim hOldPen         As Long
    
    With retObj 'g_Obj(lngNum)
    
        If g_intVGridNum(.intCh) = 0 Then Exit Sub
        
        'X = g_VGrid(g_intVGridNum(.intCh)).lngObjLeft * g_disp.Width - g_disp.X + 1
        X = (g_VGrid(g_intVGridNum(.intCh)).lngObjLeft - g_disp.X) * g_disp.Width + 1
        'Y = frmMain.picMain.Height \ Screen.TwipsPerPixelY - 5 - (g_Measure(.intMeasure).lngY + .lngPosition - g_disp.Y) * g_disp.Height
        Y = frmMain.picMain.ScaleHeight + OBJ_DIFF - (g_Measure(.intMeasure).lngY + .lngPosition - g_disp.Y) * g_disp.Height
        Width = GRID_WIDTH * g_disp.Width - 1
        
        Select Case .intCh
        
            Case 3, 8, 9 'BPM/STOP
            
                Text = .sngValue
            
            Case 4, 6, 7 'BGA/Layer/Poor
            
                Text = g_strBMP(.sngValue)
                
                If frmMain.mnuOptionsObjectFileName.Checked = True And Len(Text) > 0 Then
                
                    strArray = Split(Text, ".")
                    Text = Left$(Text, Len(Text) - (Len(strArray(UBound(strArray))) + 1))
                
                Else
                
                    Text = modInput.strNumConv(.sngValue)
                
                End If
            
            Case Else
            
                Text = g_strWAV(.sngValue)
                
                If frmMain.mnuOptionsObjectFileName.Checked = True And Len(Text) > 0 Then
                
                    strArray = Split(Text, ".")
                    Text = Left$(Text, Len(Text) - (Len(strArray(UBound(strArray))) + 1))
                
                Else
                
                    Text = modInput.strNumConv(.sngValue)
                
                End If
                
                If .intAtt = 2 Or (.intCh >= 51 And .intCh < 69) Then
                
                    X = X + 3
                    Width = Width - 6
                
                End If
                
                If .intAtt = 2 And .intCh >= 11 And .intCh <= 29 Then
                
                    Call modDraw.CopyObj(m_retObj(UBound(m_retObj)), retObj)
                    m_retObj(UBound(m_retObj)).intCh = .intCh + 40
                    
                    ReDim Preserve m_retObj(UBound(m_retObj) + 1)
                    
                    'Exit Sub
                
                End If
        
        End Select
    
        Select Case .intSelect
        
            Case 0, 4, 5, 6
            
                If .intCh < 10 Or .intCh > 100 Then
                
                    intLightNum = g_VGrid(g_intVGridNum(.intCh)).intLightNum
                    intShadowNum = g_VGrid(g_intVGridNum(.intCh)).intShadowNum
                    intBrushNum = g_VGrid(g_intVGridNum(.intCh)).intBrushNum
                
                ElseIf .intCh > 50 Then 'ロングノート
                
                    intLightNum = PEN_NUM.LONGNOTE_LIGHT
                    intShadowNum = PEN_NUM.LONGNOTE_SHADOW
                    intBrushNum = BRUSH_NUM.LONGNOTE
                
                Else
                
                    If .intAtt = 0 Then
                    
                        intLightNum = g_VGrid(g_intVGridNum(.intCh)).intLightNum
                        intShadowNum = g_VGrid(g_intVGridNum(.intCh)).intShadowNum
                        intBrushNum = g_VGrid(g_intVGridNum(.intCh)).intBrushNum
                    
                    Else 'If .intAtt = 1 Then
                    
                        intRet = .intCh Mod 10
                        
                        Select Case .intCh
                        
                            Case 11 To 15
                            
                                intLightNum = PEN_NUM.INV_KEY01_LIGHT + intRet - 1
                                intShadowNum = PEN_NUM.INV_KEY01_SHADOW + intRet - 1
                                intBrushNum = BRUSH_NUM.INV_KEY01 + intRet - 1
                            
                            Case 18
                            
                                intLightNum = PEN_NUM.KEY06_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY06_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY06
                            
                            Case 19
                            
                                intLightNum = PEN_NUM.KEY07_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY07_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY07
                            
                            Case 16
                            
                                intLightNum = PEN_NUM.KEY08_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY08_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY08
                            
                            Case 21 To 25
                            
                                intLightNum = PEN_NUM.INV_KEY11_LIGHT + intRet - 1
                                intShadowNum = PEN_NUM.INV_KEY11_SHADOW + intRet - 1
                                intBrushNum = BRUSH_NUM.INV_KEY11 + intRet - 1
                            
                            Case 28
                            
                                intLightNum = PEN_NUM.KEY16_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY16_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY16
                            
                            Case 29
                            
                                intLightNum = PEN_NUM.KEY17_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY17_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY17
                            
                            Case 26
                            
                                intLightNum = PEN_NUM.KEY18_LIGHT
                                intShadowNum = PEN_NUM.INV_KEY18_SHADOW
                                intBrushNum = BRUSH_NUM.INV_KEY18
                        
                        End Select
                    
                    End If
                
                End If
            
            Case 1 '通常選択
            
                intLightNum = PEN_NUM.SELECT_OBJ_LIGHT
                intShadowNum = PEN_NUM.SELECT_OBJ_SHADOW
                intBrushNum = BRUSH_NUM.SELECT_OBJ
            
            Case Else
            
                If .intSelect = 2 Then '白枠(編集モード)
                
                    intLightNum = PEN_NUM.EDIT_FRAME
                
                Else '赤枠(消去モード)
                
                    intLightNum = PEN_NUM.DELETE_FRAME
                
                End If
                
                intBrushNum = UBound(m_hBrush)
                
                hOldBrush = SelectObject(frmMain.picMain.hdc, m_hBrush(intBrushNum))
                hOldPen = SelectObject(frmMain.picMain.hdc, m_hPen(intLightNum))
                
                Call Rectangle(frmMain.picMain.hdc, X - 1, Y - OBJ_HEIGHT - 1, X + Width + 1, Y + 2)
                
                m_hPen(intLightNum) = SelectObject(frmMain.picMain.hdc, hOldPen)
                m_hBrush(intBrushNum) = SelectObject(frmMain.picMain.hdc, hOldBrush)
                
                Exit Sub
        
        End Select
    
    End With
    
    With frmMain.picMain
    
        hOldBrush = SelectObject(.hdc, m_hBrush(intBrushNum))
        hOldPen = SelectObject(.hdc, m_hPen(intLightNum))
        
        Call Rectangle(.hdc, X, Y - OBJ_HEIGHT, X + Width, Y + 1)
        
        m_hPen(intLightNum) = SelectObject(.hdc, m_hPen(intShadowNum))
        
        Call MoveToEx(.hdc, X, Y, 0)
        Call LineTo(.hdc, X + Width - 1, Y)
        Call LineTo(.hdc, X + Width - 1, Y - OBJ_HEIGHT)
        
        m_hPen(intShadowNum) = SelectObject(.hdc, hOldPen)
        m_hBrush(intBrushNum) = SelectObject(.hdc, hOldBrush)
        
        'Text = g_Obj(lngNum).lngID
        intRet = LenB(StrConv(Text, vbFromUnicode))
        
        Call GetTextExtentPoint32(.hdc, Text, intRet, retSize)
        
        Y = Y - (OBJ_HEIGHT + retSize.Height) \ 2 + 1
        
        'If g_Obj(lngNum).intSelect = 1 Then
        If retObj.intSelect = 1 Then
        
            Call SetTextColor(.hdc, &HFFFFFF)
            Call TextOut(.hdc, X + 3, Y, Text, intRet)
            Call SetTextColor(.hdc, &H0)
            Call TextOut(.hdc, X + 2, Y, Text, intRet)
        
        Else
        
            Call SetTextColor(.hdc, &H0)
            Call TextOut(.hdc, X + 3, Y, Text, intRet)
            Call SetTextColor(.hdc, &HFFFFFF)
            Call TextOut(.hdc, X + 2, Y, Text, intRet)
        
        End If
    
    End With
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "DrawObj")
End Sub

Public Sub DrawObjRect(ByVal Num As Long)
On Error GoTo Err:

    Dim X       As Long
    Dim Y       As Long
    Dim Width   As Integer
    
    With g_Obj(Num)
    
        If g_intVGridNum(.intCh) = 0 Then Exit Sub
        
        'X = g_VGrid(g_intVGridNum(.intCh)).lngObjLeft * g_disp.Width - g_disp.X + 1
        X = (g_VGrid(g_intVGridNum(.intCh)).lngObjLeft - g_disp.X) * g_disp.Width + 1
        Y = frmMain.picMain.ScaleHeight + OBJ_DIFF - (g_Measure(.intMeasure).lngY + .lngPosition - g_disp.Y) * g_disp.Height
        Width = GRID_WIDTH * g_disp.Width - 1
        
        If .intAtt = 2 Or (.intCh >= 51 And .intCh <= 69) Then
        
            X = X + 3
            Width = Width - 6
        
        End If
    
    End With
    
    With frmMain.picMain
    
        Call Rectangle(.hdc, X - 1, Y - OBJ_HEIGHT - 1, X + Width + 1, Y + 2)
    
    End With
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "DrawObjRect")
End Sub

Public Sub DrawObjMax(ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer)
On Error GoTo Err:

    Dim i       As Long
    Dim lngRet  As Long
    Dim retObj  As g_udtObj
    
    With g_Mouse
    
        .Shift = Shift
        .X = X
        .Y = Y
    
    End With
    
    If g_blnIgnoreInput Then Exit Sub
    
    Call SetObjData(retObj, X, Y) ', g_disp.X, g_disp.Y)
    
    With retObj
    
        If frmMain.tlbMenu.Buttons("Write").value = tbrPressed Then
        
            If .intCh >= 11 And .intCh <= 29 Then
            
                If Shift And vbCtrlMask Then
                
                    .intAtt = 1
                
                ElseIf Shift And vbShiftMask Then
                
                    .intCh = .intCh + 40
                    .intAtt = 2
                
                End If
            
            End If
            
            'If Shift And vbAltMask Then
            
                If frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex) Then
                
                    lngRet = 192 \ (frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex))
                    .lngPosition = (.lngPosition \ lngRet) * lngRet
                
                End If
            
            'End If
        
        End If
    
    End With
    
    If frmMain.tlbMenu.Buttons("Write").value = tbrUnpressed Then
    
        With retObj
        
            lngRet = g_Measure(.intMeasure).lngY + .lngPosition
            
            For i = UBound(g_Obj) - 1 To 0 Step -1
            
                If (g_Obj(i).intCh = .intCh) Or (.intAtt = 2 And g_Obj(i).intCh + 40 = .intCh) Then
                
                    If g_Measure(g_Obj(i).intMeasure).lngY + g_Obj(i).lngPosition + OBJ_HEIGHT / g_disp.Height >= lngRet And g_Measure(g_Obj(i).intMeasure).lngY + g_Obj(i).lngPosition <= lngRet Then
                    
                        'If frmMain.tlbMenu.Buttons("Write").value = tbrUnpressed Then
                        
                            If frmMain.tlbMenu.Buttons("Edit").value = tbrPressed Then
                            
                                .intSelect = 2
                            
                            ElseIf frmMain.tlbMenu.Buttons("Delete").value = tbrPressed Then
                            
                                .intSelect = 3
                            
                            End If
                            
                            .intAtt = g_Obj(i).intAtt
                            
                            If .intAtt = 2 Then .intCh = .intCh + 40
                            
                            .sngValue = g_Obj(i).sngValue
                            .lngPosition = g_Obj(i).lngPosition
                            .intMeasure = g_Obj(i).intMeasure
                            .lngHeight = i
                        
                        'End If
                        
                        '.lngHeight = i
                        
                        '.lngPosition = g_Obj(i).lngPosition
                        'とりあえず切っておいたよ、その代わり上に追加しておいた v1.1.7
                        '↑何のために消したのかわからねー上にバグるので復活させました v1.2.3
                        '↑これ消さないと書き込みモード時にオブジェに吸い込まれる。で、何がバグったんだっけ？ v1.3.0
                        '↑小節をまたがるオブジェに関してえらいことになる。どーしよう。 v1.3.5
                        '↓これを上に移動して解決？した？かも？ v1.3.6
                        '.intMeasure = g_Obj(i).intMeasure
                        
                        Exit For
                    
                    End If
                
                End If
            
            Next i
        
        End With
    
    End If
    
    Call DrawStatusBar(retObj, Shift)
    
    If frmMain.tlbMenu.Buttons("Write").value = tbrPressed Then
    
        If retObj.intCh <> g_Obj(UBound(g_Obj)).intCh Or retObj.intAtt <> g_Obj(UBound(g_Obj)).intAtt Or retObj.intMeasure <> g_Obj(UBound(g_Obj)).intMeasure Or retObj.lngPosition <> g_Obj(UBound(g_Obj)).lngPosition Or retObj.sngValue <> g_Obj(UBound(g_Obj)).sngValue Then
        
            Call CopyObj(g_Obj(UBound(g_Obj)), retObj)
            g_lngObjID(g_Obj(UBound(g_Obj)).lngID) = UBound(g_Obj)
        
        Else
        
            g_Obj(UBound(g_Obj)).lngHeight = retObj.lngHeight
            
            Exit Sub
        
        End If
        
    Else
    
        If retObj.intSelect <> 2 And retObj.intSelect <> 3 Then
        
            retObj.intCh = 0
            g_Obj(UBound(g_Obj)).intCh = 0
        
        End If
        
        'If retObj.intCh <> g_Obj(UBound(g_Obj)).intCh Or retObj.intAtt <> g_Obj(UBound(g_Obj)).intAtt Or g_Measure(retObj.intMeasure).lngY + retObj.lngPosition > g_Measure(g_Obj(UBound(g_Obj)).intMeasure).lngY + g_Obj(UBound(g_Obj)).lngPosition + OBJ_HEIGHT / g_disp.Height Or g_Measure(retObj.intMeasure).lngY + retObj.lngPosition < g_Measure(g_Obj(UBound(g_Obj)).intMeasure).lngY + g_Obj(UBound(g_Obj)).lngPosition Then
        If retObj.lngHeight <> g_Obj(UBound(g_Obj)).lngHeight Then
        
            If g_Obj(retObj.lngHeight).intCh Then retObj.lngPosition = g_Obj(retObj.lngHeight).lngPosition
            
            Call CopyObj(g_Obj(UBound(g_Obj)), retObj)
            g_lngObjID(g_Obj(UBound(g_Obj)).lngID) = UBound(g_Obj)
        
        Else
        
            Exit Sub
        
        End If
    
    End If
    
    'Call DrawStatusBar(UBound(g_Obj), Shift)
    
    Call frmMain.picMain.Cls
    
    If g_Obj(UBound(g_Obj)).intCh Then
    
        Call InitPen
        
        frmMain.picMain.Font.Size = 8
        
        Call DrawObj(g_Obj(UBound(g_Obj)))
        
        Call DeletePen
    
    End If
    
    If g_disp.intEffect Then Call modEasterEgg.DrawEffect
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "DrawObjMax")
End Sub

Public Sub SetObjData(ByRef retObj As g_udtObj, ByVal X As Single, ByVal Y As Single) ', ByVal g_disp.x As Long, ByVal g_disp.y As Long)

    Dim i       As Long
    Dim lngRet  As Long
    
    If X < 0 Then
    
        X = 0
    
    ElseIf X > frmMain.picMain.ScaleWidth Then
    
        X = frmMain.picMain.ScaleWidth
    
    End If
    
    'lngRet = (X + g_disp.X) / g_disp.Width
    lngRet = X / g_disp.Width + g_disp.X
    
    retObj.intCh = 8
    
    For i = 0 To UBound(g_VGrid)
    
        With g_VGrid(i)
        
            If .blnDraw = True And .intCh <> 0 Then
            
                If .lngLeft <= lngRet Then
                
                    retObj.intCh = .intCh
                
                Else
                
                    Exit For
                
                End If
            
            End If
        
        End With
    
    Next i
    
    With retObj
    
        .lngID = g_lngIDNum
        .lngHeight = UBound(g_Obj)
        
        If Y < 1 Then
        
            Y = 1
        
        ElseIf Y > frmMain.picMain.ScaleHeight + OBJ_DIFF Then
        
            Y = frmMain.picMain.ScaleHeight + OBJ_DIFF
        
        End If
        
        lngRet = (frmMain.picMain.ScaleHeight - Y + OBJ_DIFF) / g_disp.Height + g_disp.Y
        
        'For i = g_Disp.intStartMeasure To g_Disp.intEndMeasure
        For i = 0 To 999
        
            If g_Measure(i).lngY <= lngRet Then
            
                .intMeasure = i
                .lngPosition = lngRet - g_Measure(i).lngY
                
                If .lngPosition > g_Measure(i).intLen Then .lngPosition = g_Measure(i).intLen - 1
            
            Else
            
                Exit For
            
            End If
        
        Next i
        
        Select Case .intCh
        
            Case 3, 8, 9
            
                .sngValue = 0
            
            Case 4, 6, 7
            
                'If frmMain.mnuOptionsNumFF.Checked Then
                
                    '.sngValue = lngNumConv(Hex$(frmMain.lstBMP.ListIndex + 1))
                
                'Else
                
                    '.sngValue = frmMain.lstBMP.ListIndex + 1
                
                'End If
                
                .sngValue = frmMain.lngFromLong(frmMain.lstBMP.ListIndex + 1)
            
            Case Else
            
                'If frmMain.mnuOptionsNumFF.Checked Then
                
                    '.sngValue = lngNumConv(Hex$(frmMain.lstWAV.ListIndex + 1))
                
                'Else
                
                    '.sngValue = frmMain.lstWAV.ListIndex + 1
                
                'End If
                
                .sngValue = frmMain.lngFromLong(frmMain.lstWAV.ListIndex + 1)
        
        End Select
    
    End With

End Sub

'Public Sub DrawStatusBar(ByVal ObjNum As Long, ByVal Shift As Integer)
Public Sub DrawStatusBar(ByRef retObj As g_udtObj, ByVal Shift As Integer)

    Dim strRet      As String
    Dim lngRet      As Long
    Dim strArray()  As String
    
    'With g_Obj(ObjNum)
    With retObj
        
        strRet = "Position:  " & .intMeasure & g_strStatusBar(23) & "  "
        
        'If Not Shift And vbAltMask Then
        
            lngRet = frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex)
        
        'End If
        
        If lngRet Then
        
            If .intSelect > 1 And .lngPosition <> 0 Then
            
                lngRet = modInput.intGCD(.lngPosition, g_Measure(.intMeasure).intLen)
                
                If lngRet > 192 \ frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex) Then
                
                    lngRet = frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex)
                
                Else
                
                    lngRet = 192 \ lngRet
                
                End If
            
            End If
            
            strRet = strRet & .lngPosition * lngRet \ 192 & "/" & g_Measure(.intMeasure).intLen * lngRet \ 192
        
        Else
        
            strRet = strRet & .lngPosition & "/" & g_Measure(.intMeasure).intLen
        
        End If
        
        strRet = strRet & "  "
        
        Select Case .intCh
        
            Case Is > 100
            
                strRet = strRet & g_strStatusBar(1) & " " & Format$(.intCh - 100, "00")
            
            Case Is < 10
            
                strRet = strRet & g_strStatusBar(.intCh)
            
            Case 11 To 15
            
                strRet = strRet & g_strStatusBar(11) & .intCh - 10
            
            Case 16
            
                strRet = strRet & g_strStatusBar(13)
            
            Case 18, 19
            
                strRet = strRet & g_strStatusBar(11) & .intCh - 12
            
            Case 21 To 25
            
                strRet = strRet & g_strStatusBar(12) & .intCh - 20
            
            Case 26
            
                strRet = strRet & g_strStatusBar(14)
            
            Case 28, 29
            
                strRet = strRet & g_strStatusBar(12) & .intCh - 22
            
            Case 51 To 55
            
                strRet = strRet & g_strStatusBar(11) & .intCh - 50
            
            Case 56
            
                strRet = strRet & g_strStatusBar(13)
            
            Case 58, 59
            
                strRet = strRet & g_strStatusBar(11) & .intCh - 52
            
            Case 61 To 65
            
                strRet = strRet & g_strStatusBar(12) & .intCh - 60
            
            Case 66
            
                strRet = strRet & g_strStatusBar(14)
            
            Case 68, 69
            
                strRet = strRet & g_strStatusBar(12) & .intCh - 62
            
        
        End Select
        
        If .intCh >= 11 And .intCh <= 29 Then
        
            If .intAtt = 1 Then
            
                strRet = strRet & " " & g_strStatusBar(15)
            
            ElseIf .intAtt = 2 Then
            
                strRet = strRet & " " & g_strStatusBar(16)
            
            End If
        
        ElseIf .intCh >= 51 And .intCh <= 69 Then
        
            'If lngChangeMaxMeasure(.intMeasure) Then Call ChangeResolution

            strRet = strRet & " " & g_strStatusBar(16)
        
        End If
        
        frmMain.staMain.Panels("Position").Text = strRet
        
        strArray = Split(Mid$(frmMain.lstMeasureLen.List(.intMeasure), 6), "/")
        
        frmMain.staMain.Panels("Measure").Text = Right$(" " & strArray(0), 2) & "/" & Left$(strArray(1) & " ", 2)
    
    End With

End Sub

Public Sub DrawSelectArea()

    Dim i           As Long
    Dim lngRet      As Long
    Dim hOldPen     As Long
    Dim hNewPen     As Long
    Dim objBrush    As LOGBRUSH
    Dim hOldBrush   As Long
    Dim hNewBrush   As Long
    Dim retRect     As RECT
    
    hNewPen = CreatePen(PS_SOLID, 1, g_lngPenColor(PEN_NUM.EDIT_FRAME))
    hOldPen = SelectObject(frmMain.picMain.hdc, hNewPen)
    
    With objBrush
        .lbStyle = BS_NULL
        .lbColor = 0
        .lbHatch = BS_NULL
    End With
    
    'hNewBrush = CreateHatchBrush(HS_BDIAGONAL, g_lngPenColor(PEN_NUM.EDIT_FRAME))
    hNewBrush = CreateBrushIndirect(objBrush)
    hOldBrush = SelectObject(frmMain.picMain.hdc, hNewBrush)
    
    Call frmMain.picMain.Cls
    
    With retRect
    
        .Top = (g_SelectArea.Y1 - g_disp.Y) * -g_disp.Height + frmMain.picMain.ScaleHeight
        '.Left = g_SelectArea.X1 * g_disp.Width - g_disp.X
        .Left = (g_SelectArea.X1 - g_disp.X) * g_disp.Width
        .Right = g_Mouse.X
        .Bottom = g_Mouse.Y
        
        Call Rectangle(frmMain.picMain.hdc, .Left, .Top, .Right, .Bottom)
    
    End With
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intSelect = 4 Or .intSelect = 5 Then
            
                lngRet = g_Measure(.intMeasure).lngY + .lngPosition
                
                If g_disp.lngStartPos <= lngRet And g_disp.lngEndPos >= lngRet Then
                
                    Call modDraw.DrawObjRect(i)
                
                End If
            
            End If
        
        End With
    
    Next i
    
    hNewPen = SelectObject(frmMain.picMain.hdc, hOldPen)
    Call DeleteObject(hNewPen)
    
    hNewBrush = SelectObject(frmMain.picMain.hdc, hOldBrush)
    Call DeleteObject(hNewBrush)

End Sub

Public Function lngChangeMaxMeasure(ByVal intMeasure As Integer) As Long

    With g_disp
    
        If intMeasure + 16 > .intMaxMeasure Then
        
            .intMaxMeasure = intMeasure + 16
            
            If g_disp.intMaxMeasure > 999 Then .intMaxMeasure = 999
            
            'Call ChangeResolution
            lngChangeMaxMeasure = 1
        
        End If
    
    End With

End Function

Public Sub ChangeResolution()

    Dim i       As Long
    Dim intRet  As Integer
    Dim lngRet  As Long
    Dim sngRet  As Single
    
    With g_disp
    
        intRet = .intResolution
        
        For i = 0 To .intMaxMeasure
        
            lngRet = lngRet + g_Measure(i).intLen
        
        Next i
        
        'sngRet = 96 / (((64 / 4) * 1000 * 2) / (lngRet / 96))
        sngRet = lngRet / 32000
        
        Select Case sngRet
            Case Is > 48: .intResolution = 96
            Case Is > 24: .intResolution = 48
            Case Is > 12: .intResolution = 24
            Case Is > 6: .intResolution = 12
            Case Is > 3: .intResolution = 6
            Case Is > 1: .intResolution = 3
            Case Else: .intResolution = 1
        End Select
        
        If intRet = .intResolution Then Exit Sub
                
        frmMain.vsbMain.value = (frmMain.vsbMain.value / .intResolution) * intRet
    
    End With
    
    With frmMain.cboVScroll
    
        Call .Clear
        intRet = 0
        
        'For i = 0 To 6
        For i = 1 To 6
        
            'lngRet = 2 ^ (i - 1) * 3
            'If i = 0 Then lngRet = 1
            lngRet = 2 ^ i * 3
            
            If lngRet >= g_disp.intResolution Then
            
                Call .AddItem(lngRet, intRet)
                .ItemData(intRet) = lngRet \ g_disp.intResolution
                
                intRet = intRet + 1
            
            End If
        
        Next i
        
        .ListIndex = .ListCount - 2
    
    End With

End Sub

Public Sub CopyObj(ByRef destObj As g_udtObj, ByRef srcObj As g_udtObj)

    With destObj
    
        .lngID = srcObj.lngID
        .intCh = srcObj.intCh
        .lngHeight = srcObj.lngHeight
        .intMeasure = srcObj.intMeasure
        .lngPosition = srcObj.lngPosition
        .intSelect = srcObj.intSelect
        .sngValue = srcObj.sngValue
        .intAtt = srcObj.intAtt
    
    End With

End Sub

Public Sub RemoveObj(ByVal lngNum As Long)
On Error GoTo Err:

    With g_Obj(lngNum)
        g_lngObjID(.lngID) = -1
        .lngID = 0
        .intCh = 0
        .lngHeight = 0
        .intMeasure = 0
        .lngPosition = 0
        .intSelect = 0
        .sngValue = 0
        .intAtt = 0
    End With
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "RemoveObj")
End Sub

Public Sub ArrangeObj()

    Dim i       As Long
    Dim lngRet  As Long
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intCh Then
        
            Call modInput.SwapObj(lngRet, i)
            
            If i = g_Obj(UBound(g_Obj)).lngHeight Then g_Obj(UBound(g_Obj)).lngHeight = lngRet
            
            lngRet = lngRet + 1
        
        End If
    
    Next i
    
    Call CopyObj(g_Obj(lngRet), g_Obj(UBound(g_Obj)))
    
    ReDim Preserve g_Obj(lngRet)

End Sub

Public Sub MoveSelectedObj()
On Error GoTo Err:

    Dim i       As Long
    Dim j       As Long
    Dim lngRet  As Long
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intSelect Then
        
            lngRet = lngRet + 1
        
        End If
    
    Next i
    
    If lngRet = 0 Then Exit Sub
    
    j = UBound(g_Obj)
    
    ReDim Preserve g_Obj(j + lngRet)
    
    Call modInput.SwapObj(UBound(g_Obj), j)
    
    lngRet = 0
    
    For i = 0 To j - 1
    
        If g_Obj(i).intSelect Then
        
            Call modInput.SwapObj(i, j + lngRet)
            
            If i = g_Obj(UBound(g_Obj)).lngHeight Then g_Obj(UBound(g_Obj)).lngHeight = j + lngRet
            
            lngRet = lngRet + 1
        
        End If
    
    Next i
    
    Call ArrangeObj
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "MoveSelectedObj")
End Sub

Public Sub ObjSelectCancel()

    Dim i   As Long
    
    For i = 0 To UBound(g_Obj) - 1
    
        g_Obj(i).intSelect = 0
    
    Next i

End Sub

Public Sub InitPen()

    Dim i           As Long
    Dim objBrush    As LOGBRUSH
    
    'ペン生成
    
    For i = 0 To UBound(m_hPen)
    
        m_hPen(i) = CreatePen(PS_SOLID, 1, g_lngPenColor(i))
    
    Next i
    
    'ブラシ生成
    
    For i = 0 To UBound(m_hBrush) - 1
    
        m_hBrush(i) = CreateSolidBrush(g_lngBrushColor(i))
    
    Next i
    
    With objBrush
    
        .lbStyle = BS_NULL
        .lbColor = 0
        .lbHatch = BS_NULL
    
    End With
    
    m_hBrush(UBound(m_hBrush)) = CreateBrushIndirect(objBrush)

End Sub

Public Sub DeletePen()

    Dim i   As Long
    
    'ペン削除
    For i = 0 To UBound(m_hPen)
    
        Call DeleteObject(m_hPen(i))
    
    Next i
    
    'ブラシ削除
    For i = 0 To UBound(m_hBrush)
    
        Call DeleteObject(m_hBrush(i))
    
    Next i

End Sub
