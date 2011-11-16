Attribute VB_Name = "modMain"
Option Explicit

#Const MODE_DEBUG = True

Private Const INI_VERSION = 3

Public Const RELEASEDATE = "2006-12-27T15:46:39"



#Const MODE_SPEEDTEST = False

#If MODE_SPEEDTEST Then

Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

#End If

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'iniファイルに読みこむためのAPI
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long 'iniファイルを書きこむためのAPI

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'Get/SetWindowPlacement・ShellExecute 関連の定数
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

'GetWindowLong 関連の定数
Public Const GWL_STYLE = -16
Public Const GWL_EXSTYLE = -20

'GetStockObject 関連の定数
Private Const OEM_FIXED_FONT = 10
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEFAULT_GUI_FONT = 17

Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName(1 To LF_FACESIZE)    As Byte
End Type

Public Type POINTAPI
    X   As Long
    Y   As Long
End Type

Public Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type WINDOWPLACEMENT
    Length              As Long
    flags               As Long
    showCmd             As Long
    ptMinPosition       As POINTAPI
    ptMaxPosition       As POINTAPI
    rcNormalPosition    As RECT
End Type



Public Const PI = 3.14159265358979
Public Const RAD = PI / 180

Public Enum BGA_PARA
    BGA_NUM
    BGA_X1
    BGA_Y1
    BGA_X2
    BGA_Y2
    BGA_dX
    BGA_dY
End Enum

Public Enum CMD_LOG
    NONE
    OBJ_ADD
    OBJ_DEL
    OBJ_MOVE
    OBJ_CHANGE
    MSR_ADD
    MSR_DEL
    MSR_CHANGE
    WAV_CHANGE
    BMP_CHANGE
    LIST_ALIGN
    LIST_DEL
End Enum

Public g_strAppTitle    As String

Private Type m_udtMouse
    X       As Single
    Y       As Single
    Shift   As Integer
    Button  As Integer
End Type

Public g_Mouse  As m_udtMouse

Private Type m_udtDisplay
    X               As Long
    Y               As Long
    Width           As Single
    Height          As Single
    lngMaxX         As Long
    lngMaxY         As Long
    intStartMeasure As Integer
    intEndMeasure   As Integer
    lngStartPos     As Long
    lngEndPos       As Long
    intMaxMeasure   As Integer '最大表示小節
    intResolution   As Integer '分解能
    intEffect       As Integer '画面効果
End Type

Public g_disp   As m_udtDisplay

Private Type m_udtBMS
    strDir          As String   'ディレクトリ
    strFileName     As String   'BMSファイル名
    intPlayerType   As Integer  '#PLAYER
    strGenre        As String   '#GENRE
    strTitle        As String   '#TITLE
    strArtist       As String   '#ARTIST
    sngBPM          As Single   '#BPM
    lngPlayLevel    As Long     '#PLAYLEVEL
    intPlayRank     As Integer  '#RANK
    sngTotal        As Single   '#TOTAL
    intVolume       As Integer  '#VOLWAV
    strStageFile    As String   '#STAGEFILE
    blnSaveFlag     As Boolean
End Type

Public g_BMS    As m_udtBMS

Private Type m_udtVerticalLine
    blnVisible      As Boolean
    intCh           As Integer
    strText         As String
    intWidth        As Integer
    lngLeft         As Long
    lngObjLeft      As Long
    lngBackColor    As Long
    intLightNum     As Integer
    intShadowNum    As Integer
    intBrushNum     As Integer
    blnDraw         As Boolean
End Type

Public g_VGrid(61)  As m_udtVerticalLine

Public g_intVGridNum(132)   As Integer

Public Type g_udtObj
    lngID       As Long
    intCh       As Integer
    intAtt      As Integer
    intMeasure  As Integer
    lngHeight   As Long
    lngPosition As Long
    sngValue    As Single
    intSelect   As Integer
    '0・・・未選択
    '1・・・選択
    '2・・・白枠 (編集モード)
    '3・・・赤枠 (消去モード)
    '4・・・選択範囲内にあるオブジェ、選択中
    '5・・・選択範囲を展開した時に既に選択状態にあったオブジェ、選択中
    '6・・・5番かつ選択範囲内、つまり選択状態でなくなったオブジェ
End Type

Public g_Obj()  As g_udtObj

Public g_lngObjID() As Long
Public g_lngIDNum   As Long

Private Type m_udtMeasure
    intLen  As Integer
    lngY    As Long
End Type

Public g_Measure(999)   As m_udtMeasure

Public g_strWAV(1295)   As String
Public g_strBMP(1295)   As String
Public g_strBGA(1295)   As String
Public g_sngBPM(1295)   As Single
Public g_lngSTOP(1295)  As Long

Private Type m_udtSelectArea
    blnFlag As Boolean
    X1      As Single
    Y1      As Single
    X2      As Single
    Y2      As Single
End Type

Public g_SelectArea As m_udtSelectArea

Public g_strLangFileName()  As String
Public g_strThemeFileName() As String
Public g_strStatusBar(23)   As String

Public g_blnIgnoreInput     As Boolean
Public g_strAppDir          As String
Public g_strHelpFilename    As String
Public g_strFiler           As String
Public g_strRecentFiles(4)  As String

'Public g_strInputLog()  As String
'Public g_lngInputLogPos As Long
Public g_InputLog       As New clsLog

Public Type g_udtViewer
    strAppName  As String
    strAppPath  As String
    strArgAll   As String
    strArgPlay  As String
    strArgStop  As String
End Type

Public g_Viewer()   As g_udtViewer

Public Enum Message
    ERR_01
    ERR_02
    ERR_FILE_NOT_FOUND
    ERR_LOAD_CANCEL
    ERR_SAVE_ERROR
    ERR_SAVE_CANCEL
    ERR_OVERFLOW_LARGE
    ERR_OVERFLOW_SMALL
    ERR_OVERFLOW_BPM
    ERR_OVERFLOW_STOP
    ERR_APP_NOT_FOUND
    ERR_FILE_ALREADY_EXIST
    MSG_CONFIRM
    MSG_FILE_CHANGED
    MSG_INI_CHANGED
    MSG_ALIGN_LIST
    MSG_DELETE_FILE
    INPUT_BPM
    INPUT_STOP
    INPUT_RENAME
    INPUT_SIZE
    Max
End Enum

Public g_Message(Message.Max - 1)   As String

Public Sub Main()

    Dim i           As Long
    Dim wp          As WINDOWPLACEMENT
    Dim strRet      As String
    Dim intRet      As Integer
    Dim lngFFile    As Long
    
    If Right$(App.Path, 1) = "\" Then
    
        g_strAppDir = App.Path
    
    Else
    
        g_strAppDir = App.Path & "\"
    
    End If
    
    g_strAppTitle = "BMx Sequence Editor " & App.Major & "." & App.Minor & "." & App.Revision
    
    #If MODE_DEBUG = False Then
    
        Call modMessage.SubClass(frmMain.hwnd)
    
    #End If
    
    #If MODE_SPEEDTEST Then
    
        Call timeBeginPeriod(1)
    
    #End If
    
    ReDim g_strLangFileName(0)
    
    'ReDim g_strInputLog(0)
    Call g_InputLog.Clear
    
    ReDim g_Viewer(1)
    
    If Dir(g_strAppDir & "bmse_viewer.ini", vbNormal) = vbNullString Then
    
        lngFFile = FreeFile()
        
        Open g_strAppDir & "bmse_viewer.ini" For Output As #lngFFile
        
            Print #lngFFile, "uBMplay"
            Print #lngFFile, "uBMplay.exe"
            Print #lngFFile, "-P -N0 <filename>"
            Print #lngFFile, "-P -N<measure> <filename>"
            Print #lngFFile, "-S"
            Print #lngFFile,
            Print #lngFFile, "WAview"
            Print #lngFFile, "C:\Program Files\Winamp\Plugins\WAview.exe"
            Print #lngFFile, "-Lbml <filename>"
            Print #lngFFile, "-N<measure>"
            Print #lngFFile, "-S"
            Print #lngFFile,
            Print #lngFFile, "nBMplay"
            Print #lngFFile, "nbmplay.exe"
            Print #lngFFile, "-P -N0 <filename>"
            Print #lngFFile, "-P -N<measure> <filename>"
            Print #lngFFile, "-S"
            Print #lngFFile,
            Print #lngFFile, "BMEV"
            Print #lngFFile, "BMEV.exe"
            Print #lngFFile, "-P -N0 <filename>"
            Print #lngFFile, "-P -N<measure> <filename>"
            Print #lngFFile, "-S"
            Print #lngFFile,
            Print #lngFFile, "BMS Viewer"
            Print #lngFFile, "bmview.exe"
            Print #lngFFile, "-S -P -N0 <filename>"
            Print #lngFFile, "-S -P -N<measure> <filename>"
            Print #lngFFile, "-S"
        
        Close #lngFFile
    
    End If
    
    i = 0
    lngFFile = FreeFile()
    
    Open g_strAppDir & "bmse_viewer.ini" For Input As #lngFFile
    
        Do While Not EOF(lngFFile)
        
            Line Input #lngFFile, strRet
            
            Select Case i Mod 6
            
                Case 0
                
                    If Len(strRet) = 0 Then Exit Do
                    g_Viewer(UBound(g_Viewer)).strAppName = strRet
                
                Case 1
                
                    If Len(strRet) = 0 Then Exit Do
                    g_Viewer(UBound(g_Viewer)).strAppPath = strRet
                
                Case 2
                
                    g_Viewer(UBound(g_Viewer)).strArgAll = strRet
                
                Case 3
                
                    g_Viewer(UBound(g_Viewer)).strArgPlay = strRet
                
                Case 4
                
                    g_Viewer(UBound(g_Viewer)).strArgStop = strRet
                    
                    Call frmMain.cboViewer.AddItem(g_Viewer(UBound(g_Viewer)).strAppName)
                    ReDim Preserve g_Viewer(UBound(g_Viewer) + 1)
            
            End Select
            
            i = i + 1
        
        Loop
    
    Close #lngFFile
    
    ReDim Preserve g_Viewer(frmMain.cboViewer.ListCount)
    
    With frmMain
    
        If .cboViewer.ListCount = 0 Then
        
            .tlbMenu.Buttons("PlayAll").Enabled = False
            .tlbMenu.Buttons("Play").Enabled = False
            .tlbMenu.Buttons("Stop").Enabled = False
            .mnuToolsPlayAll.Enabled = False
            .mnuToolsPlay.Enabled = False
            .mnuToolsPlayStop.Enabled = False
            .cboViewer.Enabled = False
        
        End If
    
    End With
    
    'ランゲージファイル読み込み
    strRet = Dir(g_strAppDir & "lang\*.ini")
    intRet = 0
    
    Do While strRet <> ""
    
        If strGet_ini("Main", "Key", "", "lang\" & strRet) = "BMSE" Then
        
            intRet = intRet + 1
            
            ReDim Preserve g_strLangFileName(intRet)
            
            g_strLangFileName(intRet) = strRet
            
            Call Load(frmMain.mnuLanguage(intRet))
            
            With frmMain.mnuLanguage(intRet)
            
                .Caption = "&" & strGet_ini("Main", "Language", strRet, "lang\" & strRet)
                
                If .Caption = "&" Then .Caption = "&" & strRet
                
                .Visible = True
            
            End With
        
        End If
        
        strRet = Dir
    
    Loop
    
    If intRet Then
    
        frmMain.mnuLanguage(0).Visible = False
    
    Else
    
        frmMain.mnuLanguageParent.Enabled = False
    
    End If
    
    'テーマファイル読み込み
    strRet = Dir(g_strAppDir & "theme\*.ini")
    intRet = 0
    
    Do While strRet <> ""
    
        If strGet_ini("Main", "Key", "", "theme\" & strRet) = "BMSE" Then
        
            intRet = intRet + 1
            
            ReDim Preserve g_strThemeFileName(intRet)
            
            g_strThemeFileName(intRet) = strRet
            
            Load frmMain.mnuTheme(intRet)
            
            With frmMain.mnuTheme(intRet)
            
                .Caption = "&" & strGet_ini("Main", "Name", strRet, "theme\" & strRet)
                
                If .Caption = "&" Then .Caption = "&" & strRet
                
                .Visible = True
            
            End With
        
        End If
        
        strRet = Dir
    
    Loop
    
    If intRet Then
    
        frmMain.mnuTheme(0).Visible = False
    
    Else
    
        frmMain.mnuThemeParent.Enabled = False
    
    End If
    
    '初期化
    With g_BMS
    
        .intPlayerType = 1
        .strGenre = ""
        .strTitle = ""
        .strArtist = ""
        .sngBPM = Val(frmMain.txtBPM.Text)
        .lngPlayLevel = 1
        .intPlayRank = 3
        .sngTotal = 0
        .intVolume = 0
        .blnSaveFlag = True
    
    End With
    
    ReDim g_Obj(0)
    ReDim g_lngObjID(0)
    g_lngIDNum = 0
    
    For i = 0 To 256 + 64
    
        g_sngSin(i) = Sin(i * PI / 128)
    
    Next i
    
    With frmMain
    
        For i = 1 To 5
        
            Call .tlbMenu.Buttons("Open").ButtonMenus.Add(i, , "&" & i & ":")
            .tlbMenu.Buttons("Open").ButtonMenus(i).Enabled = False
            .tlbMenu.Buttons("Open").ButtonMenus(i).Visible = False
        
        Next i
        
        For i = 0 To .mnuRecentFiles.UBound
        
            '.mnuRecentFiles(i).Enabled = False
            .mnuRecentFiles(i).Visible = False
        
        Next i
        
        .mnuLineRecent.Visible = False
        .mnuHelpOpen.Enabled = False
        
        .Caption = g_strAppTitle
        
        Call LoadConfig
        
        .dlgMain.InitDir = g_strAppDir
        .dlgMain.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
        
        g_disp.lngMaxY = .vsbMain.Max
        g_disp.lngMaxX = .hsbMain.Min
        
        .hsbMain.SmallChange = OBJ_WIDTH
        .hsbMain.LargeChange = OBJ_WIDTH * 4
        
        .cboPlayer.ListIndex = g_BMS.intPlayerType - 1
        .cboPlayLevel = g_BMS.lngPlayLevel
        .cboPlayRank.ListIndex = g_BMS.intPlayRank
        
        If strGet_ini("Options", "UseOldFormat", False, "bmse.ini") Then
        
            For i = 1 To 255
            
                strRet = Right$("0" & Hex$(i), 2)
                Call .lstWAV.AddItem("#WAV" & strRet & ":", i - 1)
                Call .lstBMP.AddItem("#BMP" & strRet & ":", i - 1)
                Call .lstBGA.AddItem("#BGA" & strRet & ":", i - 1)
            
            Next i
            
            frmWindowPreview.cmdPreviewEnd.Caption = "FF"
        
        Else
        
            For i = 1 To 1295
            
                strRet = modInput.strNumConv(i)
                Call .lstWAV.AddItem("#WAV" & strRet & ":", i - 1)
                Call .lstBMP.AddItem("#BMP" & strRet & ":", i - 1)
                Call .lstBGA.AddItem("#BGA" & strRet & ":", i - 1)
            
            Next i
        
        End If
        
        For i = 0 To 999
        
            Call .lstMeasureLen.AddItem("#" & Format$(i, "000") & ":4/4", i)
            g_Measure(i).intLen = 192
        
        Next i
        
        For i = 0 To UBound(g_VGrid)
        
            With g_VGrid(i)
            
                .intCh = Choose(i + 1, 0, 8, 9, 0, 21, 16, 11, 12, 13, 14, 15, 18, 19, 16, 0, 26, 21, 22, 23, 24, 25, 28, 29, 26, 0, 4, 7, 6, 0, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 0)
                'If .intCh Then g_intVGridNum(.intCh) = i
                .blnVisible = True
                
                Select Case .intCh
                
                    Case 3, 8, 9 'BPM/STOP
                    
                        .intLightNum = PEN_NUM.BPM_LIGHT
                        .intShadowNum = PEN_NUM.BPM_SHADOW
                        .intBrushNum = BRUSH_NUM.BPM
                    
                    Case 4, 6, 7 'BGA/Layer/Poor
                    
                        .intLightNum = PEN_NUM.BGA_LIGHT
                        .intShadowNum = PEN_NUM.BGA_SHADOW
                        .intBrushNum = BRUSH_NUM.BGA
                    
                    Case 11
                    
                        .intLightNum = PEN_NUM.KEY01_LIGHT
                        .intShadowNum = PEN_NUM.KEY01_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY01
                    
                    Case 12
                    
                        .intLightNum = PEN_NUM.KEY02_LIGHT
                        .intShadowNum = PEN_NUM.KEY02_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY02
                    
                    Case 13
                    
                        .intLightNum = PEN_NUM.KEY03_LIGHT
                        .intShadowNum = PEN_NUM.KEY03_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY03
                    
                    Case 14
                    
                        .intLightNum = PEN_NUM.KEY04_LIGHT
                        .intShadowNum = PEN_NUM.KEY04_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY04
                    
                    Case 15
                    
                        .intLightNum = PEN_NUM.KEY05_LIGHT
                        .intShadowNum = PEN_NUM.KEY05_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY05
                    
                    Case 18
                    
                        .intLightNum = PEN_NUM.KEY06_LIGHT
                        .intShadowNum = PEN_NUM.KEY06_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY06
                    
                    Case 19
                    
                        .intLightNum = PEN_NUM.KEY07_LIGHT
                        .intShadowNum = PEN_NUM.KEY07_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY07
                    
                    Case 16
                    
                        .intLightNum = PEN_NUM.KEY08_LIGHT
                        .intShadowNum = PEN_NUM.KEY08_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY08
                    
                    Case 21
                    
                        .intLightNum = PEN_NUM.KEY11_LIGHT
                        .intShadowNum = PEN_NUM.KEY11_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY11
                    
                    Case 22
                    
                        .intLightNum = PEN_NUM.KEY12_LIGHT
                        .intShadowNum = PEN_NUM.KEY12_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY12
                    
                    Case 23
                    
                        .intLightNum = PEN_NUM.KEY13_LIGHT
                        .intShadowNum = PEN_NUM.KEY13_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY13
                    
                    Case 24
                    
                        .intLightNum = PEN_NUM.KEY14_LIGHT
                        .intShadowNum = PEN_NUM.KEY14_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY14
                    
                    Case 25
                    
                        .intLightNum = PEN_NUM.KEY15_LIGHT
                        .intShadowNum = PEN_NUM.KEY15_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY15
                    
                    Case 28
                    
                        .intLightNum = PEN_NUM.KEY16_LIGHT
                        .intShadowNum = PEN_NUM.KEY16_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY16
                    
                    Case 29
                    
                        .intLightNum = PEN_NUM.KEY17_LIGHT
                        .intShadowNum = PEN_NUM.KEY17_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY17
                    
                    Case 26
                    
                        .intLightNum = PEN_NUM.KEY18_LIGHT
                        .intShadowNum = PEN_NUM.KEY18_SHADOW
                        .intBrushNum = BRUSH_NUM.KEY18
                    
                    Case Is > 100 'BGM
                    
                        .intLightNum = PEN_NUM.BGM_LIGHT
                        .intShadowNum = PEN_NUM.BGM_SHADOW
                        .intBrushNum = BRUSH_NUM.BGM
                
                End Select
                
                If .intCh Then
                
                    .intWidth = GRID_WIDTH
                
                Else
                
                    .intWidth = SPACE_WIDTH
                
                End If
            
            End With
        
        Next i
        
        .lstWAV.ListIndex = 0
        .lstBMP.ListIndex = 0
        .lstBGA.ListIndex = 0
        .lstMeasureLen.ListIndex = 0
        
        .fraViewer.BorderStyle = 0
        .fraGrid.BorderStyle = 0
        .fraDispSize.BorderStyle = 0
        .fraResolution.BorderStyle = 0
        
        .fraHeader.BorderStyle = 0
        .fraMaterial.BorderStyle = 0
        
        For i = 0 To 2
        
            .fraTop(i).Top = 0
            .fraTop(i).Left = 0
            .fraTop(i).BorderStyle = 0
        
        Next i
        
        .fraTop(0).Visible = True
        .optChangeTop(0).value = True
        
        For i = 0 To 4
        
            .fraBottom(i).Top = 0
            .fraBottom(i).Left = 0
            .fraBottom(i).BorderStyle = 0
        
        Next i
        
        .fraBottom(0).Visible = True
        .optChangeBottom(0).value = True
        
        For i = 1 To 64
        
            Call .cboNumerator.AddItem(i, i - 1)
        
        Next i
        
        .cboNumerator.ListIndex = 3
        .cboDenominator.ListIndex = 0
        
        'g_Disp.intMaxMeasure = 31
        g_disp.intMaxMeasure = 0
        Call modDraw.lngChangeMaxMeasure(15)
        Call modDraw.ChangeResolution
        
        Call GetCmdLine
        
        g_BMS.blnSaveFlag = True
        
        .lstWAV.Selected(0) = True
        .lstBMP.Selected(0) = True
        .lstBGA.Selected(0) = True
        .lstMeasureLen.Selected(0) = True
        
        wp.Length = 44
        Call GetWindowPlacement(.hwnd, wp)
        
        wp.showCmd = strGet_ini("Main", "State", SW_SHOW, "bmse.ini")
        
        Call SetWindowPlacement(.hwnd, wp)
        
        g_blnIgnoreInput = False
    
    End With
    
    If strGet_ini("EasterEgg", "Tips", 0, "bmse.ini") Then
    
        With frmWindowTips
        
            .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
            .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
            
            Call .Show(vbModal, frmMain)
        
        End With
    
    End If

End Sub

Public Sub CleanUp(Optional ByVal lngErrNum As Long, Optional ByVal strErrDescription As String, Optional ByVal strErrProcedure As String)
On Error Resume Next

    Dim i       As Long
    
    #If MODE_DEBUG = False Then
    
        Call modMessage.UnSubClass(frmMain.hwnd)
    
    #End If
    
    #If MODE_SPEEDTEST Then
    
        Call timeEndPeriod(1)
    
    #End If
    
    Set g_InputLog = Nothing
    
    Call modEasterEgg.EndEffect
    
    Call SaveConfig
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
    Call lngDeleteFile(g_strAppDir & "___bmse_temp.bms")
    
    If lngErrNum <> 0 And strErrDescription <> "" And strErrProcedure <> "" Then
    
        If Len(g_BMS.strDir) = 0 Then g_BMS.strDir = g_strAppDir
        
        For i = 0 To 9999
        
            g_BMS.strFileName = "temp" & Format$(i, "0000") & ".bms"
            
            If i = 9999 Then
            
                Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)
            
            ElseIf Dir(g_BMS.strDir & g_BMS.strFileName) = vbNullString Then
            
                Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)
                Exit For
            
            End If
        
        Next i
        
        Call DebugOutput(lngErrNum, strErrDescription, strErrProcedure, True)
    
    End If
    
    End

End Sub

Public Sub DebugOutput(ByVal lngErrNum As Long, ByVal strErrDescription As String, ByVal strErrProcedure As String, Optional ByVal blnCleanUp As Boolean)

    Dim lngFFile    As Long
    Dim strError    As String
    
    lngFFile = FreeFile()
    
    Open g_strAppDir + "error.txt" For Append As #lngFFile
    
    Print #lngFFile, Date; Time; "ErrorNo." & lngErrNum & " " & strErrDescription & "@" & strErrProcedure & "/BMSE_" & App.Major & "." & App.Minor & "." & App.Revision
    
    Close #lngFFile
    
    strError = strError & "ErrorNo." & lngErrNum & " " & strErrDescription & "@" & strErrProcedure
    
    If blnCleanUp Then
    
        strError = g_Message(ERR_01) & vbCrLf & strError & vbCrLf
        strError = strError & g_Message(ERR_02) & vbCrLf
        strError = strError & g_BMS.strDir & g_BMS.strFileName
    
    End If
    
    Call frmMain.Show
    Call MsgBox(strError, vbCritical + vbOKOnly, g_strAppTitle)

End Sub

Public Function lngDeleteFile(ByVal FileName As String) As Long
On Error GoTo Err:

    Kill FileName
    
    Exit Function

Err:
    lngDeleteFile = 1
End Function

Public Function intSaveCheck() As Integer
On Error GoTo Err:

    Dim lngRet      As Long
    Dim retArray()  As String
    
    With frmMain
    
        If .cboPlayer.ListIndex + 1 <> g_BMS.intPlayerType Then g_BMS.blnSaveFlag = False
        If .txtGenre.Text <> g_BMS.strGenre Then g_BMS.blnSaveFlag = False
        If .txtTitle.Text <> g_BMS.strTitle Then g_BMS.blnSaveFlag = False
        If .txtArtist.Text <> g_BMS.strArtist Then g_BMS.blnSaveFlag = False
        If Val(.cboPlayLevel.Text) <> g_BMS.lngPlayLevel Then g_BMS.blnSaveFlag = False
        If Val(.txtBPM.Text) <> g_BMS.sngBPM Then g_BMS.blnSaveFlag = False
        
        If .cboPlayRank.ListIndex <> g_BMS.intPlayRank Then g_BMS.blnSaveFlag = False
        If Val(.txtTotal.Text) <> g_BMS.sngTotal Then g_BMS.blnSaveFlag = False
        If Val(.txtVolume.Text) <> g_BMS.intVolume Then g_BMS.blnSaveFlag = False
        If .txtStageFile.Text <> g_BMS.strStageFile Then g_BMS.blnSaveFlag = False
        'If .txtMissBMP.Text <> g_strBMP(0) Then g_BMS.blnSaveFlag = False
    
    End With
    
    If g_BMS.blnSaveFlag Then
    
        intSaveCheck = 0
        
        Exit Function
    
    End If
    
    Call frmMain.Show
    
    lngRet = MsgBox(g_Message(MSG_FILE_CHANGED), vbExclamation Or vbYesNoCancel, g_strAppTitle)
    
    Select Case lngRet
    
        Case vbYes
        
            If g_BMS.strDir <> "" And g_BMS.strFileName <> "" Then
            
                Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)
            
            Else
            
                With frmMain.dlgMain
                
                    .Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"
                        
                    .FileName = g_BMS.strFileName
                    
                    Call .ShowSave
                    
                    retArray = Split(.FileName, "\")
                    g_BMS.strDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
                    g_BMS.strFileName = retArray(UBound(retArray))
                    
                    Call CreateBMS(g_BMS.strDir & g_BMS.strFileName)
                    
                    Call RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
                    
                    .InitDir = g_BMS.strDir
                
                End With
            
            End If
        
        Case vbNo
        
            intSaveCheck = 0
        
        Case vbCancel
        
            intSaveCheck = 1
    
    End Select
    
    Exit Function

Err:

    intSaveCheck = 1

End Function

Public Sub RecentFilesRotation(ByVal strFilePath As String)

    Dim i       As Long
    Dim intRet  As Integer
    
    For i = 0 To UBound(g_strRecentFiles)
    
        If g_strRecentFiles(i) = strFilePath Then
        
            Call SubRotate(0, i, strFilePath)
            
            intRet = 1
            
            Exit For
        
        End If
    
    Next i
    
    If intRet = 0 Then Call SubRotate(0, UBound(g_strRecentFiles), strFilePath)
    
    frmMain.mnuLineRecent.Visible = True

End Sub

Private Sub SubRotate(ByVal intIndex As Integer, ByVal intEnd As Integer, ByVal strFilePath As String)

    If intIndex <> intEnd And g_strRecentFiles(intIndex) <> "" And intIndex <= UBound(g_strRecentFiles) Then
    
        Call SubRotate(intIndex + 1, intEnd, g_strRecentFiles(intIndex))
    
    End If
    
    g_strRecentFiles(intIndex) = strFilePath
    
    With frmMain.mnuRecentFiles(intIndex)
    
        .Caption = "&" & intIndex + 1 & ":" & strFilePath
        .Enabled = True
        .Visible = True
    
    End With
        
    With frmMain.tlbMenu.Buttons("Open").ButtonMenus(intIndex + 1)
    
        .Text = "&" & intIndex + 1 & ":" & strFilePath
        .Enabled = True
        .Visible = True
    
    End With

End Sub

Private Sub GetCmdLine()
On Error GoTo Err:

    Dim i               As Long
    Dim strRet          As String
    Dim strCmdArray()   As String
    Dim strArray()      As String
    Dim ReadLock        As Boolean
    Dim blnReadFlag     As Boolean
    
    strRet = Trim$(Command$())
    
    If strRet = "" Then Exit Sub
    
    ReDim strCmdArray(0)
    
    For i = 1 To Len(strRet)
    
        Select Case Asc(Mid$(strRet, i, 1))
        
            Case 32 'スペース
            
                If ReadLock = False Then
                
                    ReDim Preserve strCmdArray(UBound(strCmdArray) + 1)
                
                Else
                
                    strCmdArray(UBound(strCmdArray)) = strCmdArray(UBound(strCmdArray)) & " "
                
                End If
            
            Case 34 'ダブルクオーテーション
            
                ReadLock = Not ReadLock
            
            Case Else
            
                strCmdArray(UBound(strCmdArray)) = strCmdArray(UBound(strCmdArray)) & Mid$(strRet, i, 1)
        
        End Select
    
    Next i
    
    For i = 0 To UBound(strCmdArray)
    
        If strCmdArray(i) <> "" Then
        
            If InStr(1, strCmdArray(i), ":\") <> 0 And (UCase$(Right$(strCmdArray(i), 4)) = ".BMS" Or UCase$(Right$(strCmdArray(i), 4)) = ".BME" Or UCase$(Right$(strCmdArray(i), 4)) = ".BML" Or UCase$(Right$(strCmdArray(i), 4)) = ".PMS") Then
            
                If blnReadFlag Then
                
                    Call ShellExecute(0, "open", Chr$(34) & g_strAppDir & App.EXEName & Chr$(34), Chr$(34) & strCmdArray(i) & Chr$(34), "", SW_SHOWNORMAL)
                
                Else
                
                    strArray() = Split(strCmdArray(i), "\")
                    g_BMS.strFileName = Right$(strCmdArray(i), Len(strArray(UBound(strArray))))
                    g_BMS.strDir = Left$(strCmdArray(i), Len(strCmdArray(i)) - Len(strArray(UBound(strArray))))
                    frmMain.dlgMain.InitDir = g_BMS.strDir
                    blnReadFlag = True
                    
                    Call modInput.LoadBMS
                    Call RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
                
                End If
            
            End If
        
        End If
    
    Next i
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "GetCmdLine")
End Sub

Public Sub LoadThemeFile(ByVal strFileName As String)

    Dim strArray()      As String
    Dim i               As Long
    Dim j               As Long
    Dim lngRet          As Long
    
    strArray() = Split(strGet_ini("Main", "Background", "0,0,0", strFileName), ",")
    frmMain.picMain.BackColor = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "MeasureNum", "64,64,64", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.MEASURE_NUM) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "MeasureLine", "255,255,255", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.MEASURE_LINE) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "GridMain", "96,96,96", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.GRID_MAIN) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "GridSub", "192,192,192", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.GRID_SUB) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "VerticalMain", "255,255,255", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.VERTICAL_MAIN) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "VerticalSub", "128,128,128", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.VERTICAL_SUB) = RGB(strArray(0), strArray(1), strArray(2))
    
    strArray() = Split(strGet_ini("Main", "Info", "0,255,0", strFileName), ",")
    g_lngSystemColor(COLOR_NUM.INFO) = RGB(strArray(0), strArray(1), strArray(2))
    
    
    For i = 0 To BRUSH_NUM.Max - 1
    
        Select Case i
        
            Case BRUSH_NUM.BGM
            
                strArray() = Split(strGet_ini("BGM", "Background", "48,0,0", strFileName), ",")
                lngRet = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("BGM", "Text", "B01,B02,B03,B04,B05,B06,B07,B08,B09,B10,B11,B12,B13,B14,B15,B16,B17,B18,B19,B20,B21,B22,B23,B24,B25,B26,B27,B28,B29,B30,B31,B32", strFileName), ",")
                
                For j = 0 To 31
                
                    g_VGrid(GRID.NUM_BGM + j).strText = strArray(j)
                    g_VGrid(GRID.NUM_BGM + j).lngBackColor = lngRet
                
                Next j
                
                strArray() = Split(strGet_ini("BGM", "ObjectLight", "255,0,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BGM_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BGM", "ObjectShadow", "96,0,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BGM_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BGM", "ObjectColor", "128,0,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.BGM) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.BPM
            
                strArray() = Split(strGet_ini("BPM", "Text", "BPM,STOP", strFileName), ",")
                g_VGrid(GRID.NUM_BPM).strText = strArray(0)
                g_VGrid(GRID.NUM_STOP).strText = strArray(1)
                
                strArray() = Split(strGet_ini("BPM", "Background", "48,48,48", strFileName), ",")
                lngRet = RGB(strArray(0), strArray(1), strArray(2))
                g_VGrid(GRID.NUM_BPM).lngBackColor = lngRet
                g_VGrid(GRID.NUM_STOP).lngBackColor = lngRet
                
                strArray() = Split(strGet_ini("BPM", "ObjectLight", "192,192,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BPM_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BPM", "ObjectShadow", "128,128,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BPM_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BPM", "ObjectColor", "160,160,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.BPM) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.BGA
            
                strArray() = Split(strGet_ini("BGA", "Text", "BGA,LAYER,POOR", strFileName), ",")
                g_VGrid(GRID.NUM_BGA).strText = strArray(0)
                g_VGrid(GRID.NUM_LAYER).strText = strArray(1)
                g_VGrid(GRID.NUM_POOR).strText = strArray(2)
                
                strArray() = Split(strGet_ini("BGA", "Background", "0,24,0", strFileName), ",")
                lngRet = RGB(strArray(0), strArray(1), strArray(2))
                g_VGrid(GRID.NUM_BGA).lngBackColor = lngRet
                g_VGrid(GRID.NUM_LAYER).lngBackColor = lngRet
                g_VGrid(GRID.NUM_POOR).lngBackColor = lngRet
                
                strArray() = Split(strGet_ini("BGA", "ObjectLight", "0,255,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BGA_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BGA", "ObjectShadow", "0,96,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.BGA_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("BGA", "ObjectColor", "0,128,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.BGA) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.KEY01, BRUSH_NUM.KEY03, BRUSH_NUM.KEY05, BRUSH_NUM.KEY07
            
                lngRet = (i - BRUSH_NUM.KEY01) + 1
                
                g_VGrid(GRID.NUM_1P_1KEY + lngRet - 1).strText = strGet_ini("KEY_1P_0" & lngRet, "Text", lngRet, strFileName)
                
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "Background", "32,32,32", strFileName), ",")
                g_VGrid(GRID.NUM_1P_1KEY + lngRet - 1).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectLight", "192,192,192", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY01_LIGHT + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY01_LIGHT + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectShadow", "96,96,96", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY01_SHADOW + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY01_SHADOW + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectColor", "128,128,128", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY01 + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY01 + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
            
            Case BRUSH_NUM.KEY02, BRUSH_NUM.KEY04, BRUSH_NUM.KEY06
            
                lngRet = (i - BRUSH_NUM.KEY01) + 1
                g_VGrid(GRID.NUM_1P_1KEY + lngRet - 1).strText = strGet_ini("KEY_1P_0" & lngRet, "Text", lngRet, strFileName)
                
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "Background", "0,0,40", strFileName), ",")
                g_VGrid(GRID.NUM_1P_1KEY + lngRet - 1).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectLight", "96,96,255", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY01_LIGHT + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY01_LIGHT + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectShadow", "0,0,128", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY01_SHADOW + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY01_SHADOW + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_0" & lngRet, "ObjectColor", "0,0,255", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY01 + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY01 + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
            
            Case BRUSH_NUM.KEY08
            
                g_VGrid(GRID.NUM_1P_SC_L).strText = strGet_ini("KEY_1P_SC", "Text", "SC", strFileName)
                g_VGrid(GRID.NUM_1P_SC_R).strText = strGet_ini("KEY_1P_SC", "Text", "SC", strFileName)
                
                strArray() = Split(strGet_ini("KEY_1P_SC", "Background", "48,0,0", strFileName), ",")
                g_VGrid(GRID.NUM_1P_SC_L).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                g_VGrid(GRID.NUM_1P_SC_R).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_1P_SC", "ObjectLight", "255,96,96", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY08_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY08_LIGHT) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_SC", "ObjectShadow", "128,0,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY08_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY08_SHADOW) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_1P_SC", "ObjectColor", "255,0,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY08) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY08) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
            
            Case BRUSH_NUM.KEY11, BRUSH_NUM.KEY13, BRUSH_NUM.KEY15, BRUSH_NUM.KEY17
            
                lngRet = (i - BRUSH_NUM.KEY11) + 1
                g_VGrid(GRID.NUM_2P_1KEY + lngRet - 1).strText = strGet_ini("KEY_2P_0" & lngRet, "Text", lngRet, strFileName)
                
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "Background", "32,32,32", strFileName), ",")
                g_VGrid(GRID.NUM_2P_1KEY + lngRet - 1).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectLight", "192,192,192", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY11_LIGHT + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY11_LIGHT + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectShadow", "96,96,96", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY11_SHADOW + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY11_SHADOW + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectColor", "128,128,128", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY11 + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY11 + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                
                If i = BRUSH_NUM.KEY11 Then
                
                    g_VGrid(GRID.NUM_FOOTPEDAL).strText = strGet_ini("KEY_2P_01", "Text", lngRet, strFileName)
                    strArray() = Split(strGet_ini("KEY_2P_01", "Background", "32,32,32", strFileName), ",")
                    g_VGrid(GRID.NUM_FOOTPEDAL).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                    strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectLight", "192,192,192", strFileName), ",")
                
                End If
            
            Case BRUSH_NUM.KEY12, BRUSH_NUM.KEY14, BRUSH_NUM.KEY16
            
                lngRet = (i - BRUSH_NUM.KEY11) + 1
                g_VGrid(GRID.NUM_2P_1KEY + lngRet - 1).strText = strGet_ini("KEY_2P_0" & lngRet, "Text", lngRet, strFileName)
                
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "Background", "0,0,40", strFileName), ",")
                g_VGrid(GRID.NUM_2P_1KEY + lngRet - 1).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectLight", "96,96,255", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY11_LIGHT + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY11_LIGHT + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectShadow", "0,0,128", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY11_SHADOW + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY11_SHADOW + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_0" & lngRet, "ObjectColor", "0,0,255", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY11 + lngRet - 1) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY11 + lngRet - 1) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
            
            Case BRUSH_NUM.KEY18
            
                g_VGrid(GRID.NUM_2P_SC_L).strText = strGet_ini("KEY_2P_SC", "Text", "SC", strFileName)
                g_VGrid(GRID.NUM_2P_SC_R).strText = strGet_ini("KEY_2P_SC", "Text", "SC", strFileName)
                
                strArray() = Split(strGet_ini("KEY_2P_SC", "Background", "48,0,0", strFileName), ",")
                g_VGrid(GRID.NUM_2P_SC_L).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                g_VGrid(GRID.NUM_2P_SC_R).lngBackColor = RGB(strArray(0), strArray(1), strArray(2))
                
                strArray() = Split(strGet_ini("KEY_2P_SC", "ObjectLight", "255,96,96", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY18_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY18_LIGHT) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_SC", "ObjectShadow", "128,0,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.KEY18_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngPenColor(PEN_NUM.INV_KEY18_SHADOW) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
                strArray() = Split(strGet_ini("KEY_2P_SC", "ObjectColor", "255,0,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.KEY18) = RGB(strArray(0), strArray(1), strArray(2))
                g_lngBrushColor(BRUSH_NUM.INV_KEY18) = RGB(strArray(0) \ 2, strArray(1) \ 2, strArray(2) \ 2)
            
            Case BRUSH_NUM.LONGNOTE
            
                strArray() = Split(strGet_ini("KEY_LONGNOTE", "ObjectLight", "0,128,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.LONGNOTE_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("KEY_LONGNOTE", "ObjectShadow", "0,32,0", strFileName), ",")
                g_lngPenColor(PEN_NUM.LONGNOTE_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("KEY_LONGNOTE", "ObjectColor", "0,64,0", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.LONGNOTE) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.SELECT_OBJ
            
                strArray() = Split(strGet_ini("SELECT", "ObjectLight", "255,255,255", strFileName), ",")
                g_lngPenColor(PEN_NUM.SELECT_OBJ_LIGHT) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("SELECT", "ObjectShadow", "128,128,128", strFileName), ",")
                g_lngPenColor(PEN_NUM.SELECT_OBJ_SHADOW) = RGB(strArray(0), strArray(1), strArray(2))
                strArray() = Split(strGet_ini("SELECT", "ObjectColor", "0,255,255", strFileName), ",")
                g_lngBrushColor(BRUSH_NUM.SELECT_OBJ) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.EDIT_FRAME
            
                strArray() = Split(strGet_ini("SELECT", "EditFrame", "255,255,255", strFileName), ",")
                g_lngPenColor(PEN_NUM.EDIT_FRAME) = RGB(strArray(0), strArray(1), strArray(2))
            
            Case BRUSH_NUM.DELETE_FRAME
            
                strArray() = Split(strGet_ini("SELECT", "DeleteFrame", "255,255,255", strFileName), ",")
                g_lngPenColor(PEN_NUM.DELETE_FRAME) = RGB(strArray(0), strArray(1), strArray(2))
        
        End Select
    
    Next i

End Sub

Public Sub LoadLanguageFile(ByVal strFileName As String)

    g_strStatusBar(1) = strGet_ini("StatusBar", "CH_01", "BGM", strFileName)
    g_strStatusBar(4) = strGet_ini("StatusBar", "CH_04", "BGA", strFileName)
    g_strStatusBar(6) = strGet_ini("StatusBar", "CH_06", "BGA Poor", strFileName)
    g_strStatusBar(7) = strGet_ini("StatusBar", "CH_07", "BGA Layer", strFileName)
    g_strStatusBar(8) = strGet_ini("StatusBar", "CH_08", "BPM Change", strFileName)
    g_strStatusBar(9) = strGet_ini("StatusBar", "CH_09", "Stop Sequence", strFileName)
    g_strStatusBar(11) = strGet_ini("StatusBar", "CH_KEY_1P", "1P Key", strFileName)
    g_strStatusBar(12) = strGet_ini("StatusBar", "CH_KEY_2P", "2P Key", strFileName)
    g_strStatusBar(13) = strGet_ini("StatusBar", "CH_SCRATCH_1P", "1P Scratch", strFileName)
    g_strStatusBar(14) = strGet_ini("StatusBar", "CH_SCRATCH_2P", "2P Scratch", strFileName)
    g_strStatusBar(15) = strGet_ini("StatusBar", "CH_INVISIBLE", "(Invisible)", strFileName)
    g_strStatusBar(16) = strGet_ini("StatusBar", "CH_LONGNOTE", "(LongNote)", strFileName)
    g_strStatusBar(20) = strGet_ini("StatusBar", "MODE_EDIT", "Edit Mode", strFileName)
    g_strStatusBar(21) = strGet_ini("StatusBar", "MODE_WRITE", "Write Mode", strFileName)
    g_strStatusBar(22) = strGet_ini("StatusBar", "MODE_DELETE", "Delete Mode", strFileName)
    g_strStatusBar(23) = strGet_ini("StatusBar", "MEASURE", "Measure", strFileName)
    
    With frmMain
    
        .mnuFile.Caption = strGet_ini("Menu", "FILE", "&File", strFileName)
        .mnuFileNew.Caption = strGet_ini("Menu", "FILE_NEW", "&New", strFileName)
        .mnuFileOpen.Caption = strGet_ini("Menu", "FILE_OPEN", "&Open", strFileName)
        .mnuFileSave.Caption = strGet_ini("Menu", "FILE_SAVE", "&Save", strFileName)
        .mnuFileSaveAs.Caption = strGet_ini("Menu", "FILE_SAVE_AS", "Save &As", strFileName)
        .mnuFileOpenDirectory.Caption = strGet_ini("Menu", "FILE_OPEN_DIRECTORY", "Open &Directory", strFileName)
        '.mnuFileDeleteUnusedFile.Caption = strGet_ini("Menu", "FILE_DELETE_UNUSED_FILE", "&Delete Unused File(s)", strFileName)
        '.mnuFileNameConvert.Caption = strGet_ini("Menu", "FILE_CONVERT_FILENAME", "&Convert Filenames to [01-ZZ]", strFileName)
        '.mnuFileListAlign.Caption = strGet_ini("Menu", "FILE_ALIGN_LIST", "Rewrite &List into old format [01-FF]", strFileName)
        .mnuFileConvertWizard.Caption = strGet_ini("Menu", "FILE_CONVERT_WIZARD", "Show &Conversion Wizard", strFileName)
        .mnuFileExit.Caption = strGet_ini("Menu", "FILE_EXIT", "&Exit", strFileName)
        
        .mnuEdit.Caption = strGet_ini("Menu", "EDIT", "&Edit", strFileName)
        .mnuEditUndo.Caption = strGet_ini("Menu", "EDIT_UNDO", "&Undo", strFileName)
        .mnuEditRedo.Caption = strGet_ini("Menu", "EDIT_REDO", "&Redo", strFileName)
        .mnuEditCut.Caption = strGet_ini("Menu", "EDIT_CUT", "Cu&t", strFileName)
        .mnuEditCopy.Caption = strGet_ini("Menu", "EDIT_COPY", "&Copy", strFileName)
        .mnuEditPaste.Caption = strGet_ini("Menu", "EDIT_PASTE", "&Paste", strFileName)
        .mnuEditDelete.Caption = strGet_ini("Menu", "EDIT_DELETE", "&Delete", strFileName)
        .mnuEditSelectAll.Caption = strGet_ini("Menu", "EDIT_SELECT_ALL", "&Find/Replace/Delete", strFileName)
        .mnuEditFind.Caption = strGet_ini("Menu", "EDIT_FIND", "&Select All", strFileName)
        .mnuEditMode(0).Caption = strGet_ini("Menu", "EDIT_MODE_EDIT", "Edit &Mode", strFileName)
        .mnuEditMode(1).Caption = strGet_ini("Menu", "EDIT_MODE_WRITE", "Write &Mode", strFileName)
        .mnuEditMode(2).Caption = strGet_ini("Menu", "EDIT_MODE_DELETE", "Delete &Mode", strFileName)
        
        .mnuView.Caption = strGet_ini("Menu", "VIEW", "&View", strFileName)
        .mnuViewToolBar.Caption = strGet_ini("Menu", "VIEW_TOOL_BAR", "&Tool Bar", strFileName)
        .mnuViewDirectInput.Caption = strGet_ini("Menu", "VIEW_DIRECT_INPUT", "&Direct Input", strFileName)
        .mnuViewStatusBar.Caption = strGet_ini("Menu", "VIEW_STATUS_BAR", "&Status Bar", strFileName)
        
        .mnuOptions.Caption = strGet_ini("Menu", "OPTIONS", "&Options", strFileName)
        .mnuOptionsActiveIgnore.Caption = strGet_ini("Menu", "OPTIONS_IGNORE_ACTIVE", "&Control Unavailable When Active", strFileName)
        .mnuOptionsFileNameOnly.Caption = strGet_ini("Menu", "OPTIONS_FILE_NAME_ONLY", "Display &File Name Only", strFileName)
        .mnuOptionsVertical.Caption = strGet_ini("Menu", "OPTIONS_VERTICAL", "&Vertical Grid Info", strFileName)
        .mnuOptionsLaneBG.Caption = strGet_ini("Menu", "OPTIONS_LANE_BG", "&Background Color", strFileName)
        '.mnuOptionsSelectPreview.Caption = strGet_ini("Menu", "OPTIONS_SINGLE_SELECT_SOUND", "&Sound Upon Object Selection", strFileName)
        .mnuOptionsSelectPreview.Caption = strGet_ini("Menu", "OPTIONS_SINGLE_SELECT_PREVIEW", "&Preview Upon Object Selection", strFileName)
        .mnuOptionsObjectFileName.Caption = strGet_ini("Menu", "OPTIONS_OBJECT_FILE_NAME", "Show &Objects' File Names", strFileName)
        .mnuOptionsMoveOnGrid.Caption = strGet_ini("Menu", "OPTIONS_MOVE_ON_GRID", "Restrict Objects' &Movement Onto Grid", strFileName)
        .mnuOptionsNumFF.Caption = strGet_ini("Menu", "OPTIONS_USE_OLD_FORMAT", "&Use Old Format (01-FF)", strFileName)
        .mnuOptionsRightClickDelete.Caption = strGet_ini("Menu", "OPTIONS_RIGHT_CLICK_DELETE", "&Right Click To Delete Objects", strFileName)
        
        .mnuTools.Caption = strGet_ini("Menu", "TOOLS", "&Tools", strFileName)
        .mnuToolsPlayAll.Caption = strGet_ini("Menu", "TOOLS_PLAY_FIRST", "Play &All", strFileName)
        .mnuToolsPlay.Caption = strGet_ini("Menu", "TOOLS_PLAY", "&Play From Current Position", strFileName)
        .mnuToolsPlayStop.Caption = strGet_ini("Menu", "TOOLS_STOP", "&Stop", strFileName)
        .mnuToolsSetting.Caption = strGet_ini("Menu", "TOOLS_SETTING", "&Viewer Setting", strFileName)
        
        .mnuHelp.Caption = strGet_ini("Menu", "HELP", "&Help", strFileName)
        .mnuHelpOpen.Caption = strGet_ini("Menu", "HELP_OPEN", "&Help", strFileName)
        .mnuHelpWeb.Caption = strGet_ini("Menu", "HELP_WEB", "Open &Website", strFileName)
        .mnuHelpAbout.Caption = strGet_ini("Menu", "HELP_ABOUT", "&About BMSE", strFileName)
        
        .mnuContext.Visible = False
        .mnuContextInsertMeasure.Caption = strGet_ini("Menu", "CONTEXT_MEASURE_INSERT", "&Insert Measure", strFileName)
        .mnuContextDeleteMeasure.Caption = strGet_ini("Menu", "CONTEXT_MEASURE_DELETE", "Delete &Measure", strFileName)
        .mnuContextEditCut.Caption = strGet_ini("Menu", "EDIT_CUT", "Cu&t", strFileName)
        .mnuContextEditCopy.Caption = strGet_ini("Menu", "EDIT_COPY", "&Copy", strFileName)
        .mnuContextEditPaste.Caption = strGet_ini("Menu", "EDIT_PASTE", "&Paste", strFileName)
        .mnuContextEditDelete.Caption = strGet_ini("Menu", "EDIT_DELETE", "&Delete", strFileName)
        
        .mnuContextList.Visible = False
        .mnuContextListLoad.Caption = strGet_ini("Menu", "CONTEXT_LIST_LOAD", "&Load", strFileName)
        .mnuContextListDelete.Caption = strGet_ini("Menu", "CONTEXT_LIST_DELETE", "&Delete", strFileName)
        .mnuContextListRename.Caption = strGet_ini("Menu", "CONTEXT_LIST_RENAME", "&Rename", strFileName)
        
        .optChangeTop(0).Caption = strGet_ini("Header", "TAB_BASIC", "Basic", strFileName)
        .optChangeTop(1).Caption = strGet_ini("Header", "TAB_EXPAND", "Expand", strFileName)
        .optChangeTop(2).Caption = strGet_ini("Header", "TAB_CONFIG", "Config", strFileName)
        
        .lblPlayMode.Caption = strGet_ini("Header", "BASIC_PLAYER", "#PLAYER", strFileName)
        .cboPlayer.List(0) = strGet_ini("Header", "BASIC_PLAYER_1P", "1 Player", strFileName)
        .cboPlayer.List(1) = strGet_ini("Header", "BASIC_PLAYER_2P", "2 Player", strFileName)
        .cboPlayer.List(2) = strGet_ini("Header", "BASIC_PLAYER_DP", "Double Play", strFileName)
        .cboPlayer.List(3) = strGet_ini("Header", "BASIC_PLAYER_PMS", "9 Keys (PMS)", strFileName)
        .cboPlayer.List(4) = strGet_ini("Header", "BASIC_PLAYER_OCT", "13 Keys (Oct)", strFileName)
        .lblGenre.Caption = strGet_ini("Header", "BASIC_GENRE", "#GENRE", strFileName)
        .lblTitle.Caption = strGet_ini("Header", "BASIC_TITLE", "#TITLE", strFileName)
        .lblArtist.Caption = strGet_ini("Header", "BASIC_ARTIST", "#ARTIST", strFileName)
        .lblPlayLevel.Caption = strGet_ini("Header", "BASIC_PLAYLEVEL", "#PLAYLEVEL", strFileName)
        .lblBPM.Caption = strGet_ini("Header", "BASIC_BPM", "#BPM", strFileName)
        
        .lblPlayRank.Caption = strGet_ini("Header", "EXPAND_RANK", "#RANK", strFileName)
        .cboPlayRank.List(0) = strGet_ini("Header", "EXPAND_RANK_VERY_HARD", "Very Hard", strFileName)
        .cboPlayRank.List(1) = strGet_ini("Header", "EXPAND_RANK_HARD", "Hard", strFileName)
        .cboPlayRank.List(2) = strGet_ini("Header", "EXPAND_RANK_NORMAL", "Normal", strFileName)
        .cboPlayRank.List(3) = strGet_ini("Header", "EXPAND_RANK_EASY", "Easy", strFileName)
        .lblTotal.Caption = strGet_ini("Header", "EXPAND_TOTAL", "#TOTAL", strFileName)
        .lblVolume.Caption = strGet_ini("Header", "EXPAND_VOLWAV", "#VOLWAV", strFileName)
        .lblStageFile.Caption = strGet_ini("Header", "EXPAND_STAGEFILE", "#STAGEFILE", strFileName)
        .lblMissBMP.Caption = strGet_ini("Header", "EXPAND_BPM_MISS", "#BMP00", strFileName)
        .cmdLoadMissBMP.Caption = strGet_ini("Header", "EXPAND_SET_FILE", "...", strFileName)
        .cmdLoadStageFile.Caption = strGet_ini("Header", "EXPAND_SET_FILE", "...", strFileName)
        
        .lblDispFrame.Caption = strGet_ini("Header", "CONFIG_KEY_FRAME", "Key Frame", strFileName)
        .cboDispFrame.List(0) = strGet_ini("Header", "CONFIG_KEY_HALF", "Half", strFileName)
        .cboDispFrame.List(1) = strGet_ini("Header", "CONFIG_KEY_SEPARATE", "Separate", strFileName)
        .lblDispKey.Caption = strGet_ini("Header", "CONFIG_KEY_POSITION", "Key Position", strFileName)
        .cboDispKey.List(0) = strGet_ini("Header", "CONFIG_KEY_5KEYS", "5Keys/10Keys", strFileName)
        .cboDispKey.List(1) = strGet_ini("Header", "CONFIG_KEY_7KEYS", "7Keys/14Keys", strFileName)
        .lblDispSC1P.Caption = strGet_ini("Header", "CONFIG_SCRATCH_1P", "Scratch 1P", strFileName)
        .cboDispSC1P.List(0) = strGet_ini("Header", "CONFIG_SCRATCH_LEFT", "L", strFileName)
        .cboDispSC1P.List(1) = strGet_ini("Header", "CONFIG_SCRATCH_RIGHT", "R", strFileName)
        .lblDispSC2P.Caption = strGet_ini("Header", "CONFIG_SCRATCH_2P", "2P", strFileName)
        .cboDispSC2P.List(0) = strGet_ini("Header", "CONFIG_SCRATCH_LEFT", "L", strFileName)
        .cboDispSC2P.List(1) = strGet_ini("Header", "CONFIG_SCRATCH_RIGHT", "R", strFileName)
        
        .optChangeBottom(0).Caption = strGet_ini("Material", "TAB_WAV", "#WAV", strFileName)
        .optChangeBottom(1).Caption = strGet_ini("Material", "TAB_BMP", "#BMP", strFileName)
        .optChangeBottom(2).Caption = strGet_ini("Material", "TAB_BGA", "#BGA", strFileName)
        .optChangeBottom(3).Caption = strGet_ini("Material", "TAB_BEAT", "Beat", strFileName)
        .optChangeBottom(4).Caption = strGet_ini("Material", "TAB_EXPAND", "Expand", strFileName)
        
        .cmdSoundStop.Caption = strGet_ini("Material", "MATERIAL_STOP", "Stop", strFileName)
        .cmdSoundExcUp.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
        .cmdSoundExcDown.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
        .cmdSoundDelete.Caption = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
        .cmdSoundLoad.Caption = strGet_ini("Material", "MATERIAL_SET_FILE", "...", strFileName)
        
        .cmdBMPPreview.Caption = strGet_ini("Material", "MATERIAL_PREVIEW", "Preview", strFileName)
        .cmdBMPExcUp.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
        .cmdBMPExcDown.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
        .cmdBMPDelete.Caption = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
        .cmdBMPLoad.Caption = strGet_ini("Material", "MATERIAL_SET_FILE", "...", strFileName)
        
        .cmdBGAPreview.Caption = strGet_ini("Material", "MATERIAL_PREVIEW", "Preview", strFileName)
        .cmdBGAExcUp.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_UP", "<", strFileName)
        .cmdBGAExcDown.Caption = strGet_ini("Material", "MATERIAL_EXCHANGE_DOWN", ">", strFileName)
        .cmdBGADelete.Caption = strGet_ini("Material", "MATERIAL_DELETE", "Del", strFileName)
        .cmdBGASet.Caption = strGet_ini("Material", "MATERIAL_INPUT", "Input", strFileName)
        
        .cmdMeasureSelectAll.Caption = strGet_ini("Material", "MATERIAL_SELECT_ALL", "All", strFileName)
        
        .cmdInputMeasureLen.Caption = strGet_ini("Material", "MATERIAL_INPUT", "Input", strFileName)
        
        .lblGridMain.Caption = strGet_ini("ToolBar", "GRID_MAIN", "Grid", strFileName)
        .lblGridSub.Caption = strGet_ini("ToolBar", "GRID_SUB", "Sub", strFileName)
        .lblDispHeight.Caption = strGet_ini("ToolBar", "DISP_HEIGHT", "Height", strFileName)
        .lblDispWidth.Caption = strGet_ini("ToolBar", "DISP_WIDTH", "Width", strFileName)
        .lblVScroll.Caption = strGet_ini("ToolBar", "VSCROLL", "VScroll", strFileName)
        
        If .tlbMenu.Buttons("Edit").value = tbrPressed Then
        
            .staMain.Panels("Mode").Text = g_strStatusBar(20)
        
        ElseIf .tlbMenu.Buttons("Write").value = tbrPressed Then
        
            .staMain.Panels("Mode").Text = g_strStatusBar(21)
        
        Else
        
            .staMain.Panels("Mode").Text = g_strStatusBar(22)
        
        End If
    
    End With
    
    With frmMain.tlbMenu
    
        .Buttons("New").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_NEW", "New", strFileName)
        .Buttons("New").Description = strGet_ini("ToolBar", "TOOLTIP_NEW", "New", strFileName)
        .Buttons("Open").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_OPEN", "Open", strFileName)
        .Buttons("Open").Description = strGet_ini("ToolBar", "TOOLTIP_OPEN", "Open", strFileName)
        .Buttons("Reload").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_RELOAD", "Reload", strFileName)
        .Buttons("Reload").Description = strGet_ini("ToolBar", "TOOLTIP_RELOAD", "Reload", strFileName)
        .Buttons("Save").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_SAVE", "Save", strFileName)
        .Buttons("Save").Description = strGet_ini("ToolBar", "TOOLTIP_SAVE", "Save", strFileName)
        .Buttons("SaveAs").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_SAVE_AS", "Save As", strFileName)
        .Buttons("SaveAs").Description = strGet_ini("ToolBar", "TOOLTIP_SAVE_AS", "Save As", strFileName)
        
        .Buttons("Edit").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_EDIT", "Edit Mode", strFileName)
        .Buttons("Edit").Description = strGet_ini("ToolBar", "TOOLTIP_MODE_EDIT", "Edit Mode", strFileName)
        .Buttons("Write").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_WRITE", "Write Mode", strFileName)
        .Buttons("Write").Description = strGet_ini("ToolBar", "TOOLTIP_MODE_WRITE", "Write Mode", strFileName)
        .Buttons("Delete").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_MODE_DELETE", "Delete Mode", strFileName)
        .Buttons("Delete").Description = strGet_ini("ToolBar", "TOOLTIP_MODE_DELETE", "Delete Mode", strFileName)
        
        .Buttons("PlayAll").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_PLAY_FIRST", "Play All", strFileName)
        .Buttons("PlayAll").Description = strGet_ini("ToolBar", "TOOLTIP_PLAY_FIRST", "Play All", strFileName)
        .Buttons("Play").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_PLAY", "Play From Current Position", strFileName)
        .Buttons("Play").Description = strGet_ini("ToolBar", "TOOLTIP_PLAY", "Play From Current Position", strFileName)
        .Buttons("Stop").ToolTipText = strGet_ini("ToolBar", "TOOLTIP_STOP", "Stop", strFileName)
        .Buttons("Stop").Description = strGet_ini("ToolBar", "TOOLTIP_STOP", "Stop", strFileName)
    
    End With
    
    With frmWindowFind
    
        .Caption = strGet_ini("Find", "TITLE", "Find/Delete/Replace", strFileName)
        
        .fraSearchObject.Caption = strGet_ini("Find", "FRAME_SEARCH", "Range", strFileName)
        .fraSearchMeasure.Caption = strGet_ini("Find", "FRAME_MEASURE", "Range of measure", strFileName)
        .fraSearchNum.Caption = strGet_ini("Find", "FRAME_OBJ_NUM", "Range of object number", strFileName)
        .fraSearchGrid.Caption = strGet_ini("Find", "FRAME_GRID", "Lane", strFileName)
        .fraProcess.Caption = strGet_ini("Find", "FRAME_PROCESS", "Method", strFileName)
        
        .optSearchAll.Caption = strGet_ini("Find", "OPT_OBJ_ALL", "All object", strFileName)
        .optSearchSelect.Caption = strGet_ini("Find", "OPT_OBJ_SELECT", "Selected object", strFileName)
        .optProcessSelect.Caption = strGet_ini("Find", "OPT_PROCESS_SELECT", "Select", strFileName)
        .optProcessDelete.Caption = strGet_ini("Find", "OPT_PROCESS_DELETE", "Delete", strFileName)
        .optProcessReplace.Caption = strGet_ini("Find", "OPT_PROCESS_REPLACE", "Replace to", strFileName)
        
        .cmdInvert.Caption = strGet_ini("Find", "CMD_INVERT", "Invert", strFileName)
        .cmdReset.Caption = strGet_ini("Find", "CMD_RESET", "Reset", strFileName)
        .cmdSelect.Caption = strGet_ini("Find", "CMD_SELECT", "Select All", strFileName)
        .cmdClose.Caption = strGet_ini("Find", "CMD_CLOSE", "Close", strFileName)
        .cmdDecide.Caption = strGet_ini("Find", "CMD_DECIDE", "Run", strFileName)
        
        .lblNotice.Caption = strGet_ini("Find", "LBL_NOTICE", "This item doesn't influence BPM/STOP object", strFileName)
        .lblMeasure.Caption = strGet_ini("Find", "LBL_DASH", "to", strFileName)
        .lblNum.Caption = strGet_ini("Find", "LBL_DASH", "to", strFileName)
    
    End With
    
    With frmWindowInput
    
        .Caption = strGet_ini("Input", "TITLE", "Input Form", strFileName)
    
    End With
    
    With frmWindowViewer
    
        .Caption = strGet_ini("Viewer", "TITLE", "Viewer Setting", strFileName)
        
        .cmdViewerPath.Caption = strGet_ini("Viewer", "CMD_SET", "...", strFileName)
        .cmdAdd.Caption = strGet_ini("Viewer", "CMD_ADD", "Add", strFileName)
        .cmdDelete.Caption = strGet_ini("Viewer", "CMD_DELETE", "Delete", strFileName)
        .cmdOK.Caption = strGet_ini("Viewer", "CMD_OK", "OK", strFileName)
        .cmdCancel.Caption = strGet_ini("Viewer", "CMD_CANCEL", "Cancel", strFileName)
        
        .lblViewerName.Caption = strGet_ini("Viewer", "LBL_APP_NAME", "Player name", strFileName)
        .lblViewerPath.Caption = strGet_ini("Viewer", "LBL_APP_PATH", "Path", strFileName)
        .lblPlayAll.Caption = strGet_ini("Viewer", "LBL_ARG_PLAY_ALL", "Argument of ""Play All""", strFileName)
        .lblPlay.Caption = strGet_ini("Viewer", "LBL_ARG_PLAY", "Argument of ""Play""", strFileName)
        .lblStop.Caption = strGet_ini("Viewer", "LBL_ARG_STOP", "Argument of ""Stop""", strFileName)
        .lblNotice.Caption = Replace$(strGet_ini("Viewer", "LBL_ARG_INFO", "Syntax reference:\n<filename> File name\n<measure> Current measure", strFileName), "\n", vbCrLf)
    
    End With
    
    With frmWindowConvert
    
        .Caption = strGet_ini("Convert", "TITLE", "Conversion Wizard", strFileName)
        
        .chkDeleteUnusedFile.Caption = strGet_ini("Convert", "CHK_DELETE_LIST", "Clear unused definition from a list", strFileName)
        
        .chkDeleteFile.Caption = strGet_ini("Convert", "CHK_DELETE_FILE", "Delete unused files in this BMS folder (*)", strFileName)
        .lblExtension.Caption = strGet_ini("Convert", "LBL_EXTENSION", "Search extensions:", strFileName)
        .chkFileRecycle.Caption = strGet_ini("Convert", "CHK_RECYCLE", "Delete soon with no through recycled", strFileName)
        
        .chkListAlign.Caption = strGet_ini("Convert", "CHK_ALIGN_LIST", "Sort definition list", strFileName)
        .chkUseOldFormat.Caption = strGet_ini("Convert", "CHK_USE_OLD_FORMAT", "Use old Format [01 - FF] if possible", strFileName)
        .chkSortByName.Caption = strGet_ini("Convert", "CHK_SORT_BY_NAME", "Sorting by filename", strFileName)
        
        .chkFileNameConvert.Caption = strGet_ini("Convert", "CHK_CONVERT_FILENAME", "Change filename to list number [01 - ZZ] (*)", strFileName)
        
        .lblNotice.Caption = strGet_ini("Convert", "LBL_NOTICE", "(*) Cannot undo this command", strFileName)
        
        .cmdDecide.Caption = strGet_ini("Convert", "CMD_DECIDE", "Run", strFileName)
        .cmdCancel.Caption = strGet_ini("Convert", "CMD_CANCEL", "Cancel", strFileName)
    
    End With
    
    g_Message(ERR_01) = Replace$(strGet_ini("Message", "ERROR_MESSAGE_01", "The unexpected error occurred. Program will shut down.\nRefer to the ""error.txt"" for the details of an error.", strFileName), "\n", vbCrLf)
    g_Message(ERR_02) = Replace$(strGet_ini("Message", "ERROR_MESSAGE_02", "Temporary file is saved to...", strFileName), "\n", vbCrLf)
    g_Message(ERR_FILE_NOT_FOUND) = Replace$(strGet_ini("Message", "ERROR_FILE_NOT_FOUND", "File not found.", strFileName), "\n", vbCrLf)
    g_Message(ERR_LOAD_CANCEL) = Replace$(strGet_ini("Message", "ERROR_LOAD_CANCEL", "Loading will be aborted.", strFileName), "\n", vbCrLf)
    g_Message(ERR_SAVE_ERROR) = Replace$(strGet_ini("Message", "ERROR_SAVE_ERROR", "Error occured while saving.", strFileName), "\n", vbCrLf)
    g_Message(ERR_SAVE_CANCEL) = Replace$(strGet_ini("Message", "ERROR_SAVE_CANCEL", "Saving will be aborted.", strFileName), "\n", vbCrLf)
    g_Message(ERR_OVERFLOW_LARGE) = Replace$(strGet_ini("Message", "ERROR_OVERFLOW_LARGE", "Error:\nValue is too large.", strFileName), "\n", vbCrLf)
    g_Message(ERR_OVERFLOW_SMALL) = Replace$(strGet_ini("Message", "ERROR_OVERFLOW_SMALL", "Error:\nValue is too small.", strFileName), "\n", vbCrLf)
    g_Message(ERR_OVERFLOW_BPM) = Replace$(strGet_ini("Message", "ERROR_OVERFLOW_BPM", "You have used more than 1295 BPM change command.\nNumber of commands should be less than 1295.", strFileName), "\n", vbCrLf)
    g_Message(ERR_OVERFLOW_STOP) = Replace$(strGet_ini("Message", "ERROR_OVERFLOW_STOP", "You have used more than 1295 STOP command.\nNumber of commands should be less than 1295.", strFileName), "\n", vbCrLf)
    g_Message(ERR_APP_NOT_FOUND) = Replace$(strGet_ini("Message", "ERROR_APP_NOT_FOUND", " is not found.", strFileName), "\n", vbCrLf)
    g_Message(ERR_FILE_ALREADY_EXIST) = Replace$(strGet_ini("Message", "ERROR_FILE_ALREADY_EXIST", "File already exist.", strFileName), "\n", vbCrLf)
    
    g_Message(MSG_CONFIRM) = Replace$(strGet_ini("Message", "INFO_CONFIRM", "This command cannot be undone, OK?", strFileName), "\n", vbCrLf)
    g_Message(MSG_FILE_CHANGED) = Replace$(strGet_ini("Message", "INFO_FILE_CHANGED", "Do you want to save changes?", strFileName), "\n", vbCrLf)
    g_Message(MSG_INI_CHANGED) = Replace$(strGet_ini("Message", "INFO_INI_CHANGED", "ini format has changed.\n(All setting will reset)", strFileName), "\n", vbCrLf)
    g_Message(MSG_ALIGN_LIST) = Replace$(strGet_ini("Message", "INFO_ALIGN_LIST", "Do you want the filelist to be rewrited into the old format [01 - FF]?\n(Attention: Some programs are compatible only with old format files.)", strFileName), "\n", vbCrLf)
    g_Message(MSG_DELETE_FILE) = Replace$(strGet_ini("Message", "INFO_DELETE_FILE", "They have been deleted:", strFileName), "\n", vbCrLf)
    
    g_Message(INPUT_BPM) = Replace$(strGet_ini("Input", "INPUT_BPM", "Enter the BPM you wish to change to.\n(Decimal number can be used. Enter 0 to cancel)", strFileName), "\n", vbCrLf)
    g_Message(INPUT_STOP) = Replace$(strGet_ini("Input", "INPUT_STOP", "Enter the length of stoppage 1 corresponds to 1/192 of the measure.\n(Enter under 0 to cancel)", strFileName), "\n", vbCrLf)
    g_Message(INPUT_RENAME) = Replace$(strGet_ini("Input", "INPUT_RENAME", "Please enter new filename.", strFileName), "\n", vbCrLf)
    g_Message(INPUT_SIZE) = Replace$(strGet_ini("Input", "INPUT_SIZE", "Type your display magnification.\n(Maximum 16.00. Enter under 0 to cancel)", strFileName), "\n", vbCrLf)
    
    Dim i           As Long
    Dim SystemFont  As LOGFONT
    Dim DefaultFont As String
    
    Call GetObject(GetStockObject(DEFAULT_GUI_FONT), Len(SystemFont), SystemFont)
    
    For i = LBound(SystemFont.lfFaceName) To UBound(SystemFont.lfFaceName)
    
        DefaultFont = DefaultFont & Chr(SystemFont.lfFaceName(i))
    
    Next i
    
    DefaultFont = Trim$(DefaultFont)
    
    Call LoadFont(strGet_ini("Main", "Font", DefaultFont, strFileName), strGet_ini("Main", "FixedFont", DefaultFont, strFileName), strGet_ini("Main", "Charset", 1, strFileName))
    
    Call frmMain.Form_Resize

End Sub

Private Sub LoadFont(ByVal MainFont As String, ByVal FixedFont As String, ByVal Charset As Long)
On Error GoTo Err:

    Dim i       As Long
    Dim objCtl  As Object
    
    For i = 0 To Forms.Count - 1
    
        Forms(i).Font.Name = MainFont
        Forms(i).Font.Charset = Charset
        
        For Each objCtl In Forms(i).Controls
        
            If TypeOf objCtl Is Label Or TypeOf objCtl Is TextBox Or TypeOf objCtl Is ComboBox Or TypeOf objCtl Is CommandButton Or TypeOf objCtl Is OptionButton Or TypeOf objCtl Is CheckBox Or TypeOf objCtl Is Frame Then
            
                objCtl.Font.Name = MainFont
                objCtl.Font.Charset = Charset
            
            ElseIf TypeOf objCtl Is PictureBox Or TypeOf objCtl Is ListBox Then
            
                objCtl.Font.Name = FixedFont
                objCtl.Font.Charset = Charset
            
            End If
        
        Next objCtl
    
    Next i
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadFont")
End Sub

Public Sub LoadConfig()
On Error GoTo InitConfig:

    Dim i       As Long
    Dim wp      As WINDOWPLACEMENT
    Dim strRet  As String
    Dim lngRet  As Long
    Static lngCount As Long
    
    If strGet_ini("Main", "Key", "", "bmse.ini") <> "BMSE" Then GoTo InitConfig
    
    With frmMain
    
        strRet = strGet_ini("Main", "Language", "english.ini", "bmse.ini")
        
        For i = 1 To .mnuLanguage.UBound
        
            If strRet = g_strLangFileName(i) Then
            
                .mnuLanguage(i).Checked = True
                
                Exit For
            
            End If
        
        Next i
        
        Call Load(frmWindowAbout)
        Call Load(frmWindowFind)
        Call Load(frmWindowInput)
        Call Load(frmWindowPreview)
        Call Load(frmWindowTips)
        Call Load(frmWindowViewer)
        Call Load(frmWindowConvert)
        
        Call LoadLanguageFile("lang\" & strRet)
        
        Call frmWindowPreview.SetWindowSize
        
        If strGet_ini("Main", "ini", "", "bmse.ini") <> INI_VERSION Then
        
            Call MsgBox(g_Message(Message.MSG_INI_CHANGED), vbInformation, g_strAppTitle)
            
            GoTo InitConfig
        
        End If
        
        wp.Length = 44
        Call GetWindowPlacement(.hwnd, wp)
        
        With wp
        
            .showCmd = SW_HIDE
            
            With .rcNormalPosition
            
                .Right = strGet_ini("Main", "Width", 800, "bmse.ini")
                .Bottom = strGet_ini("Main", "Height", 600, "bmse.ini")
                '.Left = strGet_ini("Main", "X", (Screen.Width \ Screen.TwipsPerPixelX - .Right) \ 2, "bmse.ini")
                '.Top = strGet_ini("Main", "Y", (Screen.Height \ Screen.TwipsPerPixelY - .Bottom) \ 2, "bmse.ini")
                .Left = strGet_ini("Main", "X", 0, "bmse.ini")
                .Top = strGet_ini("Main", "Y", 0, "bmse.ini")
                .Right = .Left + .Right
                .Bottom = .Top + .Bottom
                
                If .Right > Screen.Width \ Screen.TwipsPerPixelX Then
                
                    .Left = Screen.Width \ Screen.TwipsPerPixelX - (.Right - .Left)
                    .Right = Screen.Width \ Screen.TwipsPerPixelX
                
                End If
                
                If .Left < 0 Then .Left = 0
                
                If .Bottom > Screen.Height \ Screen.TwipsPerPixelY Then
                
                    .Top = Screen.Height \ Screen.TwipsPerPixelY - (.Bottom - .Top)
                    .Bottom = Screen.Height \ Screen.TwipsPerPixelY
                
                End If
                
                If .Top < 0 Then .Top = 0
            
            End With
        
        End With
        
        strRet = strGet_ini("Main", "Theme", "default.ini", "bmse.ini")
        
        For i = 1 To .mnuTheme.UBound
        
            If strRet = g_strThemeFileName(i) Then
            
                .mnuTheme(i).Checked = True
                
                Exit For
            
            End If
        
        Next i
        
        Call LoadThemeFile("theme\" & strRet)
        
        g_strHelpFilename = strGet_ini("Main", "Help", "", "bmse.ini")
        g_strFiler = strGet_ini("Main", "Filer", "", "bmse.ini")
        
        If g_strHelpFilename <> "" Then
        
            .mnuHelpOpen.Enabled = True
        
        End If
        
        '.hsbDispWidth.Value = strGet_ini("View", "Width", 100, "bmse.ini")
        '.hsbDispHeight.Value = strGet_ini("View", "Height", 100, "bmse.ini")
        
        With .cboDispWidth
        
            lngRet = strGet_ini("View", "Width", 100, "bmse.ini")
            
            For i = 0 To .ListCount - 1
            
                If .ItemData(i) = lngRet Then
                
                    .ListIndex = i
                    
                    Exit For
                
                ElseIf .ItemData(i) > lngRet Then
                
                    Call .AddItem("x" & Format$(lngRet / 100, "#0.00"), i)
                    .ItemData(i) = lngRet
                    .ListIndex = i
                    
                    Exit For
                
                End If
            
            Next i
        
        End With
        
        With .cboDispHeight
        
            lngRet = strGet_ini("View", "Height", 100, "bmse.ini")
            
            For i = 0 To .ListCount - 1
            
                If .ItemData(i) = lngRet Then
                
                    .ListIndex = i
                    
                    Exit For
                
                ElseIf .ItemData(i) > lngRet Then
                
                    Call .AddItem("x" & Format$(lngRet / 100, "#0.00"), i)
                    .ItemData(i) = lngRet
                    .ListIndex = i
                    
                    Exit For
                
                End If
            
            Next i
        
        End With
        
        .cboDispGridMain.ListIndex = strGet_ini("View", "VGridMain", 1, "bmse.ini")
        .cboDispGridSub.ListIndex = strGet_ini("View", "VGridSub", 1, "bmse.ini")
        .cboDispFrame.ListIndex = strGet_ini("View", "Frame", 1, "bmse.ini")
        .cboDispKey.ListIndex = strGet_ini("View", "Key", 1, "bmse.ini")
        .cboDispSC1P.ListIndex = strGet_ini("View", "SC_1P", 1, "bmse.ini")
        .cboDispSC2P.ListIndex = strGet_ini("View", "SC_2P", 1, "bmse.ini")
        
        .mnuViewToolBar.Checked = strGet_ini("View", "ToolBar", True, "bmse.ini")
        .mnuViewDirectInput.Checked = strGet_ini("View", "DirectInput", True, "bmse.ini")
        .mnuViewStatusBar.Checked = strGet_ini("View", "StatusBar", True, "bmse.ini")
        
        If .cboViewer.ListCount Then
        
            If .cboViewer.ListCount > strGet_ini("View", "ViewerNum", 0, "bmse.ini") Then
            
                .cboViewer.ListIndex = strGet_ini("View", "ViewerNum", 0, "bmse.ini")
            
            Else
            
                .cboViewer.ListIndex = 0
            
            End If
        
        End If
        
        .mnuOptionsActiveIgnore.Checked = strGet_ini("Options", "Active", True, "bmse.ini")
        .mnuOptionsFileNameOnly.Checked = strGet_ini("Options", "FileNameOnly", False, "bmse.ini")
        .mnuOptionsVertical.Checked = strGet_ini("Options", "VerticalWriting", False, "bmse.ini")
        .mnuOptionsLaneBG.Checked = strGet_ini("Options", "LaneBG", True, "bmse.ini")
        .mnuOptionsSelectPreview.Checked = strGet_ini("Options", "SelectSound", True, "bmse.ini")
        .mnuOptionsMoveOnGrid.Checked = strGet_ini("Options", "MoveOnGrid", True, "bmse.ini")
        .mnuOptionsObjectFileName.Checked = strGet_ini("Options", "ObjectFileName", False, "bmse.ini")
        .mnuOptionsNumFF.Checked = strGet_ini("Options", "UseOldFormat", False, "bmse.ini")
        .mnuOptionsRightClickDelete.Checked = strGet_ini("Options", "RightClickDelete", False, "bmse.ini")
        
        .tlbMenu.Buttons("New").Visible = strGet_ini("ToolBar", "New", True, "bmse.ini")
        .tlbMenu.Buttons("Open").Visible = strGet_ini("ToolBar", "Open", True, "bmse.ini")
        .tlbMenu.Buttons("Reload").Visible = strGet_ini("ToolBar", "Reload", False, "bmse.ini")
        .tlbMenu.Buttons("Save").Visible = strGet_ini("ToolBar", "Save", True, "bmse.ini")
        .tlbMenu.Buttons("SaveAs").Visible = strGet_ini("ToolBar", "SaveAs", True, "bmse.ini")
        
        .tlbMenu.Buttons("SepMode").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
        .tlbMenu.Buttons("Edit").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
        .tlbMenu.Buttons("Write").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
        .tlbMenu.Buttons("Delete").Visible = strGet_ini("ToolBar", "Mode", True, "bmse.ini")
        
        .tlbMenu.Buttons("SepViewer").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
        .tlbMenu.Buttons("PlayAll").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
        .tlbMenu.Buttons("Play").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
        .tlbMenu.Buttons("Stop").Visible = strGet_ini("ToolBar", "Preview", True, "bmse.ini")
        
        .tlbMenu.Buttons("SepGrid").Visible = strGet_ini("ToolBar", "Grid", True, "bmse.ini")
        .tlbMenu.Buttons("ChangeGrid").Visible = strGet_ini("ToolBar", "Grid", True, "bmse.ini")
        
        .tlbMenu.Buttons("SepSize").Visible = strGet_ini("ToolBar", "Size", True, "bmse.ini")
        .tlbMenu.Buttons("DispSize").Visible = strGet_ini("ToolBar", "Size", True, "bmse.ini")
        
        .tlbMenu.Buttons("SepResolution").Visible = strGet_ini("ToolBar", "Resolution", False, "bmse.ini")
        .tlbMenu.Buttons("Resolution").Visible = strGet_ini("ToolBar", "Resolution", False, "bmse.ini")
        
        For i = 0 To UBound(g_strRecentFiles)
        
            g_strRecentFiles(i) = strGet_ini("RecentFiles", i, "", "bmse.ini")
            
            If Len(g_strRecentFiles(i)) Then
            
                With .mnuRecentFiles(i)
                
                    .Caption = "&" & Right$(i + 1, 1) & ":" & g_strRecentFiles(i)
                    .Enabled = True
                    .Visible = True
                
                End With
                
                With .tlbMenu.Buttons("Open").ButtonMenus(i + 1)
                
                    .Text = "&" & Right$(i + 1, 1) & ":" & g_strRecentFiles(i)
                    .Enabled = True
                    .Visible = True
                
                End With
                
                .mnuLineRecent.Visible = True
            
            End If
        
        Next i
        
        Call SetWindowPlacement(.hwnd, wp)
            
    End With
    
    Call modEasterEgg.InitEffect
    
    With frmWindowPreview
    
        .Left = strGet_ini("Preview", "X", ((frmMain.Left + frmMain.Width \ 2) - .Width \ 2), "bmse.ini")
        If .Left < 0 Then .Left = 0
        If .Left > Screen.Width Then .Left = 0
        
        .Top = strGet_ini("Preview", "Y", ((frmMain.Top + frmMain.Height \ 2) - .Height \ 2), "bmse.ini")
        If .Top < 0 Then .Top = 0
        If .Top > Screen.Height Then .Top = 0
    
    End With
    
    Exit Sub
    
InitConfig:

    lngCount = lngCount + 1
    
    If lngCount > 5 Then
    
        Call modMain.CleanUp(Err.Number, Err.Description, "LoadConfig")
    
    Else
    
        Call CreateConfig
    
    End If

End Sub

Private Sub CreateConfig()

    Call lngSet_ini("Main", "Key", Chr$(34) & "BMSE" & Chr$(34))
    Call lngSet_ini("Main", "ini", INI_VERSION)
    'Call lngSet_ini("Main", "X", (Screen.Width \ Screen.TwipsPerPixelX - 800) \ 2)
    'Call lngSet_ini("Main", "Y", (Screen.Height \ Screen.TwipsPerPixelY - 600) \ 2)
    Call lngSet_ini("Main", "X", 0)
    Call lngSet_ini("Main", "Y", 0)
    Call lngSet_ini("Main", "Width", "800")
    Call lngSet_ini("Main", "Height", "600")
    Call lngSet_ini("Main", "State", SW_SHOWNORMAL)
    Call lngSet_ini("Main", "Language", Chr$(34) & "english.ini" & Chr$(34))
    Call lngSet_ini("Main", "Theme", Chr$(34) & "default.ini" & Chr$(34))
    Call lngSet_ini("Main", "Help", Chr$(34) & Chr$(34))
    
    Call lngSet_ini("View", "Width", 100)
    Call lngSet_ini("View", "Height", 100)
    Call lngSet_ini("View", "VGridMain", 1)
    Call lngSet_ini("View", "VGridSub", 1)
    Call lngSet_ini("View", "Frame", 1)
    Call lngSet_ini("View", "Key", 1)
    Call lngSet_ini("View", "SC_1P", 0)
    Call lngSet_ini("View", "SC_2P", 1)
    
    Call lngSet_ini("View", "ToolBar", True)
    Call lngSet_ini("View", "DirectInput", True)
    Call lngSet_ini("View", "StatusBar", True)
    
    Call lngSet_ini("View", "ViewerNum", 0)
    
    Call lngSet_ini("ToolBar", "New", True)
    Call lngSet_ini("ToolBar", "Open", True)
    Call lngSet_ini("ToolBar", "Reload", False)
    Call lngSet_ini("ToolBar", "Save", True)
    Call lngSet_ini("ToolBar", "SaveAs", True)
    Call lngSet_ini("ToolBar", "Mode", True)
    Call lngSet_ini("ToolBar", "Preview", True)
    Call lngSet_ini("ToolBar", "Gird", True)
    Call lngSet_ini("ToolBar", "Size", True)
    Call lngSet_ini("ToolBar", "Resolution", False)
    
    Call lngSet_ini("Options", "Active", True)
    Call lngSet_ini("Options", "FileNameOnly", False)
    Call lngSet_ini("Options", "VerticalWriting", False)
    Call lngSet_ini("Options", "LaneBG", True)
    Call lngSet_ini("Options", "SelectSound", True)
    Call lngSet_ini("Options", "MoveOnGrid", True)
    Call lngSet_ini("Options", "ObjectFileName", False)
    Call lngSet_ini("Options", "UseOldFormat", False)
    Call lngSet_ini("Options", "RightClickDelete", False)
    
    Call lngSet_ini("Preview", "X", 0)
    Call lngSet_ini("Preview", "Y", 0)
    
    Call LoadConfig

End Sub

Public Sub SaveConfig()

    Dim i   As Long
    Dim wp As WINDOWPLACEMENT
    
    With frmMain
    
        Call lngSet_ini("Main", "Key", Chr$(34) & "BMSE" & Chr$(34))
        
        wp.Length = 44
        Call GetWindowPlacement(frmMain.hwnd, wp)
        
        With wp
        
            If wp.showCmd <> SW_SHOWMINIMIZED Then
            
                Call lngSet_ini("Main", "State", wp.showCmd)
            
            Else
            
                Call lngSet_ini("Main", "State", SW_SHOWNORMAL)
            
            End If
            
            With .rcNormalPosition
            
                Call lngSet_ini("Main", "X", .Left)
                Call lngSet_ini("Main", "Y", .Top)
                Call lngSet_ini("Main", "Width", .Right - .Left)
                Call lngSet_ini("Main", "Height", .Bottom - .Top)
            
            End With
        
        End With
        
        For i = 1 To .mnuLanguage.UBound
        
            If .mnuLanguage(i).Checked = True Then
            
                Call lngSet_ini("Main", "Language", Chr$(34) & g_strLangFileName(i) & Chr$(34))
                
                Exit For
            
            End If
        
        Next i
        
        For i = 1 To .mnuTheme.UBound
        
            If .mnuTheme(i).Checked = True Then
            
                Call lngSet_ini("Main", "Theme", Chr$(34) & g_strThemeFileName(i) & Chr$(34))
                
                Exit For
            
            End If
        
        Next i
        
        Call lngSet_ini("View", "Width", .cboDispWidth.ItemData(.cboDispWidth.ListIndex))
        Call lngSet_ini("View", "Height", .cboDispHeight.ItemData(.cboDispHeight.ListIndex))
        Call lngSet_ini("View", "VGridMain", .cboDispGridMain.ListIndex)
        Call lngSet_ini("View", "VGridSub", .cboDispGridSub.ListIndex)
        Call lngSet_ini("View", "Frame", .cboDispFrame.ListIndex)
        Call lngSet_ini("View", "Key", .cboDispKey.ListIndex)
        Call lngSet_ini("View", "SC_1P", .cboDispSC1P.ListIndex)
        Call lngSet_ini("View", "SC_2P", .cboDispSC2P.ListIndex)
        
        Call lngSet_ini("View", "ToolBar", .mnuViewToolBar.Checked)
        Call lngSet_ini("View", "DirectInput", .mnuViewDirectInput.Checked)
        Call lngSet_ini("View", "StatusBar", .mnuViewStatusBar.Checked)
        
        If .cboViewer.ListCount Then
        
            Call lngSet_ini("View", "ViewerNum", .cboViewer.ListIndex)
        
        End If
        
        Call lngSet_ini("Options", "Active", .mnuOptionsActiveIgnore.Checked)
        Call lngSet_ini("Options", "FileNameOnly", .mnuOptionsFileNameOnly.Checked)
        Call lngSet_ini("Options", "VerticalWriting", .mnuOptionsVertical.Checked)
        Call lngSet_ini("Options", "LaneBG", .mnuOptionsLaneBG.Checked)
        Call lngSet_ini("Options", "SelectSound", .mnuOptionsSelectPreview.Checked)
        Call lngSet_ini("Options", "MoveOnGrid", .mnuOptionsMoveOnGrid.Checked)
        Call lngSet_ini("Options", "ObjectFileName", .mnuOptionsObjectFileName.Checked)
        Call lngSet_ini("Options", "UseOldFormat", .mnuOptionsNumFF.Checked)
        Call lngSet_ini("Options", "RightClickDelete", .mnuOptionsRightClickDelete.Checked)
        
        Call lngSet_ini("ToolBar", "New", .tlbMenu.Buttons("New").Visible)
        Call lngSet_ini("ToolBar", "Open", .tlbMenu.Buttons("Open").Visible)
        Call lngSet_ini("ToolBar", "Reload", .tlbMenu.Buttons("Reload").Visible)
        Call lngSet_ini("ToolBar", "Save", .tlbMenu.Buttons("Save").Visible)
        Call lngSet_ini("ToolBar", "SaveAs", .tlbMenu.Buttons("SaveAs").Visible)
        Call lngSet_ini("ToolBar", "Mode", .tlbMenu.Buttons("Write").Visible)
        Call lngSet_ini("ToolBar", "Preview", .tlbMenu.Buttons("PlayAll").Visible)
        Call lngSet_ini("ToolBar", "Size", .tlbMenu.Buttons("DispSize").Visible)
        Call lngSet_ini("ToolBar", "Resolution", .tlbMenu.Buttons("Resolution").Visible)
        
        For i = 0 To UBound(g_strRecentFiles)
        
            Call lngSet_ini("RecentFiles", i, Chr$(34) & g_strRecentFiles(i) & Chr$(34))
        
        Next i
    
    End With
    
    With frmWindowPreview
    
        Call lngSet_ini("Preview", "X", .Left)
        Call lngSet_ini("Preview", "Y", .Top)
    
    End With

End Sub

Public Function lngSet_ini(ByVal strSection As String, ByVal strKey As String, ByVal strSet As String) As Long

    Dim lngRet  As Long
    
    'API呼び出し＆変数を返す
    lngRet = WritePrivateProfileString(strSection & Chr(0), strKey, strSet, g_strAppDir & "bmse.ini" & Chr(0))
    
    lngSet_ini = lngRet

End Function

Public Function strGet_ini(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String, ByVal strFileName As String) As String

    Dim strGetBuf   As String * 256 '収容するstringのバッファ
    Dim intGetLen   As Integer      '収容するstringの文字数のバッファ
    
    'バッファの初期化（256もあれば良いよね。）
    strGetBuf = Space$(256)
    
    'API呼び出し
    intGetLen = GetPrivateProfileString(strSection & Chr(0), strKey, strDefault & Chr(0), strGetBuf, 128, g_strAppDir & strFileName & Chr(0))
    
    '文字列を返す
    strGet_ini = Trim$(Left$(strGetBuf, InStr(strGetBuf, Chr(0)) - 1))
    
    If Val(strGet_ini) < 0 Then strGet_ini = 0

End Function
