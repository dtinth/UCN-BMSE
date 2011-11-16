VERSION 5.00
Begin VB.Form frmWindowConvert 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "変換ウィザード"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CheckBox chkSortByName 
      Caption         =   "ファイル名順でソートする"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CheckBox chkFileRecycle 
      Caption         =   "ごみ箱に移動しないですぐに削除する"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
   End
   Begin VB.CheckBox chkDeleteFile 
      Caption         =   "フォルダ内の使用していないファイルを削除 (*)"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox txtExtension 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1680
      TabIndex        =   5
      Text            =   "wav,mp3,bmp,jpg,gif"
      Top             =   1140
      Width           =   2895
   End
   Begin VB.CheckBox chkFileNameConvert 
      Caption         =   "ファイル名を連番 (01 - ZZ) に変換 (*)"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CheckBox chkUseOldFormat 
      Caption         =   "可能なら古いフォーマット (01 - FF) を使う"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Value           =   1  'ﾁｪｯｸ
      Width           =   4335
   End
   Begin VB.CheckBox chkListAlign 
      Caption         =   "定義リストの整列"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdDecide 
      Caption         =   "実行"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox chkDeleteUnusedFile 
      Caption         =   "使用していない #WAV・#BMP・#BGA の定義を消去"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label lblNotice 
      Caption         =   "(*)・・・この操作はやり直しができません"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblExtension 
      AutoSize        =   -1  'True
      Caption         =   "検索する拡張子:"
      Enabled         =   0   'False
      Height          =   180
      Left            =   390
      TabIndex        =   4
      Top             =   1200
      Width           =   1260
   End
End
Attribute VB_Name = "frmWindowConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type SHFILEOPSTRUCT
    hwnd                    As Long
    wFunc                   As Long
    pFrom                   As String
    pTo                     As String
    fFlags                  As Integer
    fAnyOperationsAborted   As Long
    hNameMappings           As Long
    lpszProgressTitle       As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4                      '  don't create progress/report
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Private Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Private Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Private Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

Public Sub DeleteUnusedFile()

    Dim i               As Long
    Dim lngRet          As Long
    Dim blnWAV(1295)    As Boolean
    Dim blnBMP(1295)    As Boolean
    Dim blnBGA(1295)    As Boolean
    Dim strArray()      As String
    Dim strLogArray()   As String
    
    ReDim strLogArray(0)
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            Select Case .intCh
            
                Case Is > 100
                
                    blnWAV(.sngValue) = True
                
                Case 11 To 29
                
                    blnWAV(.sngValue) = True
                
                Case 4, 6, 7
                
                    blnBMP(.sngValue) = True
                    blnBGA(.sngValue) = True
            
            End Select
        
        End With
    
    Next i
    
    For i = 0 To UBound(blnBGA)
    
        If blnBGA(i) Then
        
            If Len(g_strBGA(i)) Then
            
                strArray = Split(g_strBGA(i), " ")
                
                lngRet = modInput.lngNumConv(strArray(0))
                
                If lngRet >= 0 And lngRet <= UBound(blnBMP) Then
                
                    blnBMP(lngRet) = True
                
                End If
            
            End If
        
        End If
    
    Next i
    
    For i = 1 To 1295
    
        If Not blnWAV(i) Then
        
            If Len(g_strWAV(i)) Then
            
                strLogArray(UBound(strLogArray)) = modInput.strNumConv(CMD_LOG.LIST_DEL) & "1" & modInput.strNumConv(i) & g_strWAV(i)
                ReDim Preserve strLogArray(UBound(strLogArray) + 1)
                
                g_strWAV(i) = ""
            
            End If
        
        End If
        
        If Not blnBMP(i) Then
        
            If Len(g_strBMP(i)) Then
                
                strLogArray(UBound(strLogArray)) = modInput.strNumConv(CMD_LOG.LIST_DEL) & "2" & modInput.strNumConv(i) & g_strBMP(i)
                ReDim Preserve strLogArray(UBound(strLogArray) + 1)
                
                g_strBMP(i) = ""
            
            End If
        
        End If
        
        If Not blnBGA(i) Then
        
            If Len(g_strBGA(i)) Then
                
                strLogArray(UBound(strLogArray)) = modInput.strNumConv(CMD_LOG.LIST_DEL) & "3" & modInput.strNumConv(i) & g_strBGA(i)
                ReDim Preserve strLogArray(UBound(strLogArray) + 1)
                
                g_strBGA(i) = ""
            
            End If
        
        End If
    
    Next i
    
    If UBound(strLogArray) Then
    
        'g_strInputLog(g_lngInputLogPos) = Join(strLogArray(), ",")
        'g_lngInputLogPos = g_lngInputLogPos + 1
        'ReDim Preserve g_strInputLog(g_lngInputLogPos)
        'Call frmMain.SaveChanges
        Call g_InputLog.AddData(Join(strLogArray(), ","))
    
    End If

End Sub

Public Sub DeleteFile()

    Dim i           As Long
    Dim j           As Long
    Dim lngRet      As Long
    Dim strArray()  As String
    Dim strList()   As String
    Dim strDeleteList() As String
    Dim strRet      As String
    Dim sh          As SHFILEOPSTRUCT
    
    If Len(Trim$(txtExtension.Text)) = 0 Then Exit Sub
    
    strArray() = Split(txtExtension.Text, ",")
    
    For i = 0 To UBound(strArray)
    
        strArray(i) = UCase$(strArray(i))
    
    Next i
    
    ReDim strList(1295 + 1295 + 1)
    
    For i = 1 To 1295
    
        If Len(g_strWAV(i)) Then
        
            strList(lngRet) = UCase$(g_strWAV(i))
            lngRet = lngRet + 1
        
        End If
        
        If Len(g_strBMP(i)) Then
        
            strList(lngRet) = UCase$(g_strBMP(i))
            lngRet = lngRet + 1
        
        End If
    
    Next i
    
    If Len(frmMain.txtMissBMP.Text) Then
    
        strList(lngRet) = UCase$(Trim$(frmMain.txtMissBMP.Text))
        lngRet = lngRet + 1
    
    End If
    
    If Len(frmMain.txtStageFile.Text) Then
    
        strList(lngRet) = UCase$(Trim$(frmMain.txtStageFile.Text))
        lngRet = lngRet + 1
    
    End If
    
    If lngRet = 0 Then Exit Sub
    
    ReDim Preserve strList(lngRet - 1)
    
    ReDim strDeleteList(0)
    lngRet = 0
    
    strRet = Dir(g_BMS.strDir & "*.*", vbNormal)
    
    Do While strRet <> vbNullString
    
        For i = 0 To UBound(strArray)
        
            If strArray(i) = UCase(Mid$(strRet, InStrRev(strRet, ".") + 1)) Then
            
                For j = 0 To UBound(strList)
                
                    If UCase$(strRet) = strList(j) Then
                    
                        Exit For
                    
                    End If
                
                Next j
                
                If j = UBound(strList) + 1 Then
                
                    ReDim Preserve strDeleteList(lngRet)
                    strDeleteList(lngRet) = g_BMS.strDir & strRet
                    lngRet = lngRet + 1
                
                End If
            
            End If
        
        Next i
        
        strRet = Dir
    
    Loop
    
    If lngRet <> 0 Then
    
        If chkFileRecycle.value Then
        
            For i = 0 To UBound(strDeleteList)
            
                Call lngDeleteFile(strDeleteList(i))
            
            Next i
        
        Else
        
            With sh
            
                .hwnd = frmWindowConvert.hwnd
                .wFunc = FO_DELETE
                .pFrom = Join(strDeleteList(), Chr$(0))
                .pTo = vbNullString
                .fFlags = FOF_SILENT Or FOF_ALLOWUNDO
            
            End With
            
            Call SHFileOperation(sh)
        
        End If
        
        Call MsgBox(g_Message(MSG_DELETE_FILE) & vbCrLf & vbCrLf & Replace$(Join(strDeleteList, vbCrLf), g_BMS.strDir, ""), vbInformation, g_strAppTitle)
    
    End If

End Sub

Public Sub ListAlign()

    Dim i                   As Long
    'Dim blnUseOldFormat     As Boolean
    Dim intRet              As Integer
    Dim lngRet              As Long
    Dim lngWAV              As Long
    Dim lngBMP              As Long
    Dim lngArray(1295)      As Long
    Dim lngArrayWAV(1295)   As Long
    Dim lngArrayBMP(1295)   As Long
    Dim strArray()          As String
    Dim strLogArray()       As String
    
    '共通の処理とか
    For i = 0 To UBound(g_Obj) - 1 '空オブジェ対策
    
        With g_Obj(i)
        
            Select Case .intCh
            
                Case Is >= 11
                
                    If Len(g_strWAV(.sngValue)) = 0 Then
                    
                        g_strWAV(.sngValue) = "***"
                    
                    End If
                
                Case 4, 6, 7
                
                    If Len(g_strBMP(.sngValue)) = 0 Then
                    
                        g_strBMP(.sngValue) = "***"
                    
                    End If
                    
                    If Len(g_strBGA(.sngValue)) = 0 Then
                    
                        g_strBGA(.sngValue) = "***"
                    
                    End If
            
            End Select
        
        End With
    
    Next i
    
    For i = 1 To UBound(g_strWAV)
    
        If Len(g_strWAV(i)) Then
        
            lngWAV = lngWAV + 1
        
        End If
        
        If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then
        
            lngBMP = lngBMP + 1
        
        End If
        
        lngArrayWAV(i) = i
        lngArrayBMP(i) = i
        lngArray(i) = i
    
    Next i
    
    If (lngWAV < 256 And lngBMP < 256) And (lngWAV > 0 Or lngBMP > 0) Then
    
        'If MsgBox(g_Message(Message.MSG_ALIGN_LIST), vbYesNo + vbInformation, g_strAppTitle) = vbNo Then
        
            'blnUseOldFormat = True
        
        'Else
        
            'blnUseOldFormat = False
        
        'End If
        
        If chkUseOldFormat.value Then
        
            'blnUseOldFormat = True
            frmMain.mnuOptionsNumFF.Checked = True
        
        Else
        
            'blnUseOldFormat = False
            frmMain.mnuOptionsNumFF.Checked = False
        
        End If
    
    ElseIf lngWAV > 255 Or lngBMP > 255 Then
    
        'blnUseOldFormat = False
        frmMain.mnuOptionsNumFF.Checked = False
    
    End If
    
    'ファイル名ソート
    If chkSortByName.value Then
    
        Dim strDummy(1295) As String
        Call strQSort(g_strWAV(), strDummy(), lngArrayWAV(), 1, UBound(g_strWAV))
        Call strQSort(g_strBMP(), g_strBGA(), lngArrayBMP(), 1, UBound(g_strBMP))
    
    End If
    
    ReDim strLogArray(0)
    strLogArray(0) = ""
    
    'ここからWAV
    If lngWAV Then
    
        For i = 0 To UBound(lngArray)
        
            lngArray(i) = lngArrayWAV(i)
        
        Next i
        
        '255以下ならまず後ろに整列する
        'If blnUseOldFormat Then
        If frmMain.mnuOptionsNumFF.Checked Then
        
            lngRet = 1295
            
            For i = UBound(g_strWAV) To 1 Step -1
            
                If Len(g_strWAV(i)) Then
                
                    If i <> lngRet Then
                    
                        'g_strWAV(lngRet) = g_strWAV(i)
                        'g_strWAV(i) = ""
                        Call swapString(g_strWAV(), lngRet, i)
                        
                        'strRet = lngArray(i)
                        'lngArray(i) = lngArray(lngRet)
                        'lngArray(lngRet) = val(strRet)
                        Call swapValue(lngArray(), lngRet, i)
                    
                    End If
                    
                    lngRet = lngRet - 1
                
                End If
            
            Next i
        
        End If
        
        lngRet = 1
        intRet = 1
        
        For i = 1 To UBound(g_strWAV)
        
            If Len(g_strWAV(i)) Then
            
                'If blnUseOldFormat Then
                
                    'intRet = modInput.lngNumConv(Hex$(lngRet))
                
                'Else
                
                    'intRet = lngRet
                
                'End If
                
                intRet = frmMain.lngFromLong(lngRet)
                
                If intRet <> i Then
                
                    'g_strWAV(intRet) = g_strWAV(i)
                    'g_strWAV(i) = ""
                    Call swapString(g_strWAV(), intRet, i)
                    
                    'strRet = lngArray(intRet)
                    'lngArray(intRet) = lngArray(i)
                    'lngArray(i) = val(strRet)
                    Call swapValue(lngArray(), intRet, i)
                    
                    lngWAV = 1
                
                End If
                
                lngRet = lngRet + 1
            
            End If
        
        Next i
        
        If lngWAV Then
        
            For i = 1 To UBound(lngArray)
            
                If Len(g_strWAV(i)) <> 0 Then
                
                    lngArrayWAV(lngArray(i)) = i
                    
                    strLogArray(UBound(strLogArray)) = "1" & modInput.strNumConv(lngArray(i)) & modInput.strNumConv(i)
                    ReDim Preserve strLogArray(UBound(strLogArray) + 1)
                
                End If
            
            Next i
        
        End If
    
    End If
    
    'ここからBMP/BGA
    If lngBMP Then
    
        lngBMP = 0
        
        For i = 0 To UBound(lngArray)
        
            lngArray(i) = lngArrayBMP(i)
        
        Next i
        
        '255以下ならまず後ろに整列する
        'If blnUseOldFormat Then
        If frmMain.mnuOptionsNumFF.Checked Then
        
            lngRet = 1295
            
            For i = UBound(g_strBMP) To 1 Step -1
            
                If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then
                
                    If i <> lngRet Then
                    
                        'g_strBMP(lngRet) = g_strBMP(i)
                        'g_strBMP(i) = ""
                        'g_strBGA(lngRet) = g_strBGA(i)
                        'g_strBGA(i) = ""
                        Call swapString(g_strBMP(), lngRet, i)
                        Call swapString(g_strBGA(), lngRet, i)
                        
                        'strRet = lngArray(i)
                        'lngArray(i) = lngArray(lngRet)
                        'lngArray(lngRet) = val(strRet)
                        Call swapValue(lngArray(), i, lngRet)
                    
                    End If
                    
                    lngRet = lngRet - 1
                
                End If
            
            Next i
        
        End If
        
        lngRet = 1
        intRet = 1
        
        For i = 1 To UBound(g_strBMP)
        
            If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then
            
                'If blnUseOldFormat Then
                
                    'intRet = modInput.lngNumConv(Hex$(lngRet))
                
                'Else
                
                    'intRet = lngRet
                
                'End If
                
                intRet = frmMain.lngFromLong(lngRet)
                
                If intRet <> i Then
                
                    'g_strBMP(intRet) = g_strBMP(i)
                    'g_strBMP(i) = ""
                    'g_strBGA(intRet) = g_strBGA(i)
                    'g_strBGA(i) = ""
                    Call swapString(g_strBMP(), intRet, i)
                    Call swapString(g_strBGA(), intRet, i)
                    
                    'strRet = lngArray(intRet)
                    'lngArray(intRet) = lngArray(i)
                    'lngArray(i) = val(strRet)
                    Call swapValue(lngArray(), intRet, i)
                    
                    lngBMP = 1
                
                End If
                
                lngRet = lngRet + 1
            
            End If
        
        Next i
        
        If lngBMP Then
        
            For i = 1 To UBound(lngArray)
            
                If Len(g_strBMP(i)) <> 0 Or Len(g_strBGA(i)) <> 0 Then
                
                    lngArrayBMP(lngArray(i)) = i
                    
                    strLogArray(UBound(strLogArray)) = "2" & modInput.strNumConv(lngArray(i)) & modInput.strNumConv(i)
                    ReDim Preserve strLogArray(UBound(strLogArray) + 1)
                
                End If
            
            Next i
            
            'BGAの方も直すよー
            For i = 0 To UBound(g_strBGA)
            
                If Len(g_strBGA(i)) Then
                
                    strArray() = Split(g_strBGA(i), " ")
                    
                    If UBound(strArray) Then
                    
                        strArray(0) = modInput.strNumConv(lngArrayBMP(modInput.lngNumConv(strArray(0))), 2)
                        g_strBGA(i) = Join(strArray, " ")
                    
                    End If
                
                End If
            
            Next i
        
        End If
    
    End If
    
    '後の処理
    For i = 0 To UBound(g_strWAV)
    
        If g_strWAV(i) = "***" Then g_strWAV(i) = ""
        If g_strBMP(i) = "***" Then g_strBMP(i) = ""
        If g_strBGA(i) = "***" Then g_strBGA(i) = ""
    
    Next i
    
    If lngWAV <> 0 Or lngBMP <> 0 Then
        
        'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.LIST_ALIGN) & Join(strLogArray, "") & ","
        'g_lngInputLogPos = g_lngInputLogPos + 1
        'ReDim Preserve g_strInputLog(g_lngInputLogPos)
        'Call frmMain.SaveChanges
        Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.LIST_ALIGN) & Join(strLogArray, "") & ",")
        
        'Call RefreshList
        
        For i = 0 To UBound(g_Obj) - 1
        
            With g_Obj(i)
            
                Select Case .intCh
                
                    Case Is >= 11
                    
                        .sngValue = lngArrayWAV(.sngValue)
                    
                    Case 4, 6, 7
                    
                        .sngValue = lngArrayBMP(.sngValue)
                
                End Select
            
            End With
        
        Next i
        
        Call modDraw.Redraw
    
    End If

End Sub

Private Sub strQSort(ByRef strArray1() As String, ByRef strArray2() As String, ByRef lngArray() As Long, ByVal lngLeft As Long, ByVal lngRight As Long)

    Dim i   As Long
    Dim j   As Long

    If lngLeft >= lngRight Then Exit Sub
    
    i = lngLeft + 1
    j = lngRight
    
    Do While i <= j
    
        Do While i <= j
        
            If StrComp(strArray1(i), strArray1(lngLeft)) > 0 Then
                Exit Do
            End If
            
            i = i + 1
        
        Loop
        
        Do While i <= j
        
            If StrComp(strArray1(j), strArray1(lngLeft)) < 0 Then
                Exit Do
            End If
            
            j = j - 1
        
        Loop
        
        If i >= j Then Exit Do
        
        Call swapString(strArray1(), j, i)
        Call swapString(strArray2(), j, i)
        Call swapValue(lngArray(), j, i)
        
        i = i + 1
        j = j - 1
    
    Loop
    
    Call swapString(strArray1(), j, lngLeft)
    Call swapString(strArray2(), j, lngLeft)
    Call swapValue(lngArray(), j, lngLeft)
    
    Call strQSort(strArray1(), strArray2(), lngArray(), lngLeft, j - 1)
    Call strQSort(strArray1(), strArray2(), lngArray(), j + 1, lngRight)

End Sub

Private Sub swapString(ByRef strArray() As String, ByVal i As Long, ByVal j As Long)

    Dim str As String
    
    str = strArray(i)
    strArray(i) = strArray(j)
    strArray(j) = str

End Sub

Private Sub swapValue(ByRef lngArray() As Long, ByVal i As Long, ByVal j As Long)

    Dim value   As Long
    
    value = lngArray(i)
    lngArray(i) = lngArray(j)
    lngArray(j) = value

End Sub

Public Sub FileNameConvert()

    Dim i           As Long
    Dim j           As Long
    Dim strArray()  As String
    Dim strRet      As String
    Dim intRet      As Integer
    Dim blnRet      As Boolean
    Dim blnWAV(1295)    As Boolean
    Dim blnBMP(1295)    As Boolean
    Dim lngRet          As Long
    Dim strNameFrom()   As String
    Dim strNameTo()     As String
    Dim sh              As SHFILEOPSTRUCT
    
    ReDim strNameFrom(0)
    ReDim strNameTo(0)
    lngRet = 0
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    For i = 1 To 1295
    
        If Not blnWAV(i) Then
        
            If Len(g_strWAV(i)) <> 0 And Dir(g_BMS.strDir & g_strWAV(i), vbNormal) <> vbNullString Then
            
                strArray() = Split(g_strWAV(i), ".")
                strRet = strNumConv(i) & "." & strArray(UBound(strArray))
                
                If Dir(g_BMS.strDir & strRet, vbNormal) = vbNullString Then
                
                    blnRet = True '変換するよ
                
                Else
                
                    blnRet = False 'しないよ
                
                End If
                
                For j = i + 1 To UBound(g_strWAV)
                
                    If Not blnWAV(j) Then
                    
                        If g_strWAV(i) = g_strWAV(j) Then
                        
                            If blnRet Then g_strWAV(j) = strRet
                            
                            blnWAV(j) = True
                        
                        End If
                    
                    End If
                
                Next j
                
                If blnRet Then
                
                    'Name g_BMS.strDir & g_strWAV(i) As g_BMS.strDir & strRet
                    
                    ReDim Preserve strNameFrom(lngRet)
                    ReDim Preserve strNameTo(lngRet)
                    
                    strNameFrom(lngRet) = g_BMS.strDir & g_strWAV(i)
                    strNameTo(lngRet) = g_BMS.strDir & strRet
                    
                    lngRet = lngRet + 1
                    
                    g_strWAV(i) = strRet
                    
                    intRet = 1
                
                End If
                
                blnWAV(i) = True
            
            End If
        
        End If
        
        If Not blnBMP(i) Then
        
            If Len(g_strBMP(i)) <> 0 And Dir(g_BMS.strDir & g_strBMP(i), vbNormal) <> vbNullString Then
            
                strArray() = Split(g_strBMP(i), ".")
                strRet = strNumConv(i) & "." & strArray(UBound(strArray))
                
                If Dir(g_BMS.strDir & strRet, vbNormal) = vbNullString Then
                
                    blnRet = True '変換するよ
                
                Else
                
                    blnRet = False 'しないよ
                
                End If
                
                For j = i + 1 To UBound(g_strBMP)
                
                    If Not blnBMP(j) Then
                    
                        If g_strBMP(i) = g_strBMP(j) Then
                        
                            If blnRet Then g_strBMP(j) = strRet
                            
                            blnBMP(j) = True
                        
                        End If
                    
                    End If
                
                Next j
                
                If blnRet Then
                
                    'Name g_BMS.strDir & g_strBMP(i) As g_BMS.strDir & strRet
                    
                    ReDim Preserve strNameFrom(lngRet)
                    ReDim Preserve strNameTo(lngRet)
                    
                    strNameFrom(lngRet) = g_BMS.strDir & g_strBMP(i)
                    strNameTo(lngRet) = g_BMS.strDir & strRet
                    
                    lngRet = lngRet + 1
                    
                    g_strBMP(i) = strRet
                    
                    intRet = 1
                
                End If
                
                blnBMP(i) = True
            
            End If
        
        End If
    
    Next i
    
    If Len(Trim$(frmMain.txtMissBMP.Text)) Then
    
        If Dir(g_BMS.strDir & frmMain.txtMissBMP.Text, vbNormal) <> vbNullString Then
        
            strArray() = Split(frmMain.txtMissBMP.Text, ".")
            strRet = "00." & strArray(UBound(strArray))
            
            If Dir(g_BMS.strDir & strRet, vbNormal) = vbNullString Then
            
                'Name g_BMS.strDir & frmMain.txtMissBMP.Text As g_BMS.strDir & strRet
                
                ReDim Preserve strNameFrom(lngRet)
                ReDim Preserve strNameTo(lngRet)
                
                strNameFrom(lngRet) = g_BMS.strDir & frmMain.txtMissBMP.Text
                strNameTo(lngRet) = g_BMS.strDir & strRet
                
                lngRet = lngRet + 1
                
                frmMain.txtMissBMP.Text = strRet
                
                intRet = 1
            
            End If
        
        End If
    
    End If
    
    If lngRet Then
    
        With sh
        
            .hwnd = frmWindowConvert.hwnd
            .wFunc = FO_MOVE
            .pFrom = Join(strNameFrom(), Chr$(0))
            .pTo = Join(strNameTo(), Chr$(0))
            .fFlags = FOF_SILENT Or FOF_ALLOWUNDO Or FOF_MULTIDESTFILES
        
        End With
        
        Call SHFileOperation(sh)
    
    End If
    
    If intRet Then Call frmMain.SaveChanges

End Sub

Private Sub chkDeleteFile_Click()

    If chkDeleteFile.value Then
    
        lblExtension.Enabled = True
        txtExtension.Enabled = True
        chkFileRecycle.Enabled = True
    
    Else
    
        lblExtension.Enabled = False
        txtExtension.Enabled = False
        chkFileRecycle.Enabled = False
    
    End If

End Sub

Private Sub chkListAlign_Click()

    If chkListAlign.value Then
    
        chkUseOldFormat.Enabled = True
        chkSortByName.Enabled = True
    
    Else
    
        chkUseOldFormat.Enabled = False
        chkSortByName.Enabled = False
    
    End If

End Sub

Private Sub cmdCancel_Click()

    Call Unload(Me)

End Sub

Private Sub cmdDecide_Click()

    If chkDeleteUnusedFile.value = 0 And chkDeleteFile.value = 0 And chkListAlign.value = 0 And chkFileNameConvert.value = 0 Then Exit Sub
    
    'If chkFileNameConvert.Value Then
    
        'If MsgBox(g_Message(Message.MSG_CONFIRM), vbYesNo + vbInformation, g_strAppTitle) = vbNo Then Exit Sub
    
    'End If
    
    frmWindowConvert.Enabled = False
    
    If chkDeleteUnusedFile.value Then Call DeleteUnusedFile
    
    If chkDeleteFile.value Then Call DeleteFile
    
    If chkListAlign.value Then Call ListAlign
    
    If chkFileNameConvert.value Then Call FileNameConvert
    
    Call frmMain.RefreshList
    
    frmWindowConvert.Enabled = True
    
    Call Unload(Me)

End Sub

Private Sub Form_Activate()

    With txtExtension
    
        Call .Move(lblExtension.Left + lblExtension.Width + 60, .Top, frmWindowConvert.ScaleWidth - (lblExtension.Left + lblExtension.Width) - 180, .Height)
    
    End With
    
    If Len(g_BMS.strDir) = 0 Then
    
        chkDeleteFile.Enabled = False
        chkFileNameConvert.Enabled = False
    
    Else
    
        chkDeleteFile.Enabled = True
        chkFileNameConvert.Enabled = True
    
    End If
    
    chkDeleteUnusedFile.value = 0
    chkDeleteFile.value = 0
    chkFileRecycle.value = 0
    chkListAlign.value = 0
    chkUseOldFormat.value = 1
    chkFileNameConvert.value = 0
    
    Call cmdDecide.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    Call Me.Hide
    
    Call frmMain.picMain.SetFocus

End Sub
