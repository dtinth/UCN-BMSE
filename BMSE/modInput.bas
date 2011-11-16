Attribute VB_Name = "modInput"
Option Explicit

Private m_blnReadFlag   As Boolean
Private m_strEXInfo     As String

Private m_blnBGM()      As Boolean

Public Sub LoadBMS()
On Error GoTo Err:

    Dim i       As Long
    
    frmMain.Caption = g_strAppTitle & " - Now Loading"
    
    Call LoadBMSStart
    
    Call LoadBMSData
    
    ReDim Preserve g_Obj(UBound(g_Obj))

    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            .lngPosition = (g_Measure(.intMeasure).intLen / .lngHeight) * .lngPosition
            
            If .intCh = 3 Then 'BPM
            
                .intCh = 8
            
            ElseIf .intCh = 8 Then '拡張BPM
            
                If g_sngBPM(.sngValue) = 0 Then
                
                    .intCh = 0
                
                Else
                
                    .sngValue = g_sngBPM(.sngValue)
                
                End If
            
            ElseIf .intCh = 9 Then 'ストップシーケンス
            
                .sngValue = g_lngSTOP(.sngValue)
            
            End If
        
        End With
    
    Next i
    
    'Call QuickSort(0, UBound(g_Obj))
    
    Call LoadBMSEnd
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMS")
End Sub

Public Sub LoadBMSStart()
On Error GoTo Err:

    Dim i   As Long
    
    With frmMain
    
        For i = 0 To 1295
        
            g_strWAV(i) = ""
            g_strBMP(i) = ""
            g_strBGA(i) = ""
            g_sngBPM(i) = 0
            g_lngSTOP(i) = 0
        
        Next i
        
        .cboPlayer.ListIndex = 0
        .txtGenre.Text = ""
        .txtTitle.Text = ""
        .txtArtist.Text = ""
        .cboPlayLevel = 1
        .txtBPM.Text = "120"
        .cboPlayRank.ListIndex = 3
        .txtTotal.Text = ""
        .txtVolume.Text = ""
        .txtStageFile.Text = ""
        .txtMissBMP.Text = ""
        .lstWAV.ListIndex = 0
        .lstBMP.ListIndex = 0
        .lstBGA.ListIndex = 0
        .lstMeasureLen.ListIndex = 0
        .lstMeasureLen.Visible = False
        .txtExInfo.Text = ""
        .Enabled = False
        
        .vsbMain.value = 0
        .hsbMain.value = 0
        .cboVScroll.ListIndex = .cboVScroll.ListCount - 2
        
        For i = 0 To 999
        
            g_Measure(i).intLen = 192
            .lstMeasureLen.List(i) = "#" & Format$(i, "000") & ":4/4"
        
        Next i
    
    End With
    
    With g_BMS
    
        .intPlayerType = 1
        .strGenre = ""
        .strTitle = ""
        .strArtist = ""
        .sngBPM = 120
        .lngPlayLevel = 1
        .intPlayRank = 3
        .sngTotal = 0
        .intVolume = 0
        .strStageFile = ""
    
    End With
    
    'g_Disp.intMaxMeasure = 31
    g_disp.intMaxMeasure = 0
    Call modDraw.lngChangeMaxMeasure(15)
    Call modDraw.ChangeResolution
    
    'ReDim g_strInputLog(0)
    'g_lngInputLogPos = 0
    Call g_InputLog.Clear
    
    ReDim g_Obj(0)
    ReDim g_lngObjID(0)
    g_lngIDNum = 0
    
    m_blnReadFlag = True
    m_strEXInfo = ""
    
    ReDim m_blnBGM(31999)
    
    For i = 0 To UBound(m_blnBGM)
    
        m_blnBGM(i) = False
    
    Next i
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSStart")
End Sub

Public Sub LoadBMSEnd()
On Error GoTo Err:

    With frmMain
    
        Call modEasterEgg.LoadEffect
        
        Call frmMain.RefreshList
        
        .lstMeasureLen.Visible = True
        
        .Caption = g_strAppTitle
        
        If Len(g_BMS.strDir) Then
        
            If .mnuOptionsFileNameOnly.Checked Then
            
                .Caption = .Caption & " - " & g_BMS.strFileName
            
            Else
            
                .Caption = .Caption & " - " & g_BMS.strDir & g_BMS.strFileName
            
            End If
        
        End If
        
        Call modDraw.ChangeResolution
        
        .Enabled = True
        
        If UCase$(Right$(g_BMS.strFileName, 3)) = "PMS" Then
        
            .cboPlayer.ListIndex = 3
            g_BMS.intPlayerType = 4
        
        End If
        
        m_blnReadFlag = True
        .txtExInfo.Text = m_strEXInfo
        m_strEXInfo = ""
    
    End With
    
    Erase m_blnBGM()
    
    g_BMS.blnSaveFlag = True
    
    Call modDraw.InitVerticalLine
    
    Call frmMain.Show
    
    Call frmMain.picMain.SetFocus
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSEnd")
End Sub

Private Sub LoadBMSData()
On Error GoTo Err:

    Dim i           As Long
    Dim strArray()  As String
    Dim strRet      As String
    Dim lngFFile    As Long
    
    For i = 0 To 999
    
        g_Measure(i).intLen = 192
    
    Next i
    
    lngFFile = FreeFile()
    
    Open g_BMS.strDir & g_BMS.strFileName For Input As #lngFFile
    
        Do While Not EOF(lngFFile)
        
            DoEvents
            
            Line Input #lngFFile, strRet
            
            strArray = Split(strRet, Chr$(10))
            
            For i = 0 To UBound(strArray)
            
                If Left$(strArray(i), 1) = "#" Then Call LoadBMSDataSub(strArray(i))
            
            Next i
        
        Loop
    
    Close #lngFFile
    
    Exit Sub

Err:

    If Err.Number = 75 Then
    
        Call MsgBox(g_Message(ERR_FILE_NOT_FOUND) & vbCrLf & g_Message(ERR_LOAD_CANCEL), vbCritical, g_strAppTitle)
        
        Call LoadBMSStart
        Call LoadBMSEnd
        
        Exit Sub
    
    End If
    
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSData")

End Sub

Public Sub LoadBMSDataSub(ByVal strLineData As String, Optional ByVal blnDirectInput As Boolean)
On Error GoTo Err:

    Dim strArray()  As String
    Dim strRet      As String * 2
    Dim strParam    As String
    
    strArray() = Split(Replace(strLineData, " ", ":", 1, 1), ":")
    
    With frmMain
    
        If UBound(strArray) > 0 Then
        
            strParam = Right$(strLineData, Len(strLineData) - (Len(strArray(0)) + 1))
            
            Select Case UCase$(strArray(0))
            
                'Case "#PATH_WAV"
                    
                    'g_BMS.strDir = strParam
                
                Case "#PLAYER"
                
                    g_BMS.intPlayerType = Val(strParam)
                    .cboPlayer.ListIndex = Val(strParam) - 1
                
                Case "#GENRE"
                
                    g_BMS.strGenre = strParam
                    .txtGenre.Text = strParam
                
                Case "#TITLE"
                    g_BMS.strTitle = strParam
                    .txtTitle.Text = strParam
                
                Case "#ARTIST"
                
                    g_BMS.strArtist = strParam
                    .txtArtist.Text = strParam
                
                Case "#BPM"
                
                    g_BMS.sngBPM = Val(strParam)
                    .txtBPM.Text = Val(strParam)
                
                Case "#PLAYLEVEL"
                
                    g_BMS.lngPlayLevel = Val(strParam)
                    .cboPlayLevel = Val(strParam)
                
                Case "#RANK"
                
                    g_BMS.intPlayRank = Val(strParam)
                    
                    If g_BMS.intPlayRank < 0 Then g_BMS.intPlayRank = 0
                    
                    If g_BMS.intPlayRank > 3 Then g_BMS.intPlayRank = 3
                    
                    .cboPlayRank.ListIndex = g_BMS.intPlayRank
                
                Case "#TOTAL"
                
                    g_BMS.sngTotal = Val(strParam)
                    .txtTotal.Text = Val(strParam)
                
                Case "#VOLWAV"
                
                    g_BMS.intVolume = Val(strParam)
                    .txtVolume.Text = Val(strParam)
                
                Case "#STAGEFILE"
                
                    g_BMS.strStageFile = strParam
                    .txtStageFile.Text = strParam
                
                Case "#IF", "#RANDOM", "#RONDAM", "#ENDIF"
                
                    If blnDirectInput = False Then
                    
                        m_blnReadFlag = False
                        
                        m_strEXInfo = m_strEXInfo & strLineData & vbCrLf
                    
                    End If
                
                Case Else
                
                    strRet = UCase$(Right$(strArray(0), 2))
                    
                    Select Case UCase$(Left$(strArray(0), 4))
                    
                        Case "#WAV"
                        
                            If strRet <> "00" And blnDirectInput = False Then
                            
                                g_strWAV(lngNumConv(strRet)) = Right$(strLineData, Len(strLineData) - 7)
                                
                                If Asc(Left$(strRet, 1)) > Asc("F") Or Asc(Right$(strRet, 1)) > Asc("F") Then
                                
                                    .mnuOptionsNumFF.Checked = False
                                
                                End If
                            
                            End If
                        
                        Case "#BMP"
                        
                            If strRet <> "00" And blnDirectInput = False Then
                            
                                g_strBMP(lngNumConv(strRet)) = Right$(strLineData, Len(strLineData) - 7)
                                
                                If Asc(Left$(strRet, 1)) > Asc("F") Or Asc(Right$(strRet, 1)) > Asc("F") Then
                                
                                    .mnuOptionsNumFF.Checked = False
                                
                                End If
                            
                            Else
                            
                                .txtMissBMP.Text = Right$(strLineData, Len(strLineData) - 7)
                            
                            End If
                        
                        Case "#BGA"
                        
                            If strRet <> "00" And blnDirectInput = False Then
                            
                                g_strBGA(lngNumConv(strRet)) = Right$(strLineData, Len(strLineData) - 7)
                                
                                If Asc(Left$(strRet, 1)) > Asc("F") Or Asc(Right$(strRet, 1)) > Asc("F") Then
                                
                                    .mnuOptionsNumFF.Checked = False
                                
                                End If
                            
                            End If
                        
                        Case "#BPM"
                        
                            If strRet <> "00" And blnDirectInput = False Then
                            
                                g_sngBPM(lngNumConv(strRet)) = Right$(strLineData, Len(strLineData) - 7)
                            
                            End If
                        
                        Case Else
                        
                            If UCase$(Left$(strArray(0), 5)) = "#STOP" Then
                            
                                If strRet <> "00" And blnDirectInput = False Then
                                
                                    g_lngSTOP(lngNumConv(strRet)) = Right$(strLineData, Len(strLineData) - 8)
                                
                                End If
                            
                            ElseIf IsNumeric(Mid$(strArray(0), 2)) Then
                            
                                If m_blnReadFlag Then
                                
                                    Call LoadBMSObject(strLineData)
                                
                                Else
                                
                                    m_strEXInfo = m_strEXInfo & strLineData & vbCrLf
                                
                                End If
                            
                            Else
                            
                                m_strEXInfo = m_strEXInfo & strLineData & vbCrLf
                            
                            End If
                    
                    End Select
            
            End Select
        
        ElseIf UCase$(Left$(strLineData, 6)) = "#ENDIF" Then
        
            m_blnReadFlag = True
            
            m_strEXInfo = m_strEXInfo & strLineData & vbCrLf
        
        Else
        
            m_strEXInfo = m_strEXInfo & strLineData & vbCrLf
        
        End If
    
    End With
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSDataSub")
End Sub

Private Sub LoadBMSObject(ByVal strRet As String)
On Error GoTo Err:

    Dim i           As Long
    Dim j           As Long
    Dim intRet      As Integer
    Dim intMeasure  As Integer
    Dim intCh       As Integer
    Dim strParam    As String
    Dim lngSepaNum  As Long

    intMeasure = Val(Mid$(strRet, 2, 3))
    intCh = Val(Mid$(strRet, 5, 2))
    strParam = UCase$(Trim$(strGetParam(strRet)))
    
    lngSepaNum = Len(strParam) \ 2
    
    If intCh = 2 Then
    
        If Val(strParam) = 0 Or Val(strParam) = 1 Then Exit Sub
        
        intRet = intGCD(Int(192 * Val(strParam)), 192)
        
        If intRet <= 2 Then intRet = 3
        
        If intRet >= 48 Then intRet = 48
        
        With g_Measure(intMeasure)
        
            .intLen = Int(192 * Val(strParam))
            
            If .intLen < 3 Then .intLen = 3
            
            Do While .intLen \ intRet > 64
            
                If intRet >= 48 Then
                
                    .intLen = 3072
                    
                    Exit Do
                
                End If
                
                intRet = intRet * 2
            
            Loop
            
            frmMain.lstMeasureLen.List(intMeasure) = "#" & Format$(intMeasure, "000") & ":" & (.intLen \ intRet) & "/" & (192 \ intRet)
        
        End With
        
        Exit Sub
    
    End If
    
    If intCh = 1 Then
    
        For j = 0 To 31
        
            If m_blnBGM(intMeasure * 32 + j) = False Then
            
                m_blnBGM(intMeasure * 32 + j) = True
                intRet = 101 + j
                
                Exit For
            
            End If
        
        Next j
    
    End If
    
    For i = 1 To lngSepaNum
    
        If Mid$(strParam, i * 2 - 1, 2) <> "00" Then
        
            With g_Obj(UBound(g_Obj))
            
                .lngID = g_lngIDNum
                g_lngObjID(g_lngIDNum) = g_lngIDNum
                .lngPosition = i - 1
                .lngHeight = lngSepaNum
                .intMeasure = intMeasure
                .intCh = intCh
                
                Call modDraw.lngChangeMaxMeasure(.intMeasure)
                
                Select Case intCh
                
                    Case 1 'BGM
                    
                        .sngValue = lngNumConv(Mid$(strParam, i * 2 - 1, 2))
                        .intCh = intRet
                    
                    Case 4, 6, 7, 8, 9  'BGA,Poor,Layer,拡張BPM,ストップシーケンス
                    
                        .sngValue = lngNumConv(Mid$(strParam, i * 2 - 1, 2))
                    
                    Case 3 'BPM
                    
                        .sngValue = Val("&H" + Mid$(strParam, i * 2 - 1, 2))
                    
                    Case 11 To 16, 18, 19, 21 To 26, 28, 29 'キー音
                    
                        .sngValue = lngNumConv(Mid$(strParam, i * 2 - 1, 2))
                    
                    Case 31 To 36, 38, 39, 41 To 46, 48, 49 'キー音
                    
                        .sngValue = lngNumConv(Mid$(strParam, i * 2 - 1, 2))
                        .intCh = .intCh - 20
                        .intAtt = 1
                    
                    Case 51 To 56, 58, 59, 61 To 66, 68, 69  'キー音
                    
                        .sngValue = lngNumConv(Mid$(strParam, i * 2 - 1, 2))
                        .intCh = .intCh - 40
                        .intAtt = 2
                    
                    Case Else
                    
                        Exit Sub
                
                End Select
            
            End With
            
            ReDim Preserve g_Obj(UBound(g_Obj) + 1)
            
            g_lngIDNum = g_lngIDNum + 1
            ReDim Preserve g_lngObjID(g_lngIDNum)
        
        End If
            
    Next i
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "LoadBMSObject")
End Sub

Public Sub QuickSort(ByVal lngLeft As Long, ByVal lngRight As Long)

    Dim i   As Long
    Dim j   As Long

    If lngLeft >= lngRight Then Exit Sub
    
    i = lngLeft + 1
    j = lngRight
    
    Do While i <= j
    
        Do While i <= j
        
            If g_Obj(i).intMeasure > g_Obj(lngLeft).intMeasure Then
                Exit Do
            End If
            
            i = i + 1
        Loop
        
        Do While i <= j
        
            If g_Obj(j).intMeasure < g_Obj(lngLeft).intMeasure Then
                Exit Do
            End If
            
            j = j - 1
        
        Loop

        If i >= j Then Exit Do

        Call SwapObj(j, i)

        i = i + 1
        j = j - 1
    
    Loop

    Call SwapObj(j, lngLeft)
    Call QuickSort(lngLeft, j - 1)
    Call QuickSort(j + 1, lngRight)

End Sub

Public Sub SwapObj(ByVal Obj1Num As Long, ByVal Obj2Num As Long)

    Dim dummyObj    As g_udtObj
    
    With dummyObj
    
        .lngID = g_Obj(Obj1Num).lngID
        .intCh = g_Obj(Obj1Num).intCh
        .sngValue = g_Obj(Obj1Num).sngValue
        .intMeasure = g_Obj(Obj1Num).intMeasure
        .lngPosition = g_Obj(Obj1Num).lngPosition
        .lngHeight = g_Obj(Obj1Num).lngHeight
        .intSelect = g_Obj(Obj1Num).intSelect
        .intAtt = g_Obj(Obj1Num).intAtt
    
    End With
    
    With g_Obj(Obj1Num)
    
        g_lngObjID(.lngID) = Obj2Num
        .lngID = g_Obj(Obj2Num).lngID
        .intCh = g_Obj(Obj2Num).intCh
        .sngValue = g_Obj(Obj2Num).sngValue
        .intMeasure = g_Obj(Obj2Num).intMeasure
        .lngPosition = g_Obj(Obj2Num).lngPosition
        .lngHeight = g_Obj(Obj2Num).lngHeight
        .intSelect = g_Obj(Obj2Num).intSelect
        .intAtt = g_Obj(Obj2Num).intAtt
    
    End With
    
    With g_Obj(Obj2Num)
    
        g_lngObjID(.lngID) = Obj1Num
        .lngID = dummyObj.lngID
        .intCh = dummyObj.intCh
        .sngValue = dummyObj.sngValue
        .intMeasure = dummyObj.intMeasure
        .lngPosition = dummyObj.lngPosition
        .lngHeight = dummyObj.lngHeight
        .intSelect = dummyObj.intSelect
        .intAtt = dummyObj.intAtt
    
    End With

End Sub

Public Function lngNumConv(ByVal strNum As String) As Long

    Dim i       As Long
    Dim lngRet  As Long
    
    For i = 1 To Len(strNum)
    
        lngRet = lngRet + lngSubNumConv(Mid$(strNum, i, 1)) * (36 ^ (Len(strNum) - i))
    
    Next i
    
    lngNumConv = lngRet

End Function

Public Function lngSubNumConv(ByVal b As String) As Long

    Dim R   As Long
    
    R = Abs(Asc(UCase$(b)))
    
    If R >= 65 And R <= 90 Then 'A-Z
    
        lngSubNumConv = R - 55
    
    Else
    
        lngSubNumConv = (R - 48) Mod 36
    
    End If

End Function

Public Function strNumConv(ByVal lngNum As Long, Optional ByVal Length As Long = 2) As String

    Dim strRet  As String
    
    Do While lngNum
    
        strRet = strSubNumConv(lngNum Mod 36) & strRet
        lngNum = lngNum \ 36
    
    Loop
    
    Do While Len(strRet) < Length
    
        strRet = "0" & strRet
    
    Loop
    
    strNumConv = Right$(strRet, Length)

End Function

Public Function strSubNumConv(ByVal b As Long) As String

    Select Case b
    
        Case 0 To 9
        
            strSubNumConv = b
        
        Case Else
        
            strSubNumConv = Chr$(b + 55)
    
    End Select

End Function

Private Function strGetParam(ByVal strRet As String) As String
    
    Dim strArray()  As String
    
    strArray() = Split(strRet, ":")

    If UBound(strArray) > 0 Then
    
        strGetParam = strArray(UBound(strArray))
    
    Else
    
        strGetParam = ""
    
    End If

End Function

Public Function intGCD(ByVal m As Integer, ByVal n As Integer) As Integer

    If m <= 0 Or n <= 0 Then Exit Function
    
    If m Mod n = 0 Then
    
        intGCD = n
    
    Else
    
        intGCD = intGCD(n, m Mod n)
    
    End If
    
End Function
    
