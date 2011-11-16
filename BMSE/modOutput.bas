Attribute VB_Name = "modOutput"
Option Explicit

Public Sub CreateBMS(ByVal strOutputPath As String, Optional ByVal Flag As Integer)
On Error GoTo Err:

    Dim strObjData()    As String
    Dim blnObjData()    As Boolean
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim lngFFile        As Long
    Dim intBPMNum       As Integer
    Dim intSTOPNum      As Integer
    Dim lngRet          As Long
    Dim strRet          As String
    Dim intArray()      As Integer
    
    If Flag = 0 Then frmMain.Caption = g_strAppTitle & " - Now Saving..."
    
    frmMain.Enabled = False
    
    For i = 0 To 1295
    
        g_sngBPM(i) = 0
        g_lngSTOP(i) = 0
    
    Next i
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intCh Then
            
                If lngRet < .intMeasure Then
                    
                    lngRet = .intMeasure
                
                End If
                
                Select Case .intCh
                
                    Case 8
                
                        If .sngValue > 0 And .sngValue < 256 And .sngValue = CLng(.sngValue) Then
                        
                            .intCh = 3
                        
                        Else
                        
                            If intBPMNum > 1295 Then
                            
                                Call MsgBox(g_Message(ERR_OVERFLOW_BPM) & vbCrLf & g_Message(ERR_SAVE_CANCEL), vbCritical, g_strAppTitle)
                                
                                lngRet = i - 1
                                
                                GoTo Init:
                            
                            End If
                            
                            intBPMNum = intBPMNum + 1
                            g_sngBPM(intBPMNum) = .sngValue
                            .sngValue = intBPMNum
                        
                        End If
                
                    Case 9
                
                        If intSTOPNum > 1295 Then
                        
                            Call MsgBox(g_Message(ERR_OVERFLOW_STOP) & vbCrLf & g_Message(ERR_SAVE_CANCEL), vbCritical, g_strAppTitle)
                            
                            lngRet = i - 1
                            
                            GoTo Init:
                        
                        End If
                        
                        intSTOPNum = intSTOPNum + 1
                        g_lngSTOP(intSTOPNum) = .sngValue
                        .sngValue = intSTOPNum
                    
                    Case 11 To 29
                    
                        If .intAtt = 1 Then
                        
                            .intCh = .intCh + 20
                        
                        ElseIf .intAtt = 2 Then
                        
                            .intCh = .intCh + 40
                        
                        End If
                
                End Select
            
            End If
        
        End With
    
    Next i
    
    ReDim strObjData(132, lngRet)
    ReDim blnObjData(132, lngRet)
    
    For i = 0 To lngRet
    
        For j = 3 To 132
        
            strObjData(j, i) = String$(g_Measure(i).intLen * 2, "0")
        
        Next j
    
    Next i
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            Select Case .intCh
            
                Case Is < 0
                
                Case Is > 1000
                
                Case Is > 100
                
                    strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & modInput.strNumConv(.sngValue) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                    
                    For j = 101 To .intCh - 1
                    
                        blnObjData(j, .intMeasure) = True
                    
                    Next j
                
                Case 3
                
                    strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right$("0" & Hex$(.sngValue), 2) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                
                Case 8
                
                    If intBPMNum > 255 Then
                    
                        strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right$("0" & modInput.strNumConv(.sngValue), 2) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                    
                    Else
                    
                        strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right$("0" & Hex$(.sngValue), 2) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                    
                    End If
                
                Case 9
                
                    If intSTOPNum > 255 Then
                    
                        strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right$("0" & modInput.strNumConv(.sngValue), 2) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                    
                    Else
                    
                        strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & Right$("0" & Hex$(.sngValue), 2) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                    
                    End If
                
                Case Else
                
                    strObjData(.intCh, .intMeasure) = Left$(strObjData(.intCh, .intMeasure), .lngPosition * 2) & modInput.strNumConv(.sngValue) & Mid$(strObjData(.intCh, .intMeasure), .lngPosition * 2 + 3)
                
            End Select
            
            blnObjData(.intCh, .intMeasure) = True
        
        End With
    
    Next i
    
    For i = 0 To UBound(strObjData, 2)
    
        For j = 3 To 132
        
            If blnObjData(j, i) Then
            
                If strObjData(j, i) <> "00" Then
                
                    ReDim intArray(g_Measure(i).intLen + 1)
                    
                    intArray(0) = g_Measure(i).intLen
                    strRet = ""
                    lngRet = 1
                    
                    For k = 1 To Len(strObjData(j, i)) \ 2
                    
                        If Mid$(strObjData(j, i), k * 2 - 1, 2) = "00" Then
                        
                            strRet = strRet & "0"
                        
                        Else
                        
                            intArray(lngRet) = Len(strRet)
                            lngRet = lngRet + 1
                            strRet = "1"
                        
                        End If
                    
                    Next k
                    
                    ReDim Preserve intArray(lngRet)
                    
                    intArray(lngRet) = Len(strRet)
                    
                    lngRet = intGetMaxDev(intArray)
                    
                    If lngRet Then
                    
                        strRet = ""
                        
                        For k = 1 To Len(strObjData(j, i)) \ 2 Step lngRet
                        
                            strRet = strRet & Mid$(strObjData(j, i), k * 2 - 1, 2)
                        
                        Next k
                        
                        strObjData(j, i) = strRet
                    
                    End If
                
                End If
                
            End If
        
        Next j
    
    Next i
    
    lngFFile = FreeFile()
    
    Open strOutputPath For Output As #lngFFile
    
        With frmMain
        
            Print #lngFFile,
            Print #lngFFile, "*---------------------- HEADER FIELD"
            Print #lngFFile,
            'If Flag Then Print #lngFFile, "#PATH_WAV " & g_BMS.strDir
            
            If .cboPlayer.ListIndex > 1 Then
            
                Print #lngFFile, "#PLAYER 3"
            
            Else
            
                Print #lngFFile, "#PLAYER " & .cboPlayer.ListIndex + 1
            
            End If
            
            Print #lngFFile, "#GENRE " & Trim$(.txtGenre.Text)
            Print #lngFFile, "#TITLE " & Trim$(.txtTitle.Text)
            Print #lngFFile, "#ARTIST " & Trim$(.txtArtist.Text)
            Print #lngFFile, "#BPM " & Trim$(.txtBPM.Text)
            Print #lngFFile, "#PLAYLEVEL " & Trim$(.cboPlayLevel)
            Print #lngFFile, "#RANK " & .cboPlayRank.ListIndex
            
            If Val(.txtTotal.Text) Then Print #lngFFile, "#TOTAL " & .txtTotal.Text
            
            If Val(.txtVolume.Text) Then Print #lngFFile, "#VOLWAV " & .txtVolume.Text
            
            Print #lngFFile, "#STAGEFILE " & Trim$(.txtStageFile.Text)
            Print #lngFFile,
            
            For i = 1 To 1295
            
                If Len(g_strWAV(i)) Then
                
                    Print #lngFFile, "#WAV" & modInput.strNumConv(i) & " " & g_strWAV(i)
                
                End If
            
            Next i
            
            Print #lngFFile,
            
            If Len(Trim$(.txtMissBMP.Text)) Then
            
                Print #lngFFile, "#BMP00 " & .txtMissBMP.Text
            
            End If
            
            For i = 1 To 1295
            
                If Len(g_strBMP(i)) Then
                
                    Print #lngFFile, "#BMP" & modInput.strNumConv(i) & " " & g_strBMP(i)
                
                End If
            
            Next i
            
            Print #lngFFile,
            
            For i = 1 To 1295
            
                If Len(g_strBGA(i)) Then
                
                    Print #lngFFile, "#BGA" & modInput.strNumConv(i) & " " & g_strBGA(i)
                
                End If
            
            Next i
            
            Print #lngFFile,
            
            If intBPMNum > 255 Then
            
                For i = 1 To 1295
                
                    If g_sngBPM(i) Then
                    
                        Print #lngFFile, "#BPM" & Right$("0" & modInput.strNumConv(i), 2) & " " & g_sngBPM(i)
                    
                    End If
                
                Next i
            
            ElseIf intBPMNum Then
            
                For i = 1 To 255
                
                    If g_sngBPM(i) Then
                    
                        Print #lngFFile, "#BPM" & Right$("0" & Hex$(i), 2) & " " & g_sngBPM(i)
                    
                    End If
                
                Next i
            
            End If
            
            Print #lngFFile,
            
            If intSTOPNum > 255 Then
            
                For i = 1 To 1295
                
                    If g_lngSTOP(i) Then
                    
                        Print #lngFFile, "#STOP" & Right$("0" & modInput.strNumConv(i), 2) & " " & g_lngSTOP(i)
                    
                    End If
                
                Next i
            
            ElseIf intSTOPNum Then
            
                For i = 1 To 255
                
                    If g_lngSTOP(i) Then
                    
                        Print #lngFFile, "#STOP" & Right$("0" & Hex$(i), 2) & " " & g_lngSTOP(i)
                    
                    End If
                
                Next i
            
            End If
            
            Print #lngFFile,
            
            Print #lngFFile, .txtExInfo.Text
            
            Print #lngFFile,
        
        End With
        
        Print #lngFFile,
        Print #lngFFile, "*---------------------- MAIN DATA FIELD"
        Print #lngFFile,
        
        For i = 0 To UBound(blnObjData, 2)
        
            For j = 101 To 132
            
                If blnObjData(j, i) Then
                
                    Print #lngFFile, "#" & Format$(i, "000") & "01" & ":" & strObjData(j, i)
                
                End If
            
            Next j
            
            With g_Measure(i)
                
                If .intLen <> 192 Then
                
                    Print #lngFFile, "#" & Format$(i, "000") & "02:" & .intLen / 192
                
                End If
            
            End With
            
            For j = 3 To 99
            
                If blnObjData(j, i) Then
                
                    Print #lngFFile, "#" & Format$(i, "000") & Format$(j, "00") & ":" & strObjData(j, i)
                
                End If
            
            Next j
            
            Print #lngFFile,
        
        Next i
        
        lngRet = UBound(blnObjData, 2) + 1
        
        For i = lngRet To 999
        
            With g_Measure(i)
                
                If .intLen <> 192 Then
                
                    Print #lngFFile, "#" & Format$(i, "000") & "02:" & .intLen / 192
                
                End If
            
            End With
        
        Next i
    
    lngRet = UBound(g_Obj) - 1
    
    With g_BMS
    
        .intPlayerType = frmMain.cboPlayer.ListIndex + 1
        .strGenre = frmMain.txtGenre.Text
        .strTitle = frmMain.txtTitle.Text
        .strArtist = frmMain.txtArtist.Text
        .lngPlayLevel = Val(frmMain.cboPlayLevel.Text)
        .sngBPM = Val(frmMain.txtBPM.Text)
        
        .intPlayRank = frmMain.cboPlayRank.ListIndex
        .sngTotal = Val(frmMain.txtTotal.Text)
        .intVolume = Val(frmMain.txtVolume.Text)
        .strStageFile = frmMain.txtStageFile.Text
    
    End With
    
Init:

    Close #lngFFile
    
    For i = 0 To lngRet
    
        With g_Obj(i)
        
            Select Case .intCh
            
                Case 3
                
                    .intCh = 8
                
                Case 8
                
                    .sngValue = g_sngBPM(.sngValue)
                
                Case 9
                
                    .sngValue = g_lngSTOP(.sngValue)
                
                Case 31 To 49
                
                    .intCh = .intCh - 20
                
                Case 51 To 69
                
                    .intCh = .intCh - 40
            
            End Select
        
        End With
    
    Next i
    
    frmMain.Enabled = True
    
    If Flag = 0 Then
    
        g_BMS.blnSaveFlag = True
        
        If Len(g_BMS.strDir) Then
        
            If frmMain.mnuOptionsFileNameOnly.Checked Then
            
                frmMain.Caption = g_strAppTitle & " - " & g_BMS.strFileName
            
            Else
            
                frmMain.Caption = g_strAppTitle & " - " & g_BMS.strDir & g_BMS.strFileName
            
            End If
        
        End If
    
    End If
    
    Exit Sub

Err:
    Call MsgBox(g_Message(ERR_SAVE_ERROR) & vbCrLf & g_Message(ERR_SAVE_CANCEL) & vbCrLf & "Error No." & Err.Number & " " & Err.Description, vbCritical, g_strAppTitle)
    frmMain.Enabled = True
    frmMain.Caption = g_strAppTitle & " - " & g_BMS.strDir & g_BMS.strFileName
End Sub

Private Function intGetMaxDev(ByRef BaseValue() As Integer) As Integer

    Dim Count As Long        '配列の最大インデックス
    Dim i As Long            'カウンタ
    Dim a As Long, b As Long '最大公約数を求める2つの要素
    
    Count = UBound(BaseValue)
    a = BaseValue(0)

    '繰り返す回数は、(配列の数－1)回
    For i = 1 To Count
    
        b = BaseValue(i)
        
        If b Then
        
            Do While a <> b
            
                If a > b Then
                
                    a = a - b
                
                Else
                
                    b = b - a
                
                End If
            
            Loop
            
            '1で等しい場合、最大公約数はない
            If a = 1& Then intGetMaxDev = 0&: Exit Function
        
        End If
    
    Next i
    
    '最大公約数を返す
    intGetMaxDev = a

End Function
