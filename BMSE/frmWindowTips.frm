VERSION 5.00
Begin VB.Form frmWindowTips 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  '�Œ�°� ����޳
   Caption         =   "BMSE Tips (Sorry Japanese Language Only!!!!!!!111)"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   FillStyle       =   0  '�h��Ԃ�
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  '�Ȃ�
      Height          =   960
      Left            =   0
      Picture         =   "frmWindowTips.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   0
   End
   Begin VB.CheckBox chkNextDisp 
      Caption         =   "Launch at next startup"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3540
      Value           =   1  '����
      Width           =   2775
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "����"
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   3480
      Width           =   1515
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "����"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frmWindowTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_WORDBREAK = &H10

Dim m_sngTwipsX     As Single
Dim m_sngTwipsY     As Single

Dim m_strTips()     As String
Dim m_intTipsPos    As Integer
Dim m_lngTipsNum    As Long

Private Sub chkNextDisp_Click()

    Dim i       As Long
    Dim lngRet  As Long
    Dim lngArg  As Long
    
    For i = 0 To frmMain.mnuLanguage.UBound
    
        If frmMain.mnuLanguage(i).Checked Then
        
            If g_strLangFileName(i) <> "japanese.ini" Then
            
                Exit Sub
            
            End If
            
            Exit For
        
        End If
    
    Next i
    
    If chkNextDisp.Value = 0 Then
    
        lngRet = vbRetry
        Call Randomize
        
        Do While lngRet = vbRetry
        
            If Int(Rnd * 256) = 0 Then
            
                Call MsgBox("�悭�킩��Ȃ����Ǒ����G���[���������܂����B" & vbCrLf & "����� Tips ��\�����܂��B", vbCritical Or vbOKOnly, g_strAppTitle)
                
                chkNextDisp.Value = 1
                chkNextDisp.Enabled = False
                
                Exit Do
            
            End If
            
            lngRet = Int(Rnd * 32) + 1
            
            If lngRet Mod 32 = 0 Then
            
                lngArg = vbExclamation
            
            ElseIf lngRet Mod 16 = 0 Then
            
                lngArg = vbInformation
            
            ElseIf lngRet Mod 8 = 0 Then
            
                lngArg = vbCritical
            
            Else
            
                lngArg = vbQuestion
            
            End If
            
            If Int(Rnd * 64) = 0 Then
            
                lngArg = lngArg Or vbMsgBoxRight
            
            End If
            
            If Int(Rnd * 128) = 0 Then
            
                lngArg = lngArg Or vbMsgBoxRtlReading
            
            End If
            
            lngRet = MsgBox("�{���ɁH", vbAbortRetryIgnore Or lngArg, g_strAppTitle)
        
        Loop
        
        Select Case lngRet
        
            Case vbAbort
            
                chkNextDisp.Value = 1
        
        End Select
    
    End If

End Sub

Private Sub cmdClose_Click()

    Call Unload(Me)

End Sub

Private Sub cmdNext_Click()

    m_intTipsPos = m_intTipsPos + 1
    
    If m_intTipsPos > UBound(m_strTips) Then m_intTipsPos = 1

    frmWindowTips.Line (5400, 360)-Step(180, 150), vbWhite, BF
    
    With frmWindowTips
    
        .Font.Size = 9
        .CurrentX = 5400
        .CurrentY = 345
        
        frmWindowTips.Print Right$(" " & m_intTipsPos, 2)
        
        .Font.Size = 12
    
    End With
    
    frmWindowTips.Line (945, 720)-(6030, 3240), vbWhite, BF
    
    Call BitBlt(frmWindowTips.hdc, 240 * m_sngTwipsX / Screen.TwipsPerPixelX, 240 * m_sngTwipsY / Screen.TwipsPerPixelY, 32, 32, picIcon.hdc, 0, 32, SRCCOPY)
    
    m_lngTipsNum = 0

End Sub

Private Sub Form_Activate()

    With Me
    
        .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
        .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
    
    End With
    
    m_sngTwipsX = 15 / Screen.TwipsPerPixelX
    m_sngTwipsY = 15 / Screen.TwipsPerPixelY
    
    m_intTipsPos = 0
    
    ReDim m_strTips(0)
    
    m_strTips(0) = " ���ꂩ�� Tips ��\�����܂��B" & vbCrLf & vbCrLf & " �����̏��͂��Ȃ��� BMSE ���g�� BMS ���쐬����̂��菕�����Ă���邱�Ƃ����邩������܂���B" & vbCrLf & vbCrLf & " �u���ցv�̃{�^���������� Tips ���J�n���Ă��������B" & vbCrLf & vbCrLf & " (���̕��͈͂�x�����\������܂���)"
    
    Call AddTutorial(" BMSE �� UCN-Soft ���J�����Ă��܂��B" & vbCrLf & vbCrLf & " UCN �̗R���͒����փl�^�Ȃ̂œ����ł��I")
    Call AddTutorial(" BMSE �� BMx Sequence Editor �̗��ł��B�m��Ȃ��F�B��������L�߂悤�I")
    Call AddTutorial(" BMSE �� bms �t�@�C���Abme �t�@�C���Abml �t�@�C������� pms �t�@�C���������o�����Ƃ��ł��܂��B")
    Call AddTutorial(" bms �̐������̂� Be-Music Script �ȂǏ�������܂����A�^���͓�̂܂܂ł��B")
    Call AddTutorial(" BMSE ���g�p����ɂ́A�܂� Windows OS �̑���ɏK�n����K�v������܂��B" & vbCrLf & vbCrLf & " �}�E�X�͕Ў�Ŏ����A��ʏ�̃|�C���^�𑀍삵�܂��B�f�B�X�v���C���w�łȂ���킯�ł͂���܂���B")
    Call AddTutorial(" �I�u�W�F��z�u����ɂ̓X�N���[�������N���b�N���܂��B" & vbCrLf & vbCrLf & " ���N���b�N�̎d���ɂ��ẮA���g���� OS �̃}�j���A�������ǂ݂��������B" & vbCrLf & vbCrLf & " (BMSE �̓}�E�X���K�{�ł�)")
    Call AddTutorial(" �I�u�W�F���z�u�ł��Ȃ��H�����S���c�[���ɂȂ��Ă��܂��񂩁H")
    Call AddTutorial(" �E���ɕ\������Ă���e�L�X�g�E�{�b�N�X�ɂ͔C�ӂ̕��������͂��܂��B" & vbCrLf & vbCrLf & " ���������͂���ɂ̓L�[�{�[�h���K�v�ł��̂ŁA���g���� OS �y�ь���c�[���̃}�j���A�������ǂ݂��������B")
    Call AddTutorial(" GENRE �́u�W�������v�Ɠǂ݁A�I�Ȓ��ɕ\������邨���܂��ȋȂ̌X������͂��܂��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ����� Techno �ɂ��Ă��������B")
    Call AddTutorial(" bpm �� Beat Per Minute �̗��ŁA1��������̃r�[�g������͂��܂��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ�����400�ɂ��Ă��������B")
    Call AddTutorial(" TITLE �́u�^�C�g���v�Ɠǂ݂܂��B�p��Łu�薼�v���Ӗ����A�I�Ȓ��ɕ\�������Ȃ̑薼����͂��܂��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ����͉p�a�����������Ă������� (�p�a�����͂��߂��̏��X�ōw���\�ł�)�B")
    Call AddTutorial(" ARTIST �͒��󂷂�Ɓu�|�p�Ɓv�ƂȂ�܂����A�����ł́u��ҁv����͂��Ă��������B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ����́uDJ �c���v�Ƃ��Ă��������B��: DJ �R�c")
    Call AddTutorial(" PLAYLEVEL �́u���ʂ̓�Փx�v�ł��B�������� 1 �` 7 �� bms �� �f�t�@�N�g�X�^���_�[�h�ł��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ����̓m�[�g����100�ɂ��Ă��������B")
    Call AddTutorial(" �u��{�v�^�u�ׂ̗Ɂu�g���v�^�u����сu���v�^�u�����邱�Ƃɂ��C�Â��ł����H" & vbCrLf & vbCrLf & " �N���b�N����ΐV���Ȑݒ���s�����Ƃ��\�ɂȂ�܂��B")
    Call AddTutorial(" RANK �͒��󂵂Ă��Ӗ����ʂ��܂���B�u����̌������v�������܂��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ����� VERY HARD �ɂ��Ă��������B")
    Call AddTutorial(" ���� BMSE �� MOD �ɑΉ����Ă��܂� (���݉B���R�}���h)�B" & vbCrLf & vbCrLf & " ���̐��ǂނɂ̓V�F�A�E�G�A�t�B�[�𕥂��K�v������܂��B" & vbCrLf & vbCrLf & " ���̃\�t�g�E�F�A�͑���E�F�A�ł��B�C�ɓ��������҂ɑ������t���Ă��������B")
    Call AddTutorial(" �e���L�[�������ƁA�r���E�Q�C�c�ƃ��b�Z���W���[�Ń`���b�g���ł��܂��B")
    Call AddTutorial(" �X�N���[���̈�ԍ��ɂ���uBPM�v����сuSTOP�v���[���ɒ��ڂ��Ă��������I" & vbCrLf & vbCrLf & " ���̃��[�����N���b�N���A�P���ɔ��p�p�� (�L�[�{�[�h�̉E�[�ɂ��鋷�������݂̗̂̈���������Ă�������) ����͂��邾���ŁA�v���C���[��|�M���邱�Ƃ��ł��܂��B")
    Call AddTutorial(" BMSE �̓}�E�}�j�ɑΉ����Ă��܂���B�{������I")
    Call AddTutorial(" �Ă��Ƃ葁�� bms �����ɂ́Awav ���g�p�����ɍ��̂���Ԃł��B" & vbCrLf & vbCrLf & " �G��`�����o�ŃX�N���[���ɃI�u�W�F��z�u (���N���b�N) ����� bms �������I�ȒP�ł���H")
    Call AddTutorial(" �u��{�v�^�u�̈�ԏ�ɂ���u�v���C���[�h�v�� Double Play �ɂ��Ă݂܂��傤�B���Ղ̐����{�����A���u�����v���ʂ���邱�Ƃ��ł��܂��B" & vbCrLf & vbCrLf & " �܂��A2 Player ��I�т܂��ƁA���ۂ̃Q�[���Ō��Ղ��������ƂɃX�N���[���̒[�ɕ��􂵂ĕ\������܂��B����ɂ��A���o�I�Ȍ��ʂœ�Փx���}�㏸�����邱�Ƃ��ł��܂�� ")
    Call AddTutorial(" �u���q�v�^�u�� 3 / 6 �ɂ��Ă݂܂��傤�B�V���ȃ��Y���𓾂邱�Ƃ��ł��܂��B")
    Call AddTutorial(" ���[��5�̌��ՂƃX�N���b�`���g�p�������ʂ́ubms�v�ŁA" & vbCrLf & vbCrLf & " 7�̌��ՂƃX�N���b�`���g�p�������ʂ́ubme�v�ŁA" & vbCrLf & vbCrLf & " 4�̃}�E�X���g�p�������ʂ́ummx�v�ŕۑ����܂��傤 (���ݎ�������Ă��܂���)�B")
    Call AddTutorial(" �R�[�������݂Ȃ��� bms �����Ȃ��ł��������B�V�~���ł���\��������܂��B")
    Call AddTutorial(" TOTAL �l��ύX���邱�Ƃɂ��A�Q�[�W�̏㏸����ύX���邱�Ƃ��ł��܂��B" & vbCrLf & vbCrLf & " �ʏ� TOTAL �l�̃f�t�H���g�� bms �̎d�l�ɂ���� 200 + Total Notes �ƌ��߂��Ă��܂����A�ꕔ�d�l�ɑ����Ă��Ȃ��v���C���[������܂��̂ł����ӂ�������� ")
    Call AddTutorial(" VOLWAV �͖�������Ă��܂��񂪁AVOLume of WAVe �̗����Ǝv���܂��B" & vbCrLf & vbCrLf & " �悭�킩��Ȃ�����0�ɂ��A�^�C�g�����u4:33�v�ɂ���Ƃ悢�悤�ł��B")
    Call AddTutorial(" ����� BMSE ����V���ȋ@�\���ǉ�����܂����B" & vbCrLf & vbCrLf & " ��葽���� Tips ��ǂނ��Ƃ��ł��܂��B")
    Call AddTutorial(" ���̃\�t�g�E�F�A�͂����ɂ��o�O�̂悤�ȐU�镑�������邱�Ƃ�����܂����A" & vbCrLf & vbCrLf & " ����������͎d�l�ł�� ")
    Call AddTutorial(" ���̃E�B���h�E�̂ǂ��ł������̂ŁA15��N���b�N���Ă��������B" & vbCrLf & " ....." & vbCrLf & " ...." & vbCrLf & " ..." & vbCrLf & " .." & vbCrLf & " ." & vbCrLf & vbCrLf & " �ق�A�����N���Ȃ��ł��傤�B")
    Call AddTutorial(" �u���q�v�^�u�� 10 / 572 �ɂ��Ă݂܂��傤�B�V���ȃ��Y���𓾂邱�Ƃ��ł��܂��B")
    Call AddTutorial(" BMSE �ō쐬���ꂽ BMS �̓r�[�g���ɂ��ōĐ��ł���ۏ؂͂���܂���B")
    Call AddTutorial(" ����I�Ɍ����T�C�g���������������B" & vbCrLf & vbCrLf & " http://www.killertomatoes.com/")
    Call AddTutorial(" �����Y��ĂȂ����H")
    Call AddTutorial(" BMSE �ɃC�[�X�^�[�G�b�O�͂������܂��� (�{������I)")
    Call AddTutorial(" BMSE �ɃC�[�X�^�[�G�b�O�͂���܂��񂪁ATips ��\������E���e�N������܂��B���Ȃ��͂����������܂������H")
    Call AddTutorial(" �ŐV�ł� BMSE �������[�X����Ă��邩�m�F���Ă��������I" & vbCrLf & vbCrLf & " ���F�B�S���� BMS ������N�[���� BMSE �̂��΂炵���������Ă����悤�I")
    Call AddTutorial(" ���� Tips �̓C�[�X�^�[�G�b�O�ł��B" & vbCrLf & vbCrLf & " ��Q�Ȃ��瓭�����ɍ�������̃\�t�g�E�F�A���݂Ȃ���ɋC�ɓ����Ă���������悤�Atokonats�����]��ł��邱�Ƃł��傤�B")
    
    With frmWindowTips
    
        frmWindowTips.Line (120, 120)-Step(720, 3210), RGB(128, 128, 128), BF
        
        frmWindowTips.Line (855, 120)-Step(5265, 3210), vbWhite, BF
        
        frmWindowTips.Line (855, 615)-Step(5265, 0), vbBlack, BF
        
        .CurrentX = 960
        .CurrentY = 210
        .Font.Size = 16
        .Font.Bold = True
        
        frmWindowTips.Print "�����m�ł���..."
        
        .Font.Size = 9
        .Font.Bold = False
        .CurrentX = 5400
        .CurrentY = 345
        
        frmWindowTips.Print " 0 / " & UBound(m_strTips)
        
        .Font.Size = 12
        
        Call BitBlt(frmWindowTips.hdc, 240 * m_sngTwipsX / Screen.TwipsPerPixelX, 240 * m_sngTwipsY / Screen.TwipsPerPixelY, 32, 32, picIcon.hdc, 0, 32, SRCCOPY)
    
    End With
    
    frmWindowTips.Line (960, 720)-(6075, 3270), vbWhite, BF
    Call DrawText(frmWindowTips.hdc, m_strTips(0), LenB(StrConv(m_strTips(0), vbFromUnicode)), ddRect(63, 48, 402, 216), DT_WORDBREAK)
    m_lngTipsNum = Len(m_strTips(0))
    
    chkNextDisp.Value = 1
    
    tmrMain.Enabled = True
    
    Call cmdNext.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call lngSet_ini("EasterEgg", "Tips", chkNextDisp.Value)
    
    Cancel = True
    
    tmrMain.Enabled = False
    
    Erase m_strTips()
    
    Call frmWindowTips.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub tmrMain_Timer()

    Dim strRet  As String
    
    m_lngTipsNum = m_lngTipsNum + 1
    tmrMain.Interval = 100
    
    frmWindowTips.Line (945, 720)-(6030, 3240), vbWhite, BF
    
    If m_lngTipsNum >= Len(m_strTips(m_intTipsPos)) + 1 Then
    
        tmrMain.Interval = 250
        
        'If (m_lngTipsNum \ 2) Mod 2 Then
        If (m_lngTipsNum \ 2) And 1 Then
        
            strRet = m_strTips(m_intTipsPos)
        
        Else
        
            strRet = m_strTips(m_intTipsPos) & "_"
        
        End If
        
        'If m_lngTipsNum Mod 2 Then
        If m_lngTipsNum And 1 Then
        
            'Call BitBlt(frmWindowTips.hdc, 16 * m_sngTwipsX, 16 * m_sngTwipsY, 32, 32, picIcon.hdc, 0, 32, SRCCOPY)
            Call BitBlt(frmWindowTips.hdc, 240 * m_sngTwipsX / Screen.TwipsPerPixelX, 240 * m_sngTwipsY / Screen.TwipsPerPixelY, 32, 32, picIcon.hdc, 0, 32, SRCCOPY)
        
        Else
        
            'Call BitBlt(frmWindowTips.hdc, 16 * m_sngTwipsX, 16 * m_sngTwipsY, 32, 32, picIcon.hdc, 0, 0, SRCCOPY)
            Call BitBlt(frmWindowTips.hdc, 240 * m_sngTwipsX / Screen.TwipsPerPixelX, 240 * m_sngTwipsY / Screen.TwipsPerPixelY, 32, 32, picIcon.hdc, 0, 0, SRCCOPY)
        
        End If
        
        Call DrawText(frmWindowTips.hdc, strRet, LenB(StrConv(strRet, vbFromUnicode)), ddRect(63, 48, 402, 216), DT_WORDBREAK)
    
    Else
    
        strRet = Left$(m_strTips(m_intTipsPos), m_lngTipsNum) & "_"
        
        Call DrawText(frmWindowTips.hdc, strRet, LenB(StrConv(strRet, vbFromUnicode)), ddRect(63, 48, 402, 216), DT_WORDBREAK)
        
        Select Case Right$(strRet, 2)
        
            Case vbCrLf & "_"
            
                tmrMain.Interval = 1
            
            Case " _"
            
                tmrMain.Interval = 50
            
            Case "�A_", "(_", ")_", "�u_", "�v_", "�`_"
            
                tmrMain.Interval = 200
            
            Case "�B_", "�I_", "�H_", ":_", "/_", "._"
            
                tmrMain.Interval = 400
        
        End Select
    
    End If

End Sub

Private Function ddRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As RECT

    With ddRect
        .Top = Y1 * m_sngTwipsY
        .Left = X1 * m_sngTwipsX
        .Right = X2 * m_sngTwipsX
        .Bottom = Y2 * m_sngTwipsY
    End With

End Function

Private Sub AddTutorial(ByVal str As String)

    ReDim Preserve m_strTips(UBound(m_strTips) + 1)
    
    m_strTips(UBound(m_strTips)) = str

End Sub

