VERSION 5.00
Begin VB.Form frmWindowTips 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "BMSE Tips (Sorry Japanese Language Only!!!!!!!111)"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   FillStyle       =   0  '塗りつぶし
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
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
   StartUpPosition =   3  'Windows の既定値
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'なし
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
      Value           =   1  'ﾁｪｯｸ
      Width           =   2775
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "次へ"
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   3480
      Width           =   1515
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
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
            
                Call MsgBox("よくわからないけど多分エラーが発生しました。" & vbCrLf & "次回も Tips を表示します。", vbCritical Or vbOKOnly, g_strAppTitle)
                
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
            
            lngRet = MsgBox("本当に？", vbAbortRetryIgnore Or lngArg, g_strAppTitle)
        
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
    
    m_strTips(0) = " これから Tips を表示します。" & vbCrLf & vbCrLf & " これらの情報はあなたが BMSE を使い BMS を作成するのを手助けしてくれることもあるかもしれません。" & vbCrLf & vbCrLf & " 「次へ」のボタンを押して Tips を開始してください。" & vbCrLf & vbCrLf & " (この文章は一度しか表示されません)"
    
    Call AddTutorial(" BMSE は UCN-Soft が開発しています。" & vbCrLf & vbCrLf & " UCN の由来は超内輪ネタなので内緒です！")
    Call AddTutorial(" BMSE は BMx Sequence Editor の略です。知らない友達がいたら広めよう！")
    Call AddTutorial(" BMSE は bms ファイル、bme ファイル、bml ファイルおよび pms ファイルを書き出すことができます。")
    Call AddTutorial(" bms の正式名称は Be-Music Script など諸説ありますが、真相は謎のままです。")
    Call AddTutorial(" BMSE を使用するには、まず Windows OS の操作に習熟する必要があります。" & vbCrLf & vbCrLf & " マウスは片手で持ち、画面上のポインタを操作します。ディスプレイを指でなぞるわけではありません。")
    Call AddTutorial(" オブジェを配置するにはスクリーンを左クリックします。" & vbCrLf & vbCrLf & " 左クリックの仕方については、お使いの OS のマニュアルをお読みください。" & vbCrLf & vbCrLf & " (BMSE はマウスが必須です)")
    Call AddTutorial(" オブジェが配置できない？消しゴムツールになっていませんか？")
    Call AddTutorial(" 右側に表示されているテキスト・ボックスには任意の文字列を入力します。" & vbCrLf & vbCrLf & " 文字列を入力するにはキーボードが必要ですので、お使いの OS 及び言語ツールのマニュアルをお読みください。")
    Call AddTutorial(" GENRE は「ジャンル」と読み、選曲中に表示されるおおまかな曲の傾向を入力します。" & vbCrLf & vbCrLf & " よくわからない時は Techno にしてください。")
    Call AddTutorial(" bpm は Beat Per Minute の略で、1分あたりのビート数を入力します。" & vbCrLf & vbCrLf & " よくわからない時は400にしてください。")
    Call AddTutorial(" TITLE は「タイトル」と読みます。英語で「題名」を意味し、選曲中に表示される曲の題名を入力します。" & vbCrLf & vbCrLf & " よくわからない時は英和辞書を引いてください (英和辞書はお近くの書店で購入可能です)。")
    Call AddTutorial(" ARTIST は直訳すると「芸術家」となりますが、ここでは「作者」を入力してください。" & vbCrLf & vbCrLf & " よくわからない時は「DJ 苗字」としてください。例: DJ 山田")
    Call AddTutorial(" PLAYLEVEL は「譜面の難易度」です。だいたい 1 ～ 7 が bms の デファクトスタンダードです。" & vbCrLf & vbCrLf & " よくわからない時はノート数÷100にしてください。")
    Call AddTutorial(" 「基本」タブの隣に「拡張」タブおよび「環境」タブがあることにお気づきですか？" & vbCrLf & vbCrLf & " クリックすれば新たな設定を行うことが可能になります。")
    Call AddTutorial(" RANK は直訳しても意味が通じません。「判定の厳しさ」を現します。" & vbCrLf & vbCrLf & " よくわからない時は VERY HARD にしてください。")
    Call AddTutorial(" 実は BMSE は MOD に対応しています (現在隠しコマンド)。" & vbCrLf & vbCrLf & " この先を読むにはシェアウエアフィーを払う必要があります。" & vbCrLf & vbCrLf & " このソフトウェアは臓器ウェアです。気に入ったら作者に臓器を寄付してください。")
    Call AddTutorial(" テンキーを押すと、ビル・ゲイツとメッセンジャーでチャットができます。")
    Call AddTutorial(" スクリーンの一番左にある「BPM」および「STOP」レーンに注目してください！" & vbCrLf & vbCrLf & " このレーンをクリックし、単純に半角英数 (キーボードの右端にある狭い数字のみの領域を押下してください) を入力するだけで、プレイヤーを翻弄することができます。")
    Call AddTutorial(" BMSE はマウマニに対応していません。本当だよ！")
    Call AddTutorial(" てっとり早く bms を作るには、wav を使用せずに作るのが一番です。" & vbCrLf & vbCrLf & " 絵を描く感覚でスクリーンにオブジェを配置 (左クリック) すれば bms が完成！簡単でしょ？")
    Call AddTutorial(" 「基本」タブの一番上にある「プレイモード」を Double Play にしてみましょう。鍵盤の数が倍増し、より「太い」譜面を作ることができます。" & vbCrLf & vbCrLf & " また、2 Player を選びますと、実際のゲームで鍵盤が半分ごとにスクリーンの端に分裂して表示されます。これにより、視覚的な効果で難易度を急上昇させることができます｡ ")
    Call AddTutorial(" 「拍子」タブで 3 / 6 にしてみましょう。新たなリズムを得ることができます。")
    Call AddTutorial(" 左端の5つの鍵盤とスクラッチを使用した譜面は「bms」で、" & vbCrLf & vbCrLf & " 7つの鍵盤とスクラッチを使用した譜面は「bme」で、" & vbCrLf & vbCrLf & " 4つのマウスを使用した譜面は「mmx」で保存しましょう (現在実装されていません)。")
    Call AddTutorial(" コーラを飲みながら bms を作らないでください。シミができる可能性があります。")
    Call AddTutorial(" TOTAL 値を変更することにより、ゲージの上昇率を変更することができます。" & vbCrLf & vbCrLf & " 通常 TOTAL 値のデフォルトは bms の仕様によって 200 + Total Notes と決められていますが、一部仕様に則っていないプレイヤーもありますのでご注意ください｡ ")
    Call AddTutorial(" VOLWAV は明言されていませんが、VOLume of WAVe の略だと思われます。" & vbCrLf & vbCrLf & " よくわからない時は0にし、タイトルを「4:33」にするとよいようです。")
    Call AddTutorial(" 今回の BMSE から新たな機能が追加されました。" & vbCrLf & vbCrLf & " より多くの Tips を読むことができます。")
    Call AddTutorial(" このソフトウェアはいかにもバグのような振る舞いをすることがありますが、" & vbCrLf & vbCrLf & " しかしそれは仕様です｡ ")
    Call AddTutorial(" このウィンドウのどこでもいいので、15回クリックしてください。" & vbCrLf & " ....." & vbCrLf & " ...." & vbCrLf & " ..." & vbCrLf & " .." & vbCrLf & " ." & vbCrLf & vbCrLf & " ほら、何も起きないでしょう。")
    Call AddTutorial(" 「拍子」タブで 10 / 572 にしてみましょう。新たなリズムを得ることができます。")
    Call AddTutorial(" BMSE で作成された BMS はビート魔にやりで再生できる保証はありません。")
    Call AddTutorial(" 定期的に公式サイトをご覧ください。" & vbCrLf & vbCrLf & " http://www.killertomatoes.com/")
    Call AddTutorial(" 何か忘れてないか？")
    Call AddTutorial(" BMSE にイースターエッグはございません (本当だよ！)")
    Call AddTutorial(" BMSE にイースターエッグはありませんが、Tips を表示するウルテクがあります。あなたはもう発見しましたか？")
    Call AddTutorial(" 最新版の BMSE がリリースされているか確認してください！" & vbCrLf & vbCrLf & " お友達全員に BMS が作れるクールな BMSE のすばらしさを教えてあげよう！")
    Call AddTutorial(" この Tips はイースターエッグです。" & vbCrLf & vbCrLf & " 夜寝ながら働かずに作ったこのソフトウェアがみなさんに気に入っていただけるよう、tokonats氏が望んでいることでしょう。")
    
    With frmWindowTips
    
        frmWindowTips.Line (120, 120)-Step(720, 3210), RGB(128, 128, 128), BF
        
        frmWindowTips.Line (855, 120)-Step(5265, 3210), vbWhite, BF
        
        frmWindowTips.Line (855, 615)-Step(5265, 0), vbBlack, BF
        
        .CurrentX = 960
        .CurrentY = 210
        .Font.Size = 16
        .Font.Bold = True
        
        frmWindowTips.Print "ご存知ですか..."
        
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
            
            Case "、_", "(_", ")_", "「_", "」_", "～_"
            
                tmrMain.Interval = 200
            
            Case "。_", "！_", "？_", ":_", "/_", "._"
            
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

