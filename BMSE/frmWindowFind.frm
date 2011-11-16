VERSION 5.00
Begin VB.Form frmWindowFind 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "検索・削除・置換"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8490
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
   ScaleHeight     =   2895
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   375
      Left            =   6540
      TabIndex        =   29
      Top             =   2040
      Width           =   1875
   End
   Begin VB.CommandButton cmdDecide 
      Caption         =   "実行"
      Default         =   -1  'True
      Height          =   375
      Left            =   6540
      TabIndex        =   30
      Top             =   2460
      Width           =   1875
   End
   Begin VB.Frame fraProcess 
      Caption         =   "処理"
      Height          =   1215
      Left            =   6540
      TabIndex        =   24
      Top             =   60
      Width           =   1875
      Begin VB.TextBox txtReplace 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "01"
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton optProcessReplace 
         Caption         =   "置換"
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1635
      End
      Begin VB.OptionButton optProcessDelete 
         Caption         =   "削除"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   540
         Width           =   1635
      End
      Begin VB.OptionButton optProcessSelect 
         Caption         =   "選択"
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.Frame fraSearchGrid 
      Caption         =   "列の指定"
      Height          =   2775
      Left            =   2880
      TabIndex        =   12
      Top             =   60
      Width           =   3555
      Begin VB.CommandButton cmdInvert 
         Caption         =   "反転"
         Height          =   315
         Left            =   300
         TabIndex        =   21
         Top             =   2340
         Width           =   915
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "全解除"
         Height          =   315
         Left            =   1260
         TabIndex        =   22
         Top             =   2340
         Width           =   915
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選択"
         Height          =   315
         Left            =   2220
         TabIndex        =   23
         Top             =   2340
         Width           =   1215
      End
      Begin VB.ListBox lstGrid 
         Height          =   1740
         Index           =   2
         ItemData        =   "frmWindowFind.frx":0000
         Left            =   1800
         List            =   "frmWindowFind.frx":0007
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   18
         Top             =   540
         Width           =   795
      End
      Begin VB.ListBox lstGrid 
         Height          =   1740
         Index           =   1
         ItemData        =   "frmWindowFind.frx":000F
         Left            =   960
         List            =   "frmWindowFind.frx":002B
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   16
         Top             =   540
         Width           =   795
      End
      Begin VB.ListBox lstGrid 
         Height          =   1740
         Index           =   0
         ItemData        =   "frmWindowFind.frx":0048
         Left            =   120
         List            =   "frmWindowFind.frx":0064
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   14
         Top             =   540
         Width           =   795
      End
      Begin VB.ListBox lstGrid 
         Height          =   1740
         Index           =   3
         ItemData        =   "frmWindowFind.frx":0081
         Left            =   2640
         List            =   "frmWindowFind.frx":0094
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   20
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lblBGM 
         AutoSize        =   -1  'True
         Caption         =   "BGM"
         Height          =   180
         Left            =   1800
         TabIndex        =   17
         Top             =   300
         Width           =   270
      End
      Begin VB.Label lblPlayer2 
         AutoSize        =   -1  'True
         Caption         =   "Player 2"
         Height          =   180
         Left            =   960
         TabIndex        =   15
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblPlayer1 
         AutoSize        =   -1  'True
         Caption         =   "Player 1"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblEtc 
         AutoSize        =   -1  'True
         Caption         =   "Etc"
         Height          =   180
         Left            =   2640
         TabIndex        =   19
         Top             =   300
         Width           =   270
      End
   End
   Begin VB.Frame fraSearchNum 
      Caption         =   "オブジェ番号の指定"
      Height          =   1157
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
      Begin VB.TextBox txtNumMax 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "ZZ"
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtNumMin 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   420
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "01"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblNotice 
         Caption         =   "This item doesn't influence BPM/STOP object"
         Height          =   420
         Left            =   120
         TabIndex        =   11
         Top             =   585
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNum 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Left            =   1080
         TabIndex        =   9
         Top             =   300
         Width           =   180
      End
   End
   Begin VB.Frame fraSearchMeasure 
      Caption         =   "小節範囲の指定"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   2655
      Begin VB.TextBox txtMeasureMax 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "999"
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox txtMeasureMin 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Left            =   420
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblMeasure 
         Alignment       =   2  '中央揃え
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Left            =   1080
         TabIndex        =   5
         Top             =   300
         Width           =   180
      End
   End
   Begin VB.Frame fraSearchObject 
      Caption         =   "選択対象"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      Begin VB.OptionButton optSearchSelect 
         Caption         =   "選択されているオブジェ"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   2415
      End
      Begin VB.OptionButton optSearchAll 
         Caption         =   "全てのオブジェ"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmWindowFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strArray()    As String

Private Sub cmdClose_Click()

    Call Unload(Me)

End Sub

Private Sub cmdDecide_Click()

    Dim i   As Long
    
    ReDim m_strArray(0)
    m_strArray(0) = ""
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If (optSearchAll.value = True Or .intSelect = 1) And (Val(txtMeasureMin.Text) <= .intMeasure And Val(txtMeasureMax.Text) >= .intMeasure) And ((modInput.lngNumConv(txtNumMin.Text) <= .sngValue And modInput.lngNumConv(txtNumMax.Text) >= .sngValue) Or .intCh = 8 Or .intCh = 9) Then
            
                .intSelect = 0
                
                Select Case .intCh
                
                    Case 8 'BPM
                    
                        If lstGrid(3).Selected(0) Then Call SearchProcess(i)
                    
                    Case 9 'STOP
                    
                        If lstGrid(3).Selected(1) Then Call SearchProcess(i)
                    
                    Case 11 To 15, 31 To 35, 51 To 55 '1P 1-5
                    
                        If lstGrid(0).Selected(.intCh Mod 10 - 1) Then Call SearchProcess(i)
                    
                    Case 18, 19, 38, 39, 58, 59 '1P 6-7
                    
                        If lstGrid(0).Selected(.intCh Mod 10 - 3) Then Call SearchProcess(i)
                    
                    Case 16, 36, 56 '1P SC
                    
                        If lstGrid(0).Selected(7) Then Call SearchProcess(i)
                    
                    Case 21 To 25, 41 To 45, 61 To 65 '2P 1-5
                    
                        If lstGrid(1).Selected(.intCh Mod 10 - 1) Then Call SearchProcess(i)
                    
                    Case 28, 29, 48, 49, 68, 69 '2P 6-7
                    
                        If lstGrid(1).Selected(.intCh Mod 10 - 3) Then Call SearchProcess(i)
                    
                    Case 26, 46, 66 '2P SC
                    
                        If lstGrid(1).Selected(7) Then Call SearchProcess(i)
                    
                    Case 4 'BGA
                    
                        If lstGrid(3).Selected(2) Then Call SearchProcess(i)
                    
                    Case 7 'Layer
                    
                        If lstGrid(3).Selected(3) Then Call SearchProcess(i)
                    
                    Case 6 'Poor
                    
                        If lstGrid(3).Selected(4) Then Call SearchProcess(i)
                    
                    Case Is > 100 'BGM
                    
                        If lstGrid(2).Selected(.intCh - 101) Then Call SearchProcess(i)
                
                End Select
            
            Else
            
                .intSelect = 0
            
            End If
        
        End With
    
    Next i
    
    If optProcessDelete.value Then
    
        Call frmMain.mnuEditDelete_Click
    
    ElseIf optProcessSelect.value Then
    
        Call modDraw.MoveSelectedObj
    
    ElseIf optProcessReplace.value Then
    
        If UBound(m_strArray) Then
        
            'g_strInputLog(g_lngInputLogPos) = Join(m_strArray, ",") & ","
            'g_lngInputLogPos = g_lngInputLogPos + 1
            'ReDim Preserve g_strInputLog(g_lngInputLogPos)
            'Call frmMain.SaveChanges
            Call g_InputLog.AddData(Join(m_strArray, ",") & ",")
        
        End If
    
    End If
    
    Call modDraw.Redraw

End Sub

Private Sub cmdReset_Click()

    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To 3
    
        With lstGrid(i)
        
            .Visible = False
            
            For j = 0 To .ListCount - 1
            
                .Selected(j) = False
            
            Next j
            
            .ListIndex = 0
            .Visible = True
        
        End With
    
    Next i

End Sub

Private Sub cmdInvert_Click()

    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To 3
    
        With lstGrid(i)
        
            .Visible = False
            
            For j = 0 To .ListCount - 1
            
                .Selected(j) = Not .Selected(j)
            
            Next j
            
            .ListIndex = 0
            .Visible = True
        
        End With
    
    Next i

End Sub

Private Sub cmdSelect_Click()

    Dim i   As Long
    Dim j   As Long
    
    For i = 0 To 3
    
        With lstGrid(i)
        
            .Visible = False
            
            For j = 0 To .ListCount - 1
            
                .Selected(j) = True
            
            Next j
            
            .ListIndex = 0
            .Visible = True
        
        End With
    
    Next i

End Sub

Private Sub Form_Activate()

    Call cmdDecide.SetFocus

End Sub

Private Sub Form_Load()

    Dim i   As Long
    
    For i = 2 To 32
    
        Call lstGrid(2).AddItem(Format(i, "00"))
    
    Next i
    
    Call cmdSelect_Click
    
    For i = 0 To 3
    
        lstGrid(i).ListIndex = 0
    
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    Call frmWindowFind.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub txtMeasureMax_GotFocus()

    txtMeasureMax.SelStart = 0
    txtMeasureMax.SelLength = Len(txtMeasureMax.Text)

End Sub

Private Sub txtMeasureMin_GotFocus()

    txtMeasureMin.SelStart = 0
    txtMeasureMin.SelLength = Len(txtMeasureMin.Text)

End Sub

Private Sub txtNumMax_GotFocus()

    txtNumMax.SelStart = 0
    txtNumMax.SelLength = Len(txtNumMax.Text)

End Sub

Private Sub txtNumMin_GotFocus()

    txtNumMin.SelStart = 0
    txtNumMin.SelLength = Len(txtNumMin.Text)

End Sub

Private Sub txtReplace_GotFocus()

    optProcessReplace.value = True
    txtReplace.SelStart = 0
    txtReplace.SelLength = Len(txtReplace.Text)

End Sub

Private Sub SearchProcess(ByVal Num As Long)

    With g_Obj(Num)
    
        If optProcessReplace.value Then
        
            If .intCh <> 8 And .intCh <> 9 Then
            
                m_strArray(UBound(m_strArray)) = modInput.strNumConv(CMD_LOG.OBJ_CHANGE) & modInput.strNumConv(.lngID, 4) & modInput.strNumConv(.sngValue) & Right$("0" & txtReplace.Text, 2)
                ReDim Preserve m_strArray(UBound(m_strArray) + 1)
                .sngValue = modInput.lngNumConv(txtReplace.Text)
            
            End If
        
        End If
        
        .intSelect = 1
    
    End With

End Sub
