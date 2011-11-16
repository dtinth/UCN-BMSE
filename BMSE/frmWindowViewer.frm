VERSION 5.00
Begin VB.Form frmWindowViewer 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "Viewer Config"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
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
   ScaleHeight     =   4845
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdDelete 
      Caption         =   "削除"
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   3960
      Width           =   795
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "追加"
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   3960
      Width           =   795
   End
   Begin VB.Frame fraViewer 
      Height          =   4215
      Left            =   2340
      TabIndex        =   3
      Top             =   60
      Width           =   3555
      Begin VB.TextBox txtStop 
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   3060
         Width           =   3315
      End
      Begin VB.TextBox txtPlay 
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   3315
      End
      Begin VB.TextBox txtPlayAll 
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Top             =   1740
         Width           =   3315
      End
      Begin VB.CommandButton cmdViewerPath 
         Caption         =   "参照"
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtViewerPath 
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2715
      End
      Begin VB.TextBox txtViewerName 
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   3315
      End
      Begin VB.Label lblNotice 
         Caption         =   "lblNotice"
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   3315
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStop 
         Caption         =   "「停止」の引数"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   2820
         Width           =   3195
      End
      Begin VB.Label lblPlay 
         Caption         =   "「現在位置から再生」の引数"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   2160
         Width           =   3195
      End
      Begin VB.Label lblPlayAll 
         Caption         =   "「最初から再生」の引数"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1500
         Width           =   3195
      End
      Begin VB.Label lblViewerPath 
         Caption         =   "実行ファイルのパス"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   840
         Width           =   3195
      End
      Begin VB.Label lblViewerName 
         Caption         =   "表示する名前"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   3195
      End
   End
   Begin VB.ListBox lstViewer 
      Height          =   3660
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   4380
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   4380
      Width           =   1455
   End
End
Attribute VB_Name = "frmWindowViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_LocalViewer() As g_udtViewer
Dim m_lngViewerNum  As Long

Private Sub cmdAdd_Click()

    If Len(Trim$(txtViewerName.Text)) = 0 Then Exit Sub
    If Len(Trim$(txtViewerPath.Text)) = 0 Then Exit Sub
    
    ReDim Preserve m_LocalViewer(UBound(m_LocalViewer) + 1)
    
    With m_LocalViewer(UBound(m_LocalViewer))
    
        .strAppName = txtViewerName.Text
        .strAppPath = txtViewerPath.Text
        .strArgAll = txtPlayAll.Text
        .strArgPlay = txtPlay.Text
        .strArgStop = txtStop.Text
        
        Call lstViewer.AddItem(.strAppName)
        lstViewer.ListIndex = UBound(m_LocalViewer) - 1
    
    End With

End Sub

Private Sub cmdCancel_Click()

    Call Unload(Me)

End Sub

Private Sub cmdOK_Click()

    Dim i           As Long
    Dim lngFFile    As Long
    Dim lngRet      As Long
    
    ReDim g_Viewer(UBound(m_LocalViewer))
    
    If lstViewer.ListIndex <> -1 Then
    
        With m_LocalViewer(lstViewer.ListIndex + 1)
        
            .strAppName = txtViewerName.Text
            .strAppPath = txtViewerPath.Text
            .strArgAll = txtPlayAll.Text
            .strArgPlay = txtPlay.Text
            .strArgStop = txtStop.Text
        
        End With
        
        m_lngViewerNum = 0
    
    End If
    
    lngRet = frmMain.cboViewer.ListIndex
    If lngRet < 0 Then lngRet = 0
    
    Call frmMain.cboViewer.Clear
    
    lngFFile = FreeFile()
    
    Open g_strAppDir & "bmse_viewer.ini" For Output As #lngFFile
    
        For i = 1 To UBound(m_LocalViewer)
        
            With m_LocalViewer(i)
            
                Print #lngFFile, .strAppName
                Print #lngFFile, .strAppPath
                Print #lngFFile, .strArgAll
                Print #lngFFile, .strArgPlay
                Print #lngFFile, .strArgStop
                Print #lngFFile,
                
                Call frmMain.cboViewer.AddItem(.strAppName)
                
                g_Viewer(i).strAppName = .strAppName
                g_Viewer(i).strAppPath = .strAppPath
                g_Viewer(i).strArgAll = .strArgAll
                g_Viewer(i).strArgPlay = .strArgPlay
                g_Viewer(i).strArgStop = .strArgStop
            
            End With
        
        Next i
    
    Close #lngFFile
    
    With frmMain
    
        If .cboViewer.ListCount = 0 Then
        
            .tlbMenu.Buttons("PlayAll").Enabled = False
            .tlbMenu.Buttons("Play").Enabled = False
            .tlbMenu.Buttons("Stop").Enabled = False
            .mnuToolsPlayAll.Enabled = False
            .mnuToolsPlay.Enabled = False
            .mnuToolsPlayStop.Enabled = False
            .cboViewer.Enabled = False
        
        Else
        
            .tlbMenu.Buttons("PlayAll").Enabled = True
            .tlbMenu.Buttons("Play").Enabled = True
            .tlbMenu.Buttons("Stop").Enabled = True
            .mnuToolsPlayAll.Enabled = True
            .mnuToolsPlay.Enabled = True
            .mnuToolsPlayStop.Enabled = True
            .cboViewer.Enabled = True
            
            If frmMain.cboViewer.ListCount > lngRet Then
            
                frmMain.cboViewer.ListIndex = lngRet
            
            Else
            
                frmMain.cboViewer.ListIndex = 0
            
            End If
        
        End If
    
    End With
    
    Call Unload(Me)

End Sub

Private Sub cmdDelete_Click()

    With lstViewer
    
        If .ListIndex < 0 Then Exit Sub
        
        Call ViewerDelete(.ListIndex + 1)
        
        Call lstViewer.RemoveItem(.ListCount - 1)
        
        ReDim Preserve m_LocalViewer(UBound(m_LocalViewer) - 1)
    
    End With

End Sub

Private Sub ViewerDelete(ByVal Num As Long)

    If Num < UBound(m_LocalViewer) Then
    
        With m_LocalViewer(Num + 1)
        
            m_LocalViewer(Num).strAppName = .strAppName
            m_LocalViewer(Num).strAppPath = .strAppPath
            m_LocalViewer(Num).strArgAll = .strArgAll
            m_LocalViewer(Num).strArgPlay = .strArgPlay
            m_LocalViewer(Num).strArgStop = .strArgStop
            
            lstViewer.List(Num - 1) = lstViewer.List(Num)
        
        End With
        
        Call ViewerDelete(Num + 1)
    
    End If

End Sub

Private Sub cmdViewerPath_Click()
On Error GoTo Err:

    Dim retArray()  As String
    
    With frmMain.dlgMain
    
        .Filter = "EXE files (*.exe)|*.exe|All files (*.*)|*.*"
        .FileName = txtViewerPath.Text
        
        Call .ShowOpen
        
        txtViewerPath.Text = .FileName
        retArray = Split(.FileName, "\")
        .InitDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
    
    End With
    
    Exit Sub

Err:

End Sub

Private Sub Form_Activate()

    Dim i   As Long
    
    frmWindowViewer.Left = (frmMain.Left + frmMain.Width \ 2) - frmWindowViewer.Width \ 2
    frmWindowViewer.Top = (frmMain.Top + frmMain.Height \ 2) - frmWindowViewer.Height \ 2
    
    m_lngViewerNum = 0
    
    ReDim m_LocalViewer(UBound(g_Viewer))
    
    Call lstViewer.Clear
    
    For i = 1 To UBound(g_Viewer)
    
        With g_Viewer(i)
        
            Call lstViewer.AddItem(.strAppName)
            m_LocalViewer(i).strAppName = .strAppName
            m_LocalViewer(i).strAppPath = .strAppPath
            m_LocalViewer(i).strArgAll = .strArgAll
            m_LocalViewer(i).strArgPlay = .strArgPlay
            m_LocalViewer(i).strArgStop = .strArgStop
        
        End With
    
    Next i
    
    If lstViewer.ListCount Then
    
        With m_LocalViewer(1)
        
            txtViewerName.Text = .strAppName
            txtViewerPath.Text = .strAppPath
            txtPlayAll.Text = .strArgAll
            txtPlay.Text = .strArgPlay
            txtStop.Text = .strArgStop
        
        End With
        
        lstViewer.ListIndex = 0
    
    End If
    
    Call cmdOK.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    Erase m_LocalViewer()
    
    Call frmWindowViewer.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub lstViewer_Click()

    With m_LocalViewer(m_lngViewerNum + 1)
    
        .strAppName = txtViewerName.Text
        .strAppPath = txtViewerPath.Text
        .strArgAll = txtPlayAll.Text
        .strArgPlay = txtPlay.Text
        .strArgStop = txtStop.Text
    
    End With
    
    With m_LocalViewer(lstViewer.ListIndex + 1)
    
        txtViewerName.Text = .strAppName
        txtViewerPath.Text = .strAppPath
        txtPlayAll.Text = .strArgAll
        txtPlay.Text = .strArgPlay
        txtStop.Text = .strArgStop
    
    End With
    
    m_lngViewerNum = lstViewer.ListIndex

End Sub

Private Sub lstViewer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call lstViewer_Click

End Sub

Private Sub txtPlay_GotFocus()

    txtPlay.SelStart = 0
    txtPlay.SelLength = Len(txtPlay.Text)

End Sub

Private Sub txtPlayAll_GotFocus()

    txtPlayAll.SelStart = 0
    txtPlayAll.SelLength = Len(txtPlayAll.Text)

End Sub

Private Sub txtStop_GotFocus()

    txtStop.SelStart = 0
    txtStop.SelLength = Len(txtStop.Text)

End Sub

Private Sub txtViewerName_GotFocus()

    txtViewerName.SelStart = 0
    txtViewerName.SelLength = Len(txtViewerName.Text)

End Sub

Private Sub txtViewerPath_GotFocus()

    txtViewerPath.SelStart = 0
    txtViewerPath.SelLength = Len(txtViewerPath.Text)

End Sub
