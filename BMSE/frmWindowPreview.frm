VERSION 5.00
Begin VB.Form frmWindowPreview 
   BorderStyle     =   5  '‰Â•ÏÂ°Ù ³¨ÝÄÞ³
   Caption         =   "Preview Window"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "‚l‚r ƒSƒVƒbƒN"
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
   ScaleHeight     =   3840
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.Frame fraBGACmd 
      Height          =   315
      Left            =   3900
      TabIndex        =   20
      Top             =   3510
      Width           =   2275
      Begin VB.CommandButton cmdPreviewBack 
         Caption         =   "<"
         Height          =   315
         Left            =   480
         TabIndex        =   22
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdPreviewNext 
         Caption         =   ">"
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdPreviewEnd 
         Caption         =   "ZZ"
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdPreviewHome 
         Caption         =   "01"
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame fraBGAPara 
      Height          =   2355
      Left            =   3900
      TabIndex        =   2
      Top             =   0
      Width           =   2288
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   315
         Left            =   65
         TabIndex        =   17
         Top             =   1350
         Width           =   1001
      End
      Begin VB.CheckBox chkLock 
         Caption         =   "Lock"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   2015
         Width           =   2115
      End
      Begin VB.CheckBox chkBGLine 
         Caption         =   "BG-Line"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   1755
         Value           =   1  'Áª¯¸
         Width           =   2115
      End
      Begin VB.TextBox txtBGAPara 
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   0
         Left            =   390
         MaxLength       =   8
         TabIndex        =   4
         Text            =   "01"
         Top             =   0
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   1
         Left            =   390
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   2
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "0"
         Top             =   325
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   3
         Left            =   390
         MaxLength       =   8
         TabIndex        =   10
         Text            =   "0"
         Top             =   650
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   4
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   12
         Text            =   "0"
         Top             =   650
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   5
         Left            =   390
         MaxLength       =   8
         TabIndex        =   14
         Text            =   "0"
         Top             =   975
         Width           =   689
      End
      Begin VB.TextBox txtBGAPara 
         Height          =   270
         IMEMode         =   3  'µÌŒÅ’è
         Index           =   6
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   16
         Text            =   "0"
         Top             =   975
         Width           =   689
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "Num"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   270
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "X1"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   390
         Width           =   180
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "Y1"
         Height          =   180
         Index           =   2
         Left            =   1230
         TabIndex        =   7
         Top             =   390
         Width           =   180
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "X2"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "Y2"
         Height          =   180
         Index           =   4
         Left            =   1230
         TabIndex        =   11
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "dX"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   13
         Top             =   1035
         Width           =   180
      End
      Begin VB.Label lblBGAPara 
         AutoSize        =   -1  'True
         Caption         =   "dY"
         Height          =   180
         Index           =   6
         Left            =   1230
         TabIndex        =   15
         Top             =   1035
         Width           =   180
      End
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  '‚È‚µ
      Height          =   540
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  '‚È‚µ
      DrawWidth       =   2
      FillStyle       =   0  '“h‚è‚Â‚Ô‚µ
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   0
      OLEDropMode     =   1  'Žè“®
      ScaleHeight     =   37
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   61
      TabIndex        =   1
      Top             =   600
      Width           =   915
   End
End
Attribute VB_Name = "frmWindowPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = -20

Public Sub SetWindowSize()

    Dim retRect As RECT
    
    Call Form_Resize
    
    With retRect
    
        .Left = 0
        .Top = 0
        .Right = 256 + fraBGAPara.Width \ Screen.TwipsPerPixelX
        .Bottom = 256
    
    End With
    
    With frmWindowPreview
    
        Call AdjustWindowRectEx(retRect, GetWindowLong(.hwnd, GWL_STYLE), False, GetWindowLong(.hwnd, GWL_EXSTYLE))
        
        Call .Move(.Left, .Top, (retRect.Right - retRect.Left) * Screen.TwipsPerPixelX, (retRect.Bottom - retRect.Top) * Screen.TwipsPerPixelY)
    
    End With

End Sub

Private Sub chkBGLine_Click()

    Call picPreview_Paint

End Sub

Private Sub cmdCopy_Click()

    Dim i           As Long
    Dim strArray(6) As String
    
    For i = 0 To UBound(strArray)
    
        strArray(i) = txtBGAPara(i).Text
        
        If Len(strArray(i)) = 0 Then strArray(i) = "0"
    
    Next i
    
    Call Clipboard.Clear
    Call Clipboard.SetText(Join(strArray, " "))

End Sub

Private Sub cmdPreviewBack_Click()

    Dim i   As Long
    
    With frmMain
    
        If .optChangeBottom(2).value Then
        
            i = .lstBGA.ListIndex - 1
            
            Do While i >= 0
            
                If Len(.lstBGA.List(i)) > 8 Then
                
                    .lstBGA.ListIndex = i
                    
                    Exit Sub
                
                End If
                
                i = i - 1
            
            Loop
        
        Else
        
            i = .lstBMP.ListIndex - 1
            
            Do While i >= 0
            
                If Len(.lstBMP.List(i)) > 8 Then
                
                    .lstBMP.ListIndex = i
                    
                    Exit Sub
                
                End If
                
                i = i - 1
            
            
            Loop
        
        End If
    
    End With

End Sub

Private Sub cmdPreviewEnd_Click()

    With frmMain
    
        If .optChangeBottom(2).value Then
        
            .lstBGA.ListIndex = .lstBGA.ListCount - 1
        
        Else
        
            .lstBMP.ListIndex = .lstBMP.ListCount - 1
        
        End If
    
    End With

End Sub

Private Sub cmdPreviewHome_Click()

    With frmMain
    
        If .optChangeBottom(2).value Then
        
            .lstBGA.ListIndex = 0
        
        Else
        
            .lstBMP.ListIndex = 0
        
        End If
    
    End With

End Sub

Private Sub cmdPreviewNext_Click()

    Dim i   As Long
    
    With frmMain
    
        If .optChangeBottom(2).value Then
        
            i = .lstBGA.ListIndex + 1
            
            Do While i < .lstBGA.ListCount
            
                If Len(.lstBGA.List(i)) > 8 Then
                
                    .lstBGA.ListIndex = i
                    
                    Exit Sub
                
                End If
            
                i = i + 1
            
            Loop
        
        Else
        
            i = .lstBMP.ListIndex + 1
            
            Do While i < .lstBMP.ListCount
            
                If Len(.lstBMP.List(i)) > 8 Then
                
                    .lstBMP.ListIndex = i
                    
                    Exit Sub
                
                End If
                
                i = i + 1
            
            Loop
        
        End If
    
    End With

End Sub

Private Sub Form_Load()

    With frmWindowPreview
    
        Call .picPreview.Move(0, 0, 256 * Screen.TwipsPerPixelX, 256 * Screen.TwipsPerPixelY)
        Call .picBackBuffer.Move(0, 0, 256 * Screen.TwipsPerPixelX, 256 * Screen.TwipsPerPixelY)
        .fraBGAPara.BorderStyle = 0
        .fraBGACmd.BorderStyle = 0
    
    End With

End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim lngRet  As Long
    
    With frmWindowPreview
    
        lngRet = 120
        
        .lblBGAPara(0).Left = lngRet
        .lblBGAPara(1).Left = lngRet
        .lblBGAPara(3).Left = lngRet
        .lblBGAPara(5).Left = lngRet
        .cmdCopy.Left = lngRet
        .chkBGLine.Left = lngRet
        .chkLock.Left = lngRet
        
        .cmdPreviewHome.Left = lngRet
        .cmdPreviewBack.Left = .cmdPreviewHome.Left + .cmdPreviewHome.Width + 60
        .cmdPreviewNext.Left = .cmdPreviewBack.Left + .cmdPreviewBack.Width + 60
        .cmdPreviewEnd.Left = .cmdPreviewNext.Left + .cmdPreviewNext.Width + 60
        .fraBGACmd.Width = .cmdPreviewEnd.Left + .cmdPreviewEnd.Width + 60
        .fraBGACmd.Height = .cmdPreviewEnd.Height
        
        lngRet = lngRet + .lblBGAPara(0).Width + 60
        
        .txtBGAPara(0).Left = lngRet
        .txtBGAPara(1).Left = lngRet
        .txtBGAPara(3).Left = lngRet
        .txtBGAPara(5).Left = lngRet
        
        lngRet = lngRet + .txtBGAPara(0).Width + 180
        
        .lblBGAPara(2).Left = lngRet
        .lblBGAPara(4).Left = lngRet
        .lblBGAPara(6).Left = lngRet
        
        lngRet = lngRet + .lblBGAPara(0).Width + 60
        
        .txtBGAPara(2).Left = lngRet
        .txtBGAPara(4).Left = lngRet
        .txtBGAPara(6).Left = lngRet
        
        lngRet = lngRet + .txtBGAPara(0).Width
        
        .chkBGLine.Width = lngRet - 120
        .chkLock.Width = lngRet - 120
        
        .fraBGAPara.Width = lngRet + 60
        
        Call .picPreview.Move(0, 0, .ScaleWidth - fraBGAPara.Width, .ScaleHeight)
        Call .fraBGAPara.Move(.ScaleWidth - fraBGAPara.Width, 60)
        Call .fraBGACmd.Move(.ScaleWidth - fraBGAPara.Width, .ScaleHeight - fraBGACmd.Height - 60)
        
        Call .picPreview_Paint
    
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    Call frmWindowPreview.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub picPreview_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err:

    Dim i           As Long
    Dim strRet      As String
    
    For i = 1 To Data.Files.Count
    
        If Dir(Data.Files.Item(i), vbNormal) <> vbNullString Then
        
            strRet = Data.Files.Item(i)
            
            Call frmMain.PreviewBMP(strRet)
            
            frmWindowPreview.Caption = Right$(frmWindowPreview.Caption, Len(frmWindowPreview.Caption) - 3)
        
        End If
    
    Next i

Err:
End Sub

Public Sub picPreview_Paint()

    Dim i       As Long
    
    With picPreview
    
        Call .Cls
        
        Call BitBlt(.hdc, (.ScaleWidth \ 2 - 128) + val(txtBGAPara(BGA_dX).Text) - val(txtBGAPara(BGA_X1).Text), (.ScaleHeight \ 2 - 128) + val(txtBGAPara(BGA_dY).Text) - val(txtBGAPara(BGA_Y1).Text), picBackBuffer.ScaleWidth, picBackBuffer.ScaleHeight, picBackBuffer.hdc, 0, 0, SRCCOPY)
        
        If chkBGLine.value Then
        
            .DrawWidth = 1
            
            For i = 4 To .ScaleHeight Step 8
            
                Call MoveToEx(.hdc, 0, i, 0)
                Call LineTo(.hdc, .ScaleWidth, i)
            
            Next i
        
        End If
        
        .DrawWidth = 2
        
        Call Rectangle(.hdc, (.ScaleWidth \ 2 - 129), (.ScaleHeight \ 2 - 129), (.ScaleWidth \ 2 + 130), (.ScaleHeight \ 2 + 130))
        
        Call BitBlt(.hdc, (.ScaleWidth \ 2 - 128) + val(txtBGAPara(BGA_dX).Text), (.ScaleHeight \ 2 - 128) + val(txtBGAPara(BGA_dY).Text), lngNumField(val(txtBGAPara(BGA_X2).Text) - val(txtBGAPara(BGA_X1).Text), 0, 256), lngNumField(val(txtBGAPara(BGA_Y2).Text) - val(txtBGAPara(BGA_Y1).Text), 0, 256), picBackBuffer.hdc, val(txtBGAPara(BGA_X1).Text), val(txtBGAPara(BGA_Y1).Text), SRCCOPY)
    
    End With

End Sub

Private Sub txtBGAPara_Change(Index As Integer)

    If val(txtBGAPara(Index).Text) < 0 Then txtBGAPara(Index).Text = 0
    
    Call picPreview_Paint

End Sub

Private Sub txtBGAPara_GotFocus(Index As Integer)

    txtBGAPara(Index).SelStart = 0
    txtBGAPara(Index).SelLength = Len(txtBGAPara(Index).Text)

End Sub

Private Function lngNumField(ByVal lngNum As Long, ByVal lngMin As Long, ByVal lngMax As Long) As Long

    If lngNum < lngMin Then lngNum = lngMin
    
    If lngNum > lngMax Then lngNum = lngMax
    
    lngNumField = lngNum

End Function
