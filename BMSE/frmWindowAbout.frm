VERSION 5.00
Begin VB.Form frmWindowAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'ŒÅ’èÂ°Ù ³¨ÝÄÞ³
   Caption         =   "About BMSE"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "‚l‚r ƒSƒVƒbƒN"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWindowAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  '‚È‚µ
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   0
      Picture         =   "frmWindowAbout.frx":08CA
      ScaleHeight     =   196
      ScaleMode       =   3  'Ëß¸¾Ù
      ScaleWidth      =   526
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7890
      Begin VB.Timer tmrMain 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmWindowAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sngRaster()   As Single
Private m_lngCounter    As Long

Private Sub PrintText(ByVal Text As String, ByVal X As Long, ByVal Y As Long)

    Dim intRet  As Integer
    
    intRet = LenB(StrConv(Text, vbFromUnicode))
    
    With frmWindowAbout
    
        Call SetTextColor(.hdc, 0)
        Call TextOut(.hdc, X - 1, Y - 1, Text, intRet)
        Call TextOut(.hdc, X, Y - 1, Text, intRet)
        Call TextOut(.hdc, X + 1, Y - 1, Text, intRet)
        Call TextOut(.hdc, X - 1, Y, Text, intRet)
        Call TextOut(.hdc, X + 1, Y, Text, intRet)
        Call TextOut(.hdc, X - 1, Y + 1, Text, intRet)
        Call TextOut(.hdc, X, Y + 1, Text, intRet)
        Call TextOut(.hdc, X + 1, Y + 1, Text, intRet)
        Call SetTextColor(.hdc, 16777215)
        Call TextOut(.hdc, X, Y, Text, intRet)
    
    End With

End Sub

Private Sub Form_Activate()

    Dim i   As Long
    
    ReDim m_sngRaster(frmWindowAbout.ScaleHeight - 1)
    
    For i = 0 To UBound(m_sngRaster)
    
        m_sngRaster(i) = 0
    
    Next i
    
    m_lngCounter = 0
    
    Call Form_Paint

End Sub

Private Sub Form_Click()

    Call Unload(Me)

End Sub

Private Sub Form_Deactivate()

    Call Unload(Me)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyM
        
            tmrMain.Enabled = True
    
    End Select

End Sub

Private Sub Form_Load()

    With frmWindowAbout
    
        .Width = (.Width * Screen.TwipsPerPixelX) / 15
        .Height = (.Height * Screen.TwipsPerPixelY) / 15
        .Caption = "About: " & g_strAppTitle & " (" & RELEASEDATE & " ver.)"
    
    End With
    
    Dim retRect As RECT
    
    With retRect
    
        .Left = 0
        .Top = 0
        .Right = 526
        .Bottom = 196
    
    End With
    
    With frmWindowAbout
    
        Call AdjustWindowRectEx(retRect, GetWindowLong(.hwnd, GWL_STYLE), False, GetWindowLong(.hwnd, GWL_EXSTYLE))
        
        Call .Move(.Left, .Top, (retRect.Right - retRect.Left) * Screen.TwipsPerPixelX, (retRect.Bottom - retRect.Top) * Screen.TwipsPerPixelY)
    
    End With

End Sub

Private Sub Form_Paint()

    Dim i       As Long
    Dim strRet  As String
    Dim lngRet  As Long
    Dim sngRet  As Single
    
    With frmWindowAbout
    
        Call .Cls
        
        sngRet = m_lngCounter / 10
        If sngRet > 8 Then sngRet = 8
        
        For i = 0 To .ScaleHeight - 1
        
            'm_sngRaster(i) = m_sngRaster(i) + Sin((i + m_lngCounter) * RAD * 8)
            'm_sngRaster(i) = m_sngRaster(i) + g_sngSin(((i + m_lngCounter) * 8) And 255)
            m_sngRaster(i) = g_sngSin(((i + m_lngCounter) * 8) And 255) * sngRet
            'm_sngRaster(i) = (m_sngRaster(i) + i) And .ScaleWidth
            
            If tmrMain.Enabled Then lngRet = m_sngRaster(i)
            
            'Call StretchBlt(.hdc, lngRet - .ScaleWidth, i, .ScaleWidth, 1, picMain.hdc, 0, i, .ScaleWidth, 1, SRCCOPY)
            'Call StretchBlt(.hdc, lngRet, i, .ScaleWidth, 1, picMain.hdc, 0, i, .ScaleWidth, 1, SRCCOPY)
            Call BitBlt(.hdc, lngRet, i, .ScaleWidth, 1, picMain.hdc, 0, i, SRCCOPY)
        
        Next i
    
        'Call BitBlt(.hWnd, 0, 0, .ScaleWidth, .ScaleHeight, picMain.hWnd, 0, 0, SRCCOPY)
        
        lngRet = 0
        
        .Font.Size = 9
        .Font.Underline = False
        .Font.Bold = True
        
        'For i = 0 To UBound(g_strInputLog)
        
            'lngRet = lngRet + Len(g_strInputLog(i))
        
        'Next i
        lngRet = g_InputLog.GetBufferSize
        
        Select Case lngRet
        
            Case Is < 1024
            
                strRet = lngRet & " Byte"
            
            Case Is < 1048576
            
                strRet = Round(lngRet / 1024, 2) & " KB"
            
            Case Else
            
                strRet = Round(lngRet / 1048576, 2) & " MB"
            
        End Select
        
        strRet = "Undo Buffer Size: " & strRet
        
        Call PrintText(strRet, 1, 1)
        
        'Call PrintText("Undo Counter: " & g_lngInputLogPos & " / " & UBound(g_strInputLog), 1, 13 * (15 / Screen.TwipsPerPixelY))
        Call PrintText("Undo Counter: " & g_InputLog.GetPos & " / " & g_InputLog.Max, 1, 13 * (15 / Screen.TwipsPerPixelY))
        
        '.Font.SIZE = 12 * (Screen.TwipsPerPixelX / 15)
        '.Font.Underline = True
        
        'strRet = App.Major & "." & App.Minor & "." & App.Revision
        'lngRet = LenB(StrConv(strRet, vbFromUnicode))
        
        'Call PrintText(strRet, 251, 174)
    
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    tmrMain.Enabled = False
    
    Erase m_sngRaster()
    
    Call frmWindowAbout.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub picMain_Click()

    Call Unload(Me)

End Sub

Private Sub tmrMain_Timer()

    m_lngCounter = m_lngCounter + 1
    Call Form_Paint

End Sub
