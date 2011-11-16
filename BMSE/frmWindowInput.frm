VERSION 5.00
Begin VB.Form frmWindowInput 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "入力フォーム"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3990
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
   ScaleHeight     =   1335
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2970
      TabIndex        =   3
      Top             =   990
      Width           =   930
   End
   Begin VB.CommandButton cmdDecide 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   990
      Width           =   1335
   End
   Begin VB.TextBox txtMain 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   3855
   End
   Begin VB.Label lblMainDisp 
      Caption         =   "lblMainDisp"
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3765
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWindowInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Call Unload(frmWindowInput)

End Sub

Private Sub cmdDecide_Click()

    Call frmWindowInput.Hide
    
    Call frmMain.picMain.SetFocus

End Sub

Private Sub Form_Activate()

    txtMain.SelStart = 0
    txtMain.SelLength = Len(txtMain.Text)
    
    lblMainDisp.AutoSize = True
    
    Call lblMainDisp.Move(60, 60, frmWindowInput.Width - 60 - 60)
    Call txtMain.Move(60, lblMainDisp.Top + lblMainDisp.Height + 60, frmWindowInput.ScaleWidth - 120)
    Call cmdCancel.Move(frmWindowInput.ScaleWidth - 60 - cmdCancel.Width, txtMain.Top + txtMain.Height + 60)
    Call cmdDecide.Move(cmdCancel.Left - 60 - cmdDecide.Width, txtMain.Top + txtMain.Height + 60)
    
    With frmMain
    
        Call frmWindowInput.Move(.Left + (.Width - frmWindowInput.Width) \ 2, .Top + (.Height - frmWindowInput.Height) \ 2, frmWindowInput.Width, cmdDecide.Top + cmdDecide.Height + 60 + (frmWindowInput.Height - frmWindowInput.ScaleHeight))
    
    End With
    
    Call txtMain.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = True
    
    txtMain.Text = ""
    
    Call frmWindowInput.Hide
    
    Call frmMain.picMain.SetFocus

End Sub
