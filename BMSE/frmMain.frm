VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "BMx Sequence Editor"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   18705
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '手動
   ScaleHeight     =   7590
   ScaleWidth      =   18705
   StartUpPosition =   3  'Windows の既定値
   Visible         =   0   'False
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  '下揃え
      Height          =   315
      Left            =   0
      TabIndex        =   98
      Top             =   7275
      Width           =   18705
      _ExtentX        =   32994
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Edit Mode"
            TextSave        =   "Edit Mode"
            Key             =   "Mode"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25321
            Text            =   "Position:"
            TextSave        =   "Position:"
            Key             =   "Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "#WAV 01"
            TextSave        =   "#WAV 01"
            Key             =   "WAV"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "#BMP 01"
            TextSave        =   "#BMP 01"
            Key             =   "BMP"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "4/4"
            TextSave        =   "4/4"
            Key             =   "Measure"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSiromaru 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   3840
      Left            =   17760
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   256
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
      ScaleWidth      =   64
      TabIndex        =   99
      Top             =   480
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12120
      Top             =   1080
   End
   Begin VB.Frame fraResolution 
      Height          =   315
      Left            =   15120
      TabIndex        =   16
      Top             =   2640
      Width           =   1755
      Begin VB.ComboBox cboVScroll 
         Height          =   300
         ItemData        =   "frmMain.frx":294C
         Left            =   720
         List            =   "frmMain.frx":2953
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblVScroll 
         AutoSize        =   -1  'True
         Caption         =   "VScroll"
         Height          =   180
         Left            =   0
         TabIndex        =   17
         Top             =   60
         Width           =   630
      End
   End
   Begin VB.Frame fraDispSize 
      Height          =   315
      Left            =   12120
      TabIndex        =   11
      Top             =   2640
      Width           =   2955
      Begin VB.ComboBox cboDispHeight 
         Height          =   300
         ItemData        =   "frmMain.frx":295A
         Left            =   480
         List            =   "frmMain.frx":298C
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   13
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cboDispWidth 
         Height          =   300
         ItemData        =   "frmMain.frx":29C5
         Left            =   1740
         List            =   "frmMain.frx":29F7
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblDispWidth 
         AutoSize        =   -1  'True
         Caption         =   "幅"
         Height          =   180
         Left            =   1440
         TabIndex        =   14
         Top             =   60
         Width           =   180
      End
      Begin VB.Label lblDispHeight 
         AutoSize        =   -1  'True
         Caption         =   "高さ"
         Height          =   180
         Left            =   0
         TabIndex        =   12
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Frame fraViewer 
      Height          =   315
      Left            =   12120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
      Begin VB.ComboBox cboViewer 
         Height          =   300
         ItemData        =   "frmMain.frx":2A30
         Left            =   0
         List            =   "frmMain.frx":2A32
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12120
      Top             =   540
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   12600
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":409C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4636
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":516A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5704
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6238
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboDirectInput 
      Height          =   300
      Left            =   600
      TabIndex        =   96
      Top             =   6840
      Width           =   435
   End
   Begin VB.CommandButton cmdDirectInput 
      Caption         =   "Input"
      Height          =   315
      Left            =   1080
      TabIndex        =   97
      Top             =   6840
      Width           =   915
   End
   Begin VB.Frame fraGrid 
      Height          =   315
      Left            =   12120
      TabIndex        =   6
      Top             =   2160
      Width           =   2955
      Begin VB.ComboBox cboDispGridMain 
         Height          =   300
         ItemData        =   "frmMain.frx":67D2
         Left            =   1920
         List            =   "frmMain.frx":67F0
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
      Begin VB.ComboBox cboDispGridSub 
         Height          =   300
         ItemData        =   "frmMain.frx":6811
         Left            =   480
         List            =   "frmMain.frx":683C
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   8
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblGridSub 
         AutoSize        =   -1  'True
         Caption         =   "補助"
         Height          =   180
         Left            =   1440
         TabIndex        =   9
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblGridMain 
         AutoSize        =   -1  'True
         Caption         =   "Grid"
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Frame fraHeader 
      Height          =   2475
      Left            =   2100
      TabIndex        =   19
      Top             =   480
      Width           =   9990
      Begin VB.Frame fraTop 
         Height          =   1875
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.ComboBox cboPlayer 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":686A
            Left            =   1200
            List            =   "frmMain.frx":686C
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   25
            Top             =   120
            Width           =   1995
         End
         Begin VB.TextBox txtGenre 
            Height          =   270
            Left            =   1200
            TabIndex        =   27
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtTitle 
            Height          =   270
            Left            =   1200
            TabIndex        =   29
            Top             =   840
            Width           =   1995
         End
         Begin VB.TextBox txtArtist 
            Height          =   270
            Left            =   1200
            TabIndex        =   31
            Top             =   1200
            Width           =   1995
         End
         Begin VB.ComboBox cboPlayLevel 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":686E
            Left            =   1200
            List            =   "frmMain.frx":688D
            TabIndex        =   33
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtBPM 
            Height          =   270
            IMEMode         =   3  'ｵﾌ固定
            Left            =   2580
            TabIndex        =   35
            Text            =   "130"
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label lblPlayMode 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "プレイモード"
            Height          =   180
            Left            =   60
            TabIndex        =   24
            Top             =   180
            Width           =   1080
         End
         Begin VB.Label lblGenre 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "ジャンル"
            Height          =   180
            Left            =   420
            TabIndex        =   26
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblTitle 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "タイトル"
            Height          =   180
            Left            =   420
            TabIndex        =   28
            Top             =   900
            Width           =   720
         End
         Begin VB.Label lblArtist 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "アーティスト"
            Height          =   180
            Left            =   60
            TabIndex        =   30
            Top             =   1260
            Width           =   1080
         End
         Begin VB.Label lblPlayLevel 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "難易度表示"
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   1620
            Width           =   900
         End
         Begin VB.Label lblBPM 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "BPM"
            Height          =   180
            Left            =   2220
            TabIndex        =   34
            Top             =   1620
            Width           =   270
         End
      End
      Begin VB.Frame fraTop 
         Height          =   1875
         Index           =   1
         Left            =   3300
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.ComboBox cboPlayRank 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":68AC
            Left            =   1200
            List            =   "frmMain.frx":68BC
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   38
            Top             =   120
            Width           =   1995
         End
         Begin VB.TextBox txtTotal 
            Height          =   270
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1200
            TabIndex        =   40
            Top             =   480
            Width           =   1995
         End
         Begin VB.TextBox txtVolume 
            Height          =   270
            IMEMode         =   3  'ｵﾌ固定
            Left            =   1200
            TabIndex        =   42
            Top             =   840
            Width           =   1995
         End
         Begin VB.TextBox txtStageFile 
            Height          =   270
            Left            =   1200
            TabIndex        =   44
            Top             =   1200
            Width           =   1395
         End
         Begin VB.CommandButton cmdLoadStageFile 
            Caption         =   "参照"
            Height          =   255
            Left            =   2640
            TabIndex        =   45
            Top             =   1200
            Width           =   555
         End
         Begin VB.CommandButton cmdLoadMissBMP 
            Caption         =   "参照"
            Height          =   255
            Left            =   2640
            TabIndex        =   48
            Top             =   1560
            Width           =   555
         End
         Begin VB.TextBox txtMissBMP 
            Height          =   270
            Left            =   1200
            TabIndex        =   47
            Top             =   1560
            Width           =   1395
         End
         Begin VB.Label lblPlayRank 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "#RANK"
            Height          =   180
            Left            =   675
            TabIndex        =   37
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "#TOTAL"
            Height          =   180
            Left            =   600
            TabIndex        =   39
            Top             =   540
            Width           =   540
         End
         Begin VB.Label lblVolume 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "#VOLWAV"
            Height          =   180
            Left            =   510
            TabIndex        =   41
            Top             =   900
            Width           =   630
         End
         Begin VB.Label lblStageFile 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "#STAGEFILE"
            Height          =   180
            Left            =   225
            TabIndex        =   43
            Top             =   1260
            Width           =   900
         End
         Begin VB.Label lblMissBMP 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "#BMP00"
            Height          =   180
            Left            =   585
            TabIndex        =   46
            Top             =   1620
            Width           =   540
         End
      End
      Begin VB.Frame fraTop 
         Height          =   1875
         Index           =   2
         Left            =   6600
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.ComboBox cboDispFrame 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":68DF
            Left            =   1200
            List            =   "frmMain.frx":68E9
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   51
            Top             =   120
            Width           =   1995
         End
         Begin VB.ComboBox cboDispSC2P 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":6901
            Left            =   2460
            List            =   "frmMain.frx":690B
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   57
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox cboDispSC1P 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":6917
            Left            =   1200
            List            =   "frmMain.frx":6921
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   55
            Top             =   840
            Width           =   735
         End
         Begin VB.ComboBox cboDispKey 
            Height          =   300
            IMEMode         =   3  'ｵﾌ固定
            ItemData        =   "frmMain.frx":692D
            Left            =   1200
            List            =   "frmMain.frx":6937
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   53
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label lblDispFrame 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "キー表示"
            Height          =   180
            Left            =   420
            TabIndex        =   50
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblDispSC2P 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "2P"
            Height          =   180
            Left            =   2220
            TabIndex        =   56
            Top             =   900
            Width           =   180
         End
         Begin VB.Label lblDispSC1P 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "スクラッチ1P"
            Height          =   180
            Left            =   60
            TabIndex        =   54
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label lblDispKey 
            Alignment       =   1  '右揃え
            AutoSize        =   -1  'True
            Caption         =   "キー配置"
            Height          =   180
            Left            =   420
            TabIndex        =   52
            Top             =   540
            Width           =   720
         End
      End
      Begin VB.OptionButton optChangeTop 
         Caption         =   "基本"
         Height          =   315
         Index           =   0
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   20
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton optChangeTop 
         Caption         =   "環境"
         Height          =   315
         Index           =   2
         Left            =   1950
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   22
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton optChangeTop 
         Caption         =   "拡張"
         Height          =   315
         Index           =   1
         Left            =   975
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   21
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Frame fraMaterial 
      Height          =   4035
      Left            =   2100
      TabIndex        =   58
      Top             =   3060
      Width           =   16575
      Begin VB.OptionButton optChangeBottom 
         Caption         =   "#WAV"
         Height          =   315
         Index           =   0
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   59
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton optChangeBottom 
         Caption         =   "#BMP"
         Height          =   315
         Index           =   1
         Left            =   975
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   60
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton optChangeBottom 
         Caption         =   "#BGA"
         Height          =   315
         Index           =   2
         Left            =   1950
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   61
         Top             =   0
         Width           =   915
      End
      Begin VB.OptionButton optChangeBottom 
         Caption         =   "拍子"
         Height          =   315
         Index           =   3
         Left            =   0
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   62
         Top             =   375
         Width           =   915
      End
      Begin VB.OptionButton optChangeBottom 
         Caption         =   "拡張命令"
         Height          =   315
         Index           =   4
         Left            =   975
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   63
         Top             =   375
         Width           =   915
      End
      Begin VB.Frame fraBottom 
         Height          =   3495
         Index           =   4
         Left            =   13260
         TabIndex        =   93
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox txtExInfo 
            Height          =   3255
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  '両方
            TabIndex        =   94
            Top             =   120
            Width           =   3195
         End
      End
      Begin VB.Frame fraBottom 
         Height          =   3135
         Index           =   0
         Left            =   60
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton cmdSoundExcUp 
            Caption         =   "▲"
            Height          =   315
            Left            =   915
            TabIndex        =   67
            Top             =   2700
            Width           =   315
         End
         Begin VB.CommandButton cmdSoundExcDown 
            Caption         =   "▼"
            Height          =   315
            Left            =   1290
            TabIndex        =   68
            Top             =   2700
            Width           =   315
         End
         Begin VB.CommandButton cmdSoundDelete 
            Caption         =   "消去"
            Height          =   315
            Left            =   1890
            TabIndex        =   69
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdSoundLoad 
            Caption         =   "指定"
            Height          =   315
            Left            =   2580
            TabIndex        =   70
            Top             =   2700
            Width           =   615
         End
         Begin VB.CommandButton cmdSoundStop 
            Caption         =   "停止"
            Height          =   315
            Left            =   0
            TabIndex        =   66
            Top             =   2700
            Width           =   795
         End
         Begin VB.ListBox lstWAV 
            Height          =   2220
            Left            =   0
            OLEDropMode     =   1  '手動
            TabIndex        =   65
            Top             =   120
            Width           =   3195
         End
      End
      Begin VB.Frame fraBottom 
         Height          =   3495
         Index           =   1
         Left            =   3360
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton cmdBMPExcDown 
            Caption         =   "▼"
            Height          =   315
            Left            =   1290
            TabIndex        =   75
            Top             =   3060
            Width           =   315
         End
         Begin VB.CommandButton cmdBMPExcUp 
            Caption         =   "▲"
            Height          =   315
            Left            =   915
            TabIndex        =   74
            Top             =   3060
            Width           =   315
         End
         Begin VB.ListBox lstBMP 
            Height          =   2580
            Left            =   0
            OLEDropMode     =   1  '手動
            TabIndex        =   72
            Top             =   120
            Width           =   3195
         End
         Begin VB.CommandButton cmdBMPDelete 
            Caption         =   "消去"
            Height          =   315
            Left            =   1890
            TabIndex        =   76
            Top             =   3060
            Width           =   615
         End
         Begin VB.CommandButton cmdBMPLoad 
            Caption         =   "指定"
            Height          =   315
            Left            =   2580
            TabIndex        =   77
            Top             =   3060
            Width           =   615
         End
         Begin VB.CommandButton cmdBMPPreview 
            Caption         =   "表示"
            Height          =   315
            Left            =   0
            TabIndex        =   73
            Top             =   3060
            Width           =   795
         End
      End
      Begin VB.Frame fraBottom 
         Height          =   3495
         Index           =   2
         Left            =   6660
         TabIndex        =   78
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CommandButton cmdBGAExcDown 
            Caption         =   "▼"
            Height          =   315
            Left            =   1290
            TabIndex        =   83
            Top             =   3060
            Width           =   315
         End
         Begin VB.CommandButton cmdBGAExcUp 
            Caption         =   "▲"
            Height          =   315
            Left            =   915
            TabIndex        =   82
            Top             =   3060
            Width           =   315
         End
         Begin VB.TextBox txtBGAInput 
            Height          =   315
            Left            =   0
            TabIndex        =   80
            Top             =   2640
            Width           =   3195
         End
         Begin VB.CommandButton cmdBGAPreview 
            Caption         =   "表示"
            Height          =   315
            Left            =   0
            TabIndex        =   81
            Top             =   3075
            Width           =   795
         End
         Begin VB.CommandButton cmdBGASet 
            Caption         =   "入力"
            Height          =   315
            Left            =   2580
            TabIndex        =   85
            Top             =   3060
            Width           =   615
         End
         Begin VB.CommandButton cmdBGADelete 
            Caption         =   "消去"
            Height          =   315
            Left            =   1890
            TabIndex        =   84
            Top             =   3060
            Width           =   615
         End
         Begin VB.ListBox lstBGA 
            Height          =   2220
            Left            =   0
            TabIndex        =   79
            Top             =   120
            Width           =   3195
         End
      End
      Begin VB.Frame fraBottom 
         Height          =   3495
         Index           =   3
         Left            =   9960
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         Begin VB.ComboBox cboNumerator 
            Height          =   300
            Left            =   1125
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   89
            Top             =   3060
            Width           =   615
         End
         Begin VB.ComboBox cboDenominator 
            Height          =   300
            ItemData        =   "frmMain.frx":6957
            Left            =   1905
            List            =   "frmMain.frx":696D
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   91
            Top             =   3060
            Width           =   615
         End
         Begin VB.CommandButton cmdMeasureSelectAll 
            Caption         =   "全選択"
            Height          =   315
            Left            =   0
            TabIndex        =   88
            Top             =   3060
            Width           =   795
         End
         Begin VB.CommandButton cmdInputMeasureLen 
            Caption         =   "入力"
            Height          =   315
            Left            =   2580
            TabIndex        =   92
            Top             =   3060
            Width           =   615
         End
         Begin VB.ListBox lstMeasureLen 
            Height          =   2580
            Left            =   0
            MultiSelect     =   2  '拡張
            TabIndex        =   87
            Top             =   120
            Width           =   3195
         End
         Begin VB.Label lblFraction 
            Caption         =   "/"
            Height          =   180
            Left            =   1785
            TabIndex        =   90
            Top             =   3120
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  '塗りつぶし
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      OLEDropMode     =   1  '手動
      ScaleHeight     =   33
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
      ScaleWidth      =   57
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   855
   End
   Begin VB.HScrollBar hsbMain 
      Height          =   255
      LargeChange     =   128
      Left            =   0
      Max             =   0
      SmallChange     =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1695
   End
   Begin VB.VScrollBar vsbMain 
      Height          =   5955
      LargeChange     =   8
      Left            =   1680
      Max             =   0
      Min             =   64
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   13200
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BMS ﾌｧｲﾙ (*.bms,*.bme,*.bml)|*.bms;*.bme;*.bml"
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  '上揃え
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   18705
      _ExtentX        =   32994
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "新規作成"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "開く"
            ImageIndex      =   2
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reload"
            Description     =   "Reload"
            Object.ToolTipText     =   "再読み込み"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "上書き保存"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Description     =   "Save As"
            Object.ToolTipText     =   "名前を付けて保存"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SepMode"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Edit"
            Description     =   "Edit Mode"
            Object.ToolTipText     =   "編集"
            ImageIndex      =   6
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Write"
            Description     =   "Write Mode"
            Object.ToolTipText     =   "書込"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Delete"
            Object.ToolTipText     =   "消去"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SepViewer"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Viewer"
            Description     =   "Viewer"
            Style           =   4
            Object.Width           =   1395
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayAll"
            Description     =   "Play All"
            Object.ToolTipText     =   "最初から再生"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Play"
            Description     =   "Play"
            Object.ToolTipText     =   "現在位置から再生"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Description     =   "Stop"
            Object.ToolTipText     =   "停止"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SepGrid"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChangeGrid"
            Description     =   "Grid"
            Style           =   4
            Object.Width           =   2955
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SepSize"
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DispSize"
            Description     =   "Size"
            Style           =   4
            Object.Width           =   2955
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SepResolution"
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resolution"
            Style           =   4
            Object.Width           =   2055
         EndProperty
      EndProperty
   End
   Begin VB.Line linStatusBar 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   18660
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line linStatusBar 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   18660
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Line linToolbarBottom 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   18660
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line linHeader 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   2040
      X2              =   18720
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line linVertical 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   2040
      X2              =   2040
      Y1              =   420
      Y2              =   7340
   End
   Begin VB.Line linDirectInput 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   2040
      X2              =   0
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line linDirectInput 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   2040
      X2              =   0
      Y1              =   6855
      Y2              =   6855
   End
   Begin VB.Line linHeader 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   2040
      X2              =   18720
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line linVertical 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   2055
      X2              =   2055
      Y1              =   420
      Y2              =   7340
   End
   Begin VB.Line linToolbarBottom 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   18660
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Label lblDirectInput 
      AutoSize        =   -1  'True
      Caption         =   "Direct"
      Height          =   180
      Left            =   60
      TabIndex        =   95
      Top             =   6870
      Width           =   540
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Begin VB.Menu mnuFileNew 
         Caption         =   "mnuFileNew"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "mnuFileOpen"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "mnuFileSave"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "mnuFileSaveAs"
      End
      Begin VB.Menu mnuFileOpenDirectory 
         Caption         =   "mnuFileOpenDirectory"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLineFile 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&1:"
         Index           =   0
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&2:"
         Index           =   1
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&3:"
         Index           =   2
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&4:"
         Index           =   3
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&5:"
         Index           =   4
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&6:"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&7:"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&8:"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&9:"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "&0:"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLineRecent 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileConvertWizard 
         Caption         =   "mnuFileConvertWizard"
      End
      Begin VB.Menu mnuLineExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "mnuFileExit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "mnuEditUndo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "mnuEditRedo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuLineEdit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "mnuEditCut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "mnuEditCopy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "mnuEditPaste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "mnuEditDelete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuLineEdit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "mnuEditSelectAll"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuLineEdit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "mnuEditFind"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLineEdit4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMode 
         Caption         =   "mnuEditMode(0)"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditMode 
         Caption         =   "mnuEditMode(1)"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditMode 
         Caption         =   "mnuEditMode(2)"
         Index           =   2
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "mnuView"
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "mnuViewToolBar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewDirectInput 
         Caption         =   "mnuViewDirectInput"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "mnuViewStatusBar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "mnuOptions"
      Begin VB.Menu mnuOptionsActiveIgnore 
         Caption         =   "mnuOptionsActiveIgnore"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsFileNameOnly 
         Caption         =   "mnuOptionsFileNameOnly"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsVertical 
         Caption         =   "mnuOptionsVertical"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsLaneBG 
         Caption         =   "mnuOptionsLaneBG"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsSelectPreview 
         Caption         =   "mnuOptionsSelectPreview"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsMoveOnGrid 
         Caption         =   "mnuOptionsMoveOnGrid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsObjectFileName 
         Caption         =   "mnuOptionsObjectFileName"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsNumFF 
         Caption         =   "mnuOptionsNumFF"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsRightClickDelete 
         Caption         =   "mnuOptionsRightClickDelete"
      End
      Begin VB.Menu mnuLineOptions 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLanguageParent 
         Caption         =   "Select &Language"
         Begin VB.Menu mnuLanguage 
            Caption         =   "mnuLanguage(0)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuThemeParent 
         Caption         =   "Select &Theme"
         Begin VB.Menu mnuTheme 
            Caption         =   "mnuTheme(0)"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "mnuTools"
      Begin VB.Menu mnuToolsPlayAll 
         Caption         =   "mnuToolsPlayAll"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuToolsPlay 
         Caption         =   "mnuToolsPlay"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuToolsPlayStop 
         Caption         =   "mnuToolsPlayStop"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuLineTools 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSetting 
         Caption         =   "mnuToolsSetting"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "mnuHelp"
      Begin VB.Menu mnuHelpOpen 
         Caption         =   "mnuHelpOpen"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLineHelp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "mnuHelpWeb"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "mnuHelpAbout"
      End
   End
   Begin VB.Menu mnuContext 
      Caption         =   "mnuContext"
      Begin VB.Menu mnuContextInsertMeasure 
         Caption         =   "mnuContextInsertMeasure"
      End
      Begin VB.Menu mnuContextDeleteMeasure 
         Caption         =   "mnuContextDeleteMeasure"
      End
      Begin VB.Menu mnuContextBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContextEditCut 
         Caption         =   "mnuContextEditCut"
      End
      Begin VB.Menu mnuContextEditCopy 
         Caption         =   "mnuContextEditCopy"
      End
      Begin VB.Menu mnuContextEditPaste 
         Caption         =   "mnuContextEditPaste"
      End
      Begin VB.Menu mnuContextEditDelete 
         Caption         =   "mnuContextEditDelete"
      End
   End
   Begin VB.Menu mnuContextList 
      Caption         =   "mnuContextList"
      Begin VB.Menu mnuContextListLoad 
         Caption         =   "mnuContextListLoad"
      End
      Begin VB.Menu mnuContextListDelete 
         Caption         =   "mnuContextListDelete"
      End
      Begin VB.Menu mnuContextListRename 
         Caption         =   "mnuContextListRename"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_intScrollDir  As Integer

Private m_retObj()      As g_udtObj

Private m_blnMouseDown  As Boolean

Private m_blnPreview As Boolean

Public Function lngFromString(ByVal str As String) As Long

    If mnuOptionsNumFF.Checked Then
    
        lngFromString = Val("&H" & str)
    
    Else
    
        lngFromString = modInput.lngNumConv(str)
    
    End If

End Function

Public Function lngFromLong(ByVal value As Long) As Long

    If mnuOptionsNumFF.Checked Then
    
        lngFromLong = modInput.lngNumConv(strFromLong(value))
    
    Else
    
        lngFromLong = value
    
    End If

End Function

Public Function strFromLong(ByVal value As Long) As String

    If mnuOptionsNumFF.Checked Then
    
        strFromLong = Right$("0" & Hex$(value), 2)
    
    Else
    
        strFromLong = modInput.strNumConv(value)
    
    End If

End Function

Private Sub MoveObj(ByVal X As Single, ByVal Y As Single, ByVal Shift As Integer)
On Error GoTo Err:

    Dim i       As Long
    Dim j       As Long
    Dim k       As Long
    Dim lngRet  As Long
    Dim oldObj  As g_udtObj
    Dim newObj  As g_udtObj
    
    With newObj
    
        Call modDraw.SetObjData(newObj, X, Y) ', g_disp.X, g_disp.Y)
        
        .intCh = g_intVGridNum(.intCh)
        
        'If Not Shift And vbAltMask Then
        
            If frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex) Then
            
                lngRet = 192 \ (frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex))
                .lngPosition = (.lngPosition \ lngRet) * lngRet
                
                'If Not Shift And vbShiftMask Then
                
                    If frmMain.mnuOptionsMoveOnGrid.Checked Then
                    
                        With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                        
                            lngRet = .lngPosition - (.lngPosition \ lngRet) * lngRet
                        
                        End With
                        
                        .lngPosition = .lngPosition - lngRet
                    
                    End If
                
                'End If
            
            End If
        
        'End If
        
        .lngPosition = .lngPosition + g_Measure(.intMeasure).lngY
    
    End With
    
    Call modDraw.CopyObj(oldObj, g_Obj(g_Obj(UBound(g_Obj)).lngHeight))
    
    With oldObj
    
        .intCh = g_intVGridNum(.intCh)
        
        'If Not Shift And vbAltMask Then
        
            If frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex) Then
            
                lngRet = 192 \ frmMain.cboDispGridSub.ItemData(frmMain.cboDispGridSub.ListIndex)
                .lngPosition = (.lngPosition \ lngRet) * lngRet
            
            End If
        
        'End If
        
        .lngPosition = .lngPosition + g_Measure(.intMeasure).lngY
    
    End With
    
    'Y 軸固定移動
    If Shift And vbShiftMask Then
    
        newObj.lngPosition = oldObj.lngPosition
    
    End If
    
    If newObj.intCh <> oldObj.intCh Or newObj.lngPosition <> oldObj.lngPosition Then
    
        If newObj.intCh > oldObj.intCh Then
        
            For j = oldObj.intCh To newObj.intCh - 1
            
                If g_VGrid(j).blnDraw = True And g_VGrid(j).intCh <> 0 Then
                
                    newObj.intAtt = newObj.intAtt + 1
                
                End If
            
            Next j
        
        ElseIf newObj.intCh < oldObj.intCh Then
        
            For j = oldObj.intCh To newObj.intCh + 1 Step -1
            
                If g_VGrid(j).blnVisible = True And g_VGrid(j).intCh <> 0 Then
                
                    newObj.intAtt = newObj.intAtt + 1
                
                End If
            
            Next j
        
        End If
        
        lngRet = newObj.intCh <> oldObj.intCh And newObj.intCh <> 0 And oldObj.intCh <> 0 And newObj.intCh <> UBound(g_VGrid) And oldObj.intCh <> UBound(g_VGrid)
        
        For i = 0 To UBound(g_Obj) - 1
        
            With g_Obj(i)
            
                If .intSelect = 1 Then
                
                    .lngPosition = .lngPosition + newObj.lngPosition - oldObj.lngPosition
                    
                    Do While .lngPosition >= g_Measure(.intMeasure).intLen
                    
                        If .intMeasure < 999 Then
                        
                            .lngPosition = .lngPosition - g_Measure(.intMeasure).intLen
                            .intMeasure = .intMeasure + 1
                        
                        Else
                        
                            .intMeasure = 999
                            
                            Exit Do
                        
                        End If
                    
                    Loop
                    
                    Do While .lngPosition < 0
                    
                        If .intMeasure > 0 Then
                        
                            .lngPosition = g_Measure(.intMeasure - 1).intLen + .lngPosition
                            .intMeasure = .intMeasure - 1
                        
                        Else
                        
                            .intMeasure = 0
                            
                            Exit Do
                        
                        End If
                    
                    Loop
                    
                    If lngRet Then
                    
                        If .intCh < 0 Then
                        
                            j = .intCh
                        
                        ElseIf .intCh > 1000 Then
                        
                            j = .intCh - 1000
                        
                        Else
                        
                            j = g_intVGridNum(.intCh)
                        
                        End If
                        
                        If newObj.intCh > oldObj.intCh Then
                        
                            For k = 1 To newObj.intAtt
                            
                                Do
                                
                                    j = j + 1
                                    
                                    If j < 0 Or j > UBound(g_VGrid) Then Exit Do
                                    
                                    If g_VGrid(j).blnVisible = True And g_VGrid(j).intCh <> 0 Then
                                    
                                        Exit Do
                                    
                                    End If
                                
                                Loop
                            
                            Next k
                        
                        Else
                        
                            For k = 1 To newObj.intAtt
                            
                                Do
                                
                                    j = j - 1
                                    
                                    If j < 0 Or j > UBound(g_VGrid) Then Exit Do
                                    
                                    If g_VGrid(j).blnVisible = True And g_VGrid(j).intCh <> 0 Then
                                    
                                        Exit Do
                                    
                                    End If
                                
                                Loop
                            
                            Next k
                        
                        End If
                        
                        If j < 0 Then
                        
                            .intCh = j
                        
                        ElseIf j > UBound(g_VGrid) Then
                        
                            .intCh = 1000 + j
                        
                        Else
                        
                            .intCh = g_VGrid(j).intCh
                        
                        End If
                        
                        Select Case .intCh
                        
                            Case 8
                            
                                '何もしない
                            
                            Case 9
                            
                                .sngValue = CLng(.sngValue)
                                
                                If .sngValue < 0 Then
                                
                                    .sngValue = 1
                                
                                End If
                            
                            Case Else
                            
                                .sngValue = CLng(.sngValue)
                                
                                If .sngValue < 0 Then
                                
                                    .sngValue = 1
                                
                                ElseIf .sngValue > 1295 Then
                                
                                    .sngValue = 1295
                                
                                End If
                        
                        End Select
                    
                    End If
                
                End If
            
            End With
        
        Next i
        
        'Call modDraw.DrawStatusBar(g_Obj(UBound(g_Obj)).lngHeight, Shift)
        Call modDraw.DrawStatusBar(g_Obj(g_Obj(UBound(g_Obj)).lngHeight), Shift)
        
        'Call SaveChanges
        
        Call modDraw.Redraw
    
    End If
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "MoveObj")
End Sub

Public Sub PreviewBMP(ByVal strFileName As String)
On Error GoTo Err:

    If m_blnPreview = False Then Exit Sub
    
    With frmWindowPreview
    
        If .chkLock.value Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            '.txtBGAPara(BGA_NUM).Text = Right$("0" & Hex$(lstBMP.ListIndex + 1), 2)
            '.Caption = Right$("0" & Hex$(lstBMP.ListIndex + 1), 2) & ":" & strFileName
        
        'Else
        
            '.txtBGAPara(BGA_NUM).Text = modInput.strNumConv(lstBMP.ListIndex + 1)
            '.Caption = modInput.strNumConv(lstBMP.ListIndex + 1) & ":" & strFileName
        
        'End If
        
        .txtBGAPara(BGA_NUM).Text = strFromLong(lstBMP.ListIndex + 1)
        .Caption = .txtBGAPara(BGA_NUM).Text & ":" & strFileName
        
        If Mid$(strFileName, 2, 2) <> ":\" Then strFileName = g_BMS.strDir & strFileName
        
        If Len(strFileName) <> 0 And strFileName <> g_BMS.strDir And Dir(strFileName, vbNormal) <> vbNullString Then
        
            Set .picBackBuffer = LoadPicture(strFileName)
        
        Else
        
            Set .picBackBuffer = LoadPicture()
        
        End If
        
        If .picBackBuffer.ScaleWidth <= 256 Then
        
            .txtBGAPara(BGA_X1).Text = 0
            .txtBGAPara(BGA_X2).Text = .picBackBuffer.ScaleWidth
            .txtBGAPara(BGA_dX).Text = (256 - .picBackBuffer.ScaleWidth) \ 2
        
        Else
        
            .txtBGAPara(BGA_X1).Text = (.picBackBuffer.ScaleWidth - 256) \ 2
            .txtBGAPara(BGA_X2).Text = .txtBGAPara(BGA_X1).Text + 256
            .txtBGAPara(BGA_dX).Text = 0
        
        End If
        
        .txtBGAPara(BGA_Y1).Text = 0
        
        If .picBackBuffer.ScaleHeight >= 256 Then
        
            .txtBGAPara(BGA_Y2).Text = 256
        
        Else
        
            .txtBGAPara(BGA_Y2).Text = .picBackBuffer.ScaleHeight
        
        End If
        
        .txtBGAPara(BGA_dY).Text = 0
        
        Call .picPreview_Paint
    
    End With

Err:
End Sub

Private Sub PreviewBGA(ByVal lngFileNum As Long)
On Error GoTo Err:

    Dim strRet      As String
    Dim strArray()  As String
    
    If m_blnPreview = False Then Exit Sub
    
    With frmWindowPreview
    
        If .chkLock.value Then Exit Sub
        
        'strRet = Trim$(Mid$(lstBGA.List(lstBGA.ListIndex), 8))
        strRet = g_strBGA(lngFileNum)
        strArray = Split(strRet, " ")
        
        If Len(strRet) Then
        
            strRet = g_strBMP(modInput.lngNumConv(strArray(0)))
            
            If Mid$(strRet, 2, 2) <> ":\" Then strRet = g_BMS.strDir & strRet
            
            If Dir(strRet, vbNormal) <> vbNullString Then
            
                Set .picBackBuffer = LoadPicture(strRet)
            
            Else
            
                Set .picBackBuffer = LoadPicture()
            
            End If
        
        Else
        
            Set .picBackBuffer = LoadPicture()
        
        End If
        
        If UBound(strArray) > 5 Then
        
            .txtBGAPara(BGA_NUM).Text = strArray(0)
            .txtBGAPara(BGA_X1).Text = strArray(1)
            If .txtBGAPara(BGA_X1).Text < 0 Then .txtBGAPara(BGA_X1).Text = 0
            .txtBGAPara(BGA_Y1).Text = strArray(2)
            If .txtBGAPara(BGA_Y1).Text < 0 Then .txtBGAPara(BGA_Y1).Text = 0
            .txtBGAPara(BGA_X2).Text = strArray(3)
            If .txtBGAPara(BGA_X2).Text < 0 Then .txtBGAPara(BGA_X2).Text = 0
            .txtBGAPara(BGA_Y2).Text = strArray(4)
            If .txtBGAPara(BGA_Y2).Text < 0 Then .txtBGAPara(BGA_Y2).Text = 0
            .txtBGAPara(BGA_dX).Text = strArray(5)
            If .txtBGAPara(BGA_dX).Text < 0 Then .txtBGAPara(BGA_dX).Text = 0
            .txtBGAPara(BGA_dY).Text = strArray(6)
            If .txtBGAPara(BGA_dY).Text < 0 Then .txtBGAPara(BGA_dY).Text = 0
             
            'If mnuOptionsNumFF.Checked Then
             
                '.Caption = Right$("0" & Hex$(lstBGA.ListIndex + 1), 2) & ":" & strRet
            
            'Else
             
                '.Caption = modInput.strNumConv(lstBGA.ListIndex + 1) & ":" & strRet
            
            'End If
            
            .Caption = strFromLong(lstBGA.ListIndex + 1) & ":" & strRet
        
        End If
        
        Call .picPreview_Paint
    
    End With

Err:
End Sub

Private Sub PreviewWAV(ByVal strFileName As String)
On Error GoTo Err:

    Dim lngError    As Long
    Dim strError    As String * 256
    Dim strRet      As String
    
    If m_blnPreview = False Then Exit Sub
    
    If Mid$(strFileName, 2, 2) <> ":\" Then
    
        strFileName = g_BMS.strDir & strFileName
    
    End If
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    Dim strArray()  As String
    strArray() = Split(strFileName, ".")
    
    Select Case UCase$(strArray(UBound(strArray)))
        Case "WAV": strRet = " type WaveAudio"
        Case "MP3", "OGG": strRet = " type MPEGVideo"
    End Select
    
    lngError = mciSendString("open " & Chr$(34) & strFileName & Chr$(34) & strRet & " alias PREVIEW", vbNullString, 0, 0)
    
    If lngError Then
    
        If mciGetErrorString(lngError, strError, 256) Then
        
            strRet = Left$(strError, InStr(strError, Chr$(0)) - 1)
        
        Else
        
            strRet = "不明なエラーです。"
        
        End If
        
        'Call modMain.DebugOutput(lngError, strRet & Chr$(34) & strFileName & Chr$(34), "PreviewWAV", False)
    
    End If
    
    Call mciSendString("play PREVIEW notify", vbNullString, 0, frmMain.hwnd)

Err:
End Sub

Private Sub FormDragDrop(ByVal Data As DataObject)

    Dim i           As Long
    Dim strArray()  As String
    Dim strRet      As String
    Dim blnReadFlag As Boolean
    
    For i = 1 To Data.Files.Count
    
        If Dir(Data.Files.Item(i), vbNormal) <> vbNullString Then
        
            strRet = Data.Files.Item(i)
            
            If Right$(UCase$(strRet), 4) = ".BMS" Or Right$(UCase$(strRet), 4) = ".BME" Or Right$(UCase$(strRet), 4) = ".BML" Or Right$(UCase$(strRet), 4) = ".PMS" Then
            
                If modMain.intSaveCheck() Or blnReadFlag Then
                
                    Call ShellExecute(0, "open", Chr$(34) & g_strAppDir & App.EXEName & Chr$(34), Chr$(34) & strRet & Chr$(34), "", SW_SHOWNORMAL)
                
                Else
                
                    Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
                    
                    strArray() = Split(strRet, "\")
                    g_BMS.strFileName = Right$(strRet, Len(strArray(UBound(strArray))))
                    g_BMS.strDir = Left$(strRet, Len(strRet) - Len(strArray(UBound(strArray))))
                    dlgMain.InitDir = g_BMS.strDir
                    blnReadFlag = True
                    
                    Call modInput.LoadBMS
                    Call modMain.RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
                
                End If
            
            End If
        
        End If
    
    Next i

End Sub

Private Sub CopyToClipboard()

    Dim i           As Long
    Dim intRet      As Integer
    Dim lngRet      As Long
    Dim strArray()  As String
    
    Call Clipboard.Clear
    
    intRet = 999
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intSelect = 1 Then
            
                If .intMeasure < intRet Then intRet = .intMeasure
                
                lngRet = lngRet + 1
            
            End If
        
        End With
    
    Next i
    
    ReDim strArray(lngRet - 1)
    lngRet = 0
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intSelect = 1 Then
            
                strArray(lngRet) = Format$(.intCh, "000") & .intAtt & Format$(g_Measure(.intMeasure).lngY + .lngPosition - g_Measure(intRet).lngY, "0000000") & .sngValue
                lngRet = lngRet + 1
            
            End If
        
        End With
    
    Next i
    
    Call Clipboard.SetText("BMSE ClipBoard Object Data Format" & vbCrLf & Join(strArray, vbCrLf) & vbCrLf)

End Sub

Public Sub SaveChanges()

    g_BMS.blnSaveFlag = False
    
    frmMain.Caption = g_strAppTitle
    
    If Len(g_BMS.strDir) Then
    
        If mnuOptionsFileNameOnly.Checked Then
        
            frmMain.Caption = frmMain.Caption & " - " & g_BMS.strFileName
            
        Else
        
            frmMain.Caption = frmMain.Caption & " - " & g_BMS.strDir & g_BMS.strFileName
        
        End If
    
    End If
    
    frmMain.Caption = frmMain.Caption & " *"

End Sub

Private Function strCmdDecode(ByVal strCmd As String, ByVal strFileName As String) As String

    strCmd = Replace$(strCmd, "<filename>", Chr$(34) & strFileName & Chr$(34))
    strCmd = Replace$(strCmd, "<measure>", g_disp.intStartMeasure)
    strCmd = Replace$(strCmd, "<appdir>", g_strAppDir)
    
    strCmdDecode = strCmd

End Function

Public Sub RefreshList()

    Dim i       As Long
    Dim strRet  As String * 2
    Dim lngRet  As Long
    Dim lngIndex(2) As Long
    
    lngIndex(0) = lstWAV.ListIndex
    lngIndex(1) = lstBMP.ListIndex
    lngIndex(2) = lstBGA.ListIndex
    
    lstWAV.Visible = False
    lstBMP.Visible = False
    lstBGA.Visible = False
    
    lstWAV.Clear
    lstBMP.Clear
    lstBGA.Clear
    
    If mnuOptionsNumFF.Checked Then
    
        For i = 1 To 255
        
            'strRet = Right$("0" & Hex$(i), 2)
            'lngRet = modInput.lngNumConv(strRet)
            
            strRet = strFromLong(i)
            lngRet = lngFromLong(i)
            
            lstWAV.List(i - 1) = "#WAV" & strRet & ":" & g_strWAV(lngRet)
            lstBMP.List(i - 1) = "#BMP" & strRet & ":" & g_strBMP(lngRet)
            lstBGA.List(i - 1) = "#BGA" & strRet & ":" & g_strBGA(lngRet)
        
        Next i
        
        frmWindowPreview.cmdPreviewEnd.Caption = "FF"
    
    Else
    
        For i = 1 To 1295
        
            strRet = modInput.strNumConv(i)
            
            lstWAV.List(i - 1) = "#WAV" & strRet & ":" & g_strWAV(i)
            lstBMP.List(i - 1) = "#BMP" & strRet & ":" & g_strBMP(i)
            lstBGA.List(i - 1) = "#BGA" & strRet & ":" & g_strBGA(i)
        
        Next i
        
        frmWindowPreview.cmdPreviewEnd.Caption = "ZZ"
    
    End If
    
    m_blnPreview = False
    lstWAV.ListIndex = lngIndex(0)
    lstBMP.ListIndex = lngIndex(1)
    lstBGA.ListIndex = lngIndex(2)
    m_blnPreview = True
    
    lstWAV.Visible = True
    lstBMP.Visible = True
    lstBGA.Visible = True

End Sub

Private Sub cboDirectInput_GotFocus()

    cmdDirectInput.Default = True

End Sub

Private Sub cboDirectInput_LostFocus()

    cmdDirectInput.Default = False

End Sub

Private Sub cboDispFrame_Click()

    Call modDraw.InitVerticalLine

End Sub

Private Sub cboDispGridMain_Click()

    Call modDraw.Redraw

End Sub

Private Sub cboDispGridSub_Click()

    Call modDraw.Redraw

End Sub

Private Sub cboDispHeight_Click()

    Dim i       As Long
    Dim sngRet  As Single
    
    If frmMain.Visible = False Then Exit Sub
    
    If cboDispHeight.ListIndex = cboDispHeight.ListCount - 1 Then
    
        With frmWindowInput
        
            .lblMainDisp.Caption = g_Message(Message.INPUT_SIZE)
            .txtMain.Text = Format$(g_disp.Height, "#0.00")
            If .txtMain.Text = "100.00" Then .txtMain.Text = "1.00"
            
            Call .Show(vbModal, frmMain)
            
            sngRet = Round(Val(.txtMain.Text), 2)
            
            If sngRet <= 0 Then
            
                sngRet = g_disp.Height
            
            ElseIf sngRet > 16 Then
            
                sngRet = 16
            
            End If
            
            For i = 0 To cboDispHeight.ListCount - 1
            
                If cboDispHeight.ItemData(i) = sngRet * 100 Then
                
                    cboDispHeight.ListIndex = i
                    
                    Exit For
                
                ElseIf cboDispHeight.ItemData(i) > sngRet * 100 Then
                
                    Call cboDispHeight.AddItem("x" & Format$(sngRet, "#0.00"), i)
                    cboDispHeight.ItemData(i) = sngRet * 100
                    cboDispHeight.ListIndex = i
                    
                    Exit For
                
                End If
            
            Next i
        
        End With
    
    End If
    
    Call modDraw.Redraw

End Sub

Private Sub cboDispKey_Click()

    Call modDraw.InitVerticalLine

End Sub

Private Sub cboDispSC1P_Click()

    Call modDraw.InitVerticalLine

End Sub

Private Sub cboDispSC2P_Click()

    Call modDraw.InitVerticalLine

End Sub

Private Sub cboDispWidth_Click()

    Dim i       As Long
    Dim sngRet  As Single
    
    If frmMain.Visible = False Then Exit Sub
    
    If cboDispWidth.ListIndex = cboDispWidth.ListCount - 1 Then
    
        With frmWindowInput
        
            .lblMainDisp.Caption = g_Message(Message.INPUT_SIZE)
            .txtMain.Text = Format$(g_disp.Width, "#0.00")
            If .txtMain.Text = "100.00" Then .txtMain.Text = "1.00"
            
            Call .Show(vbModal, frmMain)
            
            sngRet = Round(Val(.txtMain.Text), 2)
            
            If sngRet <= 0 Then
            
                sngRet = g_disp.Width
            
            ElseIf sngRet > 16 Then
            
                sngRet = 16
            
            End If
            
            For i = 0 To cboDispWidth.ListCount - 1
            
                If cboDispWidth.ItemData(i) = sngRet * 100 Then
                
                    cboDispWidth.ListIndex = i
                    
                    Exit For
                
                ElseIf cboDispWidth.ItemData(i) > sngRet * 100 Then
                
                    Call cboDispWidth.AddItem("x" & Format$(sngRet, "#0.00"), i)
                    cboDispWidth.ItemData(i) = sngRet * 100
                    cboDispWidth.ListIndex = i
                    
                    Exit For
                
                End If
            
            Next i
        
        End With
    
    End If
    
    Call modDraw.Redraw

End Sub

Private Sub cboPlayer_Click()

    Call modDraw.InitVerticalLine

End Sub

Private Sub cboVScroll_Click()

    vsbMain.SmallChange = cboVScroll.ItemData(cboVScroll.ListIndex)
    vsbMain.LargeChange = frmMain.vsbMain.SmallChange * 8

End Sub

Private Sub cmdBGAExcDown_Click()

    Call cmdBMPExcDown_Click

End Sub

Private Sub cmdBGAExcUp_Click()

    Call cmdBMPExcUp_Click

End Sub

Private Sub cmdBGAPreview_Click()

    Call PreviewBGA(lstBGA.ListIndex + 1)
    
    With frmWindowPreview
    
        If Not .Visible Then
        
            'Call .SetWindowSize
            '.Left = frmMain.Left + (frmMain.Width - .Width) \ 2
            '.Top = frmMain.Top + (frmMain.Height - .Height) \ 2
            
            Call .Show(0, frmMain)
        
        Else
        
            Call Unload(frmWindowPreview)
        
        End If
    
    End With

End Sub

Private Sub cmdBMPExcDown_Click()

    Dim i           As Long
    Dim lngChangeA  As Long
    Dim lngChangeB  As Long
    Dim strRet      As String
    Dim lngIndex    As Long
    Dim strArray()  As String
    
    With lstBMP
    
        If .ListIndex = .ListCount - 1 Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            'lngChangeB = modInput.lngNumConv(Hex$(.ListIndex + 2))
            'lngChangeA = modInput.lngNumConv(Hex$(.ListIndex + 1))
        
        'Else
        
            'lngChangeB = .ListIndex + 2
            'lngChangeA = .ListIndex + 1
        
        'End If
        
        lngChangeB = lngFromLong(.ListIndex + 2)
        lngChangeA = lngFromLong(.ListIndex + 1)
        
        strRet = g_strBMP(lngChangeB)
        g_strBMP(lngChangeB) = g_strBMP(lngChangeA)
        g_strBMP(lngChangeA) = strRet
        
        lngIndex = .ListIndex + 1
    
    End With
    
    With lstBGA
    
        strRet = g_strBGA(lngChangeB)
        g_strBGA(lngChangeB) = g_strBGA(lngChangeA)
        g_strBGA(lngChangeA) = strRet
    
    End With
    
    For i = 0 To UBound(g_strBGA)
    
        If Len(g_strBGA(i)) Then
        
            strArray() = Split(g_strBGA(i), " ")
            
            If UBound(strArray) Then
            
                If modInput.lngNumConv(strArray(0)) = lngChangeB Then
                
                    strArray(0) = modInput.strNumConv(lngChangeA, 2)
                
                ElseIf modInput.lngNumConv(strArray(0)) = lngChangeA Then
                
                    strArray(0) = modInput.strNumConv(lngChangeB, 2)
                
                End If
                
                g_strBGA(i) = Join(strArray(), " ")
            
            End If
        
        End If
    
    Next i
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intCh = 4 Or .intCh = 6 Or .intCh = 7 Then
            
                If .sngValue = lngChangeA Then
                
                    .sngValue = lngChangeB
                
                ElseIf .sngValue = lngChangeB Then
                
                    .sngValue = lngChangeA
                
                End If
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.BMP_CHANGE) & modInput.strNumConv(lngChangeB) & modInput.strNumConv(lngChangeA) & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.BMP_CHANGE) & modInput.strNumConv(lngChangeB) & modInput.strNumConv(lngChangeA) & ",")
    
    Call RefreshList
    
    Call Redraw
    
    lstBMP.ListIndex = lngIndex
    lstBGA.ListIndex = lngIndex

End Sub

Private Sub cmdBMPExcUp_Click()

    Dim i           As Long
    Dim lngChangeA  As Long
    Dim lngChangeB  As Long
    Dim strRet      As String
    Dim lngIndex    As Long
    Dim strArray()  As String
    
    With lstBMP
    
        If .ListIndex = 0 Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            'lngChangeB = modInput.lngNumConv(Hex$(.ListIndex))
            'lngChangeA = modInput.lngNumConv(Hex$(.ListIndex + 1))
        
        'Else
        
            'lngChangeB = .ListIndex
            'lngChangeA = .ListIndex + 1
        
        'End If
        
        lngChangeB = lngFromLong(.ListIndex)
        lngChangeA = lngFromLong(.ListIndex + 1)
        
        strRet = g_strBMP(lngChangeB)
        g_strBMP(lngChangeB) = g_strBMP(lngChangeA)
        g_strBMP(lngChangeA) = strRet
        
        lngIndex = .ListIndex - 1
    
    End With
    
    With lstBGA
    
        strRet = g_strBGA(lngChangeB)
        g_strBGA(lngChangeB) = g_strBGA(lngChangeA)
        g_strBGA(lngChangeA) = strRet
    
    End With
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intCh = 4 Or .intCh = 6 Or .intCh = 7 Then
            
                If .sngValue = lngChangeA Then
                
                    .sngValue = lngChangeB
                
                ElseIf .sngValue = lngChangeB Then
                
                    .sngValue = lngChangeA
                
                End If
            
            End If
        
        End With
    
    Next i
    
    For i = 0 To UBound(g_strBGA)
    
        If Len(g_strBGA(i)) Then
        
            strArray() = Split(g_strBGA(i), " ")
            
            If UBound(strArray) Then
            
                If modInput.lngNumConv(strArray(0)) = lngChangeB Then
                
                    strArray(0) = modInput.strNumConv(lngChangeA, 2)
                
                ElseIf modInput.lngNumConv(strArray(0)) = lngChangeA Then
                
                    strArray(0) = modInput.strNumConv(lngChangeB, 2)
                
                End If
                
                g_strBGA(i) = Join(strArray(), " ")
            
            End If
        
        End If
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.BMP_CHANGE) & modInput.strNumConv(lngChangeB) & modInput.strNumConv(lngChangeA) & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.BMP_CHANGE) & modInput.strNumConv(lngChangeB) & modInput.strNumConv(lngChangeA) & ",")
    
    Call RefreshList
    
    Call Redraw
    
    lstBMP.ListIndex = lngIndex
    lstBGA.ListIndex = lngIndex

End Sub

Private Sub cmdDirectInput_Click()
On Error GoTo Err:

    Dim intRet  As Integer
    Dim i       As Long
    
    With cboDirectInput
    
        If Len(.Text) Then
        
            intRet = UBound(g_Obj)
            
            Call modInput.LoadBMSDataSub(.Text, True)
            
            For i = intRet To UBound(g_Obj) - 1
            
                With g_Obj(i)
                
                    .lngPosition = (g_Measure(.intMeasure).intLen / .lngHeight) * .lngPosition
                    
                    Select Case .intCh
                    
                        Case 3
                        
                            .intCh = 8
                        
                        Case 31 To 49
                        
                            .intCh = .intCh - 20
                            .intAtt = 1
                        
                        Case 51 To 69
                        
                            .intCh = .intCh - 40
                            .intAtt = 2
                    
                    End Select
                
                End With
            
            Next i
            
            Call .AddItem(.Text, 0)
            
            For i = 1 To .ListCount
            
                If .Text = .List(i) Then
                
                    Call .RemoveItem(i)
                    
                    Exit For
                
                End If
            
            Next i
            
            If .ListCount > 10 Then Call .RemoveItem(.ListCount - 1)
            
            .Text = ""
            
            Call .SetFocus
            
            Call SaveChanges
            
            Call modDraw.Redraw
        
        End If
    
    End With
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "cmdDirectImput_Click")
End Sub

Private Sub cmdMeasureSelectAll_Click()

    Dim i       As Long
    Dim intRet  As Integer
    
    With lstMeasureLen
    
        intRet = .ListIndex
        .Visible = False
        
        For i = 0 To 999
        
            .Selected(i) = True
        
        Next i
        
        .ListIndex = intRet
        .Visible = True
    
    End With

End Sub

Private Sub cmdBGADelete_Click()

    Dim strRet  As String * 7
    
    With lstBGA
    
        If Len(.List(.ListIndex)) > 7 Then
        
            strRet = "#BGA00:"
            
            'If mnuOptionsNumFF.Checked Then
            
                'Mid$(strRet, 5, 2) = Right$("0" & Hex$(.ListIndex + 1), 2)
                'g_strBGA(lngNumConv(Mid$(strRet, 5, 2))) = ""
            
            'Else
            
                'Mid$(strRet, 5, 2) = modInput.strNumConv(.ListIndex + 1)
                'g_strBGA(.ListIndex + 1) = ""
            
            'End If
            
            Mid$(strRet, 5, 2) = strFromLong(.ListIndex + 1)
            g_strBGA(lngFromLong(.ListIndex + 1)) = ""
            
            .List(.ListIndex) = strRet
            
            Call SaveChanges
        
        End If
    
    End With

End Sub

Private Sub cmdBGASet_Click()

    'Dim strRet  As String
    
    With lstBGA
    
        If Len(txtBGAInput.Text) = 0 Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            'strRet = Right$("0" & Hex$(.ListIndex + 1), 2)
            '.List(.ListIndex) = "#BGA" & strRet & ":" & txtBGAInput.Text
            'g_strBGA(modInput.lngNumConv(strRet)) = txtBGAInput.Text
        
        'Else
        
            '.List(.ListIndex) = "#BGA" & modInput.strNumConv(.ListIndex + 1) & ":" & txtBGAInput.Text
            'g_strBGA(.ListIndex + 1) = txtBGAInput.Text
        
        'End If
        
        .List(.ListIndex) = "#BGA" & strFromLong(.ListIndex + 1) & ":" & txtBGAInput.Text
        g_strBGA(lngFromLong(.ListIndex + 1)) = txtBGAInput.Text
        
        txtBGAInput.Text = ""
        
        Call SaveChanges
    
    End With
    
    With frmWindowPreview
    
        If .Visible Then
        
            Call PreviewBGA(lstBGA.ListIndex + 1)
            Call .Show(0, frmMain)
        
        End If
    
    End With
    
End Sub

Private Sub cmdBmpDelete_Click()

    Dim strRet  As String * 7
    
    With lstBMP
    
        If Len(.List(.ListIndex)) > 7 Then
    
            strRet = "#BMP00:"
            
            'If mnuOptionsNumFF.Checked Then
            
                'Mid$(strRet, 5, 2) = Right$("0" & Hex$(.ListIndex + 1), 2)
                'g_strBMP(lngNumConv(Mid$(strRet, 5, 2))) = ""
            
            'Else
            
                'Mid$(strRet, 5, 2) = modInput.strNumConv(.ListIndex + 1)
                'g_strBMP(.ListIndex + 1) = ""
            
            'End If
            
            Mid$(strRet, 5, 2) = strFromLong(.ListIndex + 1)
            g_strBMP(lngFromLong(.ListIndex + 1)) = ""
            
            .List(.ListIndex) = strRet
            
            Call SaveChanges
        
        End If
    
    End With

End Sub

Private Sub cmdBmpLoad_Click()
On Error GoTo Err:

    Dim retArray()  As String
    'Dim strRet      As String * 2
    
    With dlgMain
    
        .Filter = "Image files (*.bmp,*.jpg,*.gif)|*.bmp;*.jpg;*.gif|All files (*.*)|*.*"
        .FileName = Mid$(lstBMP.List(lstBMP.ListIndex), 8)
        
        Call .ShowOpen
        
        retArray = Split(.FileName, "\")
        
        'If mnuOptionsNumFF.Checked Then
        
            'strRet = Right$("0" & Hex$(lstBMP.ListIndex + 1), 2)
            'lstBMP.List(lstBMP.ListIndex) = "#BMP" & strRet & ":" & retArray(UBound(retArray))
            'g_strBMP(modInput.lngNumConv(strRet)) = retArray(UBound(retArray))
        
        'Else
        
            'lstBMP.List(lstBMP.ListIndex) = "#BMP" & modInput.strNumConv(lstBMP.ListIndex + 1) & ":" & retArray(UBound(retArray))
            'g_strBMP(lstBMP.ListIndex + 1) = retArray(UBound(retArray))
        
        'End If
        
        lstBMP.List(lstBMP.ListIndex) = "#BMP" & strFromLong(lstBMP.ListIndex + 1) & ":" & retArray(UBound(retArray))
        g_strBMP(lngFromLong(lstBMP.ListIndex + 1)) = retArray(UBound(retArray))
        
        .InitDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
        
        Call SaveChanges
    
    End With
    
    With frmWindowPreview
    
        If .Visible Then
        
            Call PreviewBMP(Mid$(lstBMP.List(lstBMP.ListIndex), 8))
            Call .Show(0, frmMain)
        
        End If
    
    End With
    
    Exit Sub

Err:
End Sub

Private Sub cmdInputMeasureLen_Click()

    Dim i           As Long
    Dim lngRet      As Long
    Dim strArray()  As String
    Dim retObj      As g_udtObj
    
    With lstMeasureLen
    
        For i = 0 To 999
        
            If .Selected(i) Then
            
                lngRet = 1
                
                Exit For
            
            End If
        
        Next i
        
        If lngRet = 0 Then Exit Sub
        
        ReDim strArray(0)
        
        .Visible = False
        lngRet = 0
        
        For i = 0 To 999
        
            If .Selected(i) Then
            
                .List(i) = "#" & Format$(i, "000") & ":" & cboNumerator.Text & "/" & cboDenominator.Text
                lngRet = (192 / cboDenominator.Text) * cboNumerator.Text
                
                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.MSR_CHANGE) & modInput.strNumConv(i) & Right$("00" & Hex$(g_Measure(i).intLen), 3) & Right$("00" & Hex$(lngRet), 3)
                ReDim Preserve strArray(UBound(strArray) + 1)
                
                g_Measure(i).intLen = lngRet
            
            End If
        
        Next i
        
        .Visible = True
    
    End With
    
    For i = 0 To UBound(g_Obj) - 1
    
        Call modDraw.CopyObj(retObj, g_Obj(i))
        
        With retObj
        
            Do While .lngPosition >= g_Measure(.intMeasure).intLen
            
                .lngPosition = .lngPosition - g_Measure(.intMeasure).intLen
                .intMeasure = .intMeasure + 1
            
            Loop
        
        End With
        
        With g_Obj(i)
        
            If retObj.intMeasure > 999 Then
            
                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                ReDim Preserve strArray(UBound(strArray) + 1)
                
                Call modDraw.RemoveObj(i)
            
            Else
            
                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_MOVE) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & Right$("0" & Hex$(retObj.intCh), 2) & modInput.strNumConv(retObj.intMeasure) & modInput.strNumConv(retObj.lngPosition, 3)
                ReDim Preserve strArray(UBound(strArray) + 1)
                
                Call CopyObj(g_Obj(i), retObj)
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(Join(strArray, ",") & ",")
    
    Call modDraw.ArrangeObj
    
    Call modDraw.ChangeResolution
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub cmdBMPPreview_Click()

    Call PreviewBMP(Mid$(lstBMP.List(lstBMP.ListIndex), 8))
    
    With frmWindowPreview
    
        If Not .Visible Then
        
            'Call .SetWindowSize
            '.Left = frmMain.Left + (frmMain.Width - .Width) \ 2
            '.Top = frmMain.Top + (frmMain.Height - .Height) \ 2
            
            Call .Show(0, frmMain)
        
        Else
        
            Call Unload(frmWindowPreview)
        
        End If
    
    End With

End Sub

Private Sub cmdLoadMissBMP_Click()
On Error GoTo Err:

    Dim retArray()  As String
    
    With dlgMain
    
        .Filter = "Image files (*.bmp,*.jpg)|*.bmp;*.jpg|All files (*.*)|*.*"
        .FileName = txtStageFile.Text
        
        Call .ShowOpen
        
        retArray = Split(.FileName, "\")
        txtMissBMP.Text = retArray(UBound(retArray))
        dlgMain.InitDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
    
    End With
    
    Exit Sub

Err:
End Sub

Private Sub cmdLoadStageFile_Click()
On Error GoTo Err:

    Dim retArray()  As String
    
    With dlgMain
    
        .Filter = "Image files (*.bmp,*.jpg)|*.bmp;*.jpg|All files (*.*)|*.*"
        .FileName = txtStageFile.Text
        
        Call .ShowOpen
        
        retArray = Split(.FileName, "\")
        txtStageFile.Text = retArray(UBound(retArray))
        dlgMain.InitDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
    
    End With
    
    Exit Sub

Err:
End Sub

Private Sub cmdSoundDelete_Click()

    Dim strRet  As String * 7
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    With lstWAV
    
        If Len(.List(.ListIndex)) > 7 Then
        
            strRet = "#WAV00:"
                    
            'If mnuOptionsNumFF.Checked Then
            
                'Mid$(strRet, 5, 2) = Right$("0" & Hex$(.ListIndex + 1), 2)
                'g_strWAV(lngNumConv(Mid$(strRet, 5, 2))) = ""
            
            'Else
            
                'Mid$(strRet, 5, 2) = modInput.strNumConv(.ListIndex + 1)
                'g_strWAV(.ListIndex + 1) = ""
            
            'End If
            
            Mid$(strRet, 5, 2) = strFromLong(.ListIndex + 1)
            g_strWAV(lngFromLong(.ListIndex + 1)) = ""
            
            .List(.ListIndex) = strRet
            
            Call SaveChanges
        
        End If
    
    End With

End Sub

Private Sub cmdSoundExcDown_Click()

    Dim i       As Long
    Dim lngRet  As Long
    Dim intRet  As Integer
    Dim strRet  As String
    
    With lstWAV
    
        If .ListIndex = .ListCount - 1 Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            'intRet = lngNumConv(Hex$(.ListIndex + 2))
            'lngRet = lngNumConv(Hex$(.ListIndex + 1))
        
        'Else
        
            'intRet = .ListIndex + 2
            'lngRet = .ListIndex + 1
        
        'End If
        
        intRet = lngFromLong(.ListIndex + 2)
        lngRet = lngFromLong(.ListIndex + 1)
        
        strRet = g_strWAV(intRet)
        g_strWAV(intRet) = g_strWAV(lngRet)
        g_strWAV(lngRet) = strRet
        
        .List(.ListIndex + 1) = ""
        .ListIndex = .ListIndex + 1
        
        .List(.ListIndex) = "#WAV" & strNumConv(intRet) & ":" & g_strWAV(intRet)
        .List(.ListIndex - 1) = "#WAV" & strNumConv(lngRet) & ":" & g_strWAV(lngRet)
    
    End With
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intCh >= 11 Then
            
                If .sngValue = lngRet Then
                
                    .sngValue = intRet
                
                ElseIf .sngValue = intRet Then
                
                    .sngValue = lngRet
                
                End If
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.WAV_CHANGE) & modInput.strNumConv(intRet) & modInput.strNumConv(lngRet) & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.WAV_CHANGE) & modInput.strNumConv(intRet) & modInput.strNumConv(lngRet) & ",")
    
    Call Redraw

End Sub

Private Sub cmdSoundExcUp_Click()

    Dim i       As Long
    Dim lngRet  As Long
    Dim intRet  As Integer
    Dim strRet  As String
    
    With lstWAV
    
        If .ListIndex = 0 Then Exit Sub
        
        'If mnuOptionsNumFF.Checked Then
        
            'intRet = lngNumConv(Hex$(.ListIndex))
            'lngRet = lngNumConv(Hex$(.ListIndex + 1))
        
        'Else
        
            'intRet = .ListIndex
            'lngRet = .ListIndex + 1
        
        'End If
        
        intRet = lngFromLong(.ListIndex)
        lngRet = lngFromLong(.ListIndex + 1)
        
        strRet = g_strWAV(intRet)
        g_strWAV(intRet) = g_strWAV(lngRet)
        g_strWAV(lngRet) = strRet
        
        .List(.ListIndex - 1) = ""
        .ListIndex = .ListIndex - 1
        
        .List(.ListIndex) = "#WAV" & strNumConv(intRet) & ":" & g_strWAV(intRet)
        .List(.ListIndex + 1) = "#WAV" & strNumConv(lngRet) & ":" & g_strWAV(lngRet)
    
    End With
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intCh >= 11 Then
            
                If .sngValue = lngRet Then
                
                    .sngValue = intRet
                
                ElseIf .sngValue = intRet Then
                
                    .sngValue = lngRet
                
                End If
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.WAV_CHANGE) & modInput.strNumConv(intRet) & modInput.strNumConv(lngRet) & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.WAV_CHANGE) & modInput.strNumConv(intRet) & modInput.strNumConv(lngRet) & ",")
    
    Call Redraw

End Sub

Private Sub cmdSoundLoad_Click()
On Error GoTo Err:

    Dim retArray()  As String
    'Dim strRet      As String
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    With dlgMain
    
        .Filter = "Sound files (*.wav,*.mp3)|*.wav;*.mp3|All files (*.*)|*.*"
        .FileName = Mid$(lstWAV.List(lstWAV.ListIndex), 8)
        
        Call dlgMain.ShowOpen
        
        retArray = Split(.FileName, "\")
        
        'If mnuOptionsNumFF.Checked Then
        
            'strRet = Right$("0" & Hex$(lstWAV.ListIndex + 1), 2)
            'lstWAV.List(lstWAV.ListIndex) = "#WAV" & strRet & ":" & retArray(UBound(retArray))
            'g_strWAV(modInput.lngNumConv(strRet)) = retArray(UBound(retArray))
        
        'Else
        
            'lstWAV.List(lstWAV.ListIndex) = "#WAV" & modInput.strNumConv(lstWAV.ListIndex + 1) & ":" & retArray(UBound(retArray))
            'g_strWAV(lstWAV.ListIndex + 1) = retArray(UBound(retArray))
        
        'End If
        
        lstWAV.List(lstWAV.ListIndex) = "#WAV" & strFromLong(lstWAV.ListIndex + 1) & ":" & retArray(UBound(retArray))
        g_strWAV(lngFromLong(lstWAV.ListIndex + 1)) = retArray(UBound(retArray))
        
        .InitDir = Left$(dlgMain.FileName, Len(dlgMain.FileName) - Len(retArray(UBound(retArray))))
        
        Call SaveChanges
    
    End With

Err:
End Sub

Private Sub cmdSoundStop_Click()

    Call mciSendString("close PREVIEW", vbNullString, 0, 0)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i   As Long
    Dim j   As Long
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Exit Sub
        
        End If
    
    End If
    
    'Shift が押されていたら3回繰り返すよ
    If Shift And vbShiftMask Then j = 2
    
    For i = 0 To j
    
        Select Case KeyCode
        
            Case vbKeyAdd '+
            
                If optChangeBottom(0).value = True Then
                
                    If lstWAV.ListIndex <> lstWAV.ListCount - 1 Then
                    
                        If Shift And vbCtrlMask Then
                        
                            Call cmdSoundExcDown_Click
                        
                        Else
                        
                            lstWAV.ListIndex = lstWAV.ListIndex + 1
                        
                        End If
                    
                    End If
                
                ElseIf optChangeBottom(1).value = True Or optChangeBottom(2).value = True Then
                
                    If lstBMP.ListIndex <> lstBMP.ListCount - 1 Then
                    
                        If Shift And vbCtrlMask Then
                        
                            Call cmdBMPExcDown_Click
                        
                        Else
                        
                            lstBMP.ListIndex = lstBMP.ListIndex + 1
                        
                        End If
                    
                    End If
                
                End If
                
                Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, Shift)
            
            Case vbKeySubtract '-
            
                If optChangeBottom(0).value = True Then
                
                    If lstWAV.ListIndex <> 0 Then
                    
                        If Shift And vbCtrlMask Then
                        
                            Call cmdSoundExcUp_Click
                        
                        Else
                        
                            lstWAV.ListIndex = lstWAV.ListIndex - 1
                        
                        End If
                    
                    End If
                
                ElseIf optChangeBottom(1).value = True Or optChangeBottom(2).value = True Then
                
                    If lstBMP.ListIndex <> 0 Then
                    
                        If Shift And vbCtrlMask Then
                        
                            Call cmdBMPExcUp_Click
                       
                        Else
                        
                            lstBMP.ListIndex = lstBMP.ListIndex - 1
                        
                        End If
                    
                    End If
                
                End If
                
                Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, Shift)
        
        End Select
    
    Next i
    
    Call modEasterEgg.KeyCheck(KeyCode, Shift)
    
End Sub

Private Sub Form_Load()

    m_blnPreview = True

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call FormDragDrop(Data)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If modMain.intSaveCheck() Then
    
        Cancel = True
    
    Else
    
        Call modMain.CleanUp
    
    End If

End Sub

Public Sub Form_Resize()
On Error Resume Next

    Dim i       As Long
    Dim lngRet  As Long
    
    Dim lngLineWidth            As Long
    Dim lngLineHeight           As Long
    Dim lngToolBarHeight        As Long
    Dim lngDirectInputHeight    As Long
    Dim lngStatusBarHeight      As Long
    
    Const PADDING = 60                                      '各パディングの大きさ
    Const TOOLBAR_HEIGHT = 390 + PADDING                    'ツールバーの高さ
    Const SCROLLBAR_SIZE = 255                              'スクロールバーの大きさ
    Const COLUMN_HEIGHT = 315                               '各カラムの高さ
    Const FRAME_WIDTH = 3255                                'フレームの幅
    Const FRAME_TOP_HEIGHT = 2190                           'ヘッダフレームの高さ
    Const FRAME_BOTTOM_TOP = 630                            'ボトムフレームのY位置。タブボタンの大きさ
    Const FRAME_BOTTOM_BUTTONS_HEIGHT = COLUMN_HEIGHT       '消去とか入力とかのボタン
    
    With frmMain
    
        If .WindowState = vbMinimized Then Exit Sub
        
        lngLineWidth = 2 * Screen.TwipsPerPixelX
        lngLineHeight = 2 * Screen.TwipsPerPixelY
        If mnuViewToolBar.Checked Then lngToolBarHeight = TOOLBAR_HEIGHT
        If mnuViewDirectInput.Checked Then lngDirectInputHeight = COLUMN_HEIGHT + PADDING * 2
        If mnuViewStatusBar.Checked Then lngStatusBarHeight = staMain.Height + Screen.TwipsPerPixelY * 2
        
        staMain.Visible = mnuViewStatusBar.Checked
        
        linToolbarBottom(0).X1 = 0
        linToolbarBottom(0).X2 = .ScaleWidth
        linToolbarBottom(0).Y1 = lngToolBarHeight
        linToolbarBottom(0).Y2 = linToolbarBottom(0).Y1
        linToolbarBottom(1).X1 = 0
        linToolbarBottom(1).X2 = .ScaleWidth
        linToolbarBottom(1).Y1 = linToolbarBottom(0).Y1 + Screen.TwipsPerPixelY
        linToolbarBottom(1).Y2 = linToolbarBottom(0).Y2 + Screen.TwipsPerPixelY
        
        linVertical(0).X1 = .ScaleWidth - FRAME_WIDTH - PADDING - lngLineWidth
        linVertical(0).X2 = linVertical(0).X1
        linVertical(1).X1 = linVertical(0).X1 + Screen.TwipsPerPixelX
        linVertical(1).X2 = linVertical(1).X1
        
        linVertical(0).Y1 = lngToolBarHeight
        linVertical(0).Y2 = .ScaleHeight - lngStatusBarHeight
        linVertical(1).Y1 = linVertical(0).Y1
        linVertical(1).Y2 = linVertical(0).Y2
        
        linHeader(0).X1 = linVertical(0).X1
        linHeader(0).X2 = .ScaleWidth
        linHeader(0).Y1 = lngToolBarHeight + PADDING + FRAME_TOP_HEIGHT + PADDING
        linHeader(0).Y2 = linHeader(0).Y1
        linHeader(1).X1 = linHeader(0).X1
        linHeader(1).X2 = linHeader(0).X2
        linHeader(1).Y1 = linHeader(0).Y1 + Screen.TwipsPerPixelY
        linHeader(1).Y2 = linHeader(0).Y2 + Screen.TwipsPerPixelY
        
        linDirectInput(0).X1 = 0
        linDirectInput(0).X2 = linVertical(0).X1
        linDirectInput(0).Y1 = .ScaleHeight - lngStatusBarHeight - PADDING - COLUMN_HEIGHT - PADDING
        linDirectInput(0).Y2 = linDirectInput(0).Y1
        linDirectInput(1).X1 = linDirectInput(0).X1
        linDirectInput(1).X2 = linDirectInput(0).X2
        linDirectInput(1).Y1 = linDirectInput(0).Y1 + Screen.TwipsPerPixelY
        linDirectInput(1).Y2 = linDirectInput(0).Y2 + Screen.TwipsPerPixelY
        
        linStatusBar(0).X1 = 0
        linStatusBar(0).X2 = .ScaleWidth
        linStatusBar(0).Y1 = .ScaleHeight - lngStatusBarHeight
        linStatusBar(0).Y2 = .ScaleHeight - lngStatusBarHeight
        linStatusBar(1).X1 = linStatusBar(0).X1
        linStatusBar(1).X2 = linStatusBar(0).X2
        linStatusBar(1).Y1 = linStatusBar(0).Y1 + Screen.TwipsPerPixelY
        linStatusBar(1).Y2 = linStatusBar(0).Y2 + Screen.TwipsPerPixelY
        
        linStatusBar(0).Visible = mnuViewStatusBar.Checked
        linStatusBar(1).Visible = mnuViewStatusBar.Checked
        
        tlbMenu.Visible = mnuViewToolBar.Checked
        fraViewer.Visible = mnuViewToolBar.Checked
        fraGrid.Visible = mnuViewToolBar.Checked
        fraDispSize.Visible = mnuViewToolBar.Checked
        
        lngRet = .ScaleWidth - FRAME_WIDTH - PADDING - lngLineWidth - PADDING - SCROLLBAR_SIZE
        
        Call vsbMain.Move(lngRet, _
            lngToolBarHeight + PADDING, _
            SCROLLBAR_SIZE, _
            .ScaleHeight - lngToolBarHeight - PADDING - lngStatusBarHeight - lngDirectInputHeight - SCROLLBAR_SIZE - PADDING)
        
        Call hsbMain.Move(0, _
            .ScaleHeight - lngStatusBarHeight - lngDirectInputHeight - SCROLLBAR_SIZE - PADDING, _
            lngRet, _
            SCROLLBAR_SIZE)
        
        Call picMain.Move(0, _
            lngToolBarHeight + PADDING, _
            lngRet, _
            .ScaleHeight - lngToolBarHeight - PADDING - lngStatusBarHeight - lngDirectInputHeight - SCROLLBAR_SIZE - PADDING)
        
        linDirectInput(0).Visible = mnuViewDirectInput.Checked
        linDirectInput(1).Visible = mnuViewDirectInput.Checked
        lblDirectInput.Visible = mnuViewDirectInput.Checked
        cboDirectInput.Visible = mnuViewDirectInput.Checked
        cmdDirectInput.Visible = mnuViewDirectInput.Checked
        
        Call lblDirectInput.Move(PADDING, _
            .ScaleHeight - lngStatusBarHeight - PADDING - (COLUMN_HEIGHT + lblDirectInput.Height) \ 2)
        
        Call cmdDirectInput.Move(.ScaleWidth - FRAME_WIDTH - cmdDirectInput.Width - PADDING * 2 - lngLineWidth, _
            .ScaleHeight - lngStatusBarHeight - PADDING - cmdDirectInput.Height)
        
        Call cboDirectInput.Move(lblDirectInput.Left + lblDirectInput.Width + PADDING, _
            .ScaleHeight - lngStatusBarHeight - PADDING - (COLUMN_HEIGHT + cboDirectInput.Height) \ 2, _
            cmdDirectInput.Left - lblDirectInput.Width - PADDING * 3)
        
        With tlbMenu.Buttons("Viewer")
            .Width = fraViewer.Width
            Call fraViewer.Move(.Left + PADDING, .Top + PADDING, .Width)
            Call fraViewer.ZOrder(0)
        End With
        
        With tlbMenu.Buttons("ChangeGrid")
            lblGridMain.Left = PADDING
            cboDispGridSub.Left = lblGridMain.Left + lblGridMain.Width + PADDING
            lblGridSub.Left = cboDispGridSub.Left + cboDispGridSub.Width + PADDING * 3
            cboDispGridMain.Left = lblGridSub.Left + lblGridSub.Width + PADDING
            fraGrid.Width = cboDispGridMain.Left + cboDispGridMain.Width + PADDING
            .Width = fraGrid.Width
            Call fraGrid.Move(.Left, .Top + PADDING, .Width)
            Call fraGrid.ZOrder(0)
        End With
        
        With tlbMenu.Buttons("DispSize")
            lblDispHeight.Left = PADDING
            cboDispHeight.Left = lblDispHeight.Left + lblDispHeight.Width + PADDING
            lblDispWidth.Left = cboDispHeight.Left + cboDispHeight.Width + PADDING * 3
            cboDispWidth.Left = lblDispWidth.Left + lblDispWidth.Width + PADDING
            fraDispSize.Width = cboDispWidth.Left + cboDispWidth.Width + PADDING
            .Width = fraDispSize.Width
            Call fraDispSize.Move(.Left, .Top + PADDING, .Width)
            Call fraDispSize.ZOrder(0)
        End With
        
        With tlbMenu.Buttons("Resolution")
            lblVScroll.Left = PADDING
            cboVScroll.Left = lblVScroll.Left + lblVScroll.Width + PADDING
            fraResolution.Width = cboVScroll.Left + cboVScroll.Width + PADDING
            Call fraResolution.Move(.Left, .Top + PADDING, .Width)
            Call fraResolution.ZOrder(0)
        End With
        
        Call .picMain.SetFocus
    
    End With
    
    With fraHeader
    
        Call fraHeader.Move(frmMain.ScaleWidth - FRAME_WIDTH, _
            lngToolBarHeight + PADDING, _
            FRAME_WIDTH, _
            FRAME_TOP_HEIGHT)
        
        For i = 0 To 2
        
            fraTop(i).Top = COLUMN_HEIGHT
        
        Next i
    
    End With
    
    With fraMaterial
    
        Call fraMaterial.Move(frmMain.ScaleWidth - FRAME_WIDTH, _
            FRAME_TOP_HEIGHT + PADDING + lngLineHeight + PADDING + lngToolBarHeight + PADDING, _
            FRAME_WIDTH, _
            frmMain.ScaleHeight - lngToolBarHeight - PADDING - fraHeader.Height - lngLineHeight - PADDING - lngStatusBarHeight - PADDING)
        
        lngRet = .Height - FRAME_BOTTOM_TOP
        
        For i = 0 To 4
        
            Call fraBottom(i).Move(0, _
                FRAME_BOTTOM_TOP, _
                FRAME_WIDTH, _
                lngRet)
        
        Next i
        
        lngRet = .Height - FRAME_BOTTOM_BUTTONS_HEIGHT - FRAME_BOTTOM_TOP - PADDING * 4
        lstWAV.Height = lngRet
        lstBMP.Height = lngRet
        lstMeasureLen.Height = lngRet
        txtBGAInput.Top = .Height - txtBGAInput.Height - FRAME_BOTTOM_BUTTONS_HEIGHT - FRAME_BOTTOM_TOP - PADDING * 2
        lstBGA.Height = lngRet - txtBGAInput.Height - PADDING
        txtExInfo.Height = .Height - FRAME_BOTTOM_TOP - PADDING * 3
        
        lngRet = fraBottom(0).Height - COLUMN_HEIGHT - PADDING
        cmdSoundStop.Top = lngRet
        cmdSoundExcUp.Top = lngRet
        cmdSoundExcDown.Top = lngRet
        cmdSoundDelete.Top = lngRet
        cmdSoundLoad.Top = lngRet
        cmdBMPPreview.Top = lngRet
        cmdBMPExcUp.Top = lngRet
        cmdBMPExcDown.Top = lngRet
        cmdBMPDelete.Top = lngRet
        cmdBMPLoad.Top = lngRet
        cmdBGAPreview.Top = lngRet
        cmdBGAExcUp.Top = lngRet
        cmdBGAExcDown.Top = lngRet
        cmdBGADelete.Top = lngRet
        cmdBGASet.Top = lngRet
        cmdMeasureSelectAll.Top = lngRet
        cboNumerator.Top = lngRet
        cboDenominator.Top = lngRet
        lblFraction.Top = lngRet + PADDING
        cmdInputMeasureLen.Top = lngRet
    
    End With
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub hsbMain_Change()

    With g_disp
    
        .X = hsbMain.value
    
    End With
    
    Call modDraw.Redraw
    
    'Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, g_Mouse.Shift)
    'スクロール＆オブジェ移動実現のため

End Sub

Private Sub hsbMain_Scroll()

    Call hsbMain_Change

End Sub

Private Sub lstBGA_Click()

    'If mnuOptionsNumFF.Checked Then
    
        'staMain.Panels("BMP").Text = "#BMP " & Right$("0" & Hex$(lstBGA.ListIndex + 1), 2)
    
    'Else
    
        'staMain.Panels("BMP").Text = "#BMP " & modInput.strNumConv(lstBGA.ListIndex + 1)
    
    'End If
    
    If optChangeBottom(2).value Then lstBMP.ListIndex = lstBGA.ListIndex
    
    staMain.Panels("BMP").Text = "#BMP " & strFromLong(lstBGA.ListIndex + 1)
    
    txtBGAInput.Text = Mid$(lstBGA.List(lstBGA.ListIndex), 8)
    
    If frmWindowPreview.Visible Then
    
        Call PreviewBGA(lstBGA.ListIndex + 1)
    
    End If

End Sub

Private Sub lstBMP_Click()

    'If mnuOptionsNumFF.Checked Then
    
        'staMain.Panels("BMP").Text = "#BMP " & Right$("0" & Hex$(lstBMP.ListIndex + 1), 2)
    
    'Else
    
        'staMain.Panels("BMP").Text = "#BMP " & modInput.strNumConv(lstBMP.ListIndex + 1)
    
    'End If
    
    If optChangeBottom(1).value Then lstBGA.ListIndex = lstBMP.ListIndex
    
    staMain.Panels("BMP").Text = "#BMP " & strFromLong(lstBMP.ListIndex + 1)
    
    If frmWindowPreview.Visible Then
    
        Call PreviewBMP(Mid$(lstBMP.List(lstBMP.ListIndex), 8))
    
    End If

End Sub

Private Sub lstBMP_DblClick()

    Call cmdBmpLoad_Click

End Sub

Private Sub lstBMP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        Call PopupMenu(mnuContextList)
    
    End If

End Sub

Private Sub lstBMP_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i           As Long
    Dim j           As Long
    Dim strRet      As String
    Dim strArray()  As String
    
    With lstBMP
    
        j = .ListIndex
        .Visible = False
        
        For i = 1 To Data.Files.Count
        
            If Dir(Data.Files.Item(i), vbNormal) <> vbNullString Then
            
                strRet = Data.Files.Item(i)
                
                'If Right$(UCase$(strRet), 4) = ".BMP" Or Right$(UCase$(strRet), 4) = ".JPG" Or Right$(UCase$(strRet), 4) = ".GIF" Then
                
                    Do Until Len(.List(j)) < 8
                    
                        j = j + 1
                        If j >= lstBMP.ListCount Then Exit For
                    
                    Loop
                    
                    strArray() = Split(strRet, "\")
                    
                    'If mnuOptionsNumFF.Checked Then
                    
                        'strRet = Right$("0" & Hex$(j + 1), 2)
                        '.List(j) = "#BMP" & strRet & ":" & strArray(UBound(strArray))
                        'g_strBMP(lngNumConv(strRet)) = strArray(UBound(strArray))
                    
                    'Else
                    
                        '.List(j) = "#BMP" & modInput.strNumConv(j + 1) & ":" & strArray(UBound(strArray))
                        'g_strBMP(j + 1) = strArray(UBound(strArray))
                    
                    'End If
                    
                    .List(j) = "#BMP" & strFromLong(j + 1) & ":" & strArray(UBound(strArray))
                    g_strBMP(lngFromLong(j + 1)) = strArray(UBound(strArray))
                    
                    j = j + 1
                    
                    If j >= lstBMP.ListCount Then Exit For
                
                'End If
            
            End If
        
        Next i
        
        .Visible = True
    
    End With
    
    Call frmMain.Show

End Sub

Private Sub lstMeasureLen_Click()
On Error Resume Next

    Dim strArray()  As String
    
    strArray() = Split(Mid$(lstMeasureLen.List(lstMeasureLen.ListIndex), 6), "/")
    
    cboNumerator.ListIndex = strArray(0) - 1
    
    cboDenominator.ListIndex = Log(strArray(1)) / Log(2) - 2

End Sub

Private Sub lstMeasureLen_DblClick()

    Dim i   As Long
    
    With lstMeasureLen
    
        .Visible = False
        
        For i = 0 To 999
        
            .Selected(i) = True
        
        Next i
        
        .ListIndex = 0
        .Visible = True
    
    End With

End Sub

Private Sub lstWAV_Click()

    Dim strRet  As String
    
    'If mnuOptionsNumFF.Checked Then
    
        'strRet = Right$("0" & Hex$(lstWAV.ListIndex + 1), 2)
    
    'Else
    
        'strRet = modInput.strNumConv(lstWAV.ListIndex + 1)
    
    'End If
    
    strRet = strFromLong(lstWAV.ListIndex + 1)
    
    staMain.Panels("WAV").Text = "#WAV " & strRet
    
    strRet = Mid$(lstWAV.List(lstWAV.ListIndex), 8)
    
    If strRet = "" Then Exit Sub
    If Dir(g_BMS.strDir & strRet) = vbNullString Then Exit Sub
    
    Call PreviewWAV(strRet)

End Sub

Private Sub lstWAV_DblClick()

    Call cmdSoundLoad_Click

End Sub

Private Sub lstWAV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
    
        Call PopupMenu(mnuContextList)
    
    End If

End Sub

Private Sub lstWAV_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err:

    Dim i           As Long
    Dim j           As Long
    Dim strRet      As String
    Dim strArray()  As String
    
    With lstWAV
    
        j = .ListIndex
        .Visible = False
        
        For i = 1 To Data.Files.Count
        
            If Dir(Data.Files.Item(i), vbNormal) <> vbNullString Then
            
                strRet = Data.Files.Item(i)
                
                'If Right$(UCase$(strRet), 4) = ".WAV" Or Right$(UCase$(strRet), 4) = ".MP3" Then
                
                    Do Until Len(.List(j)) < 8
                    
                        j = j + 1
                        If j >= lstWAV.ListCount Then Exit For
                    
                    Loop
                    
                    strArray() = Split(strRet, "\")
                    
                    'If mnuOptionsNumFF.Checked Then
                    
                        'strRet = Right$("0" & Hex$(j + 1), 2)
                        '.List(j) = "#WAV" & strRet & ":" & strArray(UBound(strArray))
                        'g_strWAV(lngNumConv(strRet)) = strArray(UBound(strArray))
                    
                    'Else
                    
                        '.List(j) = "#WAV" & modInput.strNumConv(j + 1) & ":" & strArray(UBound(strArray))
                        'g_strWAV(j + 1) = strArray(UBound(strArray))
                    
                    'End If
                    
                    .List(j) = "#WAV" & strFromLong(j + 1) & ":" & strArray(UBound(strArray))
                    g_strWAV(lngFromLong(j + 1)) = strArray(UBound(strArray))
                    
                    Call SaveChanges
                    
                    j = j + 1
                    
                    If j >= lstWAV.ListCount Then Exit For
                
                'End If
            
            End If
        
        Next i
        
        .Visible = True
    
    End With
    
    Call frmMain.Show
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "lstWAV_OLEDragDrop")
End Sub

Private Sub mnuContextDeleteMeasure_Click()

    Dim i           As Long
    Dim intRet      As Integer
    Dim strArray()  As String
    
    For i = 0 To 999
    
        'If g_Measure(i).lngY >= g_disp.Y + picMain.ScaleHeight - g_Mouse.Y - 6 Then
        If g_Measure(i).lngY >= g_disp.Y + (picMain.ScaleHeight - g_Mouse.Y) / g_disp.Height - 1 Then
        
            intRet = i - 1
            
            Exit For
        
        End If
    
    Next i
    
    ReDim strArray(0)
    
    strArray(0) = modInput.strNumConv(CMD_LOG.MSR_DEL) & modInput.strNumConv(intRet) & Right$("00" & Hex$(g_Measure(intRet).intLen), 3)
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intMeasure = intRet Then
            
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                
                Call modDraw.RemoveObj(i)
            
            ElseIf .intMeasure > intRet Then
            
                .intMeasure = .intMeasure - 1
            
            End If
        
        End With
    
    Next i
    
    lstMeasureLen.Visible = False
    
    For i = intRet To 998
    
        With g_Measure(i)
        
            .intLen = g_Measure(i + 1).intLen
            lstMeasureLen.List(i) = Left$(lstMeasureLen.List(i), 5) & Mid$(lstMeasureLen.List(i + 1), 6)
        
        End With
    
    Next i
    
    lstMeasureLen.List(999) = "#999:4/4"
    g_Measure(999).intLen = 192
    
    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(Join(strArray, ",") & ",")
    
    lstMeasureLen.Visible = True
    
    Call modDraw.ArrangeObj
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub mnuContextEditCopy_Click()

    Call mnuEditCopy_Click

End Sub

Private Sub mnuContextEditCut_Click()

    Call mnuEditCut_Click

End Sub

Private Sub mnuContextEditDelete_Click()

    Call mnuEditDelete_Click

End Sub

Private Sub mnuContextEditPaste_Click()

    Call mnuEditPaste_Click

End Sub

Private Sub mnuContextInsertMeasure_Click()

    Dim i           As Long
    Dim intRet      As Integer
    Dim strArray()  As String
    
    lstMeasureLen.Visible = False
    
    For i = 998 To 0 Step -1
    
        With g_Measure(i)
        
            'If .lngY < g_disp.Y + picMain.ScaleHeight - g_Mouse.Y - 6 Then
            If .lngY < g_disp.Y + (picMain.ScaleHeight - g_Mouse.Y) / g_disp.Height - 1 Then
            
                lstMeasureLen.List(i) = "#" & Format$(i, "000") & ":4/4"
                .intLen = 192
                intRet = i
                
                Exit For
            
            End If
            
            lstMeasureLen.List(i) = Left$(lstMeasureLen.List(i), 5) & Mid$(lstMeasureLen.List(i - 1), 6)
            .intLen = g_Measure(i - 1).intLen
        
        End With
    
    Next i
    
    ReDim strArray(0)
    
    strArray(0) = modInput.strNumConv(CMD_LOG.MSR_ADD) & modInput.strNumConv(intRet) & Right$("00" & Hex$(g_Measure(999).intLen), 3)
    
    lstMeasureLen.Visible = True
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intMeasure = 999 Then
            
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                
                Call modDraw.RemoveObj(i)
            
            ElseIf .intMeasure >= intRet Then
            
                .intMeasure = .intMeasure + 1
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(Join(strArray, ",") & ",")
    
    Call modDraw.ArrangeObj
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub mnuContextListDelete_Click()

    If optChangeBottom(0).value Then
    
        Call cmdSoundDelete_Click
    
    ElseIf optChangeBottom(1).value Then
    
        Call cmdBmpDelete_Click
    
    End If

End Sub

Private Sub mnuContextListLoad_Click()

    If optChangeBottom(0).value Then
    
        Call cmdSoundLoad_Click
    
    ElseIf optChangeBottom(1).value Then
    
        Call cmdBmpLoad_Click
    
    End If

End Sub

Private Sub mnuContextListRename_Click()

    Dim strRet  As String
    
    If optChangeBottom(0).value Then
    
        With lstWAV
        
            Call mciSendString("close PREVIEW", vbNullString, 0, 0)
            strRet = Mid$(.List(.ListIndex), 8)
            
            If Len(.List(.ListIndex)) > 8 Then
            
                If Dir(g_BMS.strDir & strRet, vbNormal) <> vbNullString Then
                
                    With frmWindowInput
                    
                        .lblMainDisp.Caption = g_Message(Message.INPUT_RENAME)
                        .txtMain.Text = strRet
                        
                        Call .Show(vbModal, frmMain)
                    
                    End With
                    
                    If strRet <> frmWindowInput.txtMain.Text And Len(frmWindowInput.txtMain.Text) <> 0 Then
                    
                        If Dir(g_BMS.strDir & frmWindowInput.txtMain.Text, vbNormal) = vbNullString Then
                        
                            Name g_BMS.strDir & strRet As g_BMS.strDir & frmWindowInput.txtMain.Text
                            
                            .List(.ListIndex) = Left$(.List(.ListIndex), 7) & frmWindowInput.txtMain.Text
                            g_strWAV(lngFromLong(.ListIndex + 1)) = frmWindowInput.txtMain.Text
                        
                        Else
                        
                            Call MsgBox(g_Message(Message.ERR_FILE_ALREADY_EXIST), vbCritical, g_strAppTitle)
                        
                        End If
                    
                    End If
                
                Else
                
                    Call MsgBox(g_Message(Message.ERR_FILE_NOT_FOUND) & vbCrLf & g_BMS.strDir & strRet, vbCritical, g_strAppTitle)
                
                End If
            
            End If
        
        End With
    
    ElseIf optChangeBottom(1).value Then
    
        With lstBMP
        
            strRet = Mid$(.List(.ListIndex), 8)
            
            If Len(.List(.ListIndex)) > 8 Then
            
                If Dir(g_BMS.strDir & strRet, vbNormal) <> vbNullString Then
                
                    With frmWindowInput
                    
                        .lblMainDisp.Caption = g_Message(Message.INPUT_RENAME)
                        .txtMain.Text = strRet
                        
                        Call .Show(vbModal, frmMain)
                    
                    End With
                    
                    If strRet <> frmWindowInput.txtMain.Text And Len(frmWindowInput.txtMain.Text) <> 0 Then
                    
                        If Dir(g_BMS.strDir & frmWindowInput.txtMain.Text, vbNormal) = vbNullString Then
                        
                            Name g_BMS.strDir & strRet As g_BMS.strDir & frmWindowInput.txtMain.Text
                            
                            .List(.ListIndex) = Left$(.List(.ListIndex), 7) & frmWindowInput.txtMain.Text
                            g_strBMP(lngFromLong(.ListIndex + 1)) = frmWindowInput.txtMain.Text
                        
                        Else
                        
                            Call MsgBox(g_Message(Message.ERR_FILE_ALREADY_EXIST), vbCritical, g_strAppTitle)
                        
                        End If
                    
                    End If
                
                Else
                
                    Call MsgBox(g_Message(Message.ERR_FILE_NOT_FOUND) & vbCrLf & g_BMS.strDir & strRet, vbCritical, g_strAppTitle)
                
                End If
            
            End If
        
        End With
    
    End If

End Sub

Private Sub mnuEditFind_Click()

    With frmWindowFind
    
        If Not .Visible Then
        
            .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
            .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
        
        End If
        
        Call .Show(0, frmMain)
    
    End With

End Sub

Private Sub mnuEditRedo_Click()

    Dim i           As Long
    Dim j           As Long
    Dim intRet      As Integer
    Dim lngRet      As Long
    Dim strRet      As String
    Dim strArray()  As String
    Dim lngArrayWAV(1295)   As Long
    Dim lngArrayBMP(1295)   As Long
    Dim strArrayWAV(1295)   As String
    Dim strArrayBMP(1295)   As String
    Dim strArrayBGA(1295)   As String
    Dim strArrayParamBGA()  As String
    Dim blnRefreshList      As Boolean
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Call SendMessage(Screen.ActiveControl.hwnd, WM_UNDO, 0, 0)
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Call SendMessage(Screen.ActiveControl.hwnd, WM_UNDO, 0, 0)
            
            Exit Sub
        
        End If
    
    End If
    
    'If g_lngInputLogPos = UBound(g_strInputLog) Then Exit Sub
    If g_InputLog.GetPos = g_InputLog.Max Then Exit Sub
    
    Call modDraw.ObjSelectCancel
    
    'g_lngInputLogPos = g_lngInputLogPos + 1
    Call g_InputLog.Forward
    
    'strArray() = Split(g_strInputLog(g_lngInputLogPos), ",")
    strArray() = Split(g_InputLog.GetData, ",")
    
    For i = 0 To UBound(strArray) - 1
    
        Select Case modInput.lngNumConv(Left$(strArray(i), 2))
        
            Case CMD_LOG.OBJ_ADD
            
                ReDim Preserve g_Obj(UBound(g_Obj) + 1)
                
                Call modInput.SwapObj(UBound(g_Obj), UBound(g_Obj) - 1)
                
                With g_Obj(UBound(g_Obj) - 1)
                
                    .lngID = modInput.lngNumConv(Mid$(strArray(i), 3, 4)) '
                    g_lngObjID(.lngID) = UBound(g_Obj) - 1
                    .intCh = Val("&H" & Mid$(strArray(i), 7, 2)) '
                    .intAtt = Mid$(strArray(i), 9, 1) '
                    .intMeasure = modInput.lngNumConv(Mid$(strArray(i), 10, 2)) '
                    .lngPosition = modInput.lngNumConv(Mid$(strArray(i), 12, 3)) '
                    .sngValue = Mid$(strArray(i), 15) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.OBJ_DEL
            
                Call modDraw.RemoveObj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
            
            Case CMD_LOG.OBJ_MOVE
            
                With g_Obj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
                
                    .intCh = Val("&H" & Mid$(strArray(i), 14, 2)) '
                    .intMeasure = modInput.lngNumConv(Mid$(strArray(i), 16, 2)) '
                    .lngPosition = modInput.lngNumConv(Mid$(strArray(i), 18, 3)) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.OBJ_CHANGE
            
                With g_Obj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
                
                    .sngValue = modInput.lngNumConv(Mid$(strArray(i), 9, 2)) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.MSR_ADD
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                For j = 999 To lngRet + 1 Step -1
                
                    g_Measure(j).intLen = g_Measure(j - 1).intLen
                    lstMeasureLen.List(j) = Left$(lstMeasureLen.List(j), 5) & Mid$(lstMeasureLen.List(j - 1), 6)
                
                Next j
                
                g_Measure(lngRet).intLen = 192
                lstMeasureLen.List(lngRet) = Left$(lstMeasureLen.List(lngRet), 5) & "4/4"
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intMeasure >= lngRet Then
                        
                            .intMeasure = .intMeasure + 1
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.MSR_DEL
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                For j = lngRet + 1 To 998
                
                    g_Measure(j).intLen = g_Measure(j + 1).intLen
                    lstMeasureLen.List(j) = Left$(lstMeasureLen.List(j), 5) & Mid$(lstMeasureLen.List(j + 1), 6)
                
                Next j
                
                g_Measure(999).intLen = 192
                lstMeasureLen.List(999) = "#999:4/4"
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intMeasure >= lngRet Then
                        
                            .intMeasure = .intMeasure - 1
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.MSR_CHANGE
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                g_Measure(lngRet).intLen = Val("&H" & Mid$(strArray(i), 8, 3)) '
                
                intRet = intGCD(g_Measure(lngRet).intLen, 192)
                If intRet <= 2 Then intRet = 3
                If intRet >= 48 Then intRet = 48
                lstMeasureLen.List(lngRet) = Left$(lstMeasureLen.List(lngRet), 5) & (g_Measure(lngRet).intLen / intRet) & "/" & (192 \ intRet)
            
            Case CMD_LOG.WAV_CHANGE
            
                intRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 5, 2)) '
                
                strRet = g_strWAV(intRet)
                g_strWAV(intRet) = g_strWAV(lngRet)
                g_strWAV(lngRet) = strRet
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intCh >= 11 Then
                        
                            If .sngValue = lngRet Then
                            
                                .sngValue = intRet
                            
                            ElseIf .sngValue = intRet Then
                            
                                .sngValue = lngRet
                            
                            End If
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.BMP_CHANGE
            
                intRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 5, 2)) '
                
                strRet = g_strBMP(intRet)
                g_strBMP(intRet) = g_strBMP(lngRet)
                g_strBMP(lngRet) = strRet
                
                strRet = g_strBGA(intRet)
                g_strBGA(intRet) = g_strBGA(lngRet)
                g_strBGA(lngRet) = strRet
                
                For j = 0 To UBound(g_strBGA)
                
                    If Len(g_strBGA(j)) Then
                    
                        strArrayParamBGA() = Split(g_strBGA(j), " ")
                        
                        If UBound(strArrayParamBGA) Then
                        
                            If modInput.lngNumConv(strArrayParamBGA(0)) = lngRet Then
                            
                                strArrayParamBGA(0) = modInput.strNumConv(intRet, 2)
                            
                            ElseIf modInput.lngNumConv(strArrayParamBGA(0)) = intRet Then
                            
                                strArrayParamBGA(0) = modInput.strNumConv(lngRet, 2)
                            
                            End If
                            
                            g_strBGA(j) = Join(strArrayParamBGA, " ")
                        
                        End If
                    
                    End If
                
                Next j
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intCh = 4 Or .intCh = 6 Or .intCh = 7 Then
                        
                            If .sngValue = lngRet Then
                            
                                .sngValue = intRet
                            
                            ElseIf .sngValue = intRet Then
                            
                                .sngValue = lngRet
                            
                            End If
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.LIST_ALIGN
            
                For j = 0 To UBound(lngArrayWAV)
                
                    lngArrayWAV(j) = j
                    lngArrayBMP(j) = j
                    
                    strArrayWAV(j) = g_strWAV(j)
                    strArrayBMP(j) = g_strBMP(j)
                    strArrayBGA(j) = g_strBGA(j)
                    
                    g_strWAV(j) = ""
                    g_strBMP(j) = ""
                    g_strBGA(j) = ""
                
                Next j
                
                For j = 3 To Len(strArray(i)) Step 5
                
                    lngRet = modInput.lngNumConv(Mid$(strArray(i), j + 1, 2)) '
                    intRet = modInput.lngNumConv(Mid$(strArray(i), j + 3, 2)) '
                    
                    Select Case Mid$(strArray(i), j, 1)
                    
                        Case 1 'WAV
                    
                            g_strWAV(intRet) = strArrayWAV(lngRet)
                            lngArrayWAV(lngRet) = intRet
                        
                        Case 2 'BMP
                        
                            g_strBMP(intRet) = strArrayBMP(lngRet)
                            g_strBGA(intRet) = strArrayBGA(lngRet)
                            lngArrayBMP(lngRet) = intRet
                    
                    End Select
                
                Next j
                
                For j = 0 To UBound(g_strBGA)
                
                    If Len(g_strBGA(j)) Then
                    
                        strArrayParamBGA() = Split(g_strBGA(j), " ")
                        
                        If UBound(strArrayParamBGA) Then
                        
                            strArrayParamBGA(0) = modInput.strNumConv(lngArrayBMP(modInput.lngNumConv(strArrayParamBGA(0))), 2)
                            g_strBGA(j) = Join(strArrayParamBGA(), " ")
                        
                        End If
                    
                    End If
                
                Next j
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        Select Case .intCh
                        
                            Case Is >= 11
                            
                                .sngValue = lngArrayWAV(.sngValue)
                            
                            Case 4, 6, 7
                            
                                .sngValue = lngArrayBMP(.sngValue)
                        
                        End Select
                    
                    End With
                
                Next j
            
            Case CMD_LOG.LIST_DEL
            
                Select Case Mid$(strArray(i), 3, 1)
                
                    Case 1 '#WAV
                    
                        g_strWAV(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = ""
                    
                    Case 2 '#BMP
                    
                        g_strBMP(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = ""
                    
                    Case 3 '#BGA
                    
                        g_strBGA(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = ""
                
                End Select
                
                blnRefreshList = True
        
        End Select
    
    Next i
    
    If blnRefreshList Then Call RefreshList
    
    Call modDraw.ArrangeObj
    
    Call SaveChanges
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub mnuEditUndo_Click()

    Dim i           As Long
    Dim j           As Long
    Dim intRet      As Integer
    Dim lngRet      As Long
    Dim strRet      As String
    Dim strArray()  As String
    Dim lngArrayWAV(1295)   As Long
    Dim lngArrayBMP(1295)   As Long
    Dim strArrayWAV(1295)   As String
    Dim strArrayBMP(1295)   As String
    Dim strArrayBGA(1295)   As String
    Dim strArrayParamBGA()  As String
    Dim blnRefreshList      As Boolean
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Call SendMessage(Screen.ActiveControl.hwnd, WM_UNDO, 0, 0)
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Call SendMessage(Screen.ActiveControl.hwnd, WM_UNDO, 0, 0)
            
            Exit Sub
        
        End If
    
    End If
    
    'If g_lngInputLogPos = 0 Then Exit Sub
    If g_InputLog.GetPos = 0 Then Exit Sub
    
    Call modDraw.ObjSelectCancel
    
    'strArray() = Split(g_strInputLog(g_lngInputLogPos - 1), ",")
    strArray() = Split(g_InputLog.GetData, ",")
    
    'g_lngInputLogPos = g_lngInputLogPos - 1
    Call g_InputLog.Back
    
    For i = 0 To UBound(strArray) - 1
    
        Select Case modInput.lngNumConv(Left$(strArray(i), 2))
        
            Case CMD_LOG.OBJ_ADD
            
                Call modDraw.RemoveObj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
            
            Case CMD_LOG.OBJ_DEL
            
                ReDim Preserve g_Obj(UBound(g_Obj) + 1)
                
                Call modInput.SwapObj(UBound(g_Obj), UBound(g_Obj) - 1)
                
                With g_Obj(UBound(g_Obj) - 1)
                
                    .lngID = modInput.lngNumConv(Mid$(strArray(i), 3, 4)) '
                    g_lngObjID(.lngID) = UBound(g_Obj) - 1
                    .intCh = Val("&H" & Mid$(strArray(i), 7, 2)) '
                    .intAtt = Mid$(strArray(i), 9, 1) '
                    .intMeasure = modInput.lngNumConv(Mid$(strArray(i), 10, 2)) '
                    .lngPosition = modInput.lngNumConv(Mid$(strArray(i), 12, 3)) '
                    .sngValue = Mid$(strArray(i), 15) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.OBJ_MOVE
            
                With g_Obj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
                
                    .intCh = Val("&H" & Mid$(strArray(i), 7, 2)) '
                    .intMeasure = modInput.lngNumConv(Mid$(strArray(i), 9, 2)) '
                    .lngPosition = modInput.lngNumConv(Mid$(strArray(i), 11, 3)) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.OBJ_CHANGE
            
                With g_Obj(g_lngObjID(modInput.lngNumConv(Mid$(strArray(i), 3, 4)))) '
                
                    .sngValue = modInput.lngNumConv(Mid$(strArray(i), 7, 2)) '
                    .intSelect = 1
                
                End With
            
            Case CMD_LOG.MSR_ADD
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                For j = lngRet To 998
                
                    g_Measure(j).intLen = g_Measure(j + 1).intLen
                    lstMeasureLen.List(j) = Left$(lstMeasureLen.List(j), 5) & Mid$(lstMeasureLen.List(j + 1), 6)
                
                Next j
                
                g_Measure(999).intLen = Val("&H" & Mid$(strArray(i), 5, 3)) '
                
                intRet = intGCD(g_Measure(999).intLen, 192)
                If intRet <= 2 Then intRet = 3
                If intRet >= 48 Then intRet = 48
                lstMeasureLen.List(999) = "#999:" & (g_Measure(999).intLen / intRet) & "/" & (192 \ intRet)
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intMeasure > lngRet Then
                        
                            .intMeasure = .intMeasure - 1
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.MSR_DEL
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                For j = 999 To lngRet + 1 Step -1
                
                    g_Measure(j).intLen = g_Measure(j - 1).intLen
                    lstMeasureLen.List(j) = Left$(lstMeasureLen.List(j), 5) & Mid$(lstMeasureLen.List(j - 1), 6)
                
                Next j
                
                g_Measure(lngRet).intLen = Val("&H" & Mid$(strArray(i), 5, 3)) '
                
                intRet = intGCD(g_Measure(lngRet).intLen, 192)
                If intRet <= 2 Then intRet = 3
                If intRet >= 48 Then intRet = 48
                lstMeasureLen.List(lngRet) = Left$(lstMeasureLen.List(lngRet), 5) & (g_Measure(lngRet).intLen / intRet) & "/" & (192 \ intRet)
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intMeasure >= lngRet Then
                        
                            .intMeasure = .intMeasure + 1
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.MSR_CHANGE
            
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                
                g_Measure(lngRet).intLen = Val("&H" & Mid$(strArray(i), 5, 3)) '
                
                intRet = intGCD(g_Measure(lngRet).intLen, 192)
                If intRet <= 2 Then intRet = 3
                If intRet >= 48 Then intRet = 48
                lstMeasureLen.List(lngRet) = Left$(lstMeasureLen.List(lngRet), 5) & (g_Measure(lngRet).intLen / intRet) & "/" & (192 \ intRet)
            
            Case CMD_LOG.WAV_CHANGE
            
                intRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 5, 2)) '
                
                strRet = g_strWAV(intRet)
                g_strWAV(intRet) = g_strWAV(lngRet)
                g_strWAV(lngRet) = strRet
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intCh >= 11 Then
                        
                            If .sngValue = lngRet Then
                            
                                .sngValue = intRet
                            
                            ElseIf .sngValue = intRet Then
                            
                                .sngValue = lngRet
                            
                            End If
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.BMP_CHANGE
            
                intRet = modInput.lngNumConv(Mid$(strArray(i), 3, 2)) '
                lngRet = modInput.lngNumConv(Mid$(strArray(i), 5, 2)) '
                
                strRet = g_strBMP(intRet)
                g_strBMP(intRet) = g_strBMP(lngRet)
                g_strBMP(lngRet) = strRet
                
                strRet = g_strBGA(intRet)
                g_strBGA(intRet) = g_strBGA(lngRet)
                g_strBGA(lngRet) = strRet
                
                
                For j = 0 To UBound(g_strBGA)
                
                    If Len(g_strBGA(j)) Then
                    
                        strArrayParamBGA() = Split(g_strBGA(j), " ")
                        
                        If UBound(strArrayParamBGA) Then
                        
                            If modInput.lngNumConv(strArrayParamBGA(0)) = lngRet Then
                            
                                strArrayParamBGA(0) = modInput.strNumConv(intRet, 2)
                            
                            ElseIf modInput.lngNumConv(strArrayParamBGA(0)) = intRet Then
                            
                                strArrayParamBGA(0) = modInput.strNumConv(lngRet, 2)
                            
                            End If
                            
                            g_strBGA(j) = Join(strArrayParamBGA, " ")
                        
                        End If
                    
                    End If
                
                Next j
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        If .intCh = 4 Or .intCh = 6 Or .intCh = 7 Then
                        
                            If .sngValue = lngRet Then
                            
                                .sngValue = intRet
                            
                            ElseIf .sngValue = intRet Then
                            
                                .sngValue = lngRet
                            
                            End If
                        
                        End If
                    
                    End With
                
                Next j
            
            Case CMD_LOG.LIST_ALIGN
            
                For j = 0 To UBound(lngArrayWAV)
                
                    lngArrayWAV(j) = j
                    lngArrayBMP(j) = j
                    
                    strArrayWAV(j) = g_strWAV(j)
                    strArrayBMP(j) = g_strBMP(j)
                    strArrayBGA(j) = g_strBGA(j)
                    
                    g_strWAV(j) = ""
                    g_strBMP(j) = ""
                    g_strBGA(j) = ""
                
                Next j
                
                For j = 3 To Len(strArray(i)) Step 5
                
                    intRet = modInput.lngNumConv(Mid$(strArray(i), j + 1, 2)) '古い値
                    lngRet = modInput.lngNumConv(Mid$(strArray(i), j + 3, 2)) '新しい値
                    
                    Select Case Mid$(strArray(i), j, 1)
                    
                        Case 1 'WAV
                    
                            g_strWAV(intRet) = strArrayWAV(lngRet)
                            lngArrayWAV(lngRet) = intRet
                        
                        Case 2 'BMP
                        
                            g_strBMP(intRet) = strArrayBMP(lngRet)
                            g_strBGA(intRet) = strArrayBGA(lngRet)
                            lngArrayBMP(lngRet) = intRet
                    
                    End Select
                
                Next j
                
                For j = 0 To UBound(g_strBGA)
                
                    If Len(g_strBGA(j)) Then
                    
                        strArrayParamBGA() = Split(g_strBGA(j), " ")
                        
                        If UBound(strArrayParamBGA) Then
                        
                            strArrayParamBGA(0) = modInput.strNumConv(lngArrayBMP(modInput.lngNumConv(strArrayParamBGA(0))), 2)
                            g_strBGA(j) = Join(strArrayParamBGA(), " ")
                        
                        End If
                    
                    End If
                
                Next j
                
                blnRefreshList = True
                
                For j = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(j)
                    
                        Select Case .intCh
                        
                            Case Is >= 11
                        
                                .sngValue = lngArrayWAV(.sngValue)
                            
                            Case 4, 6, 7
                            
                                .sngValue = lngArrayBMP(.sngValue)
                        
                        End Select
                    
                    End With
                
                Next j
            
            Case CMD_LOG.LIST_DEL
            
                Select Case Mid$(strArray(i), 3, 1)
                
                    Case 1 '#WAV
                    
                        g_strWAV(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = Mid$(strArray(i), 6)
                    
                    Case 2 '#BMP
                    
                        g_strBMP(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = Mid$(strArray(i), 6)
                    
                    Case 3 '#BGA
                    
                        g_strBGA(modInput.lngNumConv(Mid$(strArray(i), 4, 2))) = Mid$(strArray(i), 6)
                
                End Select
                
                blnRefreshList = True
        
        End Select
    
    Next i
    
    If blnRefreshList Then Call RefreshList
    
    Call modDraw.ArrangeObj
    
    Call SaveChanges
    
    Call modDraw.InitVerticalLine

End Sub

Private Sub mnuFileConvertWizard_Click()

    With frmWindowConvert
    
        .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
        .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
    
    End With
    
    Call frmWindowConvert.Show(vbModal, frmMain)

End Sub

Private Sub mnuFileOpenDirectory_Click()

    If Len(g_BMS.strDir) Then
    
        If Len(g_strFiler) <> 0 And Dir(g_strFiler) <> vbNullString Then '指定したファイラを使用
        
            Call ShellExecute(frmMain.hwnd, "open", Chr$(34) & g_strFiler & Chr$(34), Chr$(34) & g_BMS.strDir & Chr$(34), "", SW_SHOWNORMAL)
        
        Else 'Explorer で開く
        
            Call ShellExecute(frmMain.hwnd, "Explore", Chr$(34) & g_BMS.strDir & Chr$(34), "", "", SW_SHOWNORMAL)
        
        End If
    
    End If

End Sub

Private Sub mnuHelpAbout_Click()

    With frmWindowAbout
    
        .Left = (frmMain.Left + frmMain.Width \ 2) - .Width \ 2
        .Top = (frmMain.Top + frmMain.Height \ 2) - .Height \ 2
        
        Call .Show
    
    End With

End Sub

Private Sub mnuHelpOpen_Click()

    Call ShellExecute(0, vbNullString, g_strAppDir & g_strHelpFilename, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub mnuOptionsActiveIgnore_Click()

    mnuOptionsActiveIgnore.Checked = Not mnuOptionsActiveIgnore.Checked

End Sub

Private Sub mnuOptionsMoveOnGrid_Click()

    mnuOptionsMoveOnGrid.Checked = Not mnuOptionsMoveOnGrid.Checked

End Sub

Private Sub mnuOptionsNumFF_Click()

    mnuOptionsNumFF.Checked = Not mnuOptionsNumFF.Checked
    
    m_blnPreview = False
    lstWAV.ListIndex = 0
    lstBMP.ListIndex = 0
    lstBGA.ListIndex = 0
    m_blnPreview = True
    
    Call RefreshList

End Sub

Private Sub mnuOptionsObjectFileName_Click()

    mnuOptionsObjectFileName.Checked = Not mnuOptionsObjectFileName.Checked
    
    Call modDraw.Redraw

End Sub

Private Sub mnuOptionsRightClickDelete_Click()

    mnuOptionsRightClickDelete.Checked = Not mnuOptionsRightClickDelete.Checked

End Sub

Private Sub mnuOptionsSelectPreview_Click()

    mnuOptionsSelectPreview.Checked = Not mnuOptionsSelectPreview.Checked

End Sub

Private Sub mnuTheme_Click(Index As Integer)

    Dim i   As Long
    
    For i = 1 To mnuTheme.UBound
    
        mnuTheme(i).Checked = False
    
    Next i
    
    With mnuTheme(Index)
    
        .Checked = True
        
        Call modMain.LoadThemeFile("theme\" & g_strThemeFileName(Index))
    
    End With
    
    Call Redraw

End Sub

Private Sub mnuToolsSetting_Click()

    With frmWindowViewer
    
        .Left = frmMain.Left + (frmMain.Width - .Width) \ 2
        .Top = frmMain.Top + (frmMain.Height - .Height) \ 2
        
        Call .Show(vbModal, frmMain)
    
    End With

End Sub

Private Sub mnuViewDirectInput_Click()

    mnuViewDirectInput.Checked = Not mnuViewDirectInput.Checked
    Call Form_Resize

End Sub

Private Sub mnuOptionsLaneBG_Click()

    mnuOptionsLaneBG.Checked = Not mnuOptionsLaneBG.Checked
    
    Call modDraw.Redraw

End Sub

Private Sub mnuViewStatusBar_Click()

    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    
    Call Form_Resize

End Sub

Private Sub mnuViewToolBar_Click()

    mnuViewToolBar.Checked = Not mnuViewToolBar.Checked
    
    Call Form_Resize

End Sub

Private Sub mnuOptionsVertical_Click()

    mnuOptionsVertical.Checked = Not mnuOptionsVertical.Checked
    
    Call modDraw.Redraw

End Sub

Private Sub mnuEditCopy_Click()
On Error GoTo Err:

    Dim i       As Long
    Dim intRet  As Integer
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Call Clipboard.Clear
        Call Clipboard.SetText(Screen.ActiveControl.SelText)
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Call Clipboard.Clear
            Call Clipboard.SetText(Screen.ActiveControl.SelText)
            
            Exit Sub
        
        End If
    
    End If
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intSelect = 1 Then
        
            intRet = 1
            
            Exit For
        
        End If
    
    Next i
    
    If intRet = 0 Then Exit Sub
    
    Call CopyToClipboard
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intSelect Then g_Obj(i).intSelect = 0
    
    Next i
    
    Call modDraw.Redraw
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "mnuEditCopy_Click")
End Sub

Private Sub mnuEditCut_Click()
On Error GoTo Err:

    Dim i       As Long
    Dim intRet  As Integer
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Call Clipboard.Clear
        Call Clipboard.SetText(Screen.ActiveControl.SelText)
        Screen.ActiveControl.SelText = ""
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Call Clipboard.Clear
            Call Clipboard.SetText(Screen.ActiveControl.SelText)
            Screen.ActiveControl.SelText = ""
            
            Exit Sub
        
        End If
    
    End If
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intSelect = 1 Then
        
            intRet = 1
            
            Exit For
        
        End If
    
    Next i
    
    If intRet = 0 Then Exit Sub
    
    Call CopyToClipboard
    
    Call mnuEditDelete_Click
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "mnuEditCut_Click")
End Sub

Public Sub mnuEditDelete_Click()
On Error GoTo Err:

    Dim i           As Long
    Dim lngRet      As Long
    Dim strArray()  As String
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        If Len(Screen.ActiveControl.SelText) = 0 Then
        
            Screen.ActiveControl.SelLength = 1
        
        End If
        
        Screen.ActiveControl.SelText = ""
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            If Len(Screen.ActiveControl.SelText) = 0 Then
            
                Screen.ActiveControl.SelLength = 1
            
            End If
            
            Screen.ActiveControl.SelText = ""
            
            Exit Sub
        
        End If
    
    End If
    
    For i = 0 To UBound(g_Obj) - 1
    
        If g_Obj(i).intSelect Then
    
            lngRet = lngRet + 1
    
        End If
    
    Next i
    
    If lngRet = 0 Then Exit Sub
    
    'g_strInputLog(g_lngInputLogPos) = ""
    ReDim strArray(lngRet - 1)
    lngRet = 0
    
    For i = 0 To UBound(g_Obj) - 1
    
        With g_Obj(i)
        
            If .intSelect Then
            
                strArray(lngRet) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                lngRet = lngRet + 1
                
                Call modDraw.RemoveObj(i)
            
            End If
        
        End With
    
    Next i
    
    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(Join(strArray, ",") & ",")
    
    Call modDraw.ArrangeObj
    
    Call modDraw.Redraw
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "mnuEditDelete_Click")
End Sub

Private Sub mnuEditMode_Click(Index As Integer)

    tlbMenu.Buttons("Edit").value = tbrUnpressed
    tlbMenu.Buttons("Write").value = tbrUnpressed
    tlbMenu.Buttons("Delete").value = tbrUnpressed
    tlbMenu.Buttons(Index + 7).value = tbrPressed
    
    Select Case Index
    
        Case 0: staMain.Panels("Mode").Text = g_strStatusBar(20)
        
        Case 1: staMain.Panels("Mode").Text = g_strStatusBar(21)
        
        Case 2: staMain.Panels("Mode").Text = g_strStatusBar(22)
    
    End Select
    
End Sub

Private Sub mnuEditPaste_Click()
On Error GoTo Err:

    Dim i           As Long
    Dim j           As Long
    Dim lngArg      As Long
    Dim strArray()  As String
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Screen.ActiveControl.SelText = Clipboard.GetText()
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Screen.ActiveControl.SelText = Clipboard.GetText()
            
            Exit Sub
        
        End If
    
    End If
    
    Call modDraw.ObjSelectCancel
    
    For i = g_disp.intStartMeasure To 999
    
        If g_disp.Y <= g_Measure(i).lngY Then
        
            g_disp.intStartMeasure = i
            
            Exit For
        
        End If
    
    Next i
    
    strArray() = Split(Clipboard.GetText, vbCrLf)
    
    If UBound(strArray) < 2 Then Exit Sub
    
    If strArray(0) <> "BMSE ClipBoard Object Data Format" Then Exit Sub
    
    For i = 1 To UBound(strArray) - 1
    
        With g_Obj(UBound(g_Obj))
        
            .lngID = g_lngIDNum
            g_lngObjID(g_lngIDNum) = UBound(g_Obj)
            g_lngIDNum = g_lngIDNum + 1
            ReDim Preserve g_lngObjID(g_lngIDNum)
            .intCh = Left$(strArray(i), 3)
            .intAtt = Mid$(strArray(i), 4, 1)
            .lngPosition = Val(Mid$(strArray(i), 5, 7)) + g_Measure(g_disp.intStartMeasure).lngY
            .sngValue = Val(Mid$(strArray(i), 12))
            .lngHeight = 0
            .intSelect = 1
            
            For j = 0 To 999
            
                If .lngPosition < g_Measure(j).lngY Then
                
                    Exit For
                
                Else
                
                    .intMeasure = j
                
                End If
            
            Next j
            
            .lngPosition = .lngPosition - g_Measure(.intMeasure).lngY
            
            strArray(i - 1) = modInput.strNumConv(CMD_LOG.OBJ_ADD) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
            
            If modDraw.lngChangeMaxMeasure(.intMeasure) Then lngArg = 1
        
        End With
        
        If g_Obj(UBound(g_Obj)).lngPosition < g_Measure(g_Obj(UBound(g_Obj)).intMeasure).intLen Then ReDim Preserve g_Obj(UBound(g_Obj) + 1)
    
    Next i
    
    If lngArg Then Call modDraw.ChangeResolution
    
    ReDim Preserve strArray(UBound(strArray) - 2)
    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
    'g_lngInputLogPos = g_lngInputLogPos + 1
    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
    'Call SaveChanges
    Call g_InputLog.AddData(Join(strArray, ",") & ",")
    
    Call modDraw.Redraw
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "mnuEditPaste_Click")
End Sub

Private Sub mnuFileExit_Click()

    If modMain.intSaveCheck() Then Exit Sub
    
    Call modMain.CleanUp

End Sub

Private Sub mnuOptionsFileNameOnly_Click()

    mnuOptionsFileNameOnly.Checked = Not mnuOptionsFileNameOnly.Checked
    
    frmMain.Caption = g_strAppTitle
    
    If Len(g_BMS.strDir) Then
    
        If mnuOptionsFileNameOnly.Checked Then
        
            frmMain.Caption = frmMain.Caption & " - " & g_BMS.strFileName
        
        Else
        
            frmMain.Caption = frmMain.Caption & " - " & g_BMS.strDir & g_BMS.strFileName
        
        End If
    
    End If
    
    If Not g_BMS.blnSaveFlag Then frmMain.Caption = frmMain.Caption & " *"

End Sub

Private Sub mnuHelpWeb_Click()

    Call ShellExecute(0, vbNullString, "http://ucn.tokonats.net/", vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub mnuLanguage_Click(Index As Integer)

    Dim i   As Long
    
    For i = 1 To mnuLanguage.UBound
    
        mnuLanguage(i).Checked = False
    
    Next i
    
    With mnuLanguage(Index)
    
        .Checked = True
        
        Call modMain.LoadLanguageFile("lang\" & g_strLangFileName(Index))
    
    End With
    
    Call modDraw.Redraw

End Sub

Private Sub mnuFileNew_Click()
    
    If modMain.intSaveCheck() Then Exit Sub
    
    frmMain.Caption = g_strAppTitle & " - Now Initializing"
    
    Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
    
    With g_BMS
    
        .strDir = ""
        .strFileName = ""
    
    End With
    
    Call modInput.LoadBMSStart
    
    Call modInput.LoadBMSEnd

End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo Err:

    Dim retArray()  As String
    
    If modMain.intSaveCheck() Then Exit Sub
    
    With dlgMain
    
        .Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"
        
        .FileName = g_BMS.strFileName
        
        Call .ShowOpen
        
        Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
        
        retArray = Split(.FileName, "\")
        g_BMS.strDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
        g_BMS.strFileName = retArray(UBound(retArray))
        
        Call modInput.LoadBMS
        
        Call modMain.RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
        
        .InitDir = g_BMS.strDir
    
    End With

Err:
End Sub

Private Sub mnuToolsPlay_Click()

    Dim strFileName As String
    Dim strPath     As String
    
    If Mid$(g_Viewer(cboViewer.ListIndex + 1).strAppPath, 2, 2) <> ":\" Then
    
        strPath = g_strAppDir & g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    Else
    
        strPath = g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    End If
    
    If Dir(strPath, vbNormal) = vbNullString Then
    
        Call MsgBox(strPath & " " & g_Message(ERR_APP_NOT_FOUND), vbCritical, g_strAppTitle)
        Exit Sub
    
    End If
    
    If Len(g_BMS.strDir) Then
    
        strFileName = g_BMS.strDir & "___bmse_temp.bms" & vbNullString
    
    Else
    
        strFileName = g_strAppDir & "___bmse_temp.bms" & vbNullString
    
    End If
    
    Call modOutput.CreateBMS(strFileName, 1)
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    With g_Viewer(cboViewer.ListIndex + 1)
    
        Call ShellExecute(0, "open", Chr$(34) & strPath & Chr$(34), strCmdDecode(.strArgPlay, strFileName), "", SW_SHOWNORMAL)
    
    End With

End Sub

Private Sub mnuToolsPlayAll_Click()

    Dim strFileName As String
    Dim strPath     As String
    
    If Mid$(g_Viewer(cboViewer.ListIndex + 1).strAppPath, 2, 2) <> ":\" Then
    
        strPath = g_strAppDir & g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    Else
    
        strPath = g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    End If
    
    If Dir(strPath, vbNormal) = vbNullString Then
    
        Call MsgBox(strPath & " " & g_Message(ERR_APP_NOT_FOUND), vbCritical, g_strAppTitle)
        Exit Sub
    
    End If
    
    If Len(g_BMS.strDir) Then
    
        strFileName = g_BMS.strDir & "___bmse_temp.bms" & vbNullString
    
    Else
    
        strFileName = g_strAppDir & "___bmse_temp.bms" & vbNullString
    
    End If
    
    Call modOutput.CreateBMS(strFileName, 1)
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    With g_Viewer(cboViewer.ListIndex + 1)
    
        Call ShellExecute(0, "open", Chr$(34) & strPath & Chr$(34), strCmdDecode(.strArgAll, strFileName), "", SW_SHOWNORMAL)
    
    End With

End Sub

Private Sub mnuToolsPlayStop_Click()

    Dim strFileName As String
    Dim strPath     As String
    
    If Mid$(g_Viewer(cboViewer.ListIndex + 1).strAppPath, 2, 2) <> ":\" Then
    
        strPath = g_strAppDir & g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    Else
    
        strPath = g_Viewer(cboViewer.ListIndex + 1).strAppPath
    
    End If
    
    If Dir(strPath, vbNormal) = vbNullString Then
    
        Call MsgBox(strPath & " " & g_Message(ERR_APP_NOT_FOUND), vbCritical, g_strAppTitle)
        Exit Sub
    
    End If
    
    If Len(g_BMS.strDir) Then
    
        strFileName = g_BMS.strDir & "___bmse_temp.bms" & vbNullString
    
    Else
    
        strFileName = g_strAppDir & "___bmse_temp.bms" & vbNullString
    
    End If
    
    Call mciSendString("close PREVIEW", vbNullString, 0, 0)
    
    With g_Viewer(cboViewer.ListIndex + 1)
    
        Call ShellExecute(0, "open", Chr$(34) & strPath & Chr$(34), strCmdDecode(.strArgStop, strFileName), "", SW_SHOWNORMAL)
    
    End With

End Sub

Private Sub mnuRecentFiles_Click(Index As Integer)

    Dim retArray()  As String
    
    If modMain.intSaveCheck() Then Exit Sub
    
    If Dir(Mid$(mnuRecentFiles(Index).Caption, 4)) = vbNullString Then
    
        Call MsgBox(g_Message(ERR_FILE_NOT_FOUND) & vbCrLf & Mid$(mnuRecentFiles(Index).Caption, 4) & vbCrLf & g_Message(ERR_LOAD_CANCEL), vbCritical, g_strAppTitle)
        Exit Sub
    
    End If
    
    Call lngDeleteFile(g_BMS.strDir & "___bmse_temp.bms")
    
    retArray = Split(mnuRecentFiles(Index).Caption, "\")
    g_BMS.strDir = Mid$(mnuRecentFiles(Index).Caption, 4, Len(mnuRecentFiles(Index).Caption) - Len(retArray(UBound(retArray))) - 3)
    g_BMS.strFileName = retArray(UBound(retArray))
    
    dlgMain.InitDir = g_BMS.strDir
    
    Call modMain.RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
    
    Call modInput.LoadBMS

End Sub

Private Sub mnuFileSave_Click()

    If g_BMS.strDir <> "" And g_BMS.strFileName <> "" Then
    
        Call modOutput.CreateBMS(g_BMS.strDir & g_BMS.strFileName)
    
    Else
    
        Call mnuFileSaveAs_Click
    
    End If

End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo Err:

    Dim retArray()  As String
    
    With dlgMain
    
        .Filter = "BMS files (*.bms,*.bme,*.bml,*.pms)|*.bms;*.bme;*.bml;*.pms|All files (*.*)|*.*"
        
        .FileName = g_BMS.strFileName
        
        Call .ShowSave
        
        retArray = Split(.FileName, "\")
        g_BMS.strDir = Left$(.FileName, Len(.FileName) - Len(retArray(UBound(retArray))))
        g_BMS.strFileName = retArray(UBound(retArray))
        
        Call modOutput.CreateBMS(g_BMS.strDir & g_BMS.strFileName)
        
        Call modMain.RecentFilesRotation(g_BMS.strDir & g_BMS.strFileName)
        
        .InitDir = g_BMS.strDir
    
    End With

Err:
End Sub

Private Sub mnuEditSelectAll_Click()

    Dim i   As Long
    
    If TypeOf Screen.ActiveControl Is TextBox Then
    
        Screen.ActiveControl.SelStart = 0
        Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
        
        Exit Sub
    
    ElseIf TypeOf Screen.ActiveControl Is ComboBox Then
    
        If Screen.ActiveControl.Style = 0 Then
        
            Screen.ActiveControl.SelStart = 0
            Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
            
            Exit Sub
        
        End If
    
    End If
    
    For i = 0 To UBound(g_Obj) - 1
    
        g_Obj(i).intSelect = 1
    
    Next i
    
    Call modDraw.Redraw

End Sub

Private Sub optChangeBottom_Click(Index As Integer)

    Dim i   As Long
    
    For i = fraBottom.LBound To fraBottom.UBound
    
        fraBottom(i).Visible = False
    
    Next i
    
    fraBottom(Index).Visible = True

End Sub

Private Sub optChangeTop_Click(Index As Integer)

    Dim i   As Long
    
    For i = fraTop.LBound To fraTop.UBound
    
        fraTop(i).Visible = False
    
    Next i
    
    fraTop(Index).Visible = True

End Sub

Private Sub picMain_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim lngRet  As Long
    Dim intRet  As Integer
    Dim blnRet  As Boolean
    
    blnRet = True
    
    g_Mouse.Shift = Shift
    
    lngRet = vsbMain.value
    intRet = hsbMain.value
    
    Select Case KeyCode
    
        Case vbKeyControl, vbKeyShift ', vbKeyMenu
        
            If tlbMenu.Buttons("Write").value = tbrUnpressed Then blnRet = False
        
        Case vbKeyNumpad0
        
            cboDispGridSub.ListIndex = cboDispGridSub.ListCount - 1
        
        Case vbKeyNumpad1
        
            cboDispGridSub.ListIndex = 1
        
        Case vbKeyNumpad2
        
            cboDispGridSub.ListIndex = 2
        
        Case vbKeyNumpad3
        
            cboDispGridSub.ListIndex = 3
        
        Case vbKeyNumpad4
        
            cboDispGridSub.ListIndex = 6
        
        Case vbKeyNumpad5
        
            cboDispGridSub.ListIndex = 7
        
        Case vbKeyNumpad6
        
            cboDispGridSub.ListIndex = 8
        
        Case vbKeyHome
        
            lngRet = vsbMain.Max
        
        Case vbKeyEnd
        
            lngRet = vsbMain.Min
        
        Case vbKeyPageUp
        
            lngRet = lngRet + vsbMain.LargeChange
        
        Case vbKeyPageDown
        
            lngRet = lngRet - vsbMain.LargeChange
        
        Case vbKeyUp
        
            lngRet = lngRet + vsbMain.SmallChange
        
        Case vbKeyDown
        
            lngRet = lngRet - vsbMain.SmallChange
        
        Case vbKeyLeft
        
            intRet = intRet - hsbMain.SmallChange
        
        Case vbKeyRight
        
            intRet = intRet + hsbMain.SmallChange
        
        Case Else
        
            blnRet = False
            
    End Select
    
    With vsbMain
    
        If lngRet > .Min Then
        
            .value = .Min
        
        ElseIf lngRet < 0 Then
        
            .value = 0
        
        Else
        
            .value = lngRet
        
        End If
    
    End With
    
    With hsbMain
    
        If intRet < 0 Then
        
            .value = 0
        
        ElseIf intRet > .Max Then
        
            .value = .Max
        
        Else
        
            .value = intRet
        
        End If
    
    End With
    
    If blnRet Then
    
        If frmMain.tlbMenu.Buttons("Write").value = tbrUnpressed Then
        
            If g_SelectArea.blnFlag = True Or (g_Obj(UBound(g_Obj)).intCh <> 0 And g_Mouse.Button <> 0) Then
            
                Call picMain_MouseMove(vbLeftButton, Shift, g_Mouse.X, g_Mouse.Y)
            
            End If
        
        Else
        
            Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, Shift)
        
        End If
    
    End If

End Sub

Private Sub picMain_KeyUp(KeyCode As Integer, Shift As Integer)

    g_Mouse.Shift = Shift
    
    Select Case KeyCode
    
        Case vbKeyControl, vbKeyShift
        
            If tlbMenu.Buttons("Write").value = tbrPressed Then
            
                Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, Shift)
            
            End If
    
    End Select

End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err:

    Dim strRet      As String
    'Dim intNum      As Long
    Dim lngRet      As Long
    Dim i           As Long
    Dim retObj      As g_udtObj
    Dim strArray()  As String
    
    If g_blnIgnoreInput Then Exit Sub
    
    m_blnMouseDown = True
    
    If Button = vbLeftButton Then
    
        If tlbMenu.Buttons("Delete").value = tbrPressed Then
        
            If g_Obj(UBound(g_Obj)).intCh Then
            
                '/// Undo
                With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                
                    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                    'g_lngInputLogPos = g_lngInputLogPos + 1
                    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
                    'Call SaveChanges
                    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ",")
                
                End With
                
                Call modDraw.RemoveObj(g_Obj(UBound(g_Obj)).lngHeight)
                
                Call modDraw.ArrangeObj
                
                Call RemoveObj(UBound(g_Obj))
            
            End If
            
            Call modDraw.ObjSelectCancel
            Call modDraw.Redraw
        
        ElseIf tlbMenu.Buttons("Edit").value = tbrPressed Then
        
            If g_Obj(UBound(g_Obj)).intCh <> 0 Then 'オブジェのあるところで押したっぽいよ
            
                With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                
                    If cboDispGridSub.ItemData(cboDispGridSub.ListIndex) Then
                    
                        lngRet = 192 \ (cboDispGridSub.ItemData(cboDispGridSub.ListIndex))
                        lngRet = .lngPosition - (.lngPosition \ lngRet) * lngRet
                    
                    End If
                
                End With
                
                If g_Obj(g_Obj(UBound(g_Obj)).lngHeight).intSelect Then '複数選択っぽいよ
                
                    If Shift And vbCtrlMask Then
                    
                        Call modDraw.CopyObj(retObj, g_Obj(UBound(g_Obj)))
                        
                        'ReDim strArray(intNum - 1)
                        ReDim strArray(0)
                        'intNum = 0
                        
                        For i = 0 To UBound(g_Obj) - 1
                        
                            If g_Obj(i).intSelect Then
                            
                                With g_Obj(i)
                                
                                    Call modDraw.CopyObj(g_Obj(UBound(g_Obj)), g_Obj(i))
                                    g_Obj(UBound(g_Obj)).lngID = g_lngIDNum
                                    
                                    strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_ADD) & modInput.strNumConv(g_lngIDNum, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                                    'intNum = intNum + 1
                                    ReDim Preserve strArray(UBound(strArray) + 1)
                                    
                                    g_lngObjID(g_lngIDNum) = UBound(g_Obj)
                                    g_lngIDNum = g_lngIDNum + 1
                                    ReDim Preserve g_lngObjID(g_lngIDNum)
                                    
                                    .intSelect = 0
                                
                                End With
                                
                                If i = retObj.lngHeight Then retObj.lngHeight = UBound(g_Obj)
                                
                                ReDim Preserve g_Obj(UBound(g_Obj) + 1)
                            
                            End If
                        
                        Next i
                        
                        If UBound(strArray) Then
                        
                            ReDim Preserve strArray(UBound(strArray) - 1)
                            
                            'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
                            'g_lngInputLogPos = g_lngInputLogPos + 1
                            'ReDim Preserve g_strInputLog(g_lngInputLogPos)
                            Call g_InputLog.AddData(Join(strArray, ",") & ",")
                            
                            Call modDraw.CopyObj(g_Obj(UBound(g_Obj)), retObj)
                        
                        End If
                    
                    End If
                    
                    ReDim m_retObj(0)
                    
                    For i = 0 To UBound(g_Obj) - 1
                    
                        With g_Obj(i)
                        
                            If .intSelect Then
                            
                                Call modDraw.CopyObj(m_retObj(UBound(m_retObj)), g_Obj(i))
                                
                                .lngHeight = UBound(m_retObj)
                                
                                ReDim Preserve m_retObj(UBound(m_retObj) + 1)
                            
                            End If
                        
                        End With
                    
                    Next i
                    
                    Call modDraw.CopyObj(m_retObj(UBound(m_retObj)), g_Obj(g_Obj(UBound(g_Obj)).lngHeight))
                    
                    With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                    
                        If mnuOptionsSelectPreview.Checked = True And ((.intCh >= 11 And .intCh <= 29) Or .intCh > 100) Then
                        
                            strRet = g_strWAV(.sngValue)
                            
                            If strRet <> "" And Dir(g_BMS.strDir & strRet) <> vbNullString Then
                            
                                Call PreviewWAV(strRet)
                            
                            End If
                        
                        End If
                        
                        'オブジェをグリッドにあわせる
                        'If Not Shift And vbShiftMask Then
                        
                            'If lngRet <> 0 And mnuOptionsMoveOnGrid.Checked = True Then
                            
                                'For i = 0 To UBound(g_Obj) - 1
                                
                                    'With g_Obj(i)
                                    
                                        'If .intSelect Then
                                        
                                            '.lngPosition = .lngPosition - lngRet
                                            
                                            'Do While .lngPosition < 0 And .intMeasure <> 0
                                            
                                                '.intMeasure = .intMeasure - 1
                                                '.lngPosition = .lngPosition + g_Measure(.intMeasure).intLen
                                            
                                            'Loop
                                            
                                            'Call SaveChanges
                                        
                                        'End If
                                    
                                    'End With
                                
                                'Next i
                            
                            'End If
                        
                        'End If
                    
                    End With
                
                Else '単数選択っぽいよ
                
                    If Not Shift And vbCtrlMask Then
                    
                        Call modDraw.ObjSelectCancel
                    
                    End If
                    
                    g_Obj(g_Obj(UBound(g_Obj)).lngHeight).intSelect = 1
                    
                    Call modDraw.MoveSelectedObj
                    
                    With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                    
                        ReDim m_retObj(0)
                        
                        For i = 0 To UBound(g_Obj) - 1
                        
                            With g_Obj(i)
                            
                                Call modDraw.CopyObj(m_retObj(UBound(m_retObj)), g_Obj(i))
                                .lngHeight = UBound(m_retObj)
                                ReDim Preserve m_retObj(UBound(m_retObj) + 1)
                            
                            End With
                        
                        Next i
                        
                        Call modDraw.CopyObj(m_retObj(UBound(m_retObj)), g_Obj(g_Obj(UBound(g_Obj)).lngHeight))
                        
                        'オブジェをグリッドにあわせる
                        'If Not Shift And vbShiftMask Then
                        
                            'If lngRet <> 0 And mnuOptionsMoveOnGrid.Checked = True Then
                            
                                '.lngPosition = .lngPosition - lngRet
                                
                                'Do While .lngPosition < 0 And .intMeasure <> 0
                                
                                    '.intMeasure = .intMeasure - 1
                                    '.lngPosition = .lngPosition + g_Measure(.intMeasure).intLen
                                
                                'Loop
                                
                                'Call SaveChanges
                            
                            'End If
                        
                        'End If
                        
                        If mnuOptionsSelectPreview.Checked Then
                        
                            Select Case .intCh
                            
                                Case 11 To 29, Is > 100
                            
                                    strRet = g_strWAV(.sngValue)
                                    
                                    If strRet <> "" And Dir(g_BMS.strDir & strRet) <> vbNullString Then
                                    
                                        Call PreviewWAV(strRet)
                                    
                                    End If
                                
                                Case 4, 6, 7
                                
                                    If Len(g_strBGA(.sngValue)) Then
                                    
                                        Call PreviewBGA(.sngValue)
                                    
                                    Else
                                    
                                        strRet = g_strBMP(.sngValue)
                                        
                                        If strRet <> "" And Dir(g_BMS.strDir & strRet) <> vbNullString Then
                                        
                                            Call PreviewBMP(strRet)
                                        
                                        End If
                                    
                                    End If
                            
                            End Select
                        
                        End If
                    
                    End With
                
                End If
                
                Call modDraw.Redraw
            
            Else 'オブジェのないところで押したっぽいよ
            
                If Not Shift And vbCtrlMask Then
                
                    Call modDraw.ObjSelectCancel
                    
                    Call modDraw.Redraw
                
                Else
                
                    For i = 0 To UBound(g_Obj) - 1
                    
                        With g_Obj(i)
                        
                            If .intSelect Then
                            
                                .intSelect = 5
                            
                            End If
                        
                        End With
                    
                    Next i
                    
                    Call modDraw.Redraw
                
                End If
                
                With g_SelectArea
                
                    .blnFlag = True
                    '.X1 = (X + g_disp.X) / g_disp.Width
                    .X1 = X / g_disp.Width + g_disp.X
                    .Y1 = (picMain.ScaleHeight - Y) / g_disp.Height + g_disp.Y
                    .X2 = .X1
                    .Y2 = .Y1
                
                End With
                
                Call modDraw.DrawSelectArea
            
            End If
            
            'Call modDraw.Redraw
            If g_disp.intEffect Then Call modEasterEgg.DrawEffect
        
        Else 'If tlbMenu.Buttons("Write").Value = tbrPressed Then
        
            Call modDraw.ObjSelectCancel
            Call modDraw.Redraw
            
            picMain.Font.Size = 8
            
            Call modDraw.InitPen
            Call modDraw.DrawObj(g_Obj(UBound(g_Obj)))
            Call modDraw.DeletePen
        
        End If
    
    ElseIf Button = vbRightButton Then
    
        With g_Mouse
        
            .Button = Button
            .Shift = Shift
            .X = X
            .Y = Y
        
        End With
        
        Call DrawObjMax(X, Y, Shift)
        
        'スポイト機能
        If Shift And vbShiftMask Then
        
            If g_Obj(UBound(g_Obj)).lngHeight < UBound(g_Obj) Then
            
                With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                
                    Select Case .intCh
                    
                        Case 4, 6, 7, Is > 10
                        
                            Dim temp    As Long
                            Dim str     As String
                            
                            If mnuOptionsNumFF.Checked Then
                            
                                str = modInput.strNumConv(.sngValue)
                                
                                'もし 01-FF じゃなかったら 01-ZZ 表示に移行する
                                'ASCII 文字セットでは 0-9 < A-Z < a-z
                                If Asc(Left$(str, 1)) > Asc("F") Or Asc(Right$(str, 1)) > Asc("F") Then
                                
                                    mnuOptionsNumFF_Click
                                    temp = .sngValue
                                
                                Else
                                
                                    temp = Val("&H" & str)
                                
                                End If
                            
                            Else
                            
                                temp = .sngValue
                            
                            End If
                            
                            m_blnPreview = False
                            
                            If .intCh > 10 Then
                            
                                lstWAV.ListIndex = temp - 1
                            
                            Else
                            
                                If optChangeBottom(2).value Then
                                
                                    lstBGA.ListIndex = temp - 1
                                
                                Else
                                
                                    lstBMP.ListIndex = temp - 1
                                
                                End If
                            
                            End If
                            
                            m_blnPreview = True
                    
                    End Select
                    
                    Exit Sub
                
                End With
            
            End If
        
        End If
        
        If mnuOptionsRightClickDelete.Checked Then
        
            If g_Obj(UBound(g_Obj)).lngHeight < UBound(g_Obj) Then
            
                With g_Obj(g_Obj(UBound(g_Obj)).lngHeight)
                
                    'g_strInputLog(g_lngInputLogPos) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                    'g_lngInputLogPos = g_lngInputLogPos + 1
                    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
                    'Call SaveChanges
                    Call g_InputLog.AddData(modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ",")
                
                End With
                
                Call modDraw.RemoveObj(g_Obj(UBound(g_Obj)).lngHeight)
                
                g_Obj(UBound(g_Obj)).intCh = 0
                g_Obj(UBound(g_Obj)).lngHeight = UBound(g_Obj)
                
                Call modDraw.ArrangeObj
                
                Call modDraw.Redraw
                
                Exit Sub
            
            End If
        
        End If
        
        Call PopupMenu(mnuContext)
    
    End If
    
    g_blnIgnoreInput = False
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "picMain_MouseDown")
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err:

    Dim i           As Long
    Dim lngRet      As Long
    Dim strRet      As String
    Dim lngArg      As Long
    Dim strArray()  As String
    
    m_intScrollDir = 0
    
    If g_blnIgnoreInput Then
    
        g_blnIgnoreInput = False
        
        Exit Sub
    
    End If
    
    If Not m_blnMouseDown Then Exit Sub
    
    m_blnMouseDown = False
    
    If Button = vbLeftButton Then
    
        If tlbMenu.Buttons("Write").value = tbrPressed Then
        
            With g_Obj(UBound(g_Obj))
            
                Select Case .intCh
                
                    Case 8 'BPM
                    
                        With frmWindowInput
                        
                            .lblMainDisp.Caption = g_Message(Message.INPUT_BPM)
                            .txtMain.Text = "" 'strRet
                            
                            Call .Show(vbModal, frmMain)
                        
                        End With
                            
                        Select Case Val(frmWindowInput.txtMain.Text)
                        
                            Case 0
                            
                                Exit Sub
                            
                            Case Is > 65535
                            
                                Call MsgBox(g_Message(ERR_OVERFLOW_LARGE), vbCritical, g_strAppTitle)
                                
                                Exit Sub
                            
                            Case Is < -65535
                            
                                Call MsgBox(g_Message(ERR_OVERFLOW_SMALL), vbCritical, g_strAppTitle)
                                
                                Exit Sub
                            
                            Case Else
                            
                                .sngValue = Val(frmWindowInput.txtMain.Text)
                                Call picMain.SetFocus
                        
                        End Select
                
                    Case 9 'STOP
                    
                        With frmWindowInput
                        
                            .lblMainDisp.Caption = g_Message(Message.INPUT_STOP)
                            .txtMain.Text = "" 'strRet
                            
                            Call .Show(vbModal, frmMain)
                        
                        End With
                        
                        Select Case CLng(Val(frmWindowInput.txtMain.Text))
                        
                            Case Is <= 0
                            
                                Exit Sub
                            
                            Case Is > 65535
                            
                                Call MsgBox(g_Message(ERR_OVERFLOW_LARGE), vbCritical, g_strAppTitle)
                                
                                Exit Sub
                            
                            Case Else
                            
                                .sngValue = CLng(Val(frmWindowInput.txtMain.Text))
                                Call picMain.SetFocus
                        
                        End Select
                    
                    Case 51 To 69
                    
                        .intCh = .intCh - 40
                        .intAtt = 2
                
                End Select
                
                If .sngValue = 0 Then Exit Sub
                
                Call SaveChanges
                
                If modDraw.lngChangeMaxMeasure(.intMeasure) Then Call modDraw.ChangeResolution
            
            End With
            
            'g_strInputLog(g_lngInputLogPos) = ""
            strRet = ""
            
            For i = UBound(g_Obj) - 1 To 0 Step -1
            
                If g_Obj(i).intMeasure = g_Obj(UBound(g_Obj)).intMeasure And g_Obj(i).lngPosition = g_Obj(UBound(g_Obj)).lngPosition And g_Obj(i).intCh = g_Obj(UBound(g_Obj)).intCh Then
                
                    If g_Obj(i).intAtt \ 2 = g_Obj(UBound(g_Obj)).intAtt \ 2 Then
                    
                        'Undo
                        With g_Obj(i)
                        
                            'g_strInputLog(g_lngInputLogPos) = g_strInputLog(g_lngInputLogPos) & modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                            strRet = strRet & modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                        
                        End With
                        
                        Call modDraw.RemoveObj(i)
                    
                    End If
                
                End If
            
            Next i
            
            
            
            'Undo
            With g_Obj(UBound(g_Obj))
            
                .lngID = g_lngIDNum
                'g_strInputLog(g_lngInputLogPos) = g_strInputLog(g_lngInputLogPos) & modInput.strNumConv(CMD_LOG.OBJ_ADD) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                'g_lngInputLogPos = g_lngInputLogPos + 1
                'ReDim Preserve g_strInputLog(g_lngInputLogPos)
                strRet = strRet & modInput.strNumConv(CMD_LOG.OBJ_ADD) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue & ","
                Call g_InputLog.AddData(strRet)
                
                g_lngObjID(g_lngIDNum) = UBound(g_Obj)
                g_lngIDNum = g_lngIDNum + 1
                ReDim Preserve g_lngObjID(g_lngIDNum)
            
            End With
            
            ReDim Preserve g_Obj(UBound(g_Obj) + 1)
            
            Call modDraw.ArrangeObj
            
            Call modDraw.Redraw
        
        ElseIf tlbMenu.Buttons("Edit").value = tbrPressed Then
        
            If g_SelectArea.blnFlag Then
            
                g_SelectArea.blnFlag = False
                
                For i = 0 To UBound(g_Obj) - 1
                
                    With g_Obj(i)
                    
                        If .intSelect = 1 Or .intSelect = 4 Or .intSelect = 5 Then
                        
                            .intSelect = 1
                        
                        Else
                        
                            .intSelect = 0
                        
                        End If
                    
                    End With
                
                Next i
                
                Call modDraw.MoveSelectedObj
            
            Else '複数選択っぽいよ
            
                For i = 0 To UBound(g_Obj) - 1
                
                    If g_Obj(i).intSelect Then
                    
                        If g_Obj(g_Obj(UBound(g_Obj)).lngHeight).lngPosition + g_Measure(g_Obj(g_Obj(UBound(g_Obj)).lngHeight).intMeasure).lngY <> m_retObj(UBound(m_retObj)).lngPosition + g_Measure(m_retObj(UBound(m_retObj)).intMeasure).lngY Or g_Obj(g_Obj(UBound(g_Obj)).lngHeight).intCh <> m_retObj(UBound(m_retObj)).intCh Then
                        
                            lngRet = 1
                        
                        End If
                        
                        Exit For
                    
                    End If
                
                Next i
                
                If lngRet Then
                
                    ReDim strArray(0)
                    
                    For i = 0 To UBound(g_Obj) - 1
                    
                        With g_Obj(i)
                        
                            If .intCh <= 0 Or .intCh > 1000 Or (.intMeasure = 0 And .lngPosition < 0) Or (.intMeasure = 999 And .lngPosition > g_Measure(999).intLen) Then
                            
                                With m_retObj(g_Obj(i).lngHeight)
                                
                                    strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_DEL) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(.intCh), 2) & .intAtt & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3) & .sngValue
                                    ReDim Preserve strArray(UBound(strArray) + 1)
                                
                                End With
                                
                                Call modDraw.RemoveObj(i)
                            
                            ElseIf .intSelect Then
                            
                                strArray(UBound(strArray)) = modInput.strNumConv(CMD_LOG.OBJ_MOVE) & modInput.strNumConv(.lngID, 4) & Right$("0" & Hex$(m_retObj(.lngHeight).intCh), 2) & modInput.strNumConv(m_retObj(.lngHeight).intMeasure) & modInput.strNumConv(m_retObj(.lngHeight).lngPosition, 3) & Right$("0" & Hex$(.intCh), 2) & modInput.strNumConv(.intMeasure) & modInput.strNumConv(.lngPosition, 3)
                                ReDim Preserve strArray(UBound(strArray) + 1)
                            
                            End If
                            
                            If modDraw.lngChangeMaxMeasure(.intMeasure) Then lngArg = 1
                        
                        End With
                    
                    Next i
                    
                    If lngArg Then Call modDraw.ChangeResolution
                    
                    'g_strInputLog(g_lngInputLogPos) = Join(strArray, ",") & ","
                    'g_lngInputLogPos = g_lngInputLogPos + 1
                    'ReDim Preserve g_strInputLog(g_lngInputLogPos)
                    'Call SaveChanges
                    Call g_InputLog.AddData(Join(strArray, ",") & ",")
                
                End If
                
                Call modDraw.ArrangeObj
            
            End If
            
            Call modDraw.Redraw
        
        End If
    
    End If
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "picMain_MouseUp")
End Sub

Public Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err:

    Dim i           As Long
    Dim lngRet      As Long
    Dim retRect     As RECT
    Dim blnSelect() As Boolean
    Dim blnYAxisFixed   As Boolean
    
    'VB6 バグ対策
    If Button And vbLeftButton Then
    
        If Not m_blnMouseDown Then
        
            Exit Sub
        
        End If
    
    End If
    
    If Not g_SelectArea.blnFlag Then '選択範囲なし
    
        If Button = vbLeftButton And tlbMenu.Buttons("Edit").value = tbrPressed And g_Obj(UBound(g_Obj)).intCh <> 0 Then 'オブジェ移動中
        
            Call MoveObj(X, Y, Shift)
            
            'Y 軸固定移動
            If Shift And vbShiftMask Then blnYAxisFixed = True
        
        Else 'それ以外
        
            Call modDraw.DrawObjMax(X, Y, Shift)
        
        End If
    
    Else '選択範囲あり
        
        With g_Mouse
        
            .Button = Button
            .Shift = Shift
            .X = X
            .Y = Y
        
        End With
        
        With g_SelectArea
        
            '.X2 = (X + g_disp.X) / g_disp.Width
            .X2 = X / g_disp.Width + g_disp.X
            .Y2 = (picMain.ScaleHeight - Y) / g_disp.Height + g_disp.Y
        
        End With
        
        With retRect
        
            If g_SelectArea.X1 > g_SelectArea.X2 Then
            
                .Left = g_SelectArea.X2
                .Right = g_SelectArea.X1
            
            Else
            
                .Left = g_SelectArea.X1
                .Right = g_SelectArea.X2
            
            End If
            
            If g_SelectArea.Y1 > g_SelectArea.Y2 Then
            
                .Top = g_SelectArea.Y2
                .Bottom = g_SelectArea.Y1
            
            Else
            
                .Top = g_SelectArea.Y1
                .Bottom = g_SelectArea.Y2
            
            End If
        
        End With
        
        ReDim blnSelect(UBound(g_VGrid))
        
        For i = 0 To UBound(g_VGrid)
        
            With g_VGrid(i)
            
                blnSelect(i) = False
                
                If .blnVisible Then
                
                    If .intCh Then
                    
                        'If (.lngLeft + .intWidth >= g_SelectArea.X1 And .lngLeft <= g_SelectArea.X2) Or (.lngLeft <= g_SelectArea.X1 And .lngLeft + .intWidth >= g_SelectArea.X2) Then
                        If .lngLeft + .intWidth > retRect.Left And .lngLeft < retRect.Right Then
                        
                            blnSelect(i) = True
                        
                        End If
                    
                    End If
                
                End If
            
            End With
        
        Next i
        
        For i = 0 To UBound(g_Obj) - 1
        
            With g_Obj(i)
            
                'If g_VGrid(g_intVGridNum(.intCh)).blnSelect Then
                If blnSelect(g_intVGridNum(.intCh)) Then
                
                    lngRet = g_Measure(.intMeasure).lngY + .lngPosition
                    
                    'If (lngRet >= g_SelectArea.Y1 And lngRet <= g_SelectArea.Y2 + OBJ_HEIGHT / g_disp.Height) Or (lngRet <= g_SelectArea.Y1 And lngRet >= g_SelectArea.Y2 - OBJ_HEIGHT / g_disp.Height) Then
                    If lngRet + OBJ_HEIGHT / g_disp.Height >= retRect.Top And lngRet <= retRect.Bottom Then
                    
                        If .intSelect < 5 Then
                        
                            .intSelect = 4
                        
                        Else
                        
                            .intSelect = 6
                        
                        End If
                    
                    Else
                    
                        If .intSelect < 5 Then
                        
                            .intSelect = 0
                        
                        Else
                        
                            .intSelect = 5
                        
                        End If
                    
                    End If
                
                Else
                
                    If .intSelect < 5 Then
                    
                        .intSelect = 0
                    
                    Else
                    
                        .intSelect = 5
                    
                    End If
                
                End If
            
            End With
        
        Next i
        
        Call modDraw.DrawSelectArea
        
        If g_disp.intEffect Then Call modEasterEgg.DrawEffect
    
    End If
    
    With g_Mouse
    
        .Button = Button
        .Shift = Shift
        .X = X
        If Not blnYAxisFixed Then .Y = Y
    
    End With
    
    m_intScrollDir = 0
    
    If X < 0 Then
    
        m_intScrollDir = 20
    
    ElseIf X > picMain.ScaleWidth Then
    
        m_intScrollDir = 10
    
    End If
    
    If Not blnYAxisFixed Then
    
        If Y < 0 Then
        
            m_intScrollDir = m_intScrollDir + 1
        
        ElseIf Y > picMain.ScaleHeight Then
        
            m_intScrollDir = m_intScrollDir + 2
        
        End If
    
    End If
    
    If m_intScrollDir Then
    
        tmrMain.Enabled = True
    
    Else
    
        tmrMain.Enabled = False
    
    End If
    
    Exit Sub

Err:
    Call modMain.CleanUp(Err.Number, Err.Description, "picMain_MouseMove")
End Sub

Private Sub picMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call FormDragDrop(Data)

End Sub

Private Sub staMain_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

    If Panel.Key = "Mode" Then
    
        If tlbMenu.Buttons("Edit").value = tbrPressed Then
        
            Call mnuEditMode_Click(1)
        
        ElseIf tlbMenu.Buttons("Write").value = tbrPressed Then
        
            Call mnuEditMode_Click(2)
        
        ElseIf tlbMenu.Buttons("Delete").value = tbrPressed Then
        
            Call mnuEditMode_Click(0)
        
        End If
    
    End If

End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Err:

    Select Case Button.Key
    
        Case "New" '新規作成
        
            Call mnuFileNew_Click
        
        Case "Open" '開く
        
            Call mnuFileOpen_Click
        
        Case "Reload" '再読み込み
        
            Call mnuRecentFiles_Click(0)
        
        Case "Save" '上書き保存
        
            Call mnuFileSave_Click
        
        Case "SaveAs" '名前を付けて保存
        
            Call mnuFileSaveAs_Click
        
        Case "Edit"
        
            Call mnuEditMode_Click(0)
        
        Case "Write"
        
            Call mnuEditMode_Click(1)
        
        Case "Delete"
        
            Call mnuEditMode_Click(2)
        
        Case "PlayAll"
        
            Call mnuToolsPlayAll_Click
        
        Case "Play"
        
            Call mnuToolsPlay_Click
        
        Case "Stop"
        
            Call mnuToolsPlayStop_Click
    
    End Select

Err:
End Sub

Private Sub tlbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    If ButtonMenu.Parent = tlbMenu.Buttons("Open") Then
    
        Call mnuRecentFiles_Click(ButtonMenu.Index - 1)
    
    End If

End Sub

Private Sub tlbMenu_Change()

    Call Form_Resize

End Sub

Private Sub tmrMain_Timer()

    With hsbMain
    
        Select Case m_intScrollDir \ 10
        
            Case 1
            
                If .value + .SmallChange < .Max Then
                
                    .value = .value + .SmallChange
                
                Else
                
                    .value = .Max
                    
                    m_intScrollDir = m_intScrollDir - 10
                
                End If
            
            Case 2
            
                If .value - .SmallChange > .Min Then
                
                    .value = .value - .SmallChange
                
                Else
                
                    .value = .Min
                    
                    m_intScrollDir = m_intScrollDir - 20
                
                End If
        
        End Select
    
    End With
    
    With vsbMain
    
        Select Case m_intScrollDir Mod 10
        
            Case 1
            
                If .value + .SmallChange < .Min Then
                
                    .value = .value + .SmallChange
                
                Else
                
                    .value = .Min
                    
                    m_intScrollDir = m_intScrollDir - 1
                
                End If
            
            Case 2
            
                If .value - .SmallChange > .Max Then
                
                    .value = .value - .SmallChange
                
                Else
                
                    .value = .Max
                    
                    m_intScrollDir = m_intScrollDir - 2
                
                End If
        
        End Select
    
    End With
    
    If m_intScrollDir Then
    
        Call picMain_MouseMove(vbLeftButton, g_Mouse.Shift, g_Mouse.X, g_Mouse.Y)
        'Call MoveObj(g_Mouse.X, g_Mouse.Y, g_Mouse.Shift)
    
    End If

End Sub

Public Sub tmrEffect_Timer()

    Call picMain.Cls
    
    If g_Obj(UBound(g_Obj)).intCh Then
    
        Call modDraw.InitPen
        Call modDraw.DrawObj(g_Obj(UBound(g_Obj)))
        Call modDraw.DeletePen
    
    End If
    
    Select Case g_disp.intEffect
    
        Case RASTER, NOISE, STORM
        
            'Call modEasterEgg.RasterScroll
            'Call modEasterEgg.DrawRaster
        
        Case SNOW, SIROMARU
        
            Call modEasterEgg.FallingSnow
        
        Case SIROMARU2
        
            Call modEasterEgg.ZoomSiromaru2
        
        Case STAFFROLL, STAFFROLL2
        
            Call modEasterEgg.StaffRollScroll
            'Call modEasterEgg.DrawStaffRoll
    
    End Select
    
    If g_SelectArea.blnFlag Then Call modDraw.DrawSelectArea
    
    If g_disp.intEffect Then Call modEasterEgg.DrawEffect

End Sub

Private Sub txtArtist_Change()

    Call SaveChanges

End Sub

Private Sub txtBPM_Change()

    Call SaveChanges

End Sub

Private Sub txtExInfo_Change()

    Call SaveChanges

End Sub

Private Sub txtGenre_Change()

    Call SaveChanges

End Sub

Private Sub txtMissBMP_Change()

    Call SaveChanges

End Sub

Private Sub txtStageFile_Change()

    Call SaveChanges
End Sub

Private Sub txtTitle_Change()

    Call SaveChanges

End Sub

Private Sub txtTotal_Change()

    Call SaveChanges

End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        If txtTotal.Text = "10572" Then Call lngSet_ini("EasterEgg", "Tips", 1)
    
    End If

End Sub

Private Sub txtVolume_Change()

    Call SaveChanges

End Sub

Private Sub vsbMain_Change()
On Error Resume Next

    With g_disp
    
        '.Y = CLng(vsbMain.Value) * 96
        .Y = CLng(vsbMain.value) * .intResolution
    
    End With
    
    Call modDraw.Redraw
    
    'Call modDraw.DrawObjMax(g_Mouse.X, g_Mouse.Y, g_Mouse.Shift)
    'スクロール＆オブジェ移動実現のため

End Sub

Private Sub vsbMain_Scroll()

    Call vsbMain_Change

End Sub
