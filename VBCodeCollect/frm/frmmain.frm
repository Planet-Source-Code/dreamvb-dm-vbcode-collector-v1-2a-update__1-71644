VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmmain 
   Caption         =   "DM Sourcecode Collector"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9495
   Icon            =   "frmmain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5925
      TabIndex        =   17
      Tag             =   "SR"
      Text            =   "<Quick Search>"
      Top             =   765
      Width           =   2535
   End
   Begin Project1.dFlatButton cmdFind 
      Height          =   360
      Left            =   8520
      TabIndex        =   16
      ToolTipText     =   "Search"
      Top             =   765
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      Caption         =   ""
      Picture         =   "frmmain.frx":0CCA
   End
   Begin Project1.dEditor CodeView 
      Height          =   945
      Left            =   2640
      TabIndex        =   15
      Top             =   5010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1667
      Locked          =   -1  'True
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Tray Tray1 
      Left            =   435
      Top             =   2220
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.TextBox txtComment 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox pBar4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2640
      ScaleHeight     =   330
      ScaleWidth      =   6135
      TabIndex        =   12
      Top             =   3495
      Width           =   6135
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   45
         Picture         =   "frmmain.frx":101C
         Top             =   45
         Width           =   240
      End
      Begin VB.Image ImgUpdown 
         Height          =   180
         Left            =   5835
         MouseIcon       =   "frmmain.frx":10A5
         MousePointer    =   99  'Custom
         ToolTipText     =   "Hide"
         Top             =   60
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   13
         Top             =   60
         Width           =   870
      End
   End
   Begin VB.PictureBox pBar1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2640
      ScaleHeight     =   330
      ScaleWidth      =   2595
      TabIndex        =   10
      Top             =   420
      Width           =   2595
      Begin VB.Image ImgCodes 
         Height          =   225
         Left            =   45
         Picture         =   "frmmain.frx":11F7
         Top             =   45
         Width           =   225
      End
      Begin VB.Label lblCodes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   345
         TabIndex        =   11
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar tBar3 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   765
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_ADD"
            Object.ToolTipText     =   "New"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_EDIT"
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_DEL"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pBar3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   2595
      TabIndex        =   7
      Top             =   420
      Width           =   2595
      Begin VB.Label lblCat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   60
         Width           =   915
      End
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   1125
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1323
      _Version        =   393217
      Indentation     =   106
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar tBar2 
      Height          =   330
      Left            =   2640
      TabIndex        =   5
      Top             =   765
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_NEW"
            Object.ToolTipText     =   "New Code"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_EDIT"
            Object.ToolTipText     =   "Edit Code"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_DEL"
            Object.ToolTipText     =   "Delete Code"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_SAVE"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_CPY"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_MOVE"
            Object.ToolTipText     =   "Move"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_UP"
            Object.ToolTipText     =   "Up"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "M_DOWN"
            Object.ToolTipText     =   "Down"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   735
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   2535
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":129D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":15EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1941
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1C93
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1FE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2337
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2449
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":279B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2AED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2E3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3191
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":34E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":35F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3947
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3C99
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":433D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_NEW"
            Object.ToolTipText     =   "New Database"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_OPEN"
            Object.ToolTipText     =   "Open Database"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_COMP"
            Object.ToolTipText     =   "Compact Database"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_BACKUP"
            Object.ToolTipText     =   "Backup"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_ABOUT"
            Object.ToolTipText     =   "About"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "M_EXIT"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   7500
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16245
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pBar2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   2640
      ScaleHeight     =   330
      ScaleWidth      =   6135
      TabIndex        =   1
      Top             =   4665
      Width           =   6135
      Begin VB.Image imgCode 
         Height          =   240
         Left            =   45
         Picture         =   "frmmain.frx":468F
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblTitle2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codeview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   3
         Top             =   60
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   2370
      Left            =   2640
      TabIndex        =   0
      Top             =   1125
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4180
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image ImgTitle 
      Height          =   330
      Left            =   795
      Picture         =   "frmmain.frx":470A
      Top             =   2190
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Img1 
      Height          =   195
      Index           =   1
      Left            =   165
      Picture         =   "frmmain.frx":4D24
      Top             =   2295
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Img1 
      Height          =   195
      Index           =   0
      Left            =   165
      Picture         =   "frmmain.frx":4D77
      Top             =   2295
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Line ln3d 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   450
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line ln3d 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   450
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Database"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProps 
         Caption         =   "Properti&es..."
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCat 
      Caption         =   "&Category"
      Begin VB.Menu mnuNewCat 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuEditCat 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDelCat 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnublank5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBlank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCollase 
         Caption         =   "&Collase All"
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "Ex&pand All"
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "C&ode"
      Begin VB.Menu mnuNewCode 
         Caption         =   "New"
      End
      Begin VB.Menu mnuEditCode 
         Caption         =   "Edit"
      End
      Begin VB.Menu MnuDeleteCode 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuAbout1 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit1 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FistRun As Boolean
Private DBOpen As Boolean
Private ParentID As Long
Private TvID As Long
Private lViewID As Long
Private CodeID As Long
Private dbFile As String
Private TvText As String
Private pBarTop As Long
Private pCodeTop As Long
Private pLeft As Long
Private HasItem As Boolean
Private mButton As MouseButtonConstants
Private oWinState As Integer
Public mFilterIdx As Integer

Private Sub TvNodeExpand(ByVal Expand As Boolean)
Dim Count As Integer
    For Count = 1 To tv1.Nodes.Count
        tv1.Nodes(Count).Expanded = Expand
    Next Count
End Sub

Private Sub SaveSnipletFile(ByVal Filename As String)
Dim fp As Long
    
    'Store snipplet data
    With m_snip
        .dTitle = LstV.ListItems(lViewID).Text
        .dAuthor = LstV.ListItems(lViewID).SubItems(2)
        .dVersion = LstV.ListItems(lViewID).SubItems(1)
        .dComment = txtComment.Text
        .dCode = CodeView.Text
    End With
    
    fp = FreeFile
    Open Filename For Binary As #fp
        Put #fp, , m_snip
    Close #fp
    
    'Clear up
    With m_snip
        .dAuthor = vbNullString
        .dCode = vbNullString
        .dComment = vbNullString
        .dVersion = vbNullString
        .dTitle = vbNullString
    End With
    
End Sub

Private Sub BackUpDB(ByVal OutFile As String)
    If FindFile(OutFile) Then
        MsgBox "File already exists please choose a different name.", vbInformation, "Backup"
    Else
        'Close the database
        Call db.Close
        'Compact the database
        Call CompactDatabase(dbFile, OutFile)
        'Reopen the original database
        DBOpen = dOpenDataBase(dbFile)
    End If
End Sub

Private Sub LMoveItem(MoveUp As Boolean)
Dim idx As Integer

    idx = LstV.SelectedItem.Index
    
    If (MoveUp) Then
        idx = (idx - 1)
        If (idx <= 1) Then
            idx = 1
        End If
    Else
        idx = (idx + 1)
        If (idx >= LstV.ListItems.Count) Then
            idx = LstV.ListItems.Count
        End If
    End If
    
    Call SelectItem(idx)

End Sub

Private Sub ClickNode(ByVal Index As Integer)
On Error Resume Next
    'Select a treview node
    If (Index > tv1.Nodes.Count) Then
        Index = 1
    End If
    tv1.Nodes(Index).Selected = True
    Call tv1_Click
    Call tv1.SetFocus
End Sub

Private Sub SelectItem(ByVal Index As Integer)
Dim sItem As ListItem
    'Slect listview item
    With LstV
        .ListItems(Index).Selected = True
        .SetFocus
        Set sItem = .ListItems(Index)
    End With
    
    Call LstV_ItemClick(sItem)
    Set sItem = Nothing
End Sub

Public Function GetOpenDLGName(mFilter As String, Optional dTitle As String = "Open", _
Optional mShowSave As Boolean = False, Optional File As String = "") As String
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = dTitle
        .Filter = mFilter
        .Filename = File
        
        If (mShowSave) Then
            .ShowSave
        Else
            .ShowOpen
        End If
        mFilterIdx = .FilterIndex
        'Return filename
        GetOpenDLGName = .Filename
    End With
    
    Exit Function
    
    'Error flag
OpenErr:
    If Err.Number = cdlCancel Then
        Err.Clear
    End If
End Function

Private Sub cmdFind_Click()
Dim lItem As ListItem
Dim v
    'Find an item in the listview control
    For Each lItem In LstV.ListItems
        If InStr(1, lItem.Text, txtFind.Text, vbTextCompare) Then
            'Select the item
            Call SelectItem(lItem.Index)
            lItem.EnsureVisible
            Exit For
        End If
    Next lItem
    
    Set lItem = Nothing
    
End Sub

Private Sub Command1_Click()
    'Call ClickNode(21)
    
End Sub

Private Sub Form_Activate()
    If (Not FistRun) Then
        'Select second node
        Call ClickNode(2)
        FistRun = True
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 27) Then
        Call mnuExit_Click
    End If
End Sub

Private Sub Form_Load()
    'database filename
    dbFile = GetSetting("DMCodeBank", "cfg", "dbPath", FixPath(App.Path) & "codetips.mdb")
    'Create texture brush
    m_SrcBrush = CreatePatternBrush(ImgTitle.Picture.Handle)
    
    'Check if the app isready running
    If (App.PrevInstance) Then
        MsgBox frmmain.Caption & " is already running.", vbInformation, frmmain.Caption
        Unload frmmain
        Exit Sub
    End If
    
    'Check if database is found
    If (Not FindFile(dbFile)) Then
        'Opps not found
        If MsgBox("The main database was not found" _
            & vbCrLf & "would you like to create a new database now.", vbYesNo Or vbQuestion, frmmain.Caption) = vbNo Then
            Unload frmmain
            Exit Sub
        Else
            'Create the new database
            If CreateNewDatabase(dbFile) <> 1 Then
                MsgBox "There was an error creating the database.", vbInformation, frmmain.Caption
                Unload frmmain
                Exit Sub
            Else
                'Add database path to regedit
                SaveSetting "DMCodeBank", "cfg", "dbPath", dbFile
                MsgBox "Database saved to:" & vbCrLf & vbCrLf & dbFile, vbInformation, frmmain.Caption
            End If
        End If
    End If
    
    'Load the database
    DBOpen = dOpenDataBase(dbFile)
    'Set button state
    ButtonPress = vbCancel
    
    ImgUpdown.Picture = Img1(0).Picture
    pBarTop = pBar2.Top
    pCodeTop = CodeView.Top
    pLeft = pBar1.Left
    
    Set cmdFind.MouseIcon = ImgUpdown.MouseIcon
    
    'Setup try control
    Tray1.ToolTip = frmmain.Caption
    Set Tray1.Icon = frmmain.Icon
    
    'Try and open the default database
    If (Not DBOpen) Then
        MsgBox "There was an error opening the database.", vbInformation, frmmain.Caption
        Exit Sub
    Else
        'Load categories
        Call LoadCategories(tv1)
    End If
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    'Check if minsizeing to the tray or taskbar
    If GetSetting("DMCodeBank", "cfg", "Minsize", 1) = 1 Then
        If (frmmain.WindowState <> 1) Then
            oWinState = frmmain.WindowState
        Else
            frmmain.Visible = False
            Tray1.Visible = True
        End If
    End If
    
    'Resize controls
    cmdFind.Left = (frmmain.ScaleWidth - cmdFind.Width) - 30
    txtFind.Left = (cmdFind.Left - txtFind.Width)
    pBar1.Width = (frmmain.ScaleWidth - pBar1.Left)
    pBar2.Width = pBar1.Width
    pBar4.Width = pBar1.Width
    LstV.Width = (frmmain.ScaleWidth - LstV.Left)
    txtComment.Width = (frmmain.ScaleWidth - txtComment.Left)
    
    ln3d(0).X2 = frmmain.ScaleWidth
    ln3d(1).X2 = frmmain.ScaleWidth
    
    tv1.Height = (frmmain.ScaleHeight - sBar1.Height - tv1.Top)
    
    'Resize codeview
    CodeView.Width = pBar2.Width
    CodeView.Height = (frmmain.ScaleHeight - sBar1.Height - CodeView.Top)
    
    'Texture pictureboxes
    Call TexturePicBox(pBar1)
    Call TexturePicBox(pBar2)
    Call TexturePicBox(pBar3)
    Call TexturePicBox(pBar4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Destroy form obj
    Set frmmain = Nothing
End Sub

Private Sub ImgUpdown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static mHide As Boolean
    If (Button = vbLeftButton) Then
        If (mHide) Then
            txtComment.Visible = True
            pBar2.Top = pBarTop
            CodeView.Top = pCodeTop
            ImgUpdown.Picture = Img1(0).Picture
            ImgUpdown.ToolTipText = "Hide"
        Else
            txtComment.Visible = False
            pBar2.Top = txtComment.Top
            CodeView.Top = pBar2.Top
            ImgUpdown.Picture = Img1(1).Picture
            ImgUpdown.ToolTipText = "Show"
        End If
        
        mHide = (Not mHide)
        Call Form_Resize
    End If
End Sub

Private Sub LstV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static sSort As Integer
    sSort = (Not sSort)
    LstV.SortKey = ColumnHeader.Index - 1
    LstV.SortOrder = Abs(sSort)
    LstV.Sorted = True
End Sub

Private Sub LstV_DblClick()
    If (mButton = vbLeftButton) Then
        If (HasItem) Then Call mnuEditCode_Click
    End If
End Sub

Private Sub LstV_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim rc As Recordset
Dim vQuery As String
On Error GoTo OpenQErr:
    
    HasItem = True
    'Store Listview index
    lViewID = LstV.SelectedItem.Index
    'Enable/Display button and menu items
    MnuDeleteCode.Enabled = True
    mnuEditCode.Enabled = True
    tBar3.Buttons(2).Enabled = False
    tBar3.Buttons(3).Enabled = False
    tBar2.Buttons(2).Enabled = True
    tBar2.Buttons(3).Enabled = True
    tBar2.Buttons(5).Enabled = True
    tBar2.Buttons(6).Enabled = True
    tBar2.Buttons(8).Enabled = True
    tBar2.Buttons(10).Enabled = True
    tBar2.Buttons(11).Enabled = True
    'Extract the Sourcecode ID
    CodeID = Val(Mid(Item.Key, 3))
    
    'Build the Query string
    vQuery = "SELECT ID,sCode,sComment FROM Codes WHERE ID = " & Str(CodeID)
    'Perform the Query
    Set rc = db.OpenRecordset(vQuery)
    'Check that Query was found
    If (rc.RecordCount) Then
        'Show the sourcecode
        CodeView.Text = rc.Fields("sCode").Value
        txtComment.Text = rc.Fields("sComment").Value & ""
    End If
    
    'Clean up
    rc.Close
    Set rc = Nothing
    
    Exit Sub
    'Error flag
OpenQErr:
    If Err Then
        MsgBox Err.Description, vbInformation, frmmain.Caption
        rc.Close
        Set rc = Nothing
    End If
End Sub

Private Sub LstV_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) Then
        Call MnuDeleteCode_Click
    End If
End Sub

Private Sub LstV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mButton = Button
End Sub

Private Sub mnuAbout_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub mnuAbout1_Click()
    Call mnuAbout_Click
End Sub

Private Sub mnuCollase_Click()
    'Collase all
    Call TvNodeExpand(False)
End Sub

Private Sub mnuCompact_Click()
Dim TmpName As String
Dim sMsg As String

    'Create a temp name
    TmpName = FixPath(App.Path) & "compact.mdb"
    sMsg = "Old Filesize: " & FileLen(dbFile) & " bytes" & vbCrLf
    
    If (DBOpen) Then
        'We must first close the database
        db.Close
        'Compact the database
        Call CompactDatabase(dbFile, TmpName)
        'Delete the original database
        Call SetAttr(dbFile, vbNormal)
        Call Kill(dbFile)
        'Rename the temp database to the original one
        Name TmpName As dbFile
        'Reopen the database
        DBOpen = dOpenDataBase(dbFile)
        'Display status
        sMsg = sMsg & "New Filesize: " & FileLen(dbFile) & " bytes"
        MsgBox sMsg, vbInformation, "Compact Finished"
        sMsg = vbNullString
    End If
    
End Sub

Private Sub mnuDelCat_Click()
    If MsgBox("Are you sure you want to delete '" & TvText & "' and all it's items?", _
    vbYesNo Or vbQuestion, frmmain.Caption) = vbYes Then
        'Check if category was deleted
        If DeleteCategory(tv1.Nodes(TvID)) <> 1 Then
            MsgBox "There was an error while deleteing the category.", vbInformation, frmmain.Caption
            Exit Sub
        Else
            'Delete all the records for CatID Index
            Call DeleteRecords(ParentID)
            'Remove the node
            Call tv1.Nodes.Remove(TvID)
            'Select the first node
            Call ClickNode(1)
            'Clean up
            LstV.ListItems.Clear
            CodeView.Text = ""
        End If
    End If
End Sub

Private Sub MnuDeleteCode_Click()
    If MsgBox("Are you sure you want to delete this item?", vbYesNo Or vbQuestion, frmmain.Caption) = vbYes Then
        If DeleteCode(CodeID) <> 1 Then
            MsgBox "There was an error removeing the code.", vbInformation, frmmain.Caption
        Else
            'Remove the list item
            Call LstV.ListItems.Remove(lViewID)
            'Enabe/display menu items
            MnuDeleteCode.Enabled = False
            mnuEditCode.Enabled = False
            tBar2.Buttons(2).Enabled = False
            tBar2.Buttons(3).Enabled = False
            tBar2.Buttons(5).Enabled = False
            tBar2.Buttons(6).Enabled = False
            tBar2.Buttons(8).Enabled = False
            tBar2.Buttons(10).Enabled = False
            tBar2.Buttons(11).Enabled = False
            'Select first item
            If (LstV.ListItems.Count) Then
                Call ClickNode(TvID)
                Call SelectItem(1)
                Call lvSizeColumns(LstV)
            Else
                Call ClickNode(TvID)
            End If
        End If
    End If
End Sub

Private Sub mnuEditCat_Click()
Dim cName As String

    'Get the category name
    cName = Trim$(InputBox$("Edit Category", frmmain.Caption, TvText))

    If Len(cName) Then
        If EditCategory(ParentID, cName) <> 1 Then
            MsgBox "Error while editing category.", vbInformation, frmmain.Caption
        Else
            'Update Treeview text
            tv1.SelectedItem.Text = cName
            Call tv1_Click
        End If
    End If
End Sub

Private Sub mnuEditCode_Click()
    
    With m_tcode
        'Store code info
        .dTitle = LstV.ListItems(lViewID).Text
        .dVersion = LstV.ListItems(lViewID).SubItems(1)
        .dAuthor = LstV.ListItems(lViewID).SubItems(2)
        .dComment = txtComment.Text
        .dCodeBlock = CodeView.Text
    End With
    
    EditOp = 1 'Edit code
    frmAddCode.Show vbModal, frmmain
    'Check if the ok button was pressed
    If (ButtonPress <> vbCancel) Then
        'Edit code
        If EditCode(CodeID) <> 1 Then
            MsgBox "There was an error updateing the code.", vbInformation, frmmain.Caption
        Else
            'Select the item
            Call ClickNode(TvID)
            Call SelectItem(lViewID)
            Call lvSizeColumns(LstV)
        End If
    End If
    
    ButtonPress = vbCancel

End Sub

Private Sub mnuExit_Click()
    'Clean objects
    If (DBOpen) Then
        db.Close
    End If
    
    Set db = Nothing
    'Unload the form
    Unload frmmain
End Sub

Private Sub mnuExit1_Click()
    Call mnuExit_Click
End Sub

Private Sub mnuExpand_Click()
    'Expand All
    Call TvNodeExpand(True)
End Sub

Private Sub mnuNew_Click()
Dim lFile As String
    'Create new database
    lFile = GetOpenDLGName(Filter1, "Create Database", True)
    If Len(lFile) Then
        If CreateNewDatabase(lFile) <> 1 Then
            MsgBox "There was an error creating the database.", vbInformation, frmmain.Caption
        End If
    End If
End Sub

Private Sub mnuNewCat_Click()
Dim cName As String
    
    cName = Trim$(InputBox$("Create New Category.", frmmain.Caption))
    'Check for entered name
    If Len(cName) Then
        'Add new category
        If AddCategory(cName, ParentID) <> 1 Then
            MsgBox "The was an error adding the new category.", vbInformation, frmmain.Caption
        Else
            'Add the new node
            Call LoadCategories(tv1)
            'Select the new node
            Call ClickNode(tv1.Nodes.Count)
        End If
    End If
End Sub

Private Sub mnuNewCode_Click()
    EditOp = 0 'Add new code
    frmAddCode.Show vbModal, frmmain

    'Check if the ok button was pressed
    If (ButtonPress <> vbCancel) Then
        'Add new code
        If AddNewCode(ParentID) <> 1 Then
            MsgBox "There was an error while adding the new code.", vbInformation, frmmain.Caption
        Else
            'Reload the categories
            Call LoadCategories(tv1)
            Call ClickNode(TvID)
            'Last the new item added
            Call SelectItem(LstV.ListItems.Count)
            Call lvSizeColumns(LstV)
        End If
    End If
    
    ButtonPress = vbCancel
    
End Sub

Private Sub mnuOpen_Click()
Dim lFile As String

    'Get filename
    lFile = GetOpenDLGName(Filter1)
    If Len(lFile) Then
        'Check if the database was opened
        If (DBOpen) Then
            db.Close
        End If
        
        If Not dOpenDataBase(lFile) Then
            'Error code
        Else
            Call LoadCategories(tv1)
            'Select the first node
            Call ClickNode(2)
        End If
    End If
End Sub

Private Sub mnuOptions_Click()
    'Show options form
    frmOptions.Show vbModal, frmmain
End Sub

Private Sub mnuProps_Click()
Dim sProp As String

    'Display database properties
    sProp = "DB Version " & db.Version & vbCrLf
    sProp = sProp & "Categories " & RecordCount("Category") & vbCrLf
    sProp = sProp & "Codes " & RecordCount("Codes") & vbCrLf & vbCrLf
    sProp = sProp & "Filesize " & Format(FileLen(dbFile), "#,##0") & " bytes" & vbCrLf
    sProp = sProp & "Updatable " & db.Updatable
    MsgBox sProp, vbInformation, "Properties"
    
    sProp = vbNullString
End Sub

Private Sub mnuRestore_Click()
    Call Tray1_MouseDown(vbLeftButton)
End Sub

Private Sub mnuShow_Click()
    'Show/Hide categories
    mnuShow.Checked = (Not mnuShow.Checked)
    'Hide categories
    If Not (mnuShow.Checked) Then
        pBar3.Visible = False
        tBar3.Visible = False
        tv1.Visible = False
        pBar1.Left = 0
        pBar2.Left = 0
        pBar4.Left = 0
        tBar2.Left = 0
        CodeView.Left = 0
        LstV.Left = 0
        txtComment.Left = 0
    Else
        'Show categories
        pBar3.Visible = True
        tBar3.Visible = True
        tv1.Visible = True
        pBar1.Left = pLeft
        pBar2.Left = pLeft
        pBar4.Left = pLeft
        tBar2.Left = pLeft
        CodeView.Left = pLeft
        LstV.Left = pLeft
        txtComment.Left = pLeft
    End If
    'Resize the controls
    Call Form_Resize
    'Resize ColumnHeader
    Call lvSizeColumns(LstV)
End Sub

Private Sub pBar4_Resize()
    ImgUpdown.Left = (pBar4.ScaleWidth - ImgUpdown.Width) - 60
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lFile As String
    Select Case Button.Key
        Case "M_NEW"
            'Create new database
            Call mnuNew_Click
        Case "M_OPEN"
            'Open Database
            Call mnuOpen_Click
        Case "M_COMP"
            'Compact database
            Call mnuCompact_Click
        Case "M_BACKUP"
            'Backup
            lFile = GetOpenDLGName(Filter1, "Backup", True)
            If Len(lFile) Then
                'Do backup
                Call BackUpDB(lFile)
            End If
        Case "M_ABOUT"
            'About
            Call mnuAbout_Click
        Case "M_EXIT"
            'Exit
            Call mnuExit_Click
    End Select
End Sub

Private Sub tBar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lFile As String
    Select Case Button.Key
        Case "M_NEW"
            'Add new code
            Call mnuNewCode_Click
        Case "M_EDIT"
            'Edit code
            Call mnuEditCode_Click
        Case "M_DEL"
            'Delete code
            Call MnuDeleteCode_Click
        Case "M_SAVE"
            'Save sourcecode
            lFile = GetOpenDLGName(Filter2, "Save", True, LstV.ListItems(lViewID).Text)
            If Len(lFile) Then
                If (mFilterIdx = 1) Then
                    'Save Text file
                    Call SaveText(lFile, CodeView.Text)
                Else
                    'Save snip file
                    Call SaveSnipletFile(lFile)
                End If
                
                'Clear var
                lFile = vbNullString
            End If
        Case "M_CPY"
            'Copy code to the clipboard
            Clipboard.Clear
            Clipboard.SetText CodeView.Text
        Case "M_MOVE"
            'Get category name
            mCatName = tv1.SelectedItem.Text
            'Show move form
            frmMove.Show vbModal, frmmain
            
            If (ButtonPress = vbOK) Then
                'Move the Item
                If MoveCodeItem(mCatName, CodeID) <> 1 Then
                    MsgBox "The was an error while moving the item.", vbInformation, frmmain.Caption
                Else
                    'Select the treeview node
                    Call ClickNode(TvID)
                End If
                ButtonPress = vbCancel
            End If
            
        Case "M_UP"
            Call LMoveItem(True)
        Case "M_DOWN"
            Call LMoveItem(False)
    End Select
End Sub

Private Sub tBar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "M_ADD"
            Call mnuNewCat_Click
        Case "M_EDIT"
            Call mnuEditCat_Click
        Case "M_DEL"
            Call mnuDelCat_Click
    End Select
End Sub

Private Sub Tray1_MouseDown(Button As Integer)
    If (Button = vbLeftButton) Then
        Tray1.Visible = False
        frmmain.WindowState = oWinState
        frmmain.Visible = True
    End If
End Sub

Private Sub Tray1_MouseUp(Button As Integer)
    If (Button = vbRightButton) Then
        PopupMenu mnuTray
    End If
End Sub

Private Sub tv1_AfterLabelEdit(Cancel As Integer, NewString As String)
    If EditCategory(ParentID, NewString) <> 1 Then
        MsgBox "Error while editing category.", vbInformation, frmmain.Caption
    Else
        'Update Treeview text
        tv1.SelectedItem.Text = NewString
        Call tv1_Click
    End If
End Sub

Private Sub tv1_BeforeLabelEdit(Cancel As Integer)
    'Disable editing of top item
    Cancel = (tv1.SelectedItem.Key = "M_TOP")
End Sub

Private Sub tv1_Click()
    
    If (tv1.Nodes.Count = 0) Then
        Exit Sub
    End If
    
    HasItem = False
    TvID = tv1.SelectedItem.Index
    '
    'Enable/Disable menus and button items
    tBar3.Buttons(2).Enabled = tv1.SelectedItem.Key <> "M_TOP"
    tBar3.Buttons(3).Enabled = tBar3.Buttons(2).Enabled
    tBar2.Buttons(2).Enabled = False
    tBar2.Buttons(3).Enabled = False
    tBar2.Buttons(5).Enabled = False
    tBar2.Buttons(6).Enabled = False
    tBar2.Buttons(8).Enabled = False
    tBar2.Buttons(10).Enabled = False
    tBar2.Buttons(11).Enabled = False
    mnuEditCat.Enabled = tBar3.Buttons(2).Enabled
    mnuDelCat.Enabled = tBar3.Buttons(2).Enabled
    mnuNewCode.Enabled = tBar3.Buttons(2).Enabled
    mnuEditCode.Enabled = False
    MnuDeleteCode.Enabled = False
    tBar2.Buttons(1).Enabled = tBar3.Buttons(2).Enabled
    
    CodeView.Text = vbNullString
    txtComment.Text = vbNullString
    
    If (tv1.SelectedItem.Key) = "M_TOP" Then
        ParentID = 0
        'Enable/Disable menus and button items
        If ((tv1.Nodes.Count - 1) = 1) Then
            lblCodes.Caption = "1 item in " & tv1.Nodes(TvID).FullPath
        Else
            lblCodes.Caption = (tv1.Nodes.Count - 1) & " items in " & tv1.Nodes(TvID).FullPath
        End If
        
        mnuEditCode.Enabled = False
        MnuDeleteCode.Enabled = False
        LstV.ListItems.Clear
    Else
        'Node Text
        TvText = tv1.SelectedItem.Text
        'ParentID
        ParentID = Mid(tv1.SelectedItem.Key, 2)
        'Show Records
        Call LoadRecords(LstV, ParentID)
        'Resize ColumnHeader
        Call lvSizeColumns(LstV)
        If (LstV.ListItems.Count = 1) Then
            lblCodes.Caption = LstV.ListItems.Count & " item in " & tv1.Nodes(TvID).FullPath
        Else
            lblCodes.Caption = LstV.ListItems.Count & " items in " & tv1.Nodes(TvID).FullPath
        End If
    End If
    
End Sub

Private Sub tv1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyDelete) Then
        Call mnuDelCat_Click
    End If
End Sub

Private Sub tv1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mButton = Button
End Sub

Private Sub txtFind_Click()
    If (txtFind.Tag = "SR") Then
        txtFind.Text = ""
        txtFind.Tag = ""
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        Call cmdFind_Click
        KeyAscii = 0
    End If
    
End Sub
