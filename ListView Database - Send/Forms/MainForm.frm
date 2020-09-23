VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MainForm 
   Caption         =   "ListView Database - by Special-K"
   ClientHeight    =   8055
   ClientLeft      =   15
   ClientTop       =   30
   ClientWidth     =   11430
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "MainForm.frx":058A
   ScaleHeight     =   8055
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   290
      Left            =   1920
      Picture         =   "MainForm.frx":0E54
      ScaleHeight     =   285
      ScaleWidth      =   2595
      TabIndex        =   27
      Top             =   7210
      Width           =   2595
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   1920
      TabIndex        =   26
      Top             =   7200
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Options..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   188
      TabIndex        =   14
      Top             =   6960
      Width           =   5520
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   4680
         ScaleHeight     =   600
         ScaleWidth      =   735
         TabIndex        =   25
         Top             =   240
         Width           =   735
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   600
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1058
            ButtonWidth     =   1984
            ButtonHeight    =   1005
            Appearance      =   1
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tbSearch"
                  ImageIndex      =   6
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
      End
      Begin VB.OptionButton optstart 
         Caption         =   "Starts with"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   17
         Top             =   650
         Width           =   1095
      End
      Begin VB.OptionButton optany 
         Caption         =   "Any Letter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   16
         Top             =   650
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optmatch 
         Caption         =   "Match Whole Word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   18
         Top             =   650
         Width           =   2295
      End
      Begin VB.ComboBox combocategory 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "MainForm.frx":139E
         Left            =   120
         List            =   "MainForm.frx":13A0
         TabIndex        =   15
         Text            =   "Category..."
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   10080
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":13A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2D34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":32CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3720
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":8336
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":9CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":B65A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":CFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E97E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":10310
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":11CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":13634
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":14FC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":15CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":16584
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":17260
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":17F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":18C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":198F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1A5D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   5990
      ScaleHeight     =   600
      ScaleWidth      =   5295
      TabIndex        =   23
      Top             =   7200
      Width           =   5295
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   600
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1058
         ButtonWidth     =   2143
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refresh"
               Key             =   "tbRefresh"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tbRefreshs"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Key             =   "tbPrint"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tbPrints"
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "About"
               Key             =   "tbAbout"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "tbAbouts"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               Key             =   "tbExit"
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.PictureBox picture1 
      Height          =   255
      Left            =   7200
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   22
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Timer FormFade 
      Interval        =   25
      Left            =   10680
      Top             =   5640
   End
   Begin VB.TextBox tmpRCount 
      Height          =   285
      Left            =   9120
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Edit 
      BackColor       =   &H00008000&
      Caption         =   "Edit..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   195
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   11040
      Begin VB.TextBox EDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   960
         MaxLength       =   255
         TabIndex        =   1
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox EID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox ERemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7920
         MaxLength       =   255
         TabIndex        =   7
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox EBloodGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7440
         MaxLength       =   5
         TabIndex        =   6
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox EDesignation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6120
         MaxLength       =   255
         TabIndex        =   5
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox EDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4800
         MaxLength       =   255
         TabIndex        =   4
         Top             =   360
         Width           =   1380
      End
      Begin VB.TextBox ELastName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3480
         MaxLength       =   255
         TabIndex        =   3
         Top             =   360
         Width           =   1365
      End
      Begin VB.TextBox EFirstName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2080
         MaxLength       =   255
         TabIndex        =   2
         Top             =   360
         Width           =   1450
      End
      Begin VB.CommandButton ESave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9360
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton ECancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -75
      TabIndex        =   11
      Top             =   -75
      Width           =   11520
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Right Click on ListView for Editing..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   3555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use Search options for any specific search."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   3555
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   480
         Picture         =   "MainForm.frx":1AEAC
         Top             =   150
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   188
      TabIndex        =   20
      Top             =   1200
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image Image2 
      Height          =   60
      Left            =   -75
      Picture         =   "MainForm.frx":1BD76
      Stretch         =   -1  'True
      Top             =   885
      Width           =   11520
   End
   Begin VB.Menu lv 
      Caption         =   "Listview"
      Visible         =   0   'False
      Begin VB.Menu maddnew 
         Caption         =   "Add New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu lvsep1 
         Caption         =   "-"
      End
      Begin VB.Menu medit 
         Caption         =   "Edit"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu lvsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mdelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mdeleteselected 
         Caption         =   "Delete Selected"
         Enabled         =   0   'False
      End
      Begin VB.Menu mdeleteall 
         Caption         =   "Delete All"
         Enabled         =   0   'False
      End
      Begin VB.Menu lvsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mrefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu lvsep4 
         Caption         =   "-"
      End
      Begin VB.Menu mprintselected 
         Caption         =   "Print Selected"
         Enabled         =   0   'False
      End
      Begin VB.Menu lvsep5 
         Caption         =   "-"
      End
      Begin VB.Menu menableselection 
         Caption         =   "Enable Selection"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nearly all codes are taken from Planet-Source.com
'Compiled those to become usefull in a way to handle database in listview
'*************************************************************************
'Suggestions Welcome : kaleemullah@windowslive.com
'*************************************************************************

Option Explicit

'LISTVIEW COLUMN SIZE FIT TO TEXT WIDTH
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
    
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
    
    
'LISTVIEW ROW COLORS
Private Enum ImageSizingTypes
   [sizenone] = 0
   [sizeCheckBox]
   [sizeicon]
End Enum

Private Enum LedgerColours
  vbledgerWhite = &HF9FEFF
  vbLedgerGreen = &HD0FFCC
  vbLedgerYellow = &HE1FAFF
  vbLedgerred = &HE1E1FF
  vbLedgerGrey = &HE0E0E0
  vbLedgerBeige = &HD9F2F7
  vbLedgerSoftWhite = &HF7F7F7
  vbledgerPureWhite = &HFFFFFF
End Enum

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long 'SORTING LOCK

Dim fd As Integer 'FORM FADE INTEGER

'MASKED TEXT
Dim CT As clsTextBox
    
'LISTVIEW COLOR
'NOTE: Does'nt works with Listview Font < 10

Private Sub SetListViewLedger(lv As ListView, _
                              Bar1Color As LedgerColours, _
                              Bar2Color As LedgerColours, _
                              nSizingType As ImageSizingTypes)

   Dim iBarHeight  As Long  'HEIGHT OF 1 LINE IN THE LISTVIEW
   Dim lBarWidth   As Long  'WIDTH OF LISTVIEW
   Dim diff        As Long  'USED IN CALCULATIONS OF ROW HEIGHT
   Dim twipsy      As Long  'VARIABLE HOLDING SCREEN.TWIPSPERPICTURE1elY
   
   iBarHeight = 0
   lBarWidth = 0
   diff = 0
   
   On Local Error GoTo SetListViewColor_Error
   
   twipsy = Screen.TwipsPerPixelY
   
   If lv.View = lvwReport Then
   
     'SET UP THE LISTVIEW PROPERTIES
      With lv
        .Picture = Nothing  'CLEAR PICTURE
        .Refresh
        .Visible = 1
        .PictureAlignment = lvwTile
        lBarWidth = .Width
      End With  ' lv
        
     'SET UP THE PICTURE BOX PROPERTIES
      With picture1
         .AutoRedraw = False       'CLEAR PICTURE
         .Picture = Nothing
         .BackColor = vbWhite
         .Height = 1
         .AutoRedraw = True        'ASSURE IMAGE DRAWS
         .BorderStyle = vbBSNone   'OTHER ATTRIBUTES
         .ScaleMode = vbTwips
         .Top = MainForm.Top - 10000  'MOVE IT WAY OFF SCREEN
         .Width = ListView1.Width 'FILLS THE COLOR IN THE ROW THROUGH ITS LENGTH AND TILL LISTVIEW WIDTH
         .Visible = False
         .Font = lv.Font           'ASSURE PICTURE1 FONT MATCHED LISTVIEW FONT
         
      'MATCH PICTURE BOX FONT PROPERTIES WITH LISTVIEW
      With .Font
         .Bold = lv.Font.Bold
         .Charset = lv.Font.Charset
         .Italic = lv.Font.Italic
         .Name = lv.Font.Name
         .Strikethrough = lv.Font.Strikethrough
         .Underline = lv.Font.Underline
         .Weight = lv.Font.Weight
         .Size = lv.Font.Size
         
      End With 'FONT
         
        'CALCULATE THE HEIGHT OF LISTVIEW ROW
        
         iBarHeight = .TextHeight("W")

         iBarHeight = iBarHeight + twipsy 'THIS IS FOR TEXT ONLY
      
        'SINCE WE NEED TWO-TONE BARS, THE PICTURE BOX NEEDS TO BE TWICE AS HIGH
         .Height = iBarHeight * 2
         .Width = lBarWidth
         
        'PAINT THE TWO BARS OF COLOR AND REFRESH
         picture1.Line (0, 0)-(lBarWidth, iBarHeight), Bar1Color, BF
         picture1.Line (0, iBarHeight)-(lBarWidth, iBarHeight * 2), Bar2Color, BF
      
         .AutoSize = True
         .Refresh
         
      End With  'Picture1
     
     'SET THE LV PICTURE TO THE PICTURE1 IMAGE
          lv.Refresh
          lv.Picture = picture1.Image
      
   Else
    
      lv.Picture = Nothing
        
   End If  'lv.View = lvwReport

SetListViewColor_Exit:
On Local Error GoTo 0
Exit Sub
    
SetListViewColor_Error:

  'CLEAR THE LISTVIEW'S PICTURE AND EXIT
   With lv
      .Picture = Nothing
      .Refresh
   End With
   
   Resume SetListViewColor_Exit
    
End Sub

Public Sub FillSearch()

ListView1.Sorted = False 'RETRIEVE FROM SORTED RESULTS TO PREVENT EMPTY FIELDS ERROR

Call Adoconn 'CONNECTING TO DATABASE
Adodc1.RecordSource = "SELECT * FROM Card"
Adodc1.Refresh

ListView1.ListItems.Clear
  
If Adodc1.Recordset.RecordCount = 0 Then    'IF DATABASE EMPTY
      
    'MsgBox "No Records in Database.", vbInformation, Me.Caption
        
    Control 'REFRESH POPUP MENU
    
    AdjustColumn 'ADJUST THE COLUMN WIDTH
    
    Exit Sub
   
End If

Screen.MousePointer = vbHourglass 'DO NOT LET USER DO ANYTHING ELSE UNTIL THIS IS COMPLETED ;)

Dim x As Integer
x = 1

Do

    With ListView1
        
        On Error Resume Next
        .ListItems.Add , , Adodc1.Recordset.Fields("ID")
        .ListItems(x).SubItems(1) = Adodc1.Recordset.Fields("Date")
        .ListItems(x).SubItems(2) = Adodc1.Recordset.Fields("FirstName")
        .ListItems(x).SubItems(3) = Adodc1.Recordset.Fields("LastName")
        .ListItems(x).SubItems(4) = Adodc1.Recordset.Fields("Department")
        .ListItems(x).SubItems(5) = Adodc1.Recordset.Fields("Designation")
        .ListItems(x).SubItems(6) = Adodc1.Recordset.Fields("BloodGroup")
        .ListItems(x).SubItems(7) = Adodc1.Recordset.Fields("Remarks")
                
    End With
    
Dim t As Integer

For t = 1 To ListView1.ListItems.Count

'ListView1.ListItems(t).Bold = True 'DOES'NT WORKS WITH THE ADJUSTCOLUMN, I DON'T KNOW WHY :(
ListView1.ListItems(t).ForeColor = vbBlue 'SAME HERE :(

Next t

ListView1.ListItems(x).ListSubItems(6).ForeColor = vbRed 'WORKS FINE EVEN WITHOUT LOOPING :)
ListView1.ListItems(x).ListSubItems(6).Bold = True 'THIS TOO :)

Adodc1.Recordset.MoveNext

x = x + 1

Loop While Adodc1.Recordset.EOF = False

AdjustColumn 'ADJUST THE COLUMN WIDTH

Control 'REFRESH POPUP MENU

Screen.MousePointer = vbNormal

End Sub

Public Sub Adoconn()

On Error Resume Next
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Data.mdb" & ";Jet OLEDB:Database Password="
Adodc1.RecordSource = "SELECT * FROM Card"
Adodc1.Refresh
tmpRCount.Text = Adodc1.Recordset.RecordCount 'FOR SEARCH TOTAL RECORDS COUNT ONLY

End Sub

Public Sub AdjustColumn()

'LISTVIEW COLUMN SIZE FIT TO COLUMN HEADER LENGTH AND TEXT MAX LENGTH

    Dim col2adjust As Long

    For col2adjust = 0 To ListView1.ColumnHeaders.Count - 1
   
    Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
   
    Next


    Dim col2adjust1 As Long
   
    col2adjust1 = ListView1.ColumnHeaders.Count - 1
   
    Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
            

    Dim col2adjust2 As Long

    For col2adjust2 = 0 To ListView1.ColumnHeaders.Count - 1
   
    Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE)

    Next

End Sub

Public Sub Control()

'CHANGE POPUP MENU ENABLE DISABLE ACCORDING TO DATABASE RECORDS COUNT

If ListView1.ListItems.Count <> 0 Then

    If menableselection.Checked = False Then

        medit.Enabled = True
        mdelete.Enabled = True
        mdeleteall.Enabled = True
        menableselection.Enabled = True
        
    End If

ElseIf ListView1.ListItems.Count = 0 Then

    maddnew.Enabled = True
    medit.Enabled = False
    mdelete.Enabled = False
    mdeleteselected.Enabled = False
    mdeleteall.Enabled = False
    mrefresh.Enabled = True
    menableselection.Enabled = False
    menableselection.Checked = False
    menableselection.Caption = "Enable Selection"
    ListView1.Checkboxes = False

End If

Frame1.Caption = "Search Options..."

End Sub

Public Sub SotrListView(ByVal lstView As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional ByVal TypeSort As String = "NUMBER")

    On Error Resume Next
  
    With lstView
    
        'CHANGE MOUSE CURSOR TO HOURGLASS WHILE SORTING
        
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        LockWindowUpdate .hWnd
        
        'CHECK THE DATA TYPE OF THE COLUMN BEING SORTED AND ACT ACCORDINGLY e.g. "NUMBER","DATE","" etc.
        
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Dim lngIndex As Long
        lngIndex = ColumnHeader.Index - 1
    
        Select Case UCase$(TypeSort)
        
            Case "DATE"
        
                'SORT BY DATE
            
                strFormat = "YYYYMMDDHhNnSs"
        
                With .ListItems
                    If (lngIndex > 0) Then
                        For l = 1 To .Count
                            With .Item(l).ListSubItems(lngIndex)
                                .Tag = .Text & Chr$(0) & .Tag
                                If IsDate(.Text) Then
                                    .Text = Format(CDate(.Text), strFormat)
                                Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            'SORT THE LIST ALPHABETICALLY BY THIS COLUMN
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            'RESTORE PREVIOUS TAGS AND VALUES OF LISTVEIW
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
        Case "NUMBER"
        
            'SORT NUMERICALLY
        
            strFormat = String(30, "0") & "." & String(30, "0")
        
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(Val(.Text)) Then
                                If CDbl(Val(.Text)) >= 0 Then
                                    .Text = Format(CDbl(Val(.Text)), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(Val(.Text)), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            .Tag = .Text & Chr$(0) & .Tag
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            'SORT THE LIST ALPHABETICALLY BY THIS COLUMN
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .Item(l).ListSubItems(lngIndex)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .Item(l)
                            strData = Split(.Tag, Chr$(0))
                            .Text = strData(0)
                            .Tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else   'ASSUME SORT BY STRING Assume sort by string
            
            'PROVIDED DEFAULT
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
            
        End Select
    
        'UNLOCK THE LISTVIEW SO THAT OCX CAN UPDATE IT
        
        LockWindowUpdate 0&
        
        'RESTORE PREVIOUS MOUSE CURSOR
        
        .MousePointer = lngCursor
    
    End With
    
    Set lstView = Nothing
    Set ColumnHeader = Nothing
    
End Sub

'USED TO ENABLE NEGATIVE NUMBERS TO BE SORTED ALPHABATICALLY BY SWITCHING THE CHARACTERS
Private Function InvNumber(ByVal Number As String) As String
    
    Static i As Integer
    
    For i = 1 To Len(Number)
        
        Select Case Mid$(Number, i, 1)
            
            Case "-": Mid$(Number, i, 1) = " "
            Case "0": Mid$(Number, i, 1) = "9"
            Case "1": Mid$(Number, i, 1) = "8"
            Case "2": Mid$(Number, i, 1) = "7"
            Case "3": Mid$(Number, i, 1) = "6"
            Case "4": Mid$(Number, i, 1) = "5"
            Case "5": Mid$(Number, i, 1) = "4"
            Case "6": Mid$(Number, i, 1) = "3"
            Case "7": Mid$(Number, i, 1) = "2"
            Case "8": Mid$(Number, i, 1) = "1"
            Case "9": Mid$(Number, i, 1) = "0"
        
        End Select
    
    Next
    
    InvNumber = Number

End Function

Private Sub EBloodGroup_Change()

CT.MaskedText EBloodGroup, "@@"

End Sub

Private Sub EDate_Change()

CT.MaskedText EDate, "## / ## / ####"

End Sub

Private Sub EDate_GotFocus()

EDate.SelStart = 0
EDate.SelLength = Len(EDate.Text)

End Sub

Private Sub EID_Change()

CT.MaskedText EID, "@@@-###"

End Sub

Private Sub mdelete_Click()

Dim del
                    
del = MsgBox("Are you sure you want to DELETE (" & ListView1.SelectedItem.Text & ") record?", vbInformation + vbYesNo, Me.Caption)

If del = vbNo Then Exit Sub

Call Adoconn 'CONNECTING TO DATABASE
Adodc1.RecordSource = "SELECT * FROM Card Where ID='" & ListView1.SelectedItem.Text & "'"
Adodc1.Refresh
                  
If Adodc1.Recordset.EOF = True Then

    MsgBox "Sorry, Either the Record is deleted or is an Invalid data.", vbCritical, Me.Caption
    Exit Sub
    
End If

On Error Resume Next
Adodc1.Recordset.Delete

MsgBox "Record DELETED.", , Me.Caption

FillSearch 'REFRESH LISTVIEW WITH RECORDS

End Sub

Private Sub mdeleteall_Click()

Dim dela

dela = MsgBox("Are you sure you want to DELETE Current (" & ListView1.ListItems.Count & ") Records?", vbInformation + vbYesNo, Me.Caption)

If dela = vbNo Then Exit Sub

'COUNT TIME ELAPSED IN DELETING
Dim start_time As Single
Dim stop_time As Single
DoEvents
start_time = Timer

Dim tmplistcount As Integer
        
tmplistcount = ListView1.ListItems.Count 'TEMPORARILY STORE THE LISTVIEW COUNT

Dim da As Long

If ListView1.ListItems.Count = 0 Then Exit Sub

For da = 1 To ListView1.ListItems.Count

Call Adoconn 'CONNECTING TO DATABASE
Adodc1.RecordSource = "SELECT * FROM Card Where ID='" & ListView1.SelectedItem & "'"
Adodc1.Refresh

'IF ID EXIST, VERIFY ITS SELECTION

    If Adodc1.Recordset.Fields("ID") = ListView1.SelectedItem Then   'CORRECT ID
        
        On Error Resume Next
        Adodc1.Recordset.Delete
        Adodc1.Refresh
                                     
        FillSearch 'REFRESH LISTVIEW WITH RECORDS

    End If
       
Next da

MsgBox "All (" & tmplistcount & ") Records DELETED Successfully in " & Format$(stop_time - start_time, "0.00") & " Secs.", , Me.Caption

End Sub

Private Sub mdeleteselected_Click()

Dim ci As Integer

For ci = 1 To ListView1.ListItems.Count

If ListView1.ListItems(ci).Checked = True Then 'ALLOW THIS ACTION ONLY WHEN ANY RECORD IS CHECKED

    Dim dels

    dels = MsgBox("Are you sure you want to DELETE Selected Records?", vbInformation + vbYesNo, Me.Caption)
    
        If dels = vbNo Then Exit Sub

            'COUNT TIME ELAPSED IN DELETING
            Dim start_time As Single
            Dim stop_time As Single
            DoEvents
            start_time = Timer

            Dim s As String
            Dim i As Integer

    For i = 1 To ListView1.ListItems.Count
   
        If ListView1.ListItems(i).Checked = True Then
      
            s = ListView1.ListItems(i)  'IF AN ITEM IS CHECKED, ADD IT TO STRING S
      
            'NOW IF ANY ITEM OR ITEMS ARE CHECKED
            If s <> "" Then
        
                Call Adoconn 'CONNECTING TO DATABASE
                Adodc1.RecordSource = "SELECT * FROM Card Where ID='" & s & "'"
                Adodc1.Refresh
    
                'IF ID EXIST, VERIFY ITS SELECTION

                    If Adodc1.Recordset.Fields("ID") = s Then    'CORRECT ID

                        On Error Resume Next
                        Adodc1.Recordset.Delete
                        Adodc1.Refresh
 
                    End If
                
            End If  'If s <> "" Then

        End If  'ListView1.ListItems(i).Checked = True

    Next i

    MsgBox "All Selected Records DELETED Successfully in " & Format$(stop_time - start_time, "0.00") & " Secs.", , Me.Caption

    FillSearch 'REFRESH LISTVIEW WITH RECORDS
    
    Exit Sub 'PREVENT REPEATED ACTION

End If 'If ListView1.ListItems(ci).Checked = True Then
    
Next ci

End Sub

Private Sub medit_Click()

Edit.Move 0 / 5, ListView1.Height / 2, Edit.Width, Edit.Height

'EDITING A SELECTED RECORD

Edit.Visible = True
Edit.Caption = "Edit..."
EID.Enabled = False

'FIELD TEXT FILEDS WITH LISTVIEW COLUMN DATA

EID.Text = ListView1.SelectedItem.Text                  'FIRST COLUMN DATA
EDate.Text = ListView1.SelectedItem.SubItems(1)
EFirstName.Text = ListView1.SelectedItem.SubItems(2)
ELastName.Text = ListView1.SelectedItem.SubItems(3)
EDepartment.Text = ListView1.SelectedItem.SubItems(4)
EDesignation.Text = ListView1.SelectedItem.SubItems(5)
EBloodGroup.Text = ListView1.SelectedItem.SubItems(6)
ERemarks.Text = ListView1.SelectedItem.SubItems(7)

EDate.SetFocus

End Sub

Private Sub menableselection_Click()

If menableselection.Checked = False Then

    Dim c As Long

        If ListView1.ListItems.Count = 0 Then Exit Sub
    
    For c = 1 To ListView1.ListItems.Count
        
        ListView1.ListItems(c).Checked = False 'UNCHECK ANY PREVIOUSLY CHECKED ITEM TO PREVENT UNEXPECTED DELETION
    
    Next c

    maddnew.Enabled = False 'DISABLE ADD NEW
    medit.Enabled = False 'DISABLE EDIT
    mdelete.Enabled = False 'DISABLE DELETE
    mdeleteall.Enabled = False 'DISABLE DELETE ALL
    mdeleteselected.Enabled = True 'ENABLE DELETE SELECTED
    
    ListView1.Checkboxes = True
    menableselection.Checked = True
    menableselection.Caption = "Disable Selection"
    
    AdjustColumn 'ADJUST THE COLUMN WIDTH
    
Else

    Dim c1 As Long

        If ListView1.ListItems.Count = 0 Then Exit Sub
    
    For c1 = 1 To ListView1.ListItems.Count
        
        ListView1.ListItems(c1).Checked = False 'UNCHECK ANY PREVIOUSLY CHECKED ITEM TO PREVENT UNEXPECTED DELETION
    
    Next c1


    maddnew.Enabled = True 'ENABLE ADD NEW
    medit.Enabled = True 'ENABLE EDIT
    mdelete.Enabled = True 'ENABLE DELETE
    mdeleteall.Enabled = True 'ENABLE DELETE ALL
    mdeleteselected.Enabled = False 'DISABLE DELETE SELECTED
    
    ListView1.Checkboxes = False
    menableselection.Checked = False
    menableselection.Caption = "Enable Selection"

    AdjustColumn 'ADJUST THE COLUMN WIDTH
    
End If

End Sub

Private Sub ESave_Click()
       
    'VALIDATE EMPTY FIELDS
    If EID.Text = "" Then
    
        MsgBox "Please enter ID.", vbCritical, Me.Caption
        EID.SetFocus
        Exit Sub
        
    ElseIf EID.ForeColor = vbRed Then
    
        MsgBox "Wrong ID. Should be like ABC-123", vbCritical, Me.Caption
        EDate.SetFocus
        Exit Sub
    
    ElseIf EDate.Text = "" Then
    
        MsgBox "Please enter the Date.", vbCritical, Me.Caption
        EDate.SetFocus
        Exit Sub

    ElseIf EDate.ForeColor = vbRed Then
    
        MsgBox "Wrong Date. Should be like mm / dd / yyyy", vbCritical, Me.Caption
        EDate.SetFocus
        Exit Sub
    
    ElseIf EFirstName.Text = "" Then
    
        MsgBox "Please enter the First name.", vbCritical, Me.Caption
        EFirstName.SetFocus
        Exit Sub
    
    ElseIf ELastName.Text = "" Then
    
        MsgBox "Please enter the Last name.", vbCritical, Me.Caption
        ELastName.SetFocus
        Exit Sub
    
    ElseIf EDepartment.Text = "" Then
    
        MsgBox "Please enter the Department.", vbCritical, Me.Caption
        EDepartment.SetFocus
        Exit Sub
    
    ElseIf EDesignation.Text = "" Then
    
        MsgBox "Please enter the Designation.", vbCritical, Me.Caption
        EDesignation.SetFocus
        Exit Sub
    
    ElseIf EBloodGroup.Text = "" Then
    
        MsgBox "Please enter BloodGroup.", vbCritical, Me.Caption
        EBloodGroup.SetFocus
        Exit Sub
    
    ElseIf EBloodGroup.ForeColor = vbRed Then
    
        MsgBox "Wrong BloodGroup. Should be like A+", vbCritical, Me.Caption
        EBloodGroup.SetFocus
        Exit Sub
        
    ElseIf ERemarks.Text = "" Then
    
        MsgBox "Please enter Remarks.", vbCritical, Me.Caption
        ERemarks.SetFocus
        Exit Sub
       
    End If  'If EID.Text = "" Then
    

        'IF ALL FIELDS ARE VALID
        
            If Edit.Caption = "Add New..." Then
                
                'IF ADD NEW SELECTED (AS NEW RECORD), THEN CHECK ID FIRST IN DATABASE TO PREVENT DOUBLE ENTRY

                Call Adoconn 'CONNECTING TO DATABASE
                Adodc1.RecordSource = "SELECT * FROM Card Where ID='" & EID.Text & "'"
                Adodc1.Refresh
       
                If Adodc1.Recordset.RecordCount > 0 Then 'IF EXIST

                    MsgBox "The ID " & EID.Text & " already exists. Please choose another.", vbCritical, Me.Caption
                    EID.SetFocus

                Exit Sub
        
                Else
                
                    'PROCEED IF NEW RECORD IS BEING ADDED
            
                    Dim sav
                    
                    sav = MsgBox("Are you sure you want to SAVE this record?", vbInformation + vbYesNo, Me.Caption)

                        If sav = vbNo Then Exit Sub

                    Call Adoconn 'CONNECTING TO DATABASE
                    Adodc1.RecordSource = "Select * from Card"
           
                    On Error Resume Next
                    Adodc1.Recordset.AddNew
                        
                    Adodc1.Recordset.Fields("ID").Value = EID.Text
                    Adodc1.Recordset.Fields("Date").Value = EDate.Text
                    Adodc1.Recordset.Fields("FirstName").Value = EFirstName.Text
                    Adodc1.Recordset.Fields("LastName").Value = ELastName.Text
                    Adodc1.Recordset.Fields("Department").Value = EDepartment.Text
                    Adodc1.Recordset.Fields("Designation").Value = EDesignation.Text
                    Adodc1.Recordset.Fields("BloodGroup").Value = EBloodGroup.Text
                    Adodc1.Recordset.Fields("Remarks").Value = ERemarks.Text
                                
                    On Error Resume Next
                    Adodc1.Recordset.Update
            
                    MsgBox "Record SAVED.", , Me.Caption
                    Edit.Visible = False 'CLOSE THE EDIT FORM
            
                    FillSearch 'REFRESH LISTVIEW WITH RECORDS
                
                End If  'Adodc1.Recordset.RecordCount > 0 Then
            
            End If  'Edit.Caption = "Add New..." Then
                   
            '\\\\\\\\\\
            
            If Edit.Caption = "Edit..." Then
            
            'IF EDIT (EDITING AN EXISTING RECORD) IS SELECTED
                
                Dim upd
                    
                    upd = MsgBox("Are you sure you want to UPDATE this record?", vbInformation + vbYesNo, Me.Caption)

                        If upd = vbNo Then Exit Sub
                    
                    'FIND THE ID TO EDIT
                    
                    Call Adoconn 'CONNECTING TO DATABASE
                    Adodc1.RecordSource = "SELECT * FROM Card Where ID='" & EID.Text & "'"
                    Adodc1.Refresh
                    
                    'FILL ALL THE TEXT FIELDS WITH THE FOUND RECORD TO EDIT
                    'ID IS DISABLED TO PREVENT DOUBLE ENTRY ERROR
                    'IF ID IS TO BE CHANGED THEN FIRST DELETE THE EXISTING AND THEN USE ADD NEW
                    
                    Adodc1.Recordset.Fields("ID").Value = EID.Text
                    Adodc1.Recordset.Fields("Date").Value = EDate.Text
                    Adodc1.Recordset.Fields("FirstName").Value = EFirstName.Text
                    Adodc1.Recordset.Fields("LastName").Value = ELastName.Text
                    Adodc1.Recordset.Fields("Department").Value = EDepartment.Text
                    Adodc1.Recordset.Fields("Designation").Value = EDesignation.Text
                    Adodc1.Recordset.Fields("BloodGroup").Value = EBloodGroup.Text
                    Adodc1.Recordset.Fields("Remarks").Value = ERemarks.Text
                                
                    On Error Resume Next
                    Adodc1.Recordset.Update
            
                    MsgBox "Record UPDATED.", , Me.Caption
                    Edit.Visible = False 'CLOSE THE EDIT FORM
            
                    FillSearch 'REFRESH LISTVIEW WITH RECORDS
            
            End If  'Edit.Caption = "Edit..." Then
                   
End Sub

Private Sub Form_Resize()

        'MOVE CONTROLS WHEN FORM RESIZED
        
        If Me.Width < 11550 Or Me.Height < 8460 Then    'PREVENT NEGATIVE VALUE ERROR
        
        Exit Sub
        
        Else
        
        Frame2.Move 0, -75, Me.Width, 975   'TOP PICTURE
        
        Image2.Move 0, 885, Me.Width, 60    'TOP COLORED LINE
        
        ListView1.Move 0, 1200, Me.Width - 120, Me.Height - 2800    'LISTVIEW
        
        Frame1.Move 50, ListView1.Height + 1300, Me.Width - 6000, 975 'SEARCH FRAME
        
        Picture3.Move Me.Width - 5500, ListView1.Height + 1600, 5295, 600   'EXIT TOOLBAR
        
        txtSearch.Move 1700, Frame1.Top + 240, Frame1.Width - 2700, txtSearch.Height    'TXTSEARCH
        
        Picture2.Move 1750, txtSearch.Top + 10, txtSearch.Width - 100, Picture2.Height    'SEARCH GOOGLE LIKE PICTURE
        
        Picture4.Move Frame1.Width - 900, 265, Picture4.Width, Picture4.Height 'SEARCH BUTTON
        
        Edit.Move 0, ListView1.Height / 2, Edit.Width, Edit.Height  'EDIT FRAME
        
        Call AdjustColumn   'ADJUST COLUMN ON RESIZE
        
        End If
        
End Sub

Private Sub Form_Load()

'PROGRAM TO RUN ONLY ONE AT A TIME
If App.PrevInstance Then
    
    MsgBox "Already Open.", vbOKOnly + vbCritical, Me.Caption
    ActivatePrevInstance 'SHOW RUNNING PROGRAM
    
End If

'FORM FADE IN EFFECT
fd = 0
Transparent.ofFrm hWnd, 0

'CENTER FORM
Call CenterForm(Me)

'CONNECT TO DATABASE
Call Adoconn

With ListView1 'ASSIGN COLUMN HEADERS AS PER DATABASE RECORDS NAME
        
        On Error Resume Next
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(1).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(2).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(3).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(4).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(5).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(6).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(7).Name
        .ColumnHeaders.Add , , Adodc1.Recordset.Fields(8).Name
            
        .ColumnHeaders(1).Alignment = lvwColumnLeft
        .ColumnHeaders(2).Alignment = lvwColumnCenter
        .ColumnHeaders(3).Alignment = lvwColumnLeft
        .ColumnHeaders(4).Alignment = lvwColumnLeft
        .ColumnHeaders(5).Alignment = lvwColumnLeft
        .ColumnHeaders(6).Alignment = lvwColumnLeft
        .ColumnHeaders(7).Alignment = lvwColumnCenter
        .ColumnHeaders(8).Alignment = lvwColumnLeft
        
        On Error Resume Next
        .ColumnHeaders(8).Icon = 20
        
End With

With combocategory 'FILL COMBO WITH DATABASE FIELD'S NAME
    
    On Error Resume Next
    .AddItem (Adodc1.Recordset.Fields(1).Name)
    .AddItem (Adodc1.Recordset.Fields(2).Name)
    .AddItem (Adodc1.Recordset.Fields(3).Name)
    .AddItem (Adodc1.Recordset.Fields(4).Name)
    .AddItem (Adodc1.Recordset.Fields(5).Name)
    .AddItem (Adodc1.Recordset.Fields(6).Name)
    .AddItem (Adodc1.Recordset.Fields(7).Name)
    .AddItem (Adodc1.Recordset.Fields(8).Name)

End With


'LISTVIEW ROW COLOR
Call SetListViewLedger(ListView1, &HC0FFC0, vbWhite, sizenone)

'REFRESH LISTVIEW WITH RECORDS
FillSearch

'GOOGLE LOOKING SEARCH TEXT
Picture2.Visible = True

'MASKED TEXT
Set CT = New clsTextBox


End Sub

Private Sub Form_Unload(Cancel As Integer)

Cancel = False
End

End Sub

Private Sub FormFade_Timer()

'FORM FADE IN EFFECT Control

If fd <= 100 Then
 
    Transparent.ofFrm hWnd, fd
    fd = fd + 20

Else

     FormFade.Enabled = False
     Transparent.ofFrm hWnd, 255

End If

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

'RIGHT CLICK POPUP MENU
If Button = 2 Then Me.PopupMenu lv, , , , maddnew

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  
'COLUMN SORTING
'IF LEFT "" DOES NORMAL DEFAULT SORTING
'CASE NUMBER (COLUMNHEADER.INDEX) DEFINES THE COLUMN NUMBER
  
  Select Case ColumnHeader.Index
  
    Case 1: SotrListView ListView1, ColumnHeader, ""
    Case 2: SotrListView ListView1, ColumnHeader, "DATE"
    Case 3: SotrListView ListView1, ColumnHeader, ""
    Case 4: SotrListView ListView1, ColumnHeader, ""
    Case 5: SotrListView ListView1, ColumnHeader, ""
    Case 6: SotrListView ListView1, ColumnHeader, ""
    Case 7: SotrListView ListView1, ColumnHeader, ""
    Case 8: SotrListView ListView1, ColumnHeader, ""
    
  End Select
  
End Sub

Private Sub maddnew_Click()

'ADD NEW RECORD(S)

Edit.Visible = True
Edit.Caption = "Add New..."
EID.Enabled = True
EID.SetFocus

'CLEAR TEXT FIELDS FOR NEW ENTRIES

EID.Text = ""
EFirstName.Text = ""
ELastName.Text = ""
EDepartment.Text = ""
EDesignation.Text = ""
EBloodGroup.Text = ""
ERemarks.Text = ""

End Sub

Public Sub lvSearch()

'IF NO RECORDS IN DATABASE
        On Error Resume Next
        If Adodc1.Recordset.RecordCount = 0 Then Exit Sub

        If combocategory.Text = "Category..." Then

            MsgBox "Please select a category of search", vbCritical, Me.Caption
            combocategory.SetFocus
            Exit Sub
    
        End If
     
        If txtSearch.Text = "" Then
         
            MsgBox "Enter a search string.", vbCritical, Me.Caption
            txtSearch.SetFocus
            Exit Sub
      
        End If

        'COUNT TIME ELAPSED IN SEARCHING
        Dim start_time As Single
        Dim stop_time As Single
        DoEvents
        start_time = Timer

        'SEARCHING OPTIONS

        If optany.Value = True Then   'ANY LETTER OPTION SELECTED

            Call Adoconn 'CONNECTING TO DATABASE
            Adodc1.RecordSource = "SELECT * from Card where " & combocategory.Text & " LIKE '%" + Trim(txtSearch.Text) + "%'"
            Adodc1.Refresh

            If Adodc1.Recordset.EOF = True Then 'IF NOT FOUND
    
                MsgBox "No Match found", vbCritical, Me.Caption
                FillSearch 'REFRESH LISTVIEW WITH RECORDS
                Exit Sub

            End If
      
            'CLEAR THE LISTVIEW AND SHOW FOUND ITEM(S)
            ListView1.ListItems.Clear

            Dim x As Integer
            x = 1

            Do

                With ListView1
        
                    On Error Resume Next
                    .ListItems.Add , , Adodc1.Recordset.Fields("ID")
                    .ListItems(x).SubItems(1) = Adodc1.Recordset.Fields("Date")
                    .ListItems(x).SubItems(2) = Adodc1.Recordset.Fields("FirstName")
                    .ListItems(x).SubItems(3) = Adodc1.Recordset.Fields("LastName")
                    .ListItems(x).SubItems(4) = Adodc1.Recordset.Fields("Department")
                    .ListItems(x).SubItems(5) = Adodc1.Recordset.Fields("Designation")
                    .ListItems(x).SubItems(6) = Adodc1.Recordset.Fields("BloodGroup")
                    .ListItems(x).SubItems(7) = Adodc1.Recordset.Fields("Remarks")
        
                End With

                Dim t As Integer

                For t = 1 To ListView1.ListItems.Count

                    'ListView1.ListItems(t).Bold = True 'DOES'NT WORKS WITH THE ADJUSTCOLUMN, I DON'T KNOW WHY :(
                    ListView1.ListItems(t).ForeColor = vbBlue 'SAME HERE :(

                Next t

                ListView1.ListItems(x).ListSubItems(6).ForeColor = vbRed 'WORKS FINE EVEN WITHOUT LOOPING :)
                ListView1.ListItems(x).ListSubItems(6).Bold = True 'THIS TOO :)

                Adodc1.Recordset.MoveNext

                x = x + 1

            Loop While Adodc1.Recordset.EOF = False

            stop_time = Timer
            Frame1.Caption = ""
            Frame1.Caption = "Search Options..." & "Found (" & ListView1.ListItems.Count & ") Match(s) out of Records (" & tmpRCount.Text & ") in " & Format$(stop_time - start_time, "0.00") & " Secs."

            AdjustColumn 'ADJUST THE COLUMN WIDTH
    
        End If  'ANY LETTER OPTION SELECTED

'\\\\\\\\\\

        If optstart.Value = True Then   'STARTS WITH OPTION SELECTED

            Call Adoconn 'CONNECTING TO DATABASE
            Adodc1.RecordSource = "SELECT * FROM Card Where " & combocategory.Text & " like '" & txtSearch.Text & "%'"
            Adodc1.Refresh

                If Adodc1.Recordset.EOF = True Then 'IF NOT FOUND
    
                    MsgBox "No Match found", vbCritical, Me.Caption
                    FillSearch 'REFRESH LISTVIEW WITH RECORDS
                    Exit Sub

                End If
       
            'CLEAR THE LISTVIEW AND SHOW FOUND ITEM(S)
            ListView1.ListItems.Clear

            Dim x1 As Integer
            x1 = 1

            Do

                With ListView1
        
                    On Error Resume Next
                    .ListItems.Add , , Adodc1.Recordset.Fields("ID")
                    .ListItems(x1).SubItems(1) = Adodc1.Recordset.Fields("Date")
                    .ListItems(x1).SubItems(2) = Adodc1.Recordset.Fields("FirstName")
                    .ListItems(x1).SubItems(3) = Adodc1.Recordset.Fields("LastName")
                    .ListItems(x1).SubItems(4) = Adodc1.Recordset.Fields("Department")
                    .ListItems(x1).SubItems(5) = Adodc1.Recordset.Fields("Designation")
                    .ListItems(x1).SubItems(6) = Adodc1.Recordset.Fields("BloodGroup")
                    .ListItems(x1).SubItems(7) = Adodc1.Recordset.Fields("Remarks")
            
                End With

                Dim t1 As Integer

                For t1 = 1 To ListView1.ListItems.Count

                    'ListView1.ListItems(t).Bold = True 'DOES'NT WORKS WITH THE ADJUSTCOLUMN, I DON'T KNOW WHY :(
                    ListView1.ListItems(t1).ForeColor = vbBlue 'SAME HERE :(

                Next t1

                ListView1.ListItems(x1).ListSubItems(6).ForeColor = vbRed 'WORKS FINE EVEN WITHOUT LOOPING :)
                ListView1.ListItems(x1).ListSubItems(6).Bold = True 'THIS TOO :)

                Adodc1.Recordset.MoveNext

                x1 = x1 + 1

            Loop While Adodc1.Recordset.EOF = False

            stop_time = Timer
            Frame1.Caption = ""
            Frame1.Caption = "Search Options..." & "Found (" & ListView1.ListItems.Count & ") Match(s) out of Records (" & tmpRCount.Text & ") in " & Format$(stop_time - start_time, "0.00") & " Secs."

            AdjustColumn 'ADJUST THE COLUMN WIDTH
    
        End If  'STARTS WITH OPTION SELECTED

'\\\\\\\\\\

        If optmatch.Value = True Then   'MATCH WHOLE WORD OPTION SELECTED

            Call Adoconn 'CONNECTING TO DATABASE
            Adodc1.RecordSource = "SELECT * FROM Card Where " & combocategory.Text & "='" & txtSearch.Text & "'"
            Adodc1.Refresh

                If Adodc1.Recordset.EOF = True Then 'IF NOT FOUND
    
                    MsgBox "No Match found", vbCritical, Me.Caption
                    FillSearch 'REFRESH LISTVIEW WITH RECORDS
                    Exit Sub

                End If
       
            'CLEAR THE LISTVIEW AND SHOW FOUND ITEM(S)
            ListView1.ListItems.Clear
  
            Dim x2 As Integer
            x2 = 1

            Do

                With ListView1
        
                    On Error Resume Next
                    .ListItems.Add , , Adodc1.Recordset.Fields("ID")
                    .ListItems(x2).SubItems(1) = Adodc1.Recordset.Fields("Date")
                    .ListItems(x2).SubItems(2) = Adodc1.Recordset.Fields("FirstName")
                    .ListItems(x2).SubItems(3) = Adodc1.Recordset.Fields("LastName")
                    .ListItems(x2).SubItems(4) = Adodc1.Recordset.Fields("Department")
                    .ListItems(x2).SubItems(5) = Adodc1.Recordset.Fields("Designation")
                    .ListItems(x2).SubItems(6) = Adodc1.Recordset.Fields("BloodGroup")
                    .ListItems(x2).SubItems(7) = Adodc1.Recordset.Fields("Remarks")
            
                End With

                Dim t2 As Integer

                For t2 = 1 To ListView1.ListItems.Count

                    'ListView1.ListItems(t).Bold = True 'DOES'NT WORKS WITH THE ADJUSTCOLUMN, I DON'T KNOW WHY :(
                    ListView1.ListItems(t2).ForeColor = vbBlue 'SAME HERE :(

                Next t2

                ListView1.ListItems(x2).ListSubItems(6).ForeColor = vbRed 'WORKS FINE EVEN WITHOUT LOOPING :)
                ListView1.ListItems(x2).ListSubItems(6).Bold = True 'THIS TOO :)

                Adodc1.Recordset.MoveNext

                x2 = x2 + 1

            Loop While Adodc1.Recordset.EOF = False

            stop_time = Timer
            Frame1.Caption = ""
            Frame1.Caption = "Search Options..." & "Found (" & ListView1.ListItems.Count & ") Match(s) out of Records (" & tmpRCount.Text & ") in " & Format$(stop_time - start_time, "0.00") & " Secs."

            AdjustColumn 'ADJUST THE COLUMN WIDTH

        End If  'MATCH WHOLE WORD OPTION SELECTED

End Sub

Private Sub eid_KeyPress(KeyAscii As Integer)

If KeyAscii < 123 And KeyAscii > 96 Then

    KeyAscii = KeyAscii - 32

End If

ESave.Default = True

End Sub

Private Sub efirstname_KeyPress(KeyAscii As Integer)

If KeyAscii < 123 And KeyAscii > 96 Then

    KeyAscii = KeyAscii - 32

End If

ESave.Default = True

End Sub

Private Sub elastname_KeyPress(KeyAscii As Integer)

If KeyAscii < 123 And KeyAscii > 96 Then

    KeyAscii = KeyAscii - 32

End If

ESave.Default = True

End Sub

Private Sub edepartment_KeyPress(KeyAscii As Integer)

ESave.Default = True

End Sub

Private Sub edesignation_KeyPress(KeyAscii As Integer)

ESave.Default = True

End Sub

Private Sub ebloodgroup_KeyPress(KeyAscii As Integer)

If KeyAscii < 123 And KeyAscii > 96 Then

    KeyAscii = KeyAscii - 32

End If

ESave.Default = True

End Sub

Private Sub eremarks_KeyPress(KeyAscii As Integer)

ESave.Default = True

End Sub


Private Sub mprintselected_Click()
   
'I AM STILL WORKING ON THIS ;)

End Sub

Private Sub mrefresh_Click()

FillSearch 'REFRESH LISTVIEW WITH RECORDS

End Sub

Private Sub Picture2_Click()

Picture2.Visible = False
txtSearch.SetFocus

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "tbRefresh"

        FillSearch 'REFRESH LISTVIEW WITH RECORDS

    Case "tbPrint"
    
        'SET REPORT TO ADODC1 CURRENT RECORD POSITION (i.e. LISTVIEW), USEFULL FOR PRINTING
        Set ListViewReport.DataSource = Adodc1.Recordset

        ListViewReport.Show

    Case "tbAbout"
    
        MsgBox "Write me: kaleemullah@windowslive.com", , Me.Caption
        
    Case "tbExit"

        End

End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key

    Case "tbSearch"

        Call lvSearch
                
End Select

End Sub

Private Sub txtSearch_GotFocus()

Picture2.Visible = False
EBloodGroup.SelStart = 0
EBloodGroup.SelLength = Len(EBloodGroup.Text)

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

If KeyAscii < 123 And KeyAscii > 96 Then

    KeyAscii = KeyAscii - 32

End If

If KeyAscii = 13 Then Call lvSearch

End Sub

Private Sub combocategory_KeyPress(KeyAscii As Integer)

KeyAscii = 0 'FORCE NO EDIT IN COMBOLIST AS IT CONTAINS THE FIELD NAMES OF DATABASE

End Sub

Private Sub EBloodGroup_GotFocus()

EBloodGroup.SelStart = 0
EBloodGroup.SelLength = Len(EBloodGroup.Text)

End Sub

Private Sub ECancel_Click()

'CLOSE  THE EDIT FORM AND DO NOTHING
EID.Enabled = False
Edit.Visible = False

End Sub

Private Sub EDepartment_GotFocus()

EDepartment.SelStart = 0
EDepartment.SelLength = Len(EDepartment.Text)

End Sub

Private Sub EDesignation_GotFocus()

EDesignation.SelStart = 0
EDesignation.SelLength = Len(EDesignation.Text)

End Sub

Private Sub EFirstName_GotFocus()

EFirstName.SelStart = 0
EFirstName.SelLength = Len(EFirstName.Text)

End Sub

Private Sub EID_GotFocus()

EID.SelStart = 0
EID.SelLength = Len(EID.Text)

End Sub

Private Sub ELastName_GotFocus()

ELastName.SelStart = 0
ELastName.SelLength = Len(ELastName.Text)

End Sub

Private Sub ERemarks_GotFocus()

ERemarks.SelStart = 0
ERemarks.SelLength = Len(ERemarks.Text)

End Sub

