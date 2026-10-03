VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.4#0"; "cPopMenu.ocx"
Begin VB.Form frmBitmaps 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   6285
   ClientTop       =   6585
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5790
   Begin VB.PictureBox picVBAccel 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      Picture         =   "fBitmapTest.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin cPopMenu.PopMenu ctlMenu 
      Left            =   1500
      Top             =   2160
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin ComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4185
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      Caption         =   $"fBitmapTest.frx":0559
      Height          =   1035
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lblVBAccel 
      Caption         =   "Visit vbAccelerator - the VB Programmer's Resource - at http://vbaccelerator.com"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin ComctlLib.ImageList ilsBitmaps 
      Left            =   900
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   13
      MaskColor       =   8421376
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":061D
            Key             =   "mnuFile(1)"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":0717
            Key             =   "mnuEdit(1)"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":0811
            Key             =   "mnuEdit(0)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":090B
            Key             =   "mnuFile(0)"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":0A05
            Key             =   "mnuEdit(2)"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fBitmapTest.frx":0AFF
            Key             =   "mnuFile(3)"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   3
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   0
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmBitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
    sbrMain.Panels(1).Text = "Highlighted " & ctlMenu.Caption(ItemNumber)
    sbrMain.Panels(2).Text = "Enabled:" & bEnabled
    sbrMain.Panels(3).Text = "Separator:" & bSeparator
End Sub

Private Sub Form_Load()
Dim i As Long
    ctlMenu.ImageList = ilsBitmaps
    ctlMenu.SubClassMenu Me
    
    For i = 1 To ctlMenu.Count
        If InStr(ctlMenu.MenuKey(i), "TOP") = 0 Then
            If (ctlMenu.Caption(i) <> "-") Then
                ctlMenu.ItemIcon(i) = ilsBitmaps.ListImages(ctlMenu.MenuKey(i)).Index - 1
            End If
        End If
    Next i
    sbrMain.Panels(1).Text = "Test"
    
End Sub

