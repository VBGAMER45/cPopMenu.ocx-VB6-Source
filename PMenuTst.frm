VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{A22D979F-2684-11D2-8E21-10B404C10000}#1.6#0"; "cPopMenu.ocx"
Begin VB.Form frmTest 
   Caption         =   "PopMenu Control Demonstration"
   ClientHeight    =   6555
   ClientLeft      =   4335
   ClientTop       =   2400
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "PMenuTst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   8340
   Begin cPopMenu.PopMenu ctlPopMenu 
      Left            =   7620
      Top             =   5040
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin ComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   36
      Top             =   6240
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14182
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSpecialEffects 
      Caption         =   "Special Effects/Styles"
      Height          =   915
      Left            =   60
      TabIndex        =   29
      Top             =   4800
      Width           =   2595
      Begin VB.CheckBox chkOfficeXpStyle 
         Caption         =   "&Office XP Style"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkStyle 
         Caption         =   "Button &Select Style"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   420
         Width           =   2415
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "&Background Bitmap"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   180
         Width           =   2295
      End
   End
   Begin VB.Frame fraFindItems 
      Caption         =   "Find Menu Items in code:"
      Height          =   2535
      Left            =   2760
      TabIndex        =   26
      Top             =   3180
      Width           =   2655
      Begin VB.CommandButton cmdFindHierarchy 
         Caption         =   "Find by Hierarch&y"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   900
         Width           =   1155
      End
      Begin VB.CommandButton cmdFindKey 
         Caption         =   "Find by &Key"
         Height          =   555
         Left            =   120
         TabIndex        =   31
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdGet 
         Caption         =   "&Get Hierarchy for Index..."
         Height          =   555
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1155
      End
   End
   Begin VB.Frame fraMore 
      Caption         =   "More Demonstration Forms:"
      Height          =   1575
      Left            =   2760
      TabIndex        =   23
      Top             =   1560
      Width           =   2595
      Begin VB.CommandButton cmdMDIDemo 
         Caption         =   "&MDI Demo..."
         Height          =   555
         Left            =   180
         TabIndex        =   25
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   "Another Form..."
         Height          =   555
         Left            =   180
         TabIndex        =   24
         Top             =   840
         Width           =   1155
      End
   End
   Begin VB.Frame fraPopup 
      Caption         =   "&Show Popup Menus"
      Height          =   1515
      Left            =   2760
      TabIndex        =   20
      Top             =   0
      Width           =   2595
      Begin VB.CommandButton cmdVBPopup 
         Caption         =   "VB Popup:"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdAPIPopup 
         Caption         =   "API Popup:"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.Frame fraAddRemove 
      Caption         =   "Add/Remove to Edit Menu"
      Height          =   2115
      Left            =   60
      TabIndex        =   16
      Top             =   2640
      Width           =   2595
      Begin VB.CommandButton cmdVisible 
         Caption         =   "Make File Item Visible..."
         Height          =   555
         Left            =   1380
         TabIndex        =   35
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add to Edit Menu"
         Height          =   555
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Delete from Edit Menu"
         Enabled         =   0   'False
         Height          =   555
         Left            =   180
         TabIndex        =   18
         Top             =   900
         Width           =   1155
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert to Edit Menu"
         Height          =   555
         Left            =   180
         TabIndex        =   17
         Top             =   1500
         Width           =   1155
      End
   End
   Begin VB.Frame fraManipulate 
      Caption         =   "Manipulate Menu Items"
      Height          =   2595
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   2595
      Begin VB.CheckBox chkDefault 
         Caption         =   "Make Search Default (Bold)"
         Height          =   255
         Left            =   60
         TabIndex        =   33
         Top             =   480
         Width           =   2475
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         Height          =   555
         Left            =   1560
         ScaleHeight     =   495
         ScaleWidth      =   555
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdChangeIcon 
         Caption         =   "Change &Paste Icon"
         Height          =   555
         Left            =   300
         TabIndex        =   13
         Top             =   1920
         Width           =   1155
      End
      Begin VB.CommandButton cmdChangeCaption 
         Caption         =   "Change &Paste Caption"
         Height          =   555
         Left            =   300
         TabIndex        =   12
         Top             =   1320
         Width           =   1155
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Paste Enabled"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   1395
      End
      Begin VB.CheckBox chkNewest 
         Caption         =   "Check &Newest"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   720
         Width           =   1515
      End
      Begin VB.CheckBox chkENewest 
         Caption         =   "Enable Newest"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.Label lblCaption 
         Caption         =   "&Paste"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   1500
         Width           =   915
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information:"
      Height          =   4935
      Left            =   5460
      TabIndex        =   3
      Top             =   0
      Width           =   2835
      Begin VB.Image imgBlt 
         Height          =   480
         Index           =   3
         Left            =   60
         Picture         =   "PMenuTst.frx":030A
         Top             =   2940
         Width           =   480
      End
      Begin VB.Image imgBlt 
         Height          =   480
         Index           =   2
         Left            =   120
         Picture         =   "PMenuTst.frx":0614
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image imgBlt 
         Height          =   480
         Index           =   1
         Left            =   60
         Picture         =   "PMenuTst.frx":091E
         Top             =   1140
         Width           =   480
      End
      Begin VB.Image imgBlt 
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "PMenuTst.frx":0C28
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Caption         =   "The File, Edit and Help menu items are VB menus, whose events are responded to by the standard VB handlers."
         Height          =   795
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Caption         =   "The top menus of the other items were created in VB, with sub-items added on the fly by the pop menu control."
         Height          =   795
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Caption         =   "Directory 1 is created in code during form load.  Directory 2 is populated on the fly using the InitPopupMenu feature."
         Height          =   795
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblInfo 
         Caption         =   "The System Menu of this form has also been modified.  Close has been removed, and a new About item added."
         Height          =   855
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   2940
         Width           =   2415
      End
   End
   Begin VB.PictureBox picVBAccel 
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   60
      Picture         =   "PMenuTst.frx":0F32
      ScaleHeight     =   330
      ScaleWidth      =   1275
      TabIndex        =   1
      ToolTipText     =   "Free, Advanced source code for VB Programmers at http://vbaccelerator.com"
      Top             =   5820
      Width           =   1335
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Close"
      Height          =   435
      Left            =   5460
      TabIndex        =   0
      Top             =   5040
      Width           =   1155
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   4620
      Picture         =   "PMenuTst.frx":148B
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   28
      Top             =   3900
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label lblVBAccel 
      Caption         =   "Visit vbAccelerator - free, advanced source code for VB Programmers - at http://vbaccelerator.com"
      Height          =   405
      Left            =   1440
      TabIndex        =   2
      Top             =   5820
      Width           =   3915
   End
   Begin ComctlLib.ImageList ilsIcons 
      Left            =   6960
      Top             =   4980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":D4CD
            Key             =   "NULL"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":D7E7
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":DB01
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":DE1B
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":E135
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":E44F
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":E769
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":EA83
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":ED9D
            Key             =   "BACK"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":F0B7
            Key             =   "NEXT"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":F3D1
            Key             =   "FAVE"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":F6EB
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":FA05
            Key             =   "NET"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":FD1F
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":10039
            Key             =   "DOCUMENT"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PMenuTst.frx":10353
            Key             =   "TICK"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuF0Main 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Se&tup"
         Index           =   4
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &1"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &2"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &3"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Test Invisible &4"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   11
      End
   End
   Begin VB.Menu mnuE0MAIN 
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
         Index           =   2
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Search..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "&In Code"
      Begin VB.Menu mnuSub 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuD1Main 
      Caption         =   "&Directory1"
   End
   Begin VB.Menu mnuD2Main 
      Caption         =   "Direc&tory2"
      Begin VB.Menu mnuDir2 
         Caption         =   "<none>"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuH0MAIN 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&On the Internet..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private m_lAboutId As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Sub pAddMultiColumnMenu()
Dim lParent As Long
Dim i As Long, iPerRow As Long
Dim s As String

   With ctlPopMenu
      lParent = .MenuIndex("mnuMultiColumn")
      .ClearSubMenusOfItem lParent
      iPerRow = ilsIcons.ListImages.Count \ 4
      For i = 2 To ilsIcons.ListImages.Count
         If ((i - 2) Mod iPerRow) = 0 Then
            s = "^" '& StrConv(ilsIcons.ListImages(i).Key, vbProperCase)
         Else
            s = "" 'StrConv(ilsIcons.ListImages(i).Key, vbProperCase)
         End If
         .AddItem s, "MultiColumn" & i, , , lParent, i - 1
      Next i
   End With
End Sub

Private Sub pCreateMenuItems()
Dim lParentIndex As Long
Dim lIndex As Long
Dim lThisIndex As Long
Dim sPath As String

    With ctlPopMenu
        
        ' Set up the existing item:
        lIndex = .MenuIndex("mnuSub(0)")
        .Caption(lIndex) = "&Back"
        .HelpText(lIndex) = "Move to the previous page"
        .ItemIcon(lIndex) = plGetIconIndex("BACK")
            .AddItem "Test", "mnuNewTest(1)", , , lIndex
            .AddItem "Test2", "mnuNewTest(2)", , , lIndex
        
        ' Now add new ones to the In Code item:
        lParentIndex = .MenuIndex("mnuPop")
        lIndex = .AddItem("&Next", "mnuSub(1)", "Move to the next page", , lParentIndex, plGetIconIndex("NEXT"), True)
        .AddItem "-", "mnuSub(2)", , , lParentIndex
        .AddItem "Most &Viewed", "mnuSub(3)", "Go to the most viewed choices", , lParentIndex, plGetIconIndex("FAVE")
        .AddItem "Ne&west", "mnuSub(4)", "See the newest available items", , lParentIndex
        .AddItem "-", "mnuSub(5)", , , lParentIndex
        lIndex = .AddItem("&Test", "mnuSub(6)", "Testing sub menu items", , lParentIndex, 1)
            ' Add some sub items to the test:
            .AddItem "Sub Item &A", "mnuSubSub(0)", "Test sub item A", , lIndex, 2
            lIndex = .AddItem("Sub Item &B", "mnuSubSub(1)", "Test sub item B", , lIndex, 3)
                ' Add some sub-sub items to Sub Item &B:
                lThisIndex = .AddItem("Sub-sub Item B &1", "mnuSubSubSub(0)", "Sub-sub Item B 1", , lIndex, 1)
                .MenuDefault(lThisIndex) = True
                .AddItem "Sub-sub Item B &2", "mnuSubSubSub(1)", "Sub-sub Item B 2", , lIndex, 4
                
        ' Add some new ones to the Directory item:
        sPath = App.Path
        lParentIndex = .MenuIndex("mnuD1Main")
        pDirectoryAddItems lParentIndex, sPath
        
    End With
End Sub

Private Sub pDirectoryAddItems( _
        ByVal lParentIndex As Long, _
        ByVal sPath As String, _
        Optional ByVal bTop As Boolean = False _
    )
Dim sFiles() As String
Dim bDir() As Boolean
Dim iFileCOunt As Long
Dim iFile As Long
Dim lIndex As Long
Dim lFolderIcon As Long
Dim lDocIcon As Long
Dim sCaption As String
Dim iCount As Long
Static iTotal As Long

   If (bTop) Then
      iTotal = 0
   End If

   GetFilesInPath sPath, sFiles(), bDir(), iFileCOunt
   lFolderIcon = plGetIconIndex("FOLDER")
   lDocIcon = plGetIconIndex("DOCUMENT")
   For iFile = 1 To iFileCOunt
      sCaption = sFiles(iFile)
      iCount = iCount + 1
      iTotal = iTotal + 1
      If (iTotal > ctlPopMenu.MenuItemsPerScreen * 3) Then
         ' don't add too many...
         Exit Sub
      End If
      If (iCount > ctlPopMenu.MenuItemsPerScreen \ 2) Then
         sCaption = "|" & sCaption
         iCount = 0
      End If
      If (bDir(iFile)) Then
         lIndex = ctlPopMenu.AddItem(sCaption, , , , lParentIndex, lFolderIcon)
      Else
         lIndex = ctlPopMenu.AddItem(sCaption, , , , lParentIndex, lDocIcon)
      End If
      If (bDir(iFile)) Then
         pDirectoryAddItems lIndex, sPath & "\" & sFiles(iFile)
      End If
   Next iFile
   If (iFileCOunt = 0) Then
      ctlPopMenu.AddItem "<empty>", , , , lParentIndex, , , False
   End If
   
End Sub
Private Sub GetFilesInPath( _
        ByVal sPath As String, _
        ByRef sFiles() As String, _
        ByRef bDir() As Boolean, _
        ByRef iFileCOunt As Long _
    )
Dim sDir As String
Dim bAdd As Boolean
Dim bIsDir As Boolean

    iFileCOunt = 0
    sDir = Dir(sPath & "\*.*", vbNormal Or vbDirectory)
    Do While Len(sDir) > 0
        If (sDir <> ".") And (sDir <> "..") Then
            bIsDir = ((GetAttr(sPath & "\" & sDir) And vbDirectory) = vbDirectory)
            bAdd = False
            If Not (bIsDir) Then
               bAdd = True
            Else
               bAdd = True
            End If
            If (bAdd) Then
                iFileCOunt = iFileCOunt + 1
                ReDim Preserve sFiles(1 To iFileCOunt) As String
                ReDim Preserve bDir(1 To iFileCOunt) As Boolean
                sFiles(iFileCOunt) = sDir
                bDir(iFileCOunt) = bIsDir
                If (iFileCOunt > ctlPopMenu.MenuItemsPerScreen * 3) Then
                  ' stop - too many...
                  Exit Do
               End If
            End If
        End If
        sDir = Dir
    Loop
End Sub
    
Private Sub chkBackground_Click()
   If (chkBackground.Value = Checked) Then
      Set ctlPopMenu.BackgroundPicture = picBackground.Picture
   Else
      ctlPopMenu.ClearBackgroundPicture
   End If
End Sub

Private Sub chkDefault_Click()
   ctlPopMenu.MenuDefault("mnuEdit(4)") = chkDefault.Value * -1
End Sub

Private Sub chkEnable_Click()
   ctlPopMenu.Enabled("mnuEdit(2)") = chkEnable.Value * -1
End Sub

Private Sub chkENewest_Click()
   ctlPopMenu.Enabled("mnuSub(4)") = chkENewest.Value * -1
End Sub

Private Sub chkNewest_Click()
   ctlPopMenu.Checked("mnuSub(4)") = chkNewest.Value * -1
End Sub

Private Sub chkOfficeXpStyle_Click()
   ctlPopMenu.OfficeXpStyle = (chkOfficeXpStyle.Value = Checked)
End Sub

Private Sub chkStyle_Click()
   If chkStyle.Value = Checked Then
      ctlPopMenu.HighlightStyle = cspHighlightButton
   Else
      ctlPopMenu.HighlightStyle = cspHighlightStandard
   End If
End Sub

Private Sub cmdAdd_Click()
Dim lParent As Long
Dim lId As Long

   With ctlPopMenu
      If Not (.MenuExists("NewItem1")) Then
         ' We don't have this menu:
         lParent = .MenuIndex("mnuE0MAIN")
         .AddItem "-", "NewItem0", , , lParent
         .AddItem "Test Item 1", "NewItem1", , , lParent, 10
         lId = .AddItem("Test Item 2", "NewItem2", , , lParent, 11)
         .AddItem "Test Sub Item 2,1", "NewItem3", , , lId, 1
         .AddItem "Test Sub Item 2,2", "NewItem4", , , lId, 2
         .AddItem "Test Sub Item 2,3", "NewItem5", , , lId, 3
         .AddItem "Test Item 3", "NewItem6", , , lParent, 12
         
         Debug.Print "Add:AfterCount:" & .Count
      Else
         MsgBox "Menu items are already added.", vbInformation
      End If
   End With
   cmdRemove.Enabled = True
   cmdAdd.Enabled = False
End Sub


Private Sub cmdChangeCaption_Click()
   If (ctlPopMenu.Caption("mnuEdit(2)") = "&Paste") Then
      ctlPopMenu.Caption("mnuEdit(2)") = "Replacement Caption for &Paste"
      'ctlPopMenu.ReplaceItem "mnuEdit(2)", "Replacement Caption for &Paste"
   Else
      ctlPopMenu.Caption("mnuEdit(2)") = "&Paste"
      'ctlPopMenu.ReplaceItem "mnuEdit(2)", "&Paste"
   End If
   lblCaption.Caption = ctlPopMenu.Caption("mnuEdit(2)")
End Sub



Private Sub cmdFindHierarchy_Click()
Dim lH() As Long
Dim lCount As Long
Dim i As Long
Dim sIndex As String
Dim sI As String
Dim lIndex As Long

   sIndex = InputBox("Enter hierarchy to find item for: ", , "2,5")
   If (sIndex <> "") Then
      For i = 1 To Len(sIndex)
         If (Mid$(sIndex, i, 1) = ",") Then
            lCount = lCount + 1
            ReDim Preserve lH(1 To lCount) As Long
            lH(lCount) = CLng(sI)
            sI = ""
         Else
            sI = sI & Mid$(sIndex, i, 1)
         End If
      Next i
      If (sI <> "") Then
         lCount = lCount + 1
         ReDim Preserve lH(1 To lCount) As Long
         lH(lCount) = CLng(sI)
      End If
      
      lIndex = ctlPopMenu.IndexForMenuHierarchy(lH())
      If (lIndex > 0) Then
         MsgBox "Found at index " & lIndex & vbCrLf & "Caption: " & ctlPopMenu.Caption(lIndex) & vbCrLf & "Icon Index: " & ctlPopMenu.ItemIcon(lIndex), vbInformation
      End If
   End If
   Exit Sub
   
ErrorHandler:
   MsgBox "Couldn't interpret " & sIndex, vbInformation
   Exit Sub
End Sub

Private Sub cmdFindKey_Click()
Dim sI As String
Dim lIndex As Long
    sI = InputBox("Enter the key you wish to find", , "mnuEdit(2)")
    If (sI <> "") Then
        On Error Resume Next
        lIndex = ctlPopMenu.MenuIndex(sI)
        If (Err.Number <> 0) Then
            lIndex = -1
        End If
        If (lIndex > -1) Then
            MsgBox "Found at index " & lIndex & vbCrLf & "Caption: " & ctlPopMenu.Caption(lIndex) & vbCrLf & "Icon Index: " & ctlPopMenu.ItemIcon(lIndex), vbInformation
        Else
            MsgBox "No item with key '" & sI & "' was found.", vbInformation
        End If
    End If
End Sub

Private Sub cmdInsert_Click()
Static l As Long
   With ctlPopMenu
      .InsertItem "This has just been inserted", "mnuEdit(1)", "InsertItem" & l, , , Rnd * ilsIcons.ListImages.Count
      l = l + 1
   End With
End Sub

Private Sub cmdMDIDemo_Click()
    mfrmMDITest.Show
End Sub

Private Sub cmdMore_Click()
Dim fB As frmBitmaps
    Set fB = New frmBitmaps
    fB.Show
End Sub

Private Sub cmdRemove_Click()
Dim iItem As Long
   With ctlPopMenu
      Debug.Print "Remove:BeforeCount:" & .Count
      ' NOTE: here we loop backwards because the menu
      ' item with key "NewItem2" has a sub menu.  When you delete
      ' "NewItem2" this automatically deletes the subitems (i.e.
      ' "NewItem3","NewItem4" and "NewItem5").  So if you delete
      ' "NewItem2" then "NewItem3" thru "NewItem5" no longer exist
      ' and you get a subscript out of range error.
      ' The alternative is to only delete the items with keys "NewItem0","NewItem1",
      ' "NewItem2" and "NewItem6"
      For iItem = 6 To 0 Step -1
          .RemoveItem "NewItem" & iItem
      Next iItem
      Debug.Print "Remove:AfterCount:" & .Count
   End With
   cmdRemove.Enabled = False
   cmdAdd.Enabled = True
End Sub

Private Sub cmdUnload_Click()
    mnuFile_Click 11
End Sub

Private Sub cmdVBPopup_Click()
    Me.PopupMenu mnuPop, , cmdVBPopup.Container.left + cmdVBPopup.left, cmdVBPopup.Container.tOp + cmdVBPopup.tOp + cmdVBPopup.Height
End Sub

Private Sub cmdGet_Click()
Dim lH() As Long
Dim lR As Long
Dim l As Long
Dim sOut As String

   ReDim lH(1 To 4) As Long
   lH(1) = 3
   lH(2) = 7
   lH(3) = 2
   lH(4) = 1
   lR = ctlPopMenu.IndexForMenuHierarchy(lH())
   If (lR > 0) Then
      For l = 1 To 4
         sOut = sOut & lH(l) & ","
      Next l
      MsgBox "Index for item at hierarchy position: " & vbCrLf & left$(sOut, Len(sOut) - 1) & vbCrLf & lR & " (" & ctlPopMenu.Caption(lR) & ")", vbInformation
   Else
      MsgBox "Index not found", vbExclamation
   End If
End Sub

Private Sub cmdAPIPopup_Click()
Dim lR As Long

   With ctlPopMenu
      ' Track popup menu now built into the control:
      lR = ctlPopMenu.ShowPopupMenu(cmdAPIPopup, "mnuSubSub(1)", 0, cmdAPIPopup.Height)
        
      ' How to do it with the API:
      'Dim lIndex As Long
      'Dim hMenu As Long
      'Dim tR As RECT
      'Dim tP As POINTAPI

      'lIndex = .MenuIndex("mnuSubSub(1)")
      'If (lIndex > 0) Then
      '    hMenu = .hPopupMenu(lIndex)
      '    If (hMenu > 0) Then
      '        tP.X = (cmdAPIPopup.Container.left + cmdAPIPopup.left) \ Screen.TwipsPerPixelX
      '        tP.Y = (cmdAPIPopup.Container.tOp + cmdAPIPopup.tOp + cmdAPIPopup.Height) \ Screen.TwipsPerPixelY
      '        ClientToScreen Me.hwnd, tP
      '        lR = TrackPopupMenu(hMenu, 0, tP.X, tP.Y, 0, Me.hwnd, tR)
      '    End If
      'End If
    End With
End Sub

Private Sub cmdChangeIcon_Click()
   If ctlPopMenu.ItemIcon("mnuEdit(2)") = 6 Then
      ctlPopMenu.ItemIcon("mnuEdit(2)") = Rnd * ilsIcons.ListImages.Count
   Else
      ctlPopMenu.ItemIcon("mnuEdit(2)") = 6
   End If
   
   picIcon.Cls
   ilsIcons.ListImages(ctlPopMenu.ItemIcon("mnuEdit(2)") + 1).Draw picIcon.hDC, 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, imlTransparent
   picIcon.Refresh

End Sub

Private Sub cmdVisible_Click()
Static i As Long

   ' Make one of the invisible menu items
   ' visible again:
   If (i = 0) Then
      i = 5
   Else
      i = i + 1
   End If
   
   If (i = 9) Then
      cmdVisible.Enabled = False
   End If
      
   ' Make menu item visible:
   mnuFile(i).Visible = True
   ' Add it to the control.  When the control finds
   ' the menu you have just made Visible, it will
   ' raise a RequestNewMenuItem event, passing the
   ' caption.  You have to match this up to set
   ' the key property etc.
   ctlPopMenu.CheckForNewItems
   
End Sub

Private Sub ctlPopMenu_Click(ItemNumber As Long)
    Debug.Print "Clicked " & ItemNumber
    sbrMain.Panels(1).Text = "Clicked: " & ctlPopMenu.Caption(ItemNumber)
End Sub

Private Sub ctlPopMenu_InitPopupMenu(ParentItemNumber As Long)
Dim lIndex As Long
Dim sPath As String
Dim sFiles() As String
Dim bDir() As Boolean
Dim sPrefix As String
Dim iCount As Long
Dim iFile As Integer
Dim lParent As Long
Dim lTop As Long
Dim iNumber As Long
Static bPopulated As Boolean

   With ctlPopMenu
      If (.MenuKey(ParentItemNumber) = "mnuMultiColumn") Then
         If Not (bPopulated) Then
            pAddMultiColumnMenu
            bPopulated = True
         End If
      Else
         lTop = .UltimateParent(ParentItemNumber)
         If (.MenuKey(lTop) = "mnuD2Main") Then
            'Debug.Print
            'Debug.Print .Caption(ParentItemNumber)
            'Debug.Print "InitPopupMenu:Start:" & .Count
            Screen.MousePointer = vbHourglass
            'Debug.Print "This is Directory 2!"
            If ParentItemNumber = lTop Then
               'Debug.Print "We are at the top"
               sPath = App.Path
            Else
               'Debug.Print "Doing a sub level"
               sPath = App.Path & "\" & .HierarchyPath(ParentItemNumber, 2, "\")
            End If
            lParent = .ClearSubMenusOfItem(ParentItemNumber)
            'Debug.Print "InitPopupMenu:AfterClear:" & .Count, ParentItemNumber
            GetFilesInPath sPath, sFiles(), bDir(), iCount
            If (iCount > 0) Then
            
               If (iCount > .MenuItemsPerScreen * 3) Then
                  MsgBox "There are too many menu items in the path '" & App.Path & "'." & vbCrLf & "Only the first " & .MenuItemsPerScreen * 3 & " items will be shown.", vbInformation
                  iCount = .MenuItemsPerScreen * 3
               End If
               
               For iFile = 1 To iCount
                  If (bDir(iFile)) Then
                     lIndex = .AddItem(sPrefix & sFiles(iFile), , , , lParent, plGetIconIndex("FOLDER"))
                     iNumber = iNumber + 1
                     If (iNumber >= .MenuItemsPerScreen \ 2) Then
                        sPrefix = "|"
                        iNumber = 0
                     Else
                        sPrefix = ""
                     End If
                     'Debug.Print lIndex, sFiles(iFile)
                     .AddItem "<none>", , , , lIndex, , , False
                  End If
               Next iFile
               
               For iFile = 1 To iCount
                  If Not (bDir(iFile)) Then
                     lIndex = .AddItem(sPrefix & sFiles(iFile), , , , lParent, plGetIconIndex("DOCUMENT"))
                     iNumber = iNumber + 1
                     If (iNumber >= .MenuItemsPerScreen \ 2) Then
                        sPrefix = "|"
                        iNumber = 0
                     Else
                        sPrefix = ""
                     End If
                     'Debug.Print lIndex, sFiles(iFile)
                  End If
               Next iFile
                
            Else
               .AddItem "<empty>", , , , ParentItemNumber, , , False
            End If
            Screen.MousePointer = vbNormal
            'Debug.Print "InitPopupMenu:End:" & .Count, ParentItemNumber
            
         End If
      End If
      
   End With
End Sub

Private Sub ctlPopMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Dim sText As String
    'Debug.Print "Highlight " & ItemNumber
    If Not (bSeparator) Then
        sText = ctlPopMenu.HelpText(ItemNumber)
        If (sText = "") Then
            sText = "Highlight " & ctlPopMenu.Caption(ItemNumber) & " (Help unavailable)"
        End If
        If (bEnabled) Then
            sbrMain.Panels(1).Text = sText
        Else
            sbrMain.Panels(1).Text = sText & " (Not available)"
        End If
    Else
        sbrMain.Panels(1).Text = ""
    End If
End Sub

Private Sub ctlPopMenu_MenuExit()
    sbrMain.Panels(1).Text = "Menu Exited."
End Sub

Private Sub ctlPopMenu_RequestNewMenuDetails(sCaption As String, sKey As String, iIcon As Long, lItemData As Long, sHelpText As String, sTag As String)
Dim iPos As Long
Dim i As Long

   ' This event is fired if the cPopMenu control
   ' detects a new menu item after the CheckForNewMenuItems
   ' method is called.  Unfortunately, the control can
   ' no longer automatically match up menu items to
   ' the VB menus, so this is your only opportunity
   ' to play with the captions.
   '
   '

   ' Only file menu items can be made invisible,
   ' so we search these to see if we get a
   ' caption match to set the correct key.
   Debug.Print "NEW MENU ITEM APPEARED: ", sCaption

   For i = 6 To 9
      If (sCaption = mnuFile(i).Caption) Then
         iPos = i
         Exit For
      End If
   Next i
   
   If (iPos > 0) Then
      sKey = "mnuFile(" & iPos & ")"
      iIcon = Rnd * ilsIcons.ListImages.Count - 1
   End If
   
End Sub

Private Sub ctlPopMenu_SystemMenuClick(ItemNumber As Long)
   
   ' This event is fired when a system menu
   ' item is clicked:
   
   Debug.Print ItemNumber, m_lAboutId
   Select Case ItemNumber
   Case SC_MOVE
      sbrMain.Panels(1).Text = "Clicked on Move the Window"
   Case SC_MINIMIZE
      sbrMain.Panels(1).Text = "Clicked on Minimise the window"
   Case SC_MAXIMIZE
      sbrMain.Panels(1).Text = "Clicked on Maximise the window"
   Case SC_CLOSE
      ' Note this is removed in the demo
      sbrMain.Panels(1).Text = "Clicked Close"
   Case SC_RESTORE
      sbrMain.Panels(1).Text = "Clicked Restore"
   Case SC_SIZE
      sbrMain.Panels(1).Text = "Clicked Size"
   Case m_lAboutId
      ' Clicked the customised about item
      ' added at run-time:
      Dim lMajor As Long, lMinor As Long, lRevision As Long
      ctlPopMenu.GetVersion lMajor, lMinor, lRevision
      MsgBox "vbAccelerator IconMenu control demonstration." & vbCrLf & "Visit vbAccelerator at http://vbaccelerator.com" & vbCrLf & "Copyright © 1998 Steve McMahon" & vbCrLf & vbCrLf & "Control Version: " & lMajor & "." & lMinor & "." & lRevision, vbInformation
   End Select
End Sub

Private Sub ctlPopMenu_SystemMenuItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
    Select Case ItemNumber
    Case SC_MOVE
        sbrMain.Panels(1).Text = "Move the Window"
    Case SC_MINIMIZE
        sbrMain.Panels(1).Text = "Minimise the window to an icon"
    Case SC_MAXIMIZE
        sbrMain.Panels(1).Text = "Maximise the window's size"
    Case SC_CLOSE
        sbrMain.Panels(1).Text = "Close this application"
    Case SC_RESTORE
        sbrMain.Panels(1).Text = "Restore the window to its previous size"
    Case SC_SIZE
        sbrMain.Panels(1).Text = "Change the size of the window"
    Case m_lAboutId
        sbrMain.Panels(1).Text = "Find out about this program"
    Case Else
        sbrMain.Panels(1).Text = ""
    End Select

End Sub

Private Sub ctlPopMenu_WinIniChange()
    
    ' This is not the place to put a message box - it will
    ' lock the system!
    Debug.Print "*******************************************"
    Debug.Print "GOT A WININICHANGE EVENT"
    Debug.Print "*******************************************"
    
    
End Sub

Private Sub Form_Load()
Dim l As Long
Dim lIndex As Long
Dim lC As Long

    With ctlPopMenu
      ' Set for Office XP Styles:
      .OfficeXpStyle = True
        
      ' Remove the close item from the system menu:
      lC = .SystemMenuCount
      .SystemMenuRemoveItem lC
      ' Add a new about item to the system menu:
      m_lAboutId = .SystemMenuAppendItem("&About...")
      
      ' Associate the image list:
      .ImageList = ilsIcons
      
      ' Parse through the VB designed menu and sub class the items:
      .SubClassMenu Me
      
      lIndex = .MenuIndex("mnuEdit(2)")
      chkEnable.Value = .Enabled(lIndex) * -1
      
      ' Add the icons:
      pSetIcon "OPEN", "mnuFile(0)"
      pSetIcon "SAVE", "mnuFile(1)"
      pSetIcon "PRINT", "mnuFile(3)"
      
      pSetIcon "CUT", "mnuEdit(0)"
      pSetIcon "COPY", "mnuEdit(1)"
      pSetIcon "PASTE", "mnuEdit(2)"
      ilsIcons.ListImages(ctlPopMenu.ItemIcon("mnuEdit(2)") + 1).Draw picIcon.hDC, 4 * Screen.TwipsPerPixelX, 0.4 * Screen.TwipsPerPixelY, imlTransparent
      picIcon.Refresh
      pSetIcon "FIND", "mnuEdit(4)"
      
      pSetIcon "HELP", "mnuHelp(0)"
      pSetIcon "NET", "mnuHelp(1)"
              
      .TickIconIndex = plGetIconIndex("TICK")
    End With
    
    ' Add a whole new set of menu items and sub items to the last
    ' menu item:
    pCreateMenuItems
        
End Sub
Private Sub pSetIcon( _
        ByVal sIconKey As String, _
        ByVal sMenuKey As String _
    )
Dim lIconIndex As Long
    lIconIndex = plGetIconIndex(sIconKey)
    ctlPopMenu.ItemIcon(sMenuKey) = lIconIndex
End Sub
Private Function plGetIconIndex( _
        ByVal sKey As String _
    ) As Long
    plGetIconIndex = ilsIcons.ListImages.Item(sKey).Index - 1
End Function

Private Sub mnuEdit_Click(Index As Integer)
    MsgBox "Visual Basic Menu Edit Fired for Index:" & Index, vbInformation
End Sub

Private Sub mnuFile_Click(Index As Integer)
    If (Index = 11) Then
        If (vbYes = MsgBox("Are you sure you want to exit?", vbYesNo Or vbQuestion)) Then
            Unload Me
        End If
    Else
        MsgBox "Visual Basic Menu File Fired for Index:" & Index, vbInformation
    End If
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    If (Index < 3) Then
        MsgBox "Visual Basic Help Menu Fired for Index:" & Index, vbInformation
    Else
        ctlPopMenu_SystemMenuClick m_lAboutId
    End If
End Sub

Private Sub picIcon_Click()
Dim i As Long
Dim lIndex As Long
   For i = 0 To Controls.Count - 1
      Debug.Print Controls(i).Name,
      If TypeOf Controls(i) Is Menu Then
         Debug.Print Controls(i).Caption;
      Else
         Debug.Print
      End If
      On Error Resume Next
      lIndex = Controls(i).Index
      If (Err.Number = 0) Then
         Debug.Print lIndex
      End If
   Next i
End Sub
