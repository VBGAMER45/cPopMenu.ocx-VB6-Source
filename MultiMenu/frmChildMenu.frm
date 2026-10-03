VERSION 5.00
Begin VB.Form frmChildMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   5205
   ClientTop       =   2565
   ClientWidth     =   4680
   Icon            =   "frmChildMenu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Save &As..."
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Print"
         Index           =   6
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Pre&view"
         Index           =   7
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   "P&roperties"
         Index           =   9
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
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cu&t"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Delete"
         Index           =   6
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Select All"
         Index           =   7
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Find..."
         Index           =   9
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find &Next..."
         Index           =   10
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Replace..."
         Index           =   11
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuTestTOP 
      Caption         =   "&Test"
      Begin VB.Menu mnuTest 
         Caption         =   "&Test Sub-Menu"
         Index           =   0
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Contents"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Search for Help on..."
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmChildMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InitMenu()
Dim lParent As Long
Dim i As Long

   ' This method is called by the MDI form when it
   ' detects a new menu.  This is your opportunity
   ' to tell the PopMenu to find the new menu items
   ' and to add the icons to it.
   
   With mfrmMain.ctlPopMenu
      ' Inform the pop menu control that
      ' the menu has changed:
      .SubClassMenu Me
   
      ' Set up the icons.
      .ItemIcon("mnuFile(0)") = 41
      .ItemIcon("mnuFile(1)") = 43
      .ItemIcon("mnuFile(2)") = 61
      
      .ItemIcon("mnuFile(6)") = 53
      .ItemIcon("mnuFile(7)") = 19
      
      .ItemIcon("mnuFile(9)") = 55
      
      .ItemIcon("mnuEdit(0)") = 70
      .ItemIcon("mnuEdit(3)") = 10
      .ItemIcon("mnuEdit(4)") = 9
      .ItemIcon("mnuEdit(5)") = 48
      .ItemIcon("mnuEdit(6)") = 13
      
      .ItemIcon("mnuEdit(9)") = 3
      .ItemIcon("mnuEdit(10)") = 17
      
      .ItemIcon("mnuHelp(0)") = 26
      
      ' Add a new sub menu for testing
      ' purposes:
      lParent = .MenuIndex("mnuTest(0)")
      For i = 1 To 10
         .AddItem "Test" & i, "NewMenu" & i, , , lParent, Rnd * mfrmMain.ilsIcons.ListImages.Count
      Next i
      
   End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' If you create any new popup menus in this
   ' form, you must call this method to clear
   ' them up.  Otherwise there will be a memory
   ' leak (until the app is closed) and it is possible
   ' but unlikely you will ultimately run out of
   ' menu handles.
   mfrmMain.ctlPopMenu.UnsubclassMenu
End Sub

Private Sub mnuFile_Click(Index As Integer)
   
   Select Case Index
   Case 0
      ' New
      mfrmMain.NewWindow
   Case 4
      ' close
      Unload Me
   Case 11
      mfrmMain.EndApp
   End Select
   
End Sub
