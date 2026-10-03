VERSION 5.00
Begin VB.Form frmMDIChild 
   Caption         =   "Child Form"
   ClientHeight    =   3180
   ClientLeft      =   4710
   ClientTop       =   2685
   ClientWidth     =   4260
   Icon            =   "fChild.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   4260
   Begin VB.TextBox txtDetails 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4155
   End
End
Attribute VB_Name = "frmMDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = Me.Caption & Forms.Count
    txtDetails.Text = "vbAccelerator PopMenu control." & vbCrLf & vbCrLf & "This form demonstrates the control in an MDI form."
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With txtDetails
        .Move .left, .tOp, Me.ScaleWidth - .left * 2, Me.ScaleHeight - .tOp * 2
    End With
End Sub

