Attribute VB_Name = "mMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Public Sub Main()
   InitCommonControls
   
   Dim f As New frmTest
   frmTest.Show
   
End Sub

