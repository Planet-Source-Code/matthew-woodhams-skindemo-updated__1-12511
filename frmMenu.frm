VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   555
   ClientLeft      =   4590
   ClientTop       =   5280
   ClientWidth     =   2505
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   2505
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSkin 
         Caption         =   "&Change Skin"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'* Used when user right clicks on the  *
'* form.                               *
'***************************************

Private Sub mnuAbout_Click()
MsgBox "By Matthew Woodhams - Squash@cv.cl", vbInformation, "Please Vote!"
End Sub

Private Sub mnuExit_Click()
Unload frmMain ' Unload frmMain so it can save settings
End ' end project
End Sub

Private Sub mnuSkin_Click()
frmOpen.Show ' Shows frmOpen to change the skin
End Sub
