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
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
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
'* form. This if for frmPreview        *
'***************************************

Private Sub mnuClose_Click()
Unload frmPreview ' Close frmPreview
End Sub
