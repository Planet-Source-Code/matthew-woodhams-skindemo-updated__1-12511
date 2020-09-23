VERSION 5.00
Begin VB.Form frmMirror 
   Caption         =   "SkinDemo"
   ClientHeight    =   90
   ClientLeft      =   3855
   ClientTop       =   1635
   ClientWidth     =   2265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   2265
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   1320
   End
End
Attribute VB_Name = "frmMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'*                   NEW!                     *
'* frmMirror:                                 *
'* I made this part of the example because of *
'* a request from a guy in the PSC. frmMirror *
'* is used bacause frmMain has no boarders    *
'* and because of this it cannot be shown in  *
'* the Windows Task bar. So all it does is    *
'* simulates the programs minimize and exit   *
'* features...  It still needs some work.     *
'**********************************************

Private Sub Form_Load()
Me.Left = 999999999 ' An original way to make the form not be visible on the screen :)
'You "cannot" put me.visible = false because it will not show the form in the task bar.
End Sub

Private Sub Form_Unload(Cancel As Integer)
End ' End the program
End Sub

Private Sub Timer1_Timer()
'This timmer will make sure that when frmMirror is minimized frmMain is minimized and the same with Show...
If Me.WindowState = 1 Then
frmMain.WindowState = 1 'Makes frmMain state Minimized
frmMain.Visible = False 'If frmMirror is minimized then make frmMain minimized
End If

If Me.WindowState = 0 Then 'If the form is showing
frmMain.WindowState = 0 ' Makes frmMain state Normal
frmMain.Visible = True ' then show frmMain
Timer1.Enabled = False ' So that you can move the form
End If


End Sub
