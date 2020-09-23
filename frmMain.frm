VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "SkinDemo"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Minimize 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Minimize"
      Top             =   120
      Width           =   150
   End
   Begin VB.Label Exit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   165
      Width           =   150
   End
   Begin VB.Label MoveForm 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Right click for popup menu"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Title 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Skin Demo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Right click the form"
      Top             =   135
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit


'*****************************************************
'* Updates and fixes 11/11/00:                       *
'*  I have updated the example a bit. I updated the  *
'* move for (thanks to Joey Burgett).  I fixed the   *
'* code in the OpenSkin sub, now if the user is a    *
'* first time user it will load the default skin!    *
'* Because of a request i also made it so that       *
'* the form can be in the Windows task bar!          *
'* I also updated the SkinBuilder, now you can make  *
'* the exit and minimize buttons show or hide its    *
'* captions and i added a new feature that allows you*
'* to preview the skin in a form so that you know    *
'* what the skin will look like (this is why the     *
'* skinbuilder source is now in the "skin" directory)*
'* And i also made 4 new Skins!!                     *
'*****************************************************



'*****************************************************
'*                   Skin Demo!                      *
'*                                                   *
'*  Hi, this is an example of what i am using to skin*
'* a project of mine. I have changed it a bit, but it*
'* has most of the stuff my project has. Most of it  *
'* is commented but if you need any help contact me. *
'* Special thanks to Dos, because i used some codes  *
'* of his to make the form any shape.                *
'* Thanks a lot, hope this helps. please vote!       *
'* Contact me if you have any trouble:               *
'*                                                   *
'* Email: Squash@cv.cl                               *
'* web site:  http://www.SquashProductions.com       *
'*****************************************************

' Please visit http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=12511 to vote


' Declare Functions for RegionFromBitmap and to move the form.
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long
Public Prg As String, Sect As String ' for savesettings
Public skindir As String


Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
'***************************************
'* This part of the codes i got from an*
'* example (dos-shape) by dos. it lets *
'* your form me any shape.             *
'* Thanks!                             *
'* web site:  http://www.hider.com/dos *
'***************************************

  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hDC, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
End Function

Sub ChangeMask()
On Error Resume Next ' In case of error
' This is also part of Dos's Dos-Shape example. To update if the skin is changed
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(Mask)
  lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True)
End Sub

Private Sub Form_Load()
  On Error Resume Next ' In case of error
 Prg = "SkinDemo": Sect = "config" ' This is used for saving to registry
  skindir = (GetSetting(Prg, Sect, "Skin", skindir)) 'gets the skin from the registry. if its a first time load, it will resume next.

 skindir = (GetSetting(Prg, Sect, "Skin", skindir)) ' This gets the settings from the registry to see to load the previous skin
 OpenSkin skindir  ' opens previous skin.

frmMirror.Show
MoveForm.Top = 0 ' Put lable top most
MoveForm.Left = 0 ' Put lable at the left
MoveForm.BackStyle = 0  ' Makes label transparent
Exit Sub


End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next ' In case of error
SaveSetting Prg, Sect, "Skin", skindir ' This will save the current skin to the registry
End
End Sub

Private Sub Exit_Click()
frmMirror.Timer1.Enabled = False ' Make frmMirror's timmer not enabled so that the program can end.
Unload Me ' When you click on the exit button
End Sub

Private Sub Minimize_Click()
'***************************************
'*               UPDATE!               *
'* Now your project will be in the     *
'* windows task bar.                   *
'***************************************

On Error Resume Next ' In case of error
frmMirror.Timer1.Enabled = True
frmMirror.WindowState = 1 ' This will minimize the vb project

' If you want to minimize like before delete the above and add:
' Me.WindowState = 1
   End Sub

Private Sub MoveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************
'*               UPDATE!               *
'* Very little code update, thanks to  *
'* Joey Burgett                        *
'* ohiojedi@hotmail.com                *
'***************************************
' This will allow you to move your form without using the caption bar and to popup a menu.
 If Button = 1 Then ' Left button
    ReleaseCapture
    Call SendMessage(Me.hWnd, &HA1, 2, 0)
End If

If Button = 2 Then ' Right button
       PopupMenu frmMenu.mnuFile, 0
End If
End Sub


'***************************************
'* 'Enjoy, Please vote!                *
'***************************************

' Please visit http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=12511 to vote
