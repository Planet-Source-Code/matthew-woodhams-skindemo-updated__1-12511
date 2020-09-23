VERSION 5.00
Begin VB.Form frmPreview 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "SkinDemo"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1725
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmPreview.frx":0000
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
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'***************************************
'* NEW!                                *
'* frmPreview                          *
'* This is so that you can preview the *
'* skin you are making. it is all very *
'* similar to the main program's form  *
'* Enjoy!                              *
'***************************************


' Declare Functions for RegionFromBitmap and to move the form.
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long


Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
'***************************************
'* This part of the codes i got from an*
'* example (dos-shape) by dos.         *
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
On Error GoTo noskin ' In case there is an error
    Call ChangeMask
     MoveForm.top = 0 ' Put lable top most
MoveForm.left = 0 ' Put lable at the left
MoveForm.BackStyle = 0  ' Makes label transparent

Exit Sub

noskin:
MsgBox "An error has occured while loading skin", vbCritical, "Error"
Unload Me ' unload frmpreview
Exit Sub
End Sub


Private Sub MoveForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' This will allow you to move your form without using the caption bar and to popup a menu.
 If Button = 1 Then ' Left button
    ReleaseCapture
    Call SendMessage(Me.hWnd, &HA1, 2, 0)
End If

If Button = 2 Then ' Right button
       PopupMenu frmMenu.mnuFile, 0
End If
End Sub

'Messages for the control buttons

Private Sub Minimize_Click()
MsgBox "Minimize button cliked", vbExclamation, "Button clicked"
End Sub

Private Sub Exit_Click()
MsgBox "Exit button clicked", vbExclamation, "Button clicked"
End Sub


