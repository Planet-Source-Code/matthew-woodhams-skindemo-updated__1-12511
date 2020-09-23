Attribute VB_Name = "SkinModule"
' For Move form
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Sub OpenFile(filename As String)
On Error Resume Next

Dim no As Integer
Dim i As Integer

no = FreeFile
Open filename For Input As #no
Do
Input #no, a
Input #no, b
Input #no, c
Input #no, d
Input #no, e
Input #no, f
Input #no, g
Input #no, h
Input #no, i
Input #no, j
Input #no, k
Input #no, l
Input #no, m
Input #no, n
Input #no, o


If b = "[Skin]" Then
frmMain.txtName = c
frmMain.txtMainImage = d
frmMain.txtMaskImage = e
' Open size settings
frmMain.txtHeight = f
frmMain.txtWidth = g
' Open controls settings
frmMain.txtTopT = h
frmMain.txtLeftT = i
frmMain.txtTopE = j
frmMain.txtLeftE = k
frmMain.txtTopM = l
frmMain.txtLeftM = m
' NEW! This will look to see if the skin wants the control's captions to be shown or not.
If n = "1" Then
frmMain.NoCap.Value = True
Else
frmMain.NoCap.Value = False
End If
' Open skin comment
frmMain.txtComments = o

End If



Loop Until EOF(no)
Close #no


End Sub


Sub SaveFile(filename As String)
On Error Resume Next ' In case of error
Dim no As Integer
Dim i As Integer
no = FreeFile

Open filename For Output As #no 'Open the file

'Save title
Print #no, "SkinDemo Version 1.0 By Matthew Woodhams"
Print #no, "[Skin]"
'Save general settings
Print #no, frmMain.txtName.Text
Print #no, frmMain.txtMainImage.Text
Print #no, frmMain.txtMaskImage.Text
' Save for size
Print #no, frmMain.txtHeight.Text
Print #no, frmMain.txtWidth.Text
' Save control top/left
Print #no, frmMain.txtTopT.Text
Print #no, frmMain.txtLeftT.Text
Print #no, frmMain.txtTopE.Text
Print #no, frmMain.txtLeftE.Text
Print #no, frmMain.txtTopM.Text
Print #no, frmMain.txtLeftM.Text
' NEW! Saves if the skin has the control caption showing or not.
If frmMain.NoCap = True Then
Print #no, "1"
Else
Print #no, "0"
End If
' Save comment
Print #no, frmMain.txtComments.Text
Close #no

End Sub




Sub OpenSkin()
' This is used for frmPreview...
'On Error Resume Next ' In case of error


'Load images
frmPreview.Picture = LoadPicture(App.Path + "\" + frmMain.txtName + "\" + frmMain.txtMainImage) ' This will load the picture from the path you choose in the skin editor
frmPreview.Mask.Picture = LoadPicture(App.Path + "\" + frmMain.txtName + "\" + frmMain.txtMaskImage) ' This will load the mask from the path you choose in the skin editor
' Open size settings
frmPreview.Height = frmMain.txtHeight
frmPreview.MoveForm.Height = frmPreview.Height ' the move label
frmPreview.Width = frmMain.txtWidth
frmPreview.MoveForm.Width = frmPreview.Width ' the move label
'Open control settings
frmPreview.Title.top = frmMain.txtTopT
frmPreview.Title.left = frmMain.txtLeftT
frmPreview.Exit.top = frmMain.txtTopE
frmPreview.Exit.left = frmMain.txtLeftE
frmPreview.Minimize.top = frmMain.txtTopM
frmPreview.Minimize.left = frmMain.txtLeftM
If frmMain.NoCap = "1" Then
frmPreview.Exit.Caption = ""
frmPreview.Minimize.Caption = ""
Else
frmPreview.Exit.Caption = "X"
frmPreview.Minimize.Caption = "_"
End If


Call frmPreview.ChangeMask ' This is for updating the mask

End Sub



