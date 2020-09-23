Attribute VB_Name = "SkinModule"


Sub OpenSkin(filename As String)
On Error Resume Next ' In case of error

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

'***************************************
'*                FIXED!               *
'* Now if the user is a first time user*
'* it will load the default skin!      *
'***************************************

If a = "" Then ' If the file doesn't exist or the folder has been deleted it will load the default one.
frmMain.skindir = App.Path + "\Skins\Default\Default.skn"
OpenSkin App.Path + "\Skins\Default\Default.skn" ' Loads the default skin
Else ' else continue loading skin
If b = "[Skin]" Then
'Load images (c is the path name)
frmMain.Picture = LoadPicture(App.Path + "\Skins\" + c + "\" + d) ' This will load the picture from the path you choose in the skin editor
frmMain.Mask.Picture = LoadPicture(App.Path + "\Skins\" + c + "\" + e) ' This will load the mask from the path you choose in the skin editor
' Open size settings
frmMain.Height = f
frmMain.MoveForm.Height = frmMain.Height ' the move label
frmMain.Width = g
frmMain.MoveForm.Width = frmMain.Width ' the move label
'Open control settings
frmMain.Title.Top = h
frmMain.Title.Left = i
frmMain.Exit.Top = j
frmMain.Exit.Left = k
frmMain.Minimize.Top = l
frmMain.Minimize.Left = m
If n = "1" Then
frmMain.Exit.Caption = ""
frmMain.Minimize.Caption = ""
Else
frmMain.Exit.Caption = "X"
frmMain.Minimize.Caption = "_"
End If

End If
End If
Loop Until EOF(no)
Close #no ' close file

Call frmMain.ChangeMask ' This is for updating the mask

End Sub



