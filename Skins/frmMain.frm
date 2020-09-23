VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   6165
   ClientLeft      =   1530
   ClientTop       =   1245
   ClientWidth     =   8295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   0
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   5505
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8200
      Begin VB.CommandButton Preview 
         Appearance      =   0  'Flat
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         Picture         =   "frmMain.frx":0BD3
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Controls Caption"
         Height          =   855
         Left            =   120
         TabIndex        =   51
         Top             =   3360
         Width           =   3735
         Begin VB.OptionButton NoCap 
            Caption         =   "No caption"
            Height          =   255
            Left            =   1080
            TabIndex        =   55
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton ShowCap 
            Caption         =   "Show caption"
            Height          =   255
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CommandButton InfoCaption 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   52
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Show caption:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comments"
         Height          =   975
         Left            =   4080
         TabIndex        =   47
         Top             =   2160
         Width           =   3735
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   2535
         End
         Begin VB.CommandButton InfoComment 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   48
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Name"
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   3735
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton InfoName 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   43
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lName 
            BackStyle       =   0  'Transparent
            Caption         =   "Skin Name:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton Load 
         Appearance      =   0  'Flat
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         Picture         =   "frmMain.frx":0CD5
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Caption         =   "Controls"
         Height          =   1335
         Left            =   4080
         TabIndex        =   22
         Top             =   3240
         Width           =   3735
         Begin VB.TextBox txtLeftM 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2280
            TabIndex        =   38
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton InfoMin 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   36
            Top             =   960
            Width           =   255
         End
         Begin VB.TextBox txtTopM 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   35
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtLeftE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2280
            TabIndex        =   32
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton InfoExit 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   30
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtTopE 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   29
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtLeftT 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2280
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton InfoTitles 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtTopT 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lLeftM 
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   40
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lTopM 
            BackStyle       =   0  'Transparent
            Caption         =   "Top:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lMin 
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lLeftE 
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   34
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lTopE 
            BackStyle       =   0  'Transparent
            Caption         =   "Top:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   33
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lExit 
            BackStyle       =   0  'Transparent
            Caption         =   "Exit:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lLeftT 
            BackStyle       =   0  'Transparent
            Caption         =   "Left:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   28
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lTopT 
            BackStyle       =   0  'Transparent
            Caption         =   "Top:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   27
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lTitles 
            BackStyle       =   0  'Transparent
            Caption         =   "Title:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Form"
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   3735
         Begin VB.TextBox txtHeight 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton InfoHeight 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtWidth 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   17
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton IndoWidth 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   600
            Width           =   255
         End
         Begin VB.Label lHeight 
            BackStyle       =   0  'Transparent
            Caption         =   "Height:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lWidth 
            BackStyle       =   0  'Transparent
            Caption         =   "Width:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Images"
         Height          =   975
         Left            =   4080
         TabIndex        =   6
         Top             =   1080
         Width           =   3735
         Begin VB.CommandButton InfoMaskImage 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton OpenMaskImage 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtMaskImage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton InfoMainImage 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   10
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton OpenMainImage 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   9
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtMainImage 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lMaskImage 
            BackStyle       =   0  'Transparent
            Caption         =   "Mask Image:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lMainImage 
            BackStyle       =   0  'Transparent
            Caption         =   "Main Image:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton Save 
         Appearance      =   0  'Flat
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Picture         =   "frmMain.frx":0DD7
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "All finished skins must be in a folder in the Skin Demo Skin directory, if not it will not run  :)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   50
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Info 
         BackStyle       =   0  'Transparent
         Caption         =   "Just enter the information needed. To find out what everthing does click the ? buttons..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   46
         Top             =   480
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":0ED9
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Exit 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   8040
         TabIndex        =   5
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Resize 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   7920
         TabIndex        =   4
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Move 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Skin Maker"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'*               UPDATES               *
'* I updated the move form code (thanks*
'* to Joey Burgett) and I added a new  *
'* option to the skins, Control's capt-*
'* ion, it allows you to show or hide  *
'* the text in the exit and minimize   *
'* controls of the main program        *
'* I also made a preview form so you   *
'* can see what the skin would look    *
'* like!!!                             *
'***************************************

'***************************************
'*               SkinBuilder           *
'* This part of the example allows you *
'* to create your own skins...         *
'*                                     *
'***************************************

Private Sub Exit_Click()
End ' en project
End Sub

Private Sub Form_Load()
' Just to make the form the right size
Me.Width = Picture1.Width
Me.Height = Picture1.Height
End Sub




Private Sub Load_Click()
On Error Resume Next ' In case of error
Dim sOpen As SelectedFile
'This will open a skin for editing.
Dim filename
Dim Count As Integer

'FileDialog.sInitDir = OptDefPath
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
FileDialog.sDlgTitle = "Open Skin file"
FileDialog.sInitDir = App.Path
FileDialog.sFilter = "Skin file (*.skn)" & Chr$(0) & "*.skn"
     sOpen = ShowOpen(Me.hWnd)
      If Err.Number <> 32755 And sOpen.bCanceled = False Then
        FileList = sOpen.sLastDirectory
        For Count = 1 To sOpen.nFilesSelected
            FileList = FileList & sOpen.sFiles(Count)
        Next Count
         Screen.MousePointer = 11
        OpenFile (FileList)
     
End If
  Screen.MousePointer = 0 ' Icon normal
 

End Sub


Private Sub Move_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************
'*               UPDATE!               *
'* Very little code update, thanks to  *
'* Joey Burgett                        *
'* ohiojedi@hotmail.com                *
'***************************************

' To move form
    ReleaseCapture
    Call SendMessage(Me.hWnd, &HA1, 2, 0)
    End Sub

Private Sub OpenMainImage_Click()
On Error Resume Next ' in case of error
' This will choose your main image
Dim sOpen As SelectedFile
Dim filename

FileDialog.sInitDir = OptDefPath
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
FileDialog.sDlgTitle = "Open Skin file"
FileDialog.sFilter = "All supported files" & Chr$(0) & "*.jpg;*.bmp;*.dib;*.wmf;*.emf;*.gif" & Chr$(0) & "JPG Images(*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & "Bitmaps (*.bmp;*.dib)" & Chr$(0) & "*.bmp;*.dib" & Chr$(0) & "GIF Images(*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Metafiles (*.wmf;*.emf)" & Chr$(0) & "*.wmf;*.emf" & Chr$(0) & "All Files" & Chr$(0) & "*.*"
     sOpen = ShowOpen(Me.hWnd)
 
 
 Screen.MousePointer = 11
        txtMainImage.Text = FileDialog.sFileTitle
Screen.MousePointer = 0


End Sub



Private Sub OpenMaskImage_Click()
On Error Resume Next ' In case of error
' This will choose your mask image
Dim sOpen As SelectedFile
Dim filename

FileDialog.sInitDir = OptDefPath
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
FileDialog.sDlgTitle = "Open Skin file"
FileDialog.sFilter = "All supported files" & Chr$(0) & "*.jpg;*.bmp;*.dib;*.wmf;*.emf;*.gif" & Chr$(0) & "JPG Images(*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & "Bitmaps (*.bmp;*.dib)" & Chr$(0) & "*.bmp;*.dib" & Chr$(0) & "GIF Images(*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Metafiles (*.wmf;*.emf)" & Chr$(0) & "*.wmf;*.emf" & Chr$(0) & "All Files" & Chr$(0) & "*.*"
     sOpen = ShowOpen(Me.hWnd)
               
       
 Screen.MousePointer = 11
       txtMaskImage.Text = FileDialog.sFileTitle
Screen.MousePointer = 0

End Sub


Private Sub Preview_Click()
OpenSkin ' This will load the skin for the preview form
frmPreview.Show
End Sub

Private Sub Resize_Click()
On Error Resume Next ' in case of error
' just for when you click on the resize button
If Me.Height = 5535 Then
Me.Height = 180
Picture1.Height = 180
Else
Me.Height = 5535
Picture1.Height = 5535
End If
End Sub


Private Sub Save_Click()
Dim sSave As SelectedFile
On Error Resume Next ' In case of error
'This will save the skin.
    
    FileDialog.sFilter = "Skin file (*.skn)" & Chr$(0) & "*.skn"
    
    ' See Standard CommonDialog Flags for all options
    FileDialog.flags = OFN_HIDEREADONLY
    FileDialog.sDlgTitle = "Save skin file"
  FileDialog.sInitDir = App.Path
  FileDialog.sDefFileExt = "*.skn"
    sSave = ShowSave(Me.hWnd)
 
Screen.MousePointer = 11 'Give the user an hour glass
  
        SaveFile (FileDialog.sFileTitle) ' save skin
            

Screen.MousePointer = 0

End Sub

' Message boxes!!
Private Sub IndoWidth_Click()
MsgBox "Determines the Width of the Main form, the Player and the Time/Date", vbInformation, "Height"
End Sub


Private Sub InfoComment_Click()
MsgBox "Enter here any comments you would like to add to the .skn file.", vbInformation, "Comments"
End Sub


Private Sub InfoExit_Click()
MsgBox "Determines the Top and the left of the exit label.", vbInformation, "Exit"
End Sub

Private Sub InfoHeight_Click()
MsgBox "Determines the height of the Main form, the Player and the Time/Date", vbInformation, "Height"
End Sub

Private Sub InfoMainImage_Click()
MsgBox "Determines this is the image that the Main form, the Player and the Time/Date will have.", vbInformation, "Main Image"
End Sub

Private Sub InfoMaskImage_Click()
MsgBox "Determines this is the shape that the Main form, the Player and the Time/Date will have", vbInformation, "Mask image"
End Sub

Private Sub InfoMin_Click()
MsgBox "Determines the Top and the left of the Minimize label.", vbInformation, "Minimize"
End Sub

Private Sub InfoName_Click()
MsgBox "This is very important, this is the name of the folder where you skin file is (i.e.:C:\Skin demo directory\Skins\'The Skin name')", vbInformation, "Folder name"
End Sub


Private Sub InfoTitles_Click()
MsgBox "Determines the Top and the left of the titles", vbInformation, "Titles"
End Sub

Private Sub InfoCaption_Click()
MsgBox "Determines if the caption in the controls ( X, - ) should be visible or not.", vbInformation, "Control's caption"
End Sub
