VERSION 5.00
Begin VB.Form About 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About - Advanced Shurffle Example"
   ClientHeight    =   2160
   ClientLeft      =   1665
   ClientTop       =   2325
   ClientWidth     =   6225
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Height          =   396
      Left            =   5220
      TabIndex        =   0
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thanks to Stephen Murphy for this excelent library QCARD32.DLL."
      Height          =   465
      Left            =   2160
      TabIndex        =   4
      Top             =   1530
      Width           =   3945
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"ABOUT.frx":0000
      Height          =   885
      Left            =   2160
      TabIndex        =   3
      Top             =   630
      Width           =   3945
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E-Mail: paulomt1@clix.pt"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   330
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright by Paulo Matoso - 2002"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   30
      Width           =   2985
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_Paint()

Dim yLoc As Integer

yLoc = (About.ScaleHeight - CARDHEIGHT) / 2
DrawCard About.hwnd, 111, 5, yLoc
DrawCard About.hwnd, 112, 5 + OFFSET, yLoc
DrawCard About.hwnd, 111, 5 + (2 * OFFSET), yLoc

End Sub

