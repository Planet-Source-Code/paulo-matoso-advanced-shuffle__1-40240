VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00808000&
   Caption         =   "Advanced Shuffle Example  - by Paulo Matoso 2002"
   ClientHeight    =   4920
   ClientLeft      =   1335
   ClientTop       =   1800
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   328
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   878
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   12915
      TabIndex        =   13
      Top             =   750
      Width           =   12945
   End
   Begin VB.CommandButton cmdDec 
      Caption         =   "-"
      Height          =   255
      Left            =   9000
      TabIndex        =   12
      Top             =   4050
      Width           =   315
   End
   Begin VB.CommandButton cmdInc 
      Caption         =   "+"
      Height          =   255
      Left            =   9000
      TabIndex        =   11
      Top             =   3810
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   945
      Left            =   3090
      TabIndex        =   8
      Top             =   3930
      Width           =   855
   End
   Begin VB.TextBox lblGameNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   3870
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   60
      Width           =   3915
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Play this game number ->"
      Height          =   285
      Left            =   3990
      TabIndex        =   2
      Top             =   3930
      Width           =   2475
   End
   Begin VB.TextBox txtGameNumber 
      Height          =   285
      Left            =   6510
      MaxLength       =   13
      TabIndex        =   1
      Text            =   "0"
      Top             =   3900
      Width           =   2475
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Random game"
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      Caption         =   "You can play more than 9 trillion of diferent games with this function."
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   8025
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "This is usefull if you need to make a game with function ""Select game number for play"""
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3210
      Width           =   8025
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   $"FORM1.frx":0000
      Height          =   585
      Left            =   120
      TabIndex        =   7
      Top             =   2610
      Width           =   8025
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Selected Game number: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   5
      Top             =   90
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "This button generate a random game from 1 to 9000000000000"
      Height          =   465
      Left            =   150
      TabIndex        =   4
      Top             =   3930
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Type any number from 0 to 9000000000000 to see how the Shuffle function works. Note: the game number 0 as no shuffle, try it."
      Height          =   615
      Left            =   3990
      TabIndex        =   3
      Top             =   4260
      Width           =   4995
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' declare counters
Dim nCardsInGame As Integer
Dim Deck() As Long ' the Array for cards

'This little function take me a 12 hours to write and lots of coffes :)
Private Function FuncGetCardNumber(lTotalCardsInGame As Integer) As Integer
Dim lValue As Double

    lValue = (CDbl(gGameNumber) / 65536) / 32768
    'The var gGameNumber will change in this function, you need store the game number in other var
    'like in this example with var gOldGameNumber
    lValue = 129749 * lValue
    lValue = lValue - Int(lValue)
    lValue = (8505 * lValue) + 12345
    lValue = lValue - Int(lValue)
    gGameNumber = Int(lValue * 9000000000000#) ' you can change the value
                '9000000000000 to any other number, don´t forget to change
                'in all project
    FuncGetCardNumber = Int(lValue * lTotalCardsInGame)
End Function

Private Sub Shuffle()
Dim i As Integer
Dim Card As Integer
Dim CardValue As Integer

    
    For i = 1 To nCardsInGame 'reset the deck
        Deck(i) = i + 1
    Next i
    
    For i = 1 To nCardsInGame 'add card value
            lCardValue = FuncGetCardNumber(nCardsInGame) + 1
            lCard = Deck(i)
            Deck(i) = Deck(lCardValue)
            Deck(lCardValue) = lCard
    Next i
      

End Sub

Private Sub DrawCards()
Dim nStartY As Integer
Dim nSpacer As Integer
Dim lCount As Integer


    
Form1.Cls

Shuffle 'Shuffle the deck for selected game number

nStartY = (Form1.ScaleHeight - 2 * CARDHEIGHT) / 4
nSpacer = 15



    For i = 1 To nCardsInGame 'draw the card stored in array Deck
        DrawCard Picture1.hwnd, Deck(i), i * nSpacer, 5
    Next i

'Show game number
lblGameNumber.Text = gOldGameNumber
txtGameNumber.Text = gOldGameNumber
End Sub




Private Sub cmdRandom_Click()
'
    gGameNumber = Int(Rnd * 9000000000000# + 1)
    gOldGameNumber = gGameNumber 'store game number in this var for later use
    
    DrawCards
End Sub

Private Sub cmdSelect_Click()

gGameNumber = Val(txtGameNumber.Text) 'the selected game number
gOldGameNumber = gGameNumber 'store game number in this var for later use
    
DrawCards
End Sub

Private Sub Command1_Click()
About.Show 1
End Sub

Private Sub cmdInc_Click()
gGameNumber = Val(txtGameNumber.Text) + 1
If gGameNumber > 9000000000000# Then gGameNumber = 0
txtGameNumber.Text = gGameNumber
gOldGameNumber = gGameNumber 'store game number in this var for later use
DrawCards
End Sub

Private Sub cmdDec_Click()
gGameNumber = Val(txtGameNumber.Text) - 1
If gGameNumber < 0 Then gGameNumber = 9000000000000#
txtGameNumber.Text = gGameNumber
gOldGameNumber = gGameNumber 'store game number in this var for later use
DrawCards

End Sub

Private Sub Form_Load()
Dim nReturn As Integer

' Copy Qcard32.dll to your \windows\system directory or
   ' \windows\system32 if you use Windows XP

' try to load DLL
' a FALSE return value indicates problems, like other game use the DLL
nReturn = InitializeDeck(Form1.hwnd)
If nReturn = False Then
    MsgBox "Error loading QCard32.DLL"
    End
End If

Randomize ' initialize random numbers
nCardsInGame = 52 '52 cards in 1 deck, no jokers
ReDim Deck(1 To nCardsInGame)
gGameNumber = 0 ' game number 0, no Shuffle
End Sub

Private Sub Form_Paint()
gGameNumber = gOldGameNumber ' this is need before call the function DrawCards
                            ' or another game number is set
DrawCards
End Sub

