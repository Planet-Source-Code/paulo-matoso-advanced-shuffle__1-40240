Attribute VB_Name = "GLOBAL"


' Some useful constants
Global Const FACEDOWN = 0
Global Const FACEUP = 1
Global Const CARDWIDTH = 71
Global Const CARDHEIGHT = 96
Global Const OFFSET = 30

Global gGameNumber As Double
Global gOldGameNumber As Double


' Required in Form Load to use dll
Declare Function InitializeDeck Lib "qcard32.dll" (ByVal hwnd As Long) As Long
' Reset all card values to defaults
Declare Sub SetDefaultValues Lib "qcard32.dll" ()
' Set the currently used card back design for cards 105 to 109
Declare Sub SetCurrentBack Lib "qcard32.dll" (ByVal nIndex As Long)

' Card drawing subs
Declare Sub DrawSymbol Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nValue As Long, ByVal x As Long, ByVal y As Long)
Declare Sub DrawCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal x As Long, ByVal y As Long)
Declare Sub DrawBack Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nValue As Long, ByVal x As Long, ByVal y As Long)
Declare Sub DealCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal x As Long, ByVal y As Long)
Declare Sub RemoveCard Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long)

' Get card information functions
Declare Function GetCardColor Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardSuit Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardValue Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardStatus Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardBlocked Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function IsCardDisabled Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardX Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetCardY Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser1 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser2 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser3 Lib "qcard32.dll" (ByVal nCard As Long) As Long
Declare Function GetUser4 Lib "qcard32.dll" (ByVal nCard As Long) As Long

' Set card information subs
Declare Sub SetCardStatus Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub AdjustCardBlocked Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetCardDisabled Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetCardX Lib "qcard32.dll" (ByVal nCard As Long, ByVal x As Long)
Declare Sub SetCardY Lib "qcard32.dll" (ByVal nCard As Long, ByVal y As Long)
Declare Sub SetUser1 Lib "qcard32.dll" (ByVal nCard As Long, ByVal bValue As Long)
Declare Sub SetUser2 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetUser3 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetUser4 Lib "qcard32.dll" (ByVal nCard As Long, ByVal nValue As Long)
Declare Sub SetOffSet Lib "qcard32.dll" (ByVal nValue As Long)

' Dragging subs and functions
Declare Function InitDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Sub AbortDrag Lib "qcard32.dll" ()
Declare Sub DoDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long)
Declare Function EndDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Sub ReturnDrag Lib "qcard32.dll" (ByVal hwnd As Long, ByVal nCard As Long, ByVal nOldX As Long, ByVal nOldY As Long)
Declare Sub BlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal y As Long)
Declare Function EndBlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Sub ReturnBlockDrag Lib "qcard32.dll" (ByVal hwnd As Long, nFirst As Long, ByVal nNumCards As Long, ByVal x As Long, ByVal y As Long)
Declare Function GetFreeDestination Lib "qcard32.dll" (ByVal nSource As Long) As Long

' Undocumented functions, generally not used

' returns the number of any unblocked card which lies beneath the mouse coordinates x, y
Declare Function PointInFreeCard Lib "qcard32.dll" (ByVal x As Long, ByVal y As Long) As Long

' returns the number of any card whose top 16 (or OffSet) pixels lie beneath the mouse coordinates x, y
Declare Function PointInCardTop Lib "qcard32.dll" (ByVal x As Long, ByVal y As Long) As Long

' manually sets the active drag card for a subsequent DoDrag or BlockDrag call
Declare Sub SetActiveCard Lib "qcard32.dll" (ByVal nCard As Long)




