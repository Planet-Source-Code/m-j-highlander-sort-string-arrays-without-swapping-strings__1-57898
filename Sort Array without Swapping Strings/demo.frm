VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   2535
   ClientTop       =   2595
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6585
   Begin VB.ListBox List3 
      Height          =   5715
      Left            =   4320
      TabIndex        =   4
      Top             =   180
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fill List with Random  Strings"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   2535
   End
   Begin VB.ListBox List2 
      Height          =   5715
      Left            =   3000
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   6000
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim UnsortedTable() As String

Const MAX = 1000
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Function RndStr(ByVal StrLen)
' This function generates random strings
' The length is sprcified by the only parameter.
' Frankly, I can't think of any use for this function ;-)

Dim idx, ch, tmp

For idx = 1 To StrLen
    ch = Chr(RndInt(65, 90))
    tmp = tmp & ch
Next

RndStr = tmp

End Function
Public Function RndInt(Optional ByVal Lower As Integer = 0, Optional ByVal Upper As Integer = 0) As Integer
'Returns a random integer greater than or equal to the Lower parameter
'and less than or equal to the Upper parameter.

Const DEFAULT_MAX = 100

' if no arguments are provided, returns a randon number in the range 0 --> 100
If (Upper = 0) And (Lower = 0) Then
    Upper = DEFAULT_MAX
End If


If (Upper = 0) And (Lower <> 0) Then
    'we called the function with only one argument, so assume it is "Upper"
    Upper = Lower
    Lower = 0
End If

Randomize Timer

RndInt = Int(Rnd * (Upper - Lower + 1)) + Lower

End Function

Private Sub Command1_Click()
Dim lArray() As Long
Dim idx As Long


lArray = ShellSortLong(UnsortedTable)

List2.Clear
List3.Clear

LockWindowUpdate Me.hWnd

For idx = LBound(UnsortedTable) To UBound(UnsortedTable)
    List3.AddItem UnsortedTable(lArray(idx))
    List2.AddItem lArray(idx)
Next

LockWindowUpdate 0&

End Sub
Private Sub Command2_Click()
  Dim i As Long

LockWindowUpdate List1.hWnd

List1.Clear
ReDim UnsortedTable(1 To MAX)
For i = LBound(UnsortedTable) To UBound(UnsortedTable)
    UnsortedTable(i) = RndStr(RndInt(1, 20))
    List1.AddItem UnsortedTable(i)
Next i

LockWindowUpdate 0&

End Sub

