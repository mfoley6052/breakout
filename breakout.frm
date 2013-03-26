VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPrisoners 
      Height          =   4935
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Prisoner"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type human
    trait(1 To 3) As String
    Size As Integer
    name As String
End Type
Dim intLines As Integer
Dim prisoners() As human
Dim numPrisoners
Private Function newPrisoner(Optional ByVal numToCreate As Integer) As Boolean
Randomize Timer
Dim rand As Integer
Dim lineNum As Long
Dim lenFile As Long
Dim curLine As Long
Close #1
If numToCreate = 0 Then
    numToCreate = 1
End If
For z = 1 To numToCreate
    numPrisoners = numPrisoners + 1
    Open App.Path & "\traits.txt" For Input As #1
    ReDim Preserve prisoners(numPrisoners) As human
    For x = 1 To 3
    lenFile = 0
        Do While Not EOF(1)
            Line Input #1, prisoners(numPrisoners).trait(x)
            lenFile = lenFile + 1
        Loop
        lineNum = Int((Rnd() * lenFile) + 1)
        Seek #1, 1
        curLine = 1
         
        Do While curLine <= lineNum
            Line Input #1, prisoners(numPrisoners).trait(x)
            curLine = curLine + 1
        Loop
       ' MsgBox lineNum & " : " & Chr(34) & prisoners(numprisoners).trait(x) & Chr(34)
    Next x
    Close #1
    Open App.Path & "\yob1990.txt" For Input As #1
    lenFile = 0
        Do While Not EOF(1)
            Line Input #1, prisoners(numPrisoners).name
            lenFile = lenFile + 1
        Loop
        lineNum = Int((Rnd() * lenFile) + 1)
        Seek #1, 1
        curLine = 1
         
        Do While curLine <= lineNum
            Line Input #1, prisoners(numPrisoners).name
            curLine = curLine + 1
        Loop
    Close #1
    lstPrisoners.AddItem (prisoners(z).name & "    " & prisoners(z).trait(1) & "    " & prisoners(z).trait(2) & "    " & prisoners(z).trait(3))
Next z
End Function

Private Sub Command1_Click()
Call newPrisoner(50)
End Sub

