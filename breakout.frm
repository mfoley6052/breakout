VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEvent 
      Caption         =   "New Event"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000016&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Prisoner"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2880
      Width           =   4935
   End
   Begin VB.ListBox lstPrisoners 
      Height          =   4935
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
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
    gender As String
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
Dim temp() As String
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
    Next x
    Close #1
    Open App.Path & "\yob1990.txt" For Input As #1
    lenFile = 0
        Do While Not EOF(1)
            Line Input #1, prisoners(numPrisoners).name
         temp() = Split(prisoners(numPrisoners).name, " ")
         prisoners(numPrisoners).name = temp(0)
         prisoners(numPrisoners).gender = temp(1)
            lenFile = lenFile + 1
        Loop
        lineNum = Int((Rnd() * lenFile) + 1)
        Seek #1, 1
        curLine = 1
         
        Do While curLine <= lineNum
            Line Input #1, prisoners(numPrisoners).name
            temp() = Split(prisoners(numPrisoners).name, " ")
            prisoners(numPrisoners).name = temp(0)
            prisoners(numPrisoners).gender = temp(1)
            curLine = curLine + 1
        Loop
    Close #1
    Call eventOccur("New Prisoner arrived: " & prisoners(numPrisoners).name)
    lstPrisoners.AddItem (prisoners(numPrisoners).name & "    " & prisoners(numPrisoners).trait(1) & "    " & prisoners(numPrisoners).trait(2) & "    " & prisoners(numPrisoners).trait(3))
Next z
End Function

Private Sub cmdEvent_Click()
Call eventOccur(InputBox("Enter the event message: ", "New Event"))
End Sub

Private Sub Command1_Click()
Call newPrisoner
End Sub

Private Sub Form_Load()
Static temp As Integer
temp = 160
For x = txtUpdates.LBound To txtUpdates.UBound
    txtUpdates(x).ForeColor = RGB(temp, temp, temp)
    temp = temp - 32
    txtUpdates(x).Text = ""
Next x
End Sub

Private Function eventOccur(ByVal eventMsg As String) As Boolean
txtUpdates(4).Text = txtUpdates(3).Text
txtUpdates(3).Text = txtUpdates(2).Text
txtUpdates(2).Text = txtUpdates(1).Text
txtUpdates(1).Text = txtUpdates(0).Text
txtUpdates(0).Text = eventMsg
End Function

