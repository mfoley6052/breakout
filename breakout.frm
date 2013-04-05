VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Breakout"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   5880
      Width           =   7335
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H80000007&
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtInput 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   1215
      Left            =   7320
      TabIndex        =   8
      Top             =   5400
      Width           =   2535
   End
   Begin VB.CommandButton cmdEvent 
      BackColor       =   &H80000007&
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
      ForeColor       =   &H80000016&
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
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
      ForeColor       =   &H80000000&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3360
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtUpdates 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
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



Private Sub cmdEvent_Click()
Call eventOccur(InputBox("Enter the event message: ", "New Event"))
End Sub

Private Sub cmdSend_Click()

End Sub

Private Sub Command1_Click()
Call newPrisoner
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MsgBox (KeyCode)
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

Private Sub txtInput_Change()
If txtInput.Text <> "" Then
    cmdSend.Enabled = True
Else
    cmdSend.Enabled = False
End If
End Sub
