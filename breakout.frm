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

