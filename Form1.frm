VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   443
      TabIndex        =   1
      Top             =   1058
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   443
      TabIndex        =   0
      Top             =   218
      Width           =   1455
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   495
      Left            =   1920
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program allows a user to send a wave file the is
'compiled into the form. just right click on the OLE and
'goto Wave Sound and then locate the file and click on
'exit in the sound recorder. But i think you have to goto
'Edit and insert file.

Option Explicit

Private Sub Command1_Click()
    OLE1.Action = 7 'Start
End Sub

Private Sub Command2_Click()
    OLE1.Action = 9 'Stop
End Sub
