VERSION 5.00
Begin VB.Form Requestfrm 
   Caption         =   "Receiver"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Accept"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox fpath 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label fname 
      AutoSize        =   -1  'True
      Caption         =   "Unknown"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "File name:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bytes"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   390
   End
   Begin VB.Label fsize 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File size:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "Requestfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Close #2
Open fpath.Text For Binary Access Write As #2
Form1.w1.SendData "ok|" & fname.Caption & "|" & fsize.Caption & "|"
DoEvents
DoEvents
Me.Hide
End Sub

Private Sub Form_Load()
fpath.Text = App.Path & "\"
End Sub

