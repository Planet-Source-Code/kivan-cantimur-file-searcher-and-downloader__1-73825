VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "File Searcher"
   ClientHeight    =   3150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsDisconnected 
      Left            =   2280
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listen"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin MSWinsockLib.Winsock wMsg 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock wSendRecrd 
      Left            =   5400
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Search"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   5400
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   1920
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label txtMsg 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   45
   End
   Begin VB.Label serv 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "status"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Private Const CB_FINDSTRING = &H14C
Private blnDelete As Boolean

'---------------------------------------
Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    lastamount As Double
End Type
Private info As da
Function getfilename(ByVal filepath As String)
On Error Resume Next
Dim ta() As String
ta = Split(filepath, "\")
getfilename = ta(UBound(ta))
End Function

Private Sub Command1_Click()
serv.Caption = ""
With w1
.Close 'incase winsock1 is open for any reason
.LocalPort = 3456 'defines the port to monitor for incoming connections
.Listen 'starts the listening process
End With
With wSendRecrd
.Close 'incase winsock1 is open for any reason
.LocalPort = 3457 'defines the port to monitor for incoming connections
.Listen 'starts the listening process
End With
With wMsg
.Close
.LocalPort = 3458
.Listen
End With
With wsDisconnected
.Close
.LocalPort = 3459
.Listen
End With
Label1.Caption = "Listening..."
Command1.Enabled = False
End Sub

Private Sub Command7_Click()
txtMsg.Caption = ""
wSendRecrd.SendData txtSearch.Text
End Sub

Private Sub Form_Load()
Command7.Enabled = False
End Sub

Private Sub serv_Change()
If serv.Caption = "Server disconnected." Then
    Command7.Enabled = False
    Text2.Text = ""
    Label1.Caption = "----"
    Command1.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim e, spid, info
e = info.lastamount
e = e \ 1024 'convert bytes to KB
spid.Caption = e
info.lastamount = 0
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
DoEvents  'so the progam doesn't freezes
End Sub

Private Sub w1_ConnectionRequest(ByVal requestID As Long)
w1.Close 'closes winsock1 if it was open
w1.Accept (requestID) 'accepts the incoming connection from the client
Command7.Enabled = True
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dat As String
Dim pa, adt
w1.GetData dat, vbString
If LCase(Mid(dat, 1, 11)) = "sendrequest" Then
    Dim temparray() As String
    Dim fname As String
    Dim fsize As Double
    temparray = Split(dat, "|")
    fname = temparray(1)
    fsize = temparray(2)
    Requestfrm.fsize = fsize
    Requestfrm.fname = fname
    pa = App.Path
    If Len(pa) = 3 Then pa = Mid(pa, 1, 2)
    pa = pa & "\"
    p1.Max = fsize \ 2
    Requestfrm.fpath = pa & fname
    Requestfrm.Show 1
    Exit Sub
End If

If LCase(Mid(dat, 1, 2)) = "ok" Then
    Dim temparray2() As String
    Dim fname2 As String
    Dim fsize2 As Double
    Dim e, r
    temparray2 = Split(dat, "|")
    fname2 = temparray2(1)
    fsize2 = temparray2(2)
    If fname2 <> getfilename(info.FileToSend) Or fsize2 <> info.FileSize Then Exit Sub
    Close #1
    Open info.FileToSend For Binary Access Read As #1
        If LOF(1) = 0 Then Exit Sub
        Dim SendBuffer As String
        SendBuffer = Space$(LOF(1))
        Get #1, , SendBuffer
    Close #1
    w1.SendData SendBuffer & "/\/\ENDOFFILE/\/\"
    e = 2
    r = Timer
    Do Until Timer > r + 2  'leave the pc to send the file
    DoEvents
    Loop
    Exit Sub
End If

If LCase(Mid(dat, 1, 5)) = "notok" Then
    MsgBox "The Client Does Not Accept The File Tranfer Request"
    Exit Sub
End If

If Right(dat, 17) = "/\/\ENDOFFILE/\/\" Then
    Dim aaa As String
    aaa = Mid(dat, 1, Len(dat) - 17)
    Put #2, , aaa
    Close #2
    MsgBox "File Transfer Completed"
    p1.Value = 0
    Command2.Enabled = True
    Exit Sub
End If

adt = Len(dat) \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
info.lastamount = info.lastamount + Len(dat)
Put #2, , dat
End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
MsgBox "Error : " & Description
End Sub

Private Sub w1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
Dim adt As Double
adt = bytesSent \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
info.lastamount = info.lastamount + bytesSent
DoEvents
End Sub

Private Sub wMsg_ConnectionRequest(ByVal requestID As Long)
wMsg.Close 'closes winsock1 if it was open
wMsg.Accept (requestID) 'accepts the incoming connection from the client
Text2.Text = wMsg.RemoteHostIP 'sets text1 to the ip of the connected client
End Sub

Private Sub wMsg_DataArrival(ByVal bytesTotal As Long)
Dim dataX As String
wMsg.GetData dataX
txtMsg.Caption = dataX
End Sub

Private Sub wsDisconnected_ConnectionRequest(ByVal requestID As Long)
wsDisconnected.Close
wsDisconnected.Accept (requestID)
Text2.Text = wsDisconnected.RemoteHostIP
End Sub

Private Sub wsDisconnected_DataArrival(ByVal bytesTotal As Long)
Dim dataX As String
wsDisconnected.GetData dataX
serv.Caption = dataX
End Sub

Private Sub wSendRecrd_ConnectionRequest(ByVal requestID As Long)
wSendRecrd.Close 'closes winsock1 if it was open
wSendRecrd.Accept (requestID) 'accepts the incoming connection from the client
Text2.Text = wSendRecrd.RemoteHostIP 'sets text1 to the ip of the connected client
End Sub


