VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "File Searcher "
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock wsDs 
      Index           =   0
      Left            =   4320
      Top             =   8160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Connect"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   7320
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   2160
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsQuery 
      Index           =   0
      Left            =   1560
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   8880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   8400
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7920
      Top             =   8160
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send File"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2040
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Text            =   "127.0.0.1"
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6240
      TabIndex        =   9
      Top             =   240
      Width           =   4575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8281
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Folder"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size (Kb)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7320
      Top             =   4680
   End
   Begin VB.TextBox WhatSearch 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox SearchPath 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Text            =   "c:\"
      Top             =   240
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Search.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   420
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disconnected"
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   7680
      Width           =   990
   End
   Begin VB.Label lblRI 
      AutoSize        =   -1  'True
      Caption         =   "Remote IP"
      Height          =   195
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label lblUsersConnected 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   90
   End
   Begin VB.Label Label3 
      Caption         =   "File to Find:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Path To Investigate:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FOL = "******************"
Dim i, j, temp, pointer, flag As Integer  'timer2 variables
Dim s As String 'timer2 variable

Dim nul, nul2, effe As Integer

Dim pointer_from, pointer_to, files, counter1, counter2 As Long
'----------------------------------------------------------
'This counts how many users are connected to the server
Dim TotalUsersConnected As Integer

Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Private Const CB_FINDSTRING = &H14C
Private blnDelete As Boolean

'--------------------------------------------------- IP

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
     Private Const MIN_SOCKETS_REQD = 1
     Private Const SOCKET_ERROR = -1
     Private Const WSADescription_Len = 256
    Private Const WSASYS_Status_Len = 128

    Private Type HOSTENT
        hName As Long
        hAliases As Long
        hAddrType As Integer
        hLength As Integer
        hAddrList As Long
    End Type

    Private Type WSADATA
        wversion As Integer
        wHighVersion As Integer
        szDescription(0 To WSADescription_Len) As Byte
        szSystemStatus(0 To WSASYS_Status_Len) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpszVendorInfo As Long
    End Type

    Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
    Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
    Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
    Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

'----------------
Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    LastAmount As Double
End Type
Private info As da
Function getfilename(ByVal filepath As String)
On Error Resume Next
Dim ta() As String
ta = Split(filepath, "\")
getfilename = ta(UBound(ta))
End Function

    Function hibyte(ByVal wParam As Integer)

        hibyte = wParam \ &H100 And &HFF&

    End Function

    Function lobyte(ByVal wParam As Integer)

        lobyte = wParam And &HFF&

    End Function

Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

        iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

        If iReturn <> 0 Then
            MsgBox "Winsock.dll is not responding."
            End
        End If

        If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

            sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
            sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
            sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
            sMsg = sMsg & " is not supported by winsock.dll "
            MsgBox sMsg
               End
        End If

        'iMaxSockets is not used in winsock 2. So the following check is only
        'necessary for winsock 1. If winsock 2 is requested,
        'the following check can be skipped.

        If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
            sMsg = "This application requires a minimum of "
            sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
            MsgBox sMsg
            End
        End If
End Sub

Sub SocketsCleanup()
    Dim lReturn As Long

        lReturn = WSACleanup()

        If lReturn <> 0 Then
            MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
            End
        End If
End Sub

Private Sub Command1_Click()
Dim folder_depth As Integer

counter1 = 0
counter2 = 0
List1.Clear
ListView1.ListItems.Clear
Command1.Enabled = False
pointer = 0

'Setup our Dir1 List
Dir1 = SearchPath

'Add the path of the folder in
'which the folder will take place!
'(you could here add more paths...)
List1.AddItem (Dir1)

'set the pointers
pointer_from = 0
pointer_to = List1.ListCount - 1

folder_depth = -1

Do
DoEvents
    'analyze folders with depth=folder_depth
    folder_depth = folder_depth + 1

    'Start analysis!
    nul = analyze(pointer_from, pointer_to)

Loop Until nul = 1 'no more folders to investigate!


Command1.Enabled = True 'ready for next search
If Command1.Enabled = True And Text1.Text = "" Then
 '           On Error GoTo y
               Winsock2(SocketCount).SendData "File not found!"
                 DoEvents
                 DoEvents
                 DoEvents

               
'y:

End If
End Sub
Function analyze(ByVal pfrom, ByVal pto)
On Error Resume Next
Dim k1, k2 As Integer

'Analyze the area which pointers show
For k1 = pfrom To pto
DoEvents

    Dir1 = List1.List(k1) 'Search for folders this Path
    
    For k2 = 0 To Dir1.ListCount - 1
    DoEvents
    
        'Add the folders of this path into list1!
        List1.AddItem (Dir1.List(k2))
        
    Next

Next

'refresh pointers
pointer_from = pto + 1
pointer_to = List1.ListCount

If pointer_from <= pointer_to Then
    analyze = 0 'continue,there are folders to analyze
Else
    analyze = 1 'no folders left,end analysis!
End If

End Function

Sub statusR(ByVal st As Integer)
On Error Resume Next
Select Case st
Case 1
    'disconnected
    lblRI.Caption = "-----"
Case 2
    'connected
    lblRI.Caption = wsQuery(0).RemoteHostIP
    End Select
End Sub

Sub status(ByVal st As Integer)
'On Error Resume Next
TotalUsersConnected = 0
Select Case st
Case 1
    'disconnected
    Winsock1(0).Close
    Close #1
    Close #2
    Winsock2(0).Close
    wsQuery(0).Close
    lblRI.Caption = ""
Case 2
    'connected
    lblRI.Caption = Winsock1(0).RemoteHostIP
    Label1.Caption = "Connected"
    TotalUsersConnected = TotalUsersConnected + 1
    lblUsersConnected.Caption = "Total users connected: " & TotalUsersConnected
End Select
End Sub

Private Sub Command4_Click()
On Error Resume Next
wsDs(SocketCount).SendData "Server disconnected."
DoEvents
DoEvents
DoEvents
Winsock1(0).Close
Close #1
Close #2
Winsock2(0).Close
wsQuery(0).Close
wsDs(0).Close
Label1.Caption = "Disconnected"
lblRI.Caption = "----"
End Sub

Private Sub Command5_Click()
On Error GoTo y
Open Text1.Text For Append As #1
If LOF(1) = 0 Then
    MsgBox "The File Is Empty"
    Close #1
    Exit Sub
End If
info.FileToSend = Text1.Text
info.FileSize = LOF(1)
p1.Max = LOF(1) \ 2
Close #1
Winsock1(SocketCount).SendData "sendrequest|" & getfilename(info.FileToSend) & "|" & info.FileSize & "|"
DoEvents
DoEvents
DoEvents
y:
End Sub

Private Sub Command6_Click()
On Error Resume Next
'connect to server
With Winsock1(0)
    .Close
    .RemoteHost = Text2.Text
    .RemotePort = "3456"
    .Connect
End With
With wsQuery(0)
    .Close
    .RemoteHost = Text2.Text
    .RemotePort = "3457"
    .Connect
End With
With Winsock2(0)
    .Close
    .RemoteHost = Text2.Text
    .RemotePort = "3458"
    .Connect
End With
With wsDs(0)
    .Close
    .RemoteHost = Text2.Text
    .RemotePort = "3459"
    .Connect
End With
End Sub

Private Sub Form_Load()
Dim servID As Long
       Dim hostname As String * 256
       Dim hostent_addr As Long
       Dim host As HOSTENT
       Dim hostip_addr As Long
       Dim temp_ip_address() As Byte
       Dim i As Integer
       Dim ip_address As String

   SocketsInitialize
    If gethostname(hostname, 256) = SOCKET_ERROR Then
              MsgBox "Windows Sockets error " & Str(WSAGetLastError())
               Exit Sub
            Else
               hostname = Trim$(hostname)
           End If

           hostent_addr = gethostbyname(hostname)

          If hostent_addr = 0 Then
               MsgBox "Winsock.dll is not responding."
               Exit Sub
          End If

           RtlMoveMemory host, hostent_addr, LenB(host)
           RtlMoveMemory hostip_addr, host.hAddrList, 4

           'get all of the IP address if machine is  multi-homed

           Do
               ReDim temp_ip_address(1 To host.hLength)
               RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

               For i = 1 To host.hLength
                   ip_address = ip_address & temp_ip_address(i) & "."
               Next
               ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

               'MsgBox ip_address
                Text2.Text = ip_address
                
               ip_address = ""
               host.hAddrList = host.hAddrList + LenB(host.hAddrList)
               RtlMoveMemory hostip_addr, host.hAddrList, 4
            Loop While (hostip_addr <> 0)
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then Command5_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim e, spid, info
e = info.LastAmount
e = e \ 1024 'convert bytes to KB
spid.Caption = e
info.LastAmount = 0
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'ok here is the good stuff..
'the file search take place simultaneously
'with the folder invastigation!
'that's why we are using timer and not function.

If List1.ListCount > pointer Then


'means that there are folders
'which have been invastigated but we
'havent yet search for files inside them.

temp = List1.ListCount

'for each of these folders
'make a search for the file
'we looking for

For i = pointer To temp - 1
DoEvents

File1 = List1.List(i) 'load folder's files into file1

    For j = 0 To File1.ListCount - 1
    DoEvents
        
    
        If InStr(1, LCase(File1.List(j)), LCase(WhatSearch)) Then
        flag = 2
            'FIND WHAT YOU LOOKING FOR!
            'ADD IT IN THE LIST.
            s = File1.List(j)
 
            ListView1.ListItems.Add , , s
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = List1.List(i)
            Text1.Text = ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) & "\" & s
            On Error Resume Next
            'bad filename!
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = Round(FileLen(List1.List(i) + "\" + s) / 1024, 2)
            counter2 = counter2 + 1
        End If

    Next
Next
pointer = temp
End If

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
DoEvents  'so the progam doesn't freezes
End Sub

Private Sub WhatSearch_Change()
Text1.Text = ""
If WhatSearch.Text <> "" Then Command1_Click
End Sub

Private Sub Winsock1_Close(Index As Integer)
On Error Resume Next
status 1
End Sub

Private Sub Winsock1_Connect(Index As Integer)
On Error Resume Next
status 2
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
SocketCount = SocketCount + 1
Load Winsock1(SocketCount)
'Connection Requested
If Winsock1(0).State <> sckClosed Then Winsock1(0).Close
Winsock1(SocketCount).Accept requestID
'Connection Accepted
status 2
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim dat As String
Dim pa, adt
Dim SocketCheck As Integer

Winsock1(Index).GetData dat, vbString
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
    
    For SocketCheck = 0 To SocketCount Step 1
    Winsock1(SocketCheck).SendData SendBuffer & "/\/\ENDOFFILE/\/\"
    e = 2
    r = Timer
    Do Until Timer > r + 2  'leave the pc to send the file
    DoEvents
    Loop
    Next SocketCheck
    
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
    Exit Sub
End If

adt = Len(dat) \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
info.LastAmount = info.LastAmount + Len(dat)
Put #2, , dat
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
MsgBox "Error : " & Description
End Sub

Private Sub Winsock1_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
Dim adt As Double
adt = bytesSent \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
info.LastAmount = info.LastAmount + bytesSent
End Sub

Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim SocketCountx As Integer
SocketCountx = SocketCountx + 1
Load Winsock2(SocketCountx)
'Connection Requested
If Winsock2(0).State <> sckClosed Then Winsock2(0).Close
Winsock2(SocketCountx).Accept requestID
End Sub

Private Sub wsDs_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim SocketCountx As Integer
SocketCountx = SocketCountx + 1
Load wsDs(SocketCountx)
'Connection Requested
If wsDs(0).State <> sckClosed Then wsDs(0).Close
wsDs(SocketCountx).Accept requestID
End Sub

Private Sub wsQuery_Close(Index As Integer)
status 1
End Sub

Private Sub wsQuery_Connect(Index As Integer)
status 2
End Sub

Private Sub wsQuery_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Dim SocketCountx As Integer
SocketCountx = SocketCountx + 1
Load wsQuery(SocketCountx)
'Connection Requested
If wsQuery(0).State <> sckClosed Then wsQuery(0).Close
wsQuery(SocketCountx).Accept requestID
statusR 2
End Sub

Private Sub wsQuery_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim vData As String 'dims the data as a string
wsQuery(Index).GetData vData 'gets the incoming data from the client
WhatSearch.Text = vData
End Sub
