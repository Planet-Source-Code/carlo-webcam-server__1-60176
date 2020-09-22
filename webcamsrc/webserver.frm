VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmweb 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Webcam Server"
   ClientHeight    =   1860
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4155
   Icon            =   "webserver.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCam 
      Height          =   1335
      Left            =   1680
      ScaleHeight     =   1275
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtport 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Text            =   "4040"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdsize 
      Caption         =   "format"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "Picture size"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtms 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "30s"
      Top             =   1200
      Width           =   495
   End
   Begin VB.Timer tmrmain 
      Interval        =   1000
      Left            =   960
      Top             =   4320
   End
   Begin MSWinsockLib.Winsock mailsock 
      Left            =   1680
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrmail 
      Interval        =   2000
      Left            =   1320
      Top             =   4200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3945
      Begin VB.CommandButton cmdconnect 
         Caption         =   "Start"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin MSWinsockLib.Winsock lisSock 
         Left            =   960
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         LocalPort       =   80
      End
      Begin MSWinsockLib.Winsock sendSock 
         Left            =   960
         Top             =   2400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Port :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delay:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "URL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Status : OFF"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   2115
      End
      Begin VB.Label lblremote 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   2520
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock sockpager 
      Left            =   1920
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrstat 
      Interval        =   1000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Image imgsrc 
      Height          =   3675
      Left            =   4680
      Picture         =   "webserver.frx":058A
      Top             =   600
      Width           =   2925
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnumax 
         Caption         =   "maximize"
      End
      Begin VB.Menu mnumin 
         Caption         =   "minimize"
      End
      Begin VB.Menu mnuline 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "frmweb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////
'Webcam server

'

'Option Explicit

Dim imgFile As String, PageSent As Boolean, httpHeader As String
Public IconObject As Object
Const conMinimized = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdsize_Click()
SetCamSize
End Sub

Private Sub Form_Load()
   
'center the form
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
     

    SavePicture imgsrc.Picture, SystemDirectory & "\cam.jpg"
    
   
   PageSent = False
    
   cmdconnect.Caption = "Stop"

    
   PageSent = False
    
    
       
   httpHeader = "HTTP/1.0 200 OK" & vbCrLf
   httpHeader = httpHeader & "X -Host: webserver" & vbCrLf
   httpHeader = httpHeader & "Connection: Close" & vbCrLf
   httpHeader = httpHeader & "Content-Type: text/html" & vbCrLf & vbCrLf & vbCrLf
    
   
End Sub

Sub ExecuteFile(ByVal FFname As String, Ftype As String)
On Error Resume Next
Shell FFname, Ftype
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 'Query unload event
    
If cmdconnect.Caption = "Stop" Then
   'Make sure to disconnect from capture source - if it is connected upon termination the program can become unstable
  StopCam
   
End If
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
StopCam
End Sub

Private Sub keylog_Timer()

If GetKey Then
           
           
          txtlog = txtlog + sKeyPressed ' any keypresses ?
        End If

End Sub

Private Sub lisSock_ConnectionRequest(ByVal requestID As Long)
  sendSock.Close
  sendSock.Accept requestID
End Sub

Private Sub cmdconnect_Click()
  
  With cmdconnect
 
If cmdconnect.Caption <> "Stop" Then
   
       
   StartCam
   
        lisSock.Close
        lisSock.LocalPort = txtport.Text
        lisSock.Listen
        sendSock.Close
       .Caption = "Stop"
         Label2.Caption = "Status : Online"
         Label3.Caption = "http://" & GetInternetIP(True) & ":4040"
     
      
      Else
         lisSock.Close
         sendSock.Close
        .Caption = "Begin"
         Label2.Caption = "Status : Off"
         Label3.Caption = ""
    
    
    
      'Make sure to disconnect from capture source!!!
        StopCam
          
  End If
End With
End Sub



Private Sub sendSock_DataArrival(ByVal bytesTotal As Long)
 
    Dim requestedPage As String, strData As String, postedData As String
    Dim secondSpace As Integer, findGet As Integer, findPostedData As Integer
     
'     sckServer(Index).GetData strData
 
  'Dim strData As String
  Dim mypath As String
 Dim pict As String
 Dim ff As Long
  sendSock.GetData strData
 
 
 If InStr(1, strData, "shot") > 0 And PageSent = True Then
        pict = getImg(SystemDirectory & "\cam.jpg")  'converted camshot
        sendSock.SendData pict
        PageSent = False
       On Error Resume Next

      CamToBMP SystemDirectory & "\cam.bmp"
        
        
        picCam.Picture = LoadPicture(SystemDirectory & "\cam.bmp")
        
        
     
        
        DoEvents
        
         SAVEJPEG SystemDirectory & "\cam.jpg", 70, picCam
       DoEvents
       
       
       
       
       
       
    Kill SystemDirectory & "\cam.bmp"
           
     Else
 
    
    
    sendSock.SendData httpHeader
    sendSock.SendData "<html>" & vbCrLf
    sendSock.SendData "<head> <title>WebShot</title>"
    sendSock.SendData "<META http-equiv='Page-Enter' content='revealtrans(duration=3,transition=12'>" & vbCrLf
    sendSock.SendData "<META HTTP-EQUIV='Refresh' CONTENT=30;URL='http://" & GetInternetIP(True) & ":4040" & "/'>"
 
    
   sendSock.SendData "<script>"
   sendSock.SendData "function countDown() {"
   sendSock.SendData "count.innerHTML = countValue;"
   sendSock.SendData "countValue = countValue - 1;"
   sendSock.SendData "if (countValue >= 0) {setTimeout('countDown()', 1000);}"
   sendSock.SendData "}"
   sendSock.SendData "</script>"
    
 
    
    
    sendSock.SendData "</head>"
    sendSock.SendData "<body OnLoad= 'countValue=30; countDown()' bgcolor=1d1d26 text=FFFFFF>"
    sendSock.SendData "<center>"
        
    sendSock.SendData "<h4>Remote WebCam</h4>" & vbCrLf

    sendSock.SendData "<FONT SIZE=3>Reloading in : <font color=68876F><span id='count'></font>&nbsp;&nbsp;</span> Seconds, Please wait or hit reload!!</font><br><br>" & vbCrLf
    sendSock.SendData "<img src='shot.jpg' alt=reloading please wait..><br>" & vbCrLf & vbCrLf
  
  

    
    sendSock.SendData "<!--WebServer 1.0-->" & vbCrLf
    sendSock.SendData "</body>"
    sendSock.SendData "</html>"
    PageSent = True
 
    
 End If

End Sub

Private Sub sendSock_SendComplete()
sendSock.Close
Label2.Caption = "Status : Online"
End Sub

Private Sub Label3_Click()

Call ShellExecute(0&, vbNullString, "http://" & GetInternetIP(True) & ":4040", vbNullString, vbNullString, vbNormalFocus)

End Sub

Private Function getImg(FileName)
Dim f As Long
Dim Temp As String '-----add

f = FreeFile
Temp = ""
Open FileName For Binary As #f        ' Open file.
  Temp = Input(FileLen(FileName), #f) ' Get entire img data
Close #f
getImg = Temp
End Function


