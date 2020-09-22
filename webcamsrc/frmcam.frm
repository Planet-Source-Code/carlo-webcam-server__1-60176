VERSION 5.00
Begin VB.Form frmcam 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Capture Cam"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmcam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   " &close"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   1320
   End
   Begin VB.Image imgcam 
      BorderStyle     =   1  'Fixed Single
      Height          =   3660
      Left            =   100
      Top             =   120
      Width           =   4860
   End
End
Attribute VB_Name = "frmcam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.Clear
Unload Me
End Sub

Private Sub Form_Load()
   
   
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
      
   tmr.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
tmr.Enabled = False
  Clipboard.Clear
End Sub

Private Sub tmr_Timer()
imgcam.Picture = Clipboard.GetData

  'Clipboard.Clear
 End Sub
