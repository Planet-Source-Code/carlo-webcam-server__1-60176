Attribute VB_Name = "Modcam"


Public Const WM_CAP_DRIVER_CONNECT As Long = 1034
Public Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Public Const WM_CAP_GRAB_FRAME As Long = 1084
Public Const WM_CAP_EDIT_COPY As Long = 1054
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Public Const WM_CLOSE = &H10

Public Const WM_USER = &H400
Public Const WM_CAP_START = WM_USER
Public Const WM_CAP_FILE_SAVEAS = WM_CAP_START + 23
Public Const WM_CAP_FILE_SAVEDIB = WM_CAP_START + 25
Public Const WM_CAP_FILE_SET_CAPTURE_FILE = WM_CAP_START + 20
Public Const WM_CAP_SINGLE_FRAME_OPEN = WM_CAP_START + 70
Public Const WM_CAP_SINGLE_FRAME_CLOSE = WM_CAP_START + 71
Public Const WM_CAP_SINGLE_FRAME = WM_CAP_START + 72


Private Declare Function SendMessageAsLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAsString Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long


Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public mCapHwnd As Long

Public Sub StopCam()
    SendMessage mCapHwnd, WM_CLOSE, 0, 0
    SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
End Sub

Public Sub StartCam()
    
    
 mCapHwnd = capCreateCaptureWindow("webcam", 0, 0, 0, 320, 240, hWnd, 0)
   
   ' mCapHwnd = capCreateCaptureWindow("Webcam", 0, 0, 0, 0, 0, 0, 0)
    SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
End Sub

Public Sub SetCamSize()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0
End Sub

Public Sub SetCamSource()
    SendMessage mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0
End Sub

Public Function CamToBMP(FileName As String)
    On Error GoTo error
    CamToBMP = 0
  
    
    capCaptureSingleFrameOpen mCapHwnd
    capCaptureSingleFrame mCapHwnd
    capCaptureSingleFrameClose mCapHwnd
    capFileSaveDIB mCapHwnd, FileName
       
    Exit Function
error:
    CamToBMP = 1
End Function




Public Function capCaptureSingleFrameOpen(ByVal hCapWnd As Long) As Boolean
     capCaptureSingleFrameOpen = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME_OPEN, 0&, 0&)
End Function


Public Function capCaptureSingleFrame(hCapWnd As Long)
    capCaptureSingleFrame = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME, 0&, 0&)
End Function


Public Function capCaptureSingleFrameClose(ByVal hCapWnd As Long) As Boolean
    capCaptureSingleFrameClose = SendMessageAsLong(hCapWnd, WM_CAP_SINGLE_FRAME_CLOSE, 0&, 0&)
 End Function


Public Function capFileSaveDIB(ByVal hCapWnd As Long, ByVal FilePath As String) As Boolean
    capFileSaveDIB = SendMessageAsString(hCapWnd, WM_CAP_FILE_SAVEDIB, 0&, FilePath)
End Function


