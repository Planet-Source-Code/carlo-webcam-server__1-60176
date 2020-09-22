Attribute VB_Name = "Modkeys"
' log module
' ------------
' this module is used to get the keys currently pressed - if one IS pressed
' GetKey returns TRUE

Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer
    
Public Type VirtualKeys ' used for table including keys
   uChar As String      ' and corresponding capitalkeys
   lChar As String
End Type

Global sKeyPressed As String

Function GetKey() As Boolean
    
    Dim tmKeys(255) As VirtualKeys
    Dim bySpec(30) As Byte
    Dim iLoop As Integer
    Dim mychar
    mychar = Chr(34)
    tmKeys(110).lChar = ",": tmKeys(32).lChar = " "
    tmKeys(110).uChar = ",": tmKeys(32).uChar = " "
    bySpec(1) = 110:  bySpec(2) = 32
        
    tmKeys(186).lChar = ";": tmKeys(192).lChar = "`"
    tmKeys(186).uChar = ":": tmKeys(192).uChar = "~"
    bySpec(3) = 186: bySpec(4) = 192
    
    tmKeys(222).lChar = "'": tmKeys(219).lChar = "["
    tmKeys(222).uChar = mychar: tmKeys(219).uChar = "{"
    bySpec(5) = 222: bySpec(6) = 219
    
    tmKeys(187).lChar = "=": tmKeys(191).lChar = "/"
    tmKeys(187).uChar = "+": tmKeys(191).uChar = "?"
    bySpec(7) = 187: bySpec(8) = 191
    
    tmKeys(189).lChar = "-": tmKeys(190).lChar = "."
    tmKeys(189).uChar = "_": tmKeys(190).uChar = ":"
    bySpec(9) = 189: bySpec(10) = 190
    
    tmKeys(188).lChar = ",": tmKeys(226).lChar = "<"
    tmKeys(188).uChar = ";": tmKeys(226).uChar = ">"
    bySpec(11) = 188: bySpec(12) = 226
        
    tmKeys(106).lChar = "*": tmKeys(107).lChar = "+"
    tmKeys(106).uChar = "*": tmKeys(107).uChar = "+"
    bySpec(13) = 106: bySpec(14) = 107
        
    tmKeys(111).lChar = "/": tmKeys(109).lChar = "-"
    tmKeys(111).uChar = "/": tmKeys(109).uChar = "-"
    bySpec(15) = 111: bySpec(16) = 109
        
    tmKeys(220).lChar = "\": tmKeys(221).lChar = "]"
    tmKeys(220).uChar = "°": tmKeys(221).uChar = "}"
    bySpec(17) = 220: bySpec(18) = 221
        
    tmKeys(46).lChar = " [Del] ": tmKeys(45).lChar = " [Insert] "
    tmKeys(46).uChar = " [Del] ": tmKeys(45).uChar = " [Insert] "
    bySpec(19) = 46: bySpec(20) = 45
        
    tmKeys(36).lChar = " [Home] ": tmKeys(35).lChar = " [End] "
    tmKeys(36).uChar = " [Home] ": tmKeys(35).uChar = " [End] "
    bySpec(21) = 36: bySpec(22) = 35
        
    tmKeys(38).lChar = " [Up] ": tmKeys(40).lChar = " [Down] "
    tmKeys(38).uChar = " [Up] ": tmKeys(40).uChar = " [Down] "
    bySpec(23) = 38: bySpec(24) = 40
    
    tmKeys(37).lChar = " [Left] ": tmKeys(39).lChar = " [Rigth] "
    tmKeys(37).uChar = " [Left] ": tmKeys(39).uChar = " [Right] "
    bySpec(25) = 37: bySpec(26) = 39
        
    tmKeys(13).lChar = " [Return]": tmKeys(8).lChar = " [BCKSP] "
    tmKeys(13).uChar = " [RETN]": tmKeys(8).uChar = " [BCKSP] "
    bySpec(27) = 13: bySpec(28) = 8
        
    tmKeys(33).lChar = " [Pg Up]": tmKeys(34).lChar = " [Pg Down] "
    tmKeys(33).uChar = " [PG UP]": tmKeys(34).uChar = " [Pg Down] "
    bySpec(29) = 33: bySpec(30) = 34
        
    For iLoop = 65 To 90    ' check for A..Z and a..z
        If GetAsyncKeyState(iLoop) = -32767 Then
            sKeyPressed = IIf(CapsOn, Chr(iLoop), LCase(Chr(iLoop)))
            GoTo KeyFound
        End If
    Next iLoop
         
    For iLoop = 1 To 30     ' check if any special key is pressed
        If GetAsyncKeyState(bySpec(iLoop)) = -32767 Then
            sKeyPressed = IIf(CapsOn, tmKeys(bySpec(iLoop)).uChar, tmKeys(bySpec(iLoop)).lChar)
            GoTo KeyFound
        End If
    Next iLoop
        
    For iLoop = 0 To 9      ' check for other spec-keys
        If GetAsyncKeyState(96 + iLoop) = -32767 Then
            sKeyPressed = IIf(CapsOn, Right(Str(iLoop), 1), Right(Str(iLoop), 1))
            GoTo KeyFound
        End If
    Next iLoop
    
    For iLoop = 48 To 58    ' check for numbers 0..9
 
        If GetAsyncKeyState(iLoop) = -32767 Then

            If CapsOn Then
                If iLoop = 48 Then sKeyPressed = ")"
                If iLoop = 49 Then sKeyPressed = "!"
                If iLoop = 50 Then sKeyPressed = "@" 'Chr(34)
                If iLoop = 51 Then sKeyPressed = "#"
                If iLoop = 52 Then sKeyPressed = "$"
                If iLoop = 53 Then sKeyPressed = "%"
                If iLoop = 54 Then sKeyPressed = "^"
                If iLoop = 55 Then sKeyPressed = "&"
                If iLoop = 56 Then sKeyPressed = "*"
                If iLoop = 57 Then sKeyPressed = "("
                If iLoop = 58 Then sKeyPressed = ")"
            Else
                sKeyPressed = Chr(iLoop)
            End If
            GoTo KeyFound
        End If
    
    Next iLoop
       
    GetKey = False
    Exit Function

KeyFound:
    GetKey = True
End Function

Function CapsOn() As Boolean           ' return CapsStatus
    
    CapsOn = CBool(GetKeyState(vbKeyCapital) And 1)
    CapsOn = CapsOn Or CBool(GetAsyncKeyState(vbKeyShift))
    
End Function




