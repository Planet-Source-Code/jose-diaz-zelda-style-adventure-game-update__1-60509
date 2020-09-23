Attribute VB_Name = "GeneralCode"

'-----------------------------------------------------------------------------------------------------------
'The next four subs scroll the map in a specific direction
'they are basically the same sub, just some variables have been changed and instead of adding, we might subtract
'-----------------------------------------------------------------------------------------------------------
Public Sub ScrollMapRight()
    curScrollRow = 0
    
    'These next lines of code will do these steps
    '--hold the next screen that the player will see in a buffer
    '---only show that screen a little bit at a time to the player...specifically 4 pixels at a time
    '----draw the character on top of the screen while moving him backwards (or forwards, if we were scrolling in a different direction)
    '-----and once we have shown 16 more pixels to the user...update the current map's row by adding 1 (will subtract if we were scrolling to the left)
    '------keep looping until we have completely shown the next screen to the player
    Call frmMain.ScrollPixel(MapRow + 9, MapCol + 9, "All", surScroll.Surface)
    Do
        'this section scrolls from one screen to the next at an interval of 4 pixels...makes it look smoother
        frmMain.dx.BltSurface DBuffer2.Surface, surScroll.Surface, 0, 0, 0, curScrollRow, 320, 320
        Call frmMain.DrawChar
        frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
        frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
        curScrollRow = curScrollRow + 4
        If curScrollRow Mod 16 = 0 Then MapRow = MapRow + 1

        Player.Xpos = Player.Xpos - 4
    Loop Until curScrollRow >= 320
    
    Player.Xpos = 0
    'This is just to adjust the character's position, so that he doesnt move too far ahead or stay too far back
    Player.OldXrow = Player.Xrow + 19
    ScreenX = ScreenX + 20
    'check to see if there any tiles that are supposed to animate
    Call CheckForAnimatedTiles
    'show the player the new screen
    Call frmMain.RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)


End Sub

Public Sub ScrollMapLeft()
    curScrollRow = 320
    
    
    Call frmMain.ScrollPixel(MapRow - 11, MapCol + 9, "All", surScroll.Surface)
    Do
        frmMain.dx.BltSurface DBuffer2.Surface, surScroll.Surface, 0, 0, 0, curScrollRow, 320, 320
        Call frmMain.DrawChar
        frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
        frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
        curScrollRow = curScrollRow - 4
        If curScrollRow Mod 16 = 0 Then MapRow = MapRow - 1
        
        Player.Xpos = Player.Xpos + 4
    Loop Until curScrollRow <= 0
    
    Player.Xpos = 320 - 16
    Player.OldXrow = Player.Xrow
    
    ScreenX = ScreenX - 20
    Call CheckForAnimatedTiles
    Call frmMain.RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)
End Sub

Public Sub ScrollMapDown()
    curScrollRow = 0
    
    
    Call frmMain.ScrollPixel(MapRow + 9, MapCol + 9, "All", surScroll.Surface)
    
    Do
        frmMain.dx.BltSurface DBuffer2.Surface, surScroll.Surface, 0, 0, curScrollRow, 0, 320, 320
        Call frmMain.DrawChar
        frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
        frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
        curScrollRow = curScrollRow + 4
        If curScrollRow Mod 16 = 0 Then MapCol = MapCol + 1
            
        Player.Ypos = Player.Ypos - 4
    Loop Until curScrollRow >= 320
    
    Player.Ypos = 0
    Player.OldYcol = Player.Ycol + 19
    ScreenY = ScreenY + 20
    
    Call CheckForAnimatedTiles
    Call frmMain.RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)
    
End Sub

Public Sub ScrollMapUp()
    curScrollRow = 320
    
    
    Call frmMain.ScrollPixel(MapRow + 9, MapCol - 11, "All", surScroll.Surface)
    
    Do
        frmMain.dx.BltSurface DBuffer2.Surface, surScroll.Surface, 0, 0, curScrollRow, 0, 320, 320
        Call frmMain.DrawChar
        frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
        frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
        curScrollRow = curScrollRow - 4
        If curScrollRow Mod 16 = 0 Then MapCol = MapCol - 1
            
        Player.Ypos = Player.Ypos + 4
    Loop Until curScrollRow <= 0
    
    Player.Ypos = 320 - 22
    Player.OldYcol = Player.Ycol
    
    ScreenY = ScreenY - 20
    Call CheckForAnimatedTiles
    Call frmMain.RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)
End Sub

Public Function RoundPos(Pos As Single)
    'Make the figure a whole number which is divisible by 16
    RoundPos = Int(Pos / 16) * 16
End Function

Public Sub LoadMap(ByVal MapName As String, ByVal TileX As Integer, ByVal TileY As Integer)
    Dim songToLoad As String
    
    SetVolume (5)
    Call Curtains("Close") 'close the curtains...as u can see lol
   
    'update game properties and the position of the player on the new map
    EndGame = False
    Player.Pushing = False
    Player.Xrow = Map(TileX, TileY).LinkInfo.GotoCoord.X
    Player.Ycol = Map(TileX, TileY).LinkInfo.GotoCoord.Y
    Player.CharIDY = Map(TileX, TileY).LinkInfo.PlayerDir
    Player.CharIDX = 24
    Player.Frame = 0
    Player.Moving = False
    
    'Update Boundaries
    ScreenX = Fix(Player.Xrow / 20) * 20
    ScreenY = Fix(Player.Ycol / 20) * 20
  
    MapCol = 10 + ScreenY
    MapRow = 10 + ScreenX
    
    Player.OldXrow = 19 + ScreenX
    Player.OldYcol = 19 + ScreenY
    
    'stop the current music
    Performance.Stop Segment, Nothing, 0, 0
    'determine the real X and Y position of the player
    If Player.Xrow >= 20 And Player.Ycol >= 20 Then
        Player.Xpos = Abs(((Player.Xrow * 16) - 3) - (320 * (ScreenX \ 20)))
        Player.Ypos = Abs((Player.Ycol * 16) - (320 * (ScreenY \ 20)))
    ElseIf Player.Xrow >= 20 And Player.Ycol < 20 Then
        Player.Xpos = Abs(((Player.Xrow * 16) - 3) - (320 * (ScreenX \ 20)))
        Player.Ypos = Player.Ycol * 16
    ElseIf Player.Xrow < 20 And Player.Ycol >= 20 Then
        Player.Ypos = Abs((Player.Ycol * 16) - (320 * (ScreenY \ 20)))
         Player.Xpos = (Player.Xrow * 16) - 3
    Else
        Player.Xpos = (Player.Xrow * 16) - 3
         Player.Ypos = Player.Ycol * 16
    End If
    
    
    Player.CharIDX = 24
    opentime = frmMain.dx.dx.TickCount
    'clear the current array
    Erase Map
    If LCase(MapName) <> "town.map" Then
        Open App.Path & "\Maps\" & MapName For Binary Access Read Lock Read Write As #1
            Get #1, , MaxRow
            Get #1, , MaxCol
            'load the new array with the new maps properties
            ReDim Map(0 To MaxRow, 0 To MaxCol)
            
            Get #1, , Map
        Close #1
    Else
        Map = TownMap
    End If
        
    'find out which song we have to load...by looking at the first 4 letters of the map. If someone was to tamper with the names of the map it would screw this up(naturally)
    songToLoad = Mid(MapName, 1, 4)
    If songToLoad = "hous" Then
        Set Segment = Loader.LoadSegment(App.Path & "\Music\house.mid")
    ElseIf songToLoad = "shop" Then
        Set Segment = Loader.LoadSegment(App.Path & "\Music\shop.mid")
        Segment.SetStartPoint (10000)
    Else
        Set Segment = Loader.LoadSegment(App.Path & "\Music\forest.mid")
        Segment.SetStartPoint (5000)
    End If
    'If the loop points are both set to 0 then it will loop the entire segment.
    Segment.SetLoopPoints 0, 0
    Segment.SetRepeats 100 'repeat the song 100 times...might increase as i get deeper into the game
    Performance.PlaySegment Segment, 0, 0
    SetVolume (70) 'set the volume
    
    Call CheckForAnimatedTiles
    Call frmMain.RedrawTileMap(MapCol, MapRow, "Lower", DBuffer2.Surface)
    Call frmMain.DrawChar
    Call frmMain.RedrawTileMap(MapCol, MapRow, "Upper", DBuffer2.Surface)
    
    Call frmMain.RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)
    Call Curtains("Open") 'open the curtains
    
    frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
    
    
    
End Sub

Public Sub CheckForAnimatedTiles()

    '--------------------------------------------
    'Checks for animated water and grass tiles
    '--------------------------------------------
    Erase AnimatedWaterTiles
    Erase AnimatedGrassTiles
    Dim X, Y As Integer
    
    Dim animCount As Integer
    Dim animGcount As Integer
    ReDim AnimatedWaterTiles(animCount)
    ReDim AnimatedGrassTiles(animGcount)
    For X = ScreenX To ScreenX + 19
        For Y = ScreenY To ScreenY + 19
            If Map(X, Y).Animated = True Then 'check if any tile on the screen is animated
                If (Map(X, Y).layer(0).TileX = 64 Or Map(X, Y).layer(0).TileX = 80) Then
                    ReDim Preserve AnimatedGrassTiles(animGcount)
                    AnimatedGrassTiles(animGcount).X = X
                    AnimatedGrassTiles(animGcount).Y = Y
                    animGcount = animGcount + 1
                Else
                    ReDim Preserve AnimatedWaterTiles(animCount)
                    AnimatedWaterTiles(animCount).X = X
                    AnimatedWaterTiles(animCount).Y = Y
                    animCount = animCount + 1
                End If
            End If
        Next Y
    Next X
End Sub

Public Function CollisionDetec() As Boolean
    
    Call UpdateBoundaries
    
    If Player.CharIDY = 0 Then
        'check if his topleft corner collided with any tile
        Player.TR = (((Player.TopLeft.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.TopLeft.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
        'check if his topright corner collided with any tile
        Player.TR = (((Player.TopRight.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.TopRight.Y / 16) * 16) \ 16) + ScreenY
    
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
    ElseIf Player.CharIDY = 64 Then
        'check if his bottomright corner collided with any tile
        Player.TR = (((Player.BottomR.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.BottomR.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
        
        'check if his bottomleft corner collided with any tile
        Player.TR = (((Player.BottomL.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.BottomL.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
    ElseIf Player.CharIDY = 32 Then
        'check if his topright corner collided with any tile
        Player.TR = (((Player.TopRight.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.TopRight.Y / 16) * 16) \ 16) + ScreenY
    
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
        
        'check if his bottomright corner collided with any tile
        Player.TR = (((Player.BottomR.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.BottomR.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
    Else
        'check if his bottomleft corner collided with any tile
        Player.TR = (((Player.BottomL.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.BottomL.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
        
        'check if his topleft corner collided with any tile
        Player.TR = (((Player.TopLeft.X / 16) * 16) \ 16) + ScreenX
        Player.TC = (((Player.TopLeft.Y / 16) * 16) \ 16) + ScreenY
        
        If Map(Player.TR, Player.TC).Blocked = True Then
            CollisionDetec = True
            Exit Function
        End If
    End If

End Function

Public Sub UpdateBoundaries()
    'set the boundaries of the player in another set of variables...to be used with collision detection
    Player.TopLeft.X = Player.Xpos + 4
    Player.TopLeft.Y = Player.Ypos + 14
    Player.TopRight.X = Player.Xpos + 15
    Player.TopRight.Y = Player.Ypos + 14
    Player.BottomL.X = Player.Xpos + 4
    Player.BottomL.Y = Player.Ypos + 26
    Player.BottomR.X = Player.Xpos + 15
    Player.BottomR.Y = Player.Ypos + 26
End Sub

Public Sub Curtains(ByVal TypeOfC As String)
    Dim openedCur As Boolean
    Dim opentime As Long
    
    frmMain.dx.backbuffer.SetForeColor RGB(0, 0, 0)
    frmMain.dx.backbuffer.SetFillColor RGB(0, 0, 0)
    frmMain.dx.backbuffer.SetFillStyle (0)
    If TypeOfC = "Close" Then
        frmMain.dx.backbuffer.DrawBox PlaceScreenX, PlaceScreenY, PlaceScreenX + 160, PlaceScreenY + 320
        frmMain.dx.backbuffer.DrawBox PlaceScreenX + 160, PlaceScreenY, PlaceScreenX + 320, PlaceScreenY + 320
        frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
    Else
        'This will open and close the curtains at the same rate....i think, lol
        Do
            Dim X As Integer
            Dim j As Integer
            j = 160
            For X = 160 To 1 Step -6
                'slowly open the curtains by drawing 2 boxes and separating them from each other at the same rate
                frmMain.dx.BltSurface DBuffer2.Surface, DBuffer3.Surface, 0, 0, 0, 0, 320, 320
                
                Call frmMain.DrawChar
                Call frmMain.RedrawTileMap(MapCol, MapRow, "Upper", DBuffer2.Surface)
               
                frmMain.dx.BltSurface frmMain.dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
                frmMain.dx.backbuffer.DrawBox PlaceScreenX, PlaceScreenY, PlaceScreenX + X, PlaceScreenY + 320
                
                frmMain.dx.backbuffer.DrawBox PlaceScreenX + j, PlaceScreenY, PlaceScreenX + 320, PlaceScreenY + 320
                frmMain.dx.primary.Flip Nothing, DDFLIP_WAIT
                
                j = j + 6
            Next X
            openedCur = True
            
        Loop Until openedCur = True
    End If
    frmMain.dx.backbuffer.SetForeColor vbWhite
End Sub

Public Sub MessageHandler()
    'show the messagebox pic to the user
    frmMain.dx.BltSurface DBuffer2.Surface, surMessageBox.Pic, 28, 245, 0, 0, 263, 71
                
    lngDC = DBuffer2.Surface.GetDC
    'set the font in which we write the text with
    SetFont lngDC, "ManaSpace Regular", 13
    Select Case DirAction
        'depending the direction the player was facing...we make sure we know which way the tile he wants to interact with is facing
        Case ActionUp
            Call ShowMessage(Player.Xrow, Player.Ycol - 1)
        Case ActionRight
            Call ShowMessage(Player.Xrow + 1, Player.Ycol)
        Case ActionDown
            Call ShowMessage(Player.Xrow, Player.Ycol + 1)
        Case ActionLeft
            Call ShowMessage(Player.Xrow - 1, Player.Ycol)
    End Select
    
End Sub

Public Sub ShowMessage(X As Integer, Y As Integer)
    'these next lines will show the user the message....there can be 3 lines of text in one message box and the next couple lines are really messy so i wont bother explaining them since im gonna end up rewriting this section again...
    Dim length As Integer
    Dim phrase As String
    Dim phrase2 As String
    Dim letter As String
    Dim phrase3 As String
    length = Len(Map(X, Y).Event.Message)
    If NumLetter <= length Then
        If NumLetter > 27 And NumLetter < 56 Then
            phrase = Mid(Map(X, Y).Event.Message, 1, 27)
            phrase2 = Mid(Map(X, Y).Event.Message, 28, NumLetter - 28)
            ShowText phrase, 50, 260, lngDC, vbWhite
            ShowText phrase2, 50, 275, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            letter = Mid(Map(X, Y).Event.Message, NumLetter, 2)
            If letter <> "  " Then Beep 1000, 6
            NumLetter = NumLetter + 1
            DoneWithMessage = False
        ElseIf NumLetter > 55 And NumLetter < 83 Then
            phrase = Mid(Map(X, Y).Event.Message, 1, 27)
            phrase2 = Mid(Map(X, Y).Event.Message, 28, 27)
            phrase3 = Mid(Map(X, Y).Event.Message, 56, NumLetter - 56)
            ShowText phrase, 50, 260, lngDC, vbWhite
            ShowText phrase2, 50, 275, lngDC, vbWhite
            ShowText phrase3, 50, 290, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            letter = Mid(Map(X, Y).Event.Message, NumLetter, 2)
            If letter <> "  " Then Beep 1000, 6
            NumLetter = NumLetter + 1
            DoneWithMessage = False
        Else
            phrase = Mid(Map(X, Y).Event.Message, 1, NumLetter)
            ShowText phrase, 50, 260, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            letter = Mid(Map(X, Y).Event.Message, NumLetter, 2)
            If letter <> "  " Then Beep 1000, 6
            NumLetter = NumLetter + 1
            DoneWithMessage = False
        End If
                
    Else
        If length > 55 And length < 83 Then
            phrase = Mid(Map(X, Y).Event.Message, 1, 27)
            phrase2 = Mid(Map(X, Y).Event.Message, 28, 27)
            phrase3 = Mid(Map(X, Y).Event.Message, 56, length)
            ShowText phrase, 50, 260, lngDC, vbWhite
            ShowText phrase2, 50, 275, lngDC, vbWhite
            ShowText phrase3, 50, 290, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            DoneWithMessage = True
        ElseIf length > 27 And length < 56 Then
            phrase = Mid(Map(X, Y).Event.Message, 1, 27)
            phrase2 = Mid(Map(X, Y).Event.Message, 28, length)
            ShowText phrase, 50, 260, lngDC, vbWhite
            ShowText phrase2, 50, 275, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            DoneWithMessage = True
        Else
            phrase = Mid(Map(X, Y).Event.Message, 1, length)
            ShowText phrase, 50, 260, lngDC, vbWhite
            DBuffer2.Surface.ReleaseDC lngDC
            DoneWithMessage = True
        End If
    End If
    
End Sub
'Removes the font
Public Sub RemoveFont(lngDC As Long)

    'Returns the old font
    If lngOldFont <> 0 Then lngNewFont = SelectObject(lngDC, lngOldFont)
    DeleteObject lngNewFont
    
End Sub

'allows us to set which font to use in DX (you can use the drawtext method...but this is way more efficient since we are not locking and unlocking the surfaces everytime we need to display something
Public Sub SetFont(lngDC As Long, strFontName As String, intFontSize As Integer)

Dim nHeight As Long
Dim nWidth As Long
Dim nEscapement As Long
Dim fnWeight As Long
Dim fbItalic As Long
Dim fbUnderline As Long
Dim fbStrikeOut As Long
Dim fbCharSet As Long
Dim fbOutputPrecision As Long
Dim fbClipPrecision As Long
Dim fbQuality As Long
Dim fbPitchAndFamily As Long
Dim sFont As String

    'Sets up the new font
    sFont = strFontName
    fnWeight = FW_NORMAL
    nHeight = intFontSize
    nWidth = 0
    nEscapement = 0
    fbItalic = 0
    fbUnderline = 0
    fbStrikeOut = 0
    fbCharSet = DEFAULT_CHARSET
    fbOutputPrecision = OUT_TT_ONLY_PRECIS
    fbClipPrecision = CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS
    fbQuality = PROOF_QUALITY
    fbPitchAndFamily = TRUETYPE_FONTTYPE
    
    'Makes the new font
    lngNewFont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, sFont)
    
    'Selects the font onto the surface
    lngOldFont = SelectObject(lngDC, lngNewFont)

End Sub

Public Sub ShowText(strText As String, intX As Integer, intY As Integer, lngDC As Long, lngColour As Long, Optional lngBackColour As Long, Optional blnOpaque As Boolean)
    'Do we draw the text opaque, or transparent?
    If blnOpaque = True Then
        SetBkMode lngDC, TEXT_OPAQUE
        'If the text is opaque, the background needs a color
        SetBkColor lngDC, lngBackColour
    Else
        SetBkMode lngDC, TEXT_TRANSPARENT
    End If
    
    'Set the colour of the text
    SetTextColor lngDC, lngColour
    'Draw it!
    TextOut lngDC, intX, intY, strText, Len(strText)
End Sub

Sub SetVolume(nVolume As Byte)
'This formula allows you to specify a volume between
'0-100; similiar to a percentage
    Performance.SetMasterVolume (nVolume * 42 - 3000)
End Sub






