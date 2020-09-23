VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CharWidth = 24
'CharHeight = 32
Public dx As DX7

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'exits the program
    If KeyCode = vbKeyEscape Then
        dx.ShutDown
    End If
End Sub

Private Sub Form_Load()
  
    Set dx = New DX7
    
    PlaceScreenX = (Form1.resoluteX / 2) - 160
    PlaceScreenY = (Form1.resoluteY / 2) - 160
    
    dx.InitDX Me.hWnd, Form1.resoluteX, Form1.resoluteY, Form1.colordepth
    
    'load the charset into a DD surface
    dx.LoadFromFreeImage App.Path & "\Gfx\Charset.png", surChars.Pic, True, 0, 0

    'load the tileset into a DD surface
    dx.LoadFromFreeImage App.Path & "\Gfx\TileSet.png", surTiles.Pic, True, 0, 0
    
    'load the messagebox
    dx.LoadFromFreeImage App.Path & "\Gfx\messagebox.png", surMessageBox.Pic, True, 0, 0
    
    Set surScroll.Surface = dx.MakeNormSurface(640, 640, True, 0, 0)
    
    scrRect.Bottom = 320
    scrRect.Right = 320
    'make the surfaces that will be the buffers
    Set DBuffer2.Surface = dx.MakeNormSurface(scrRect.Right, scrRect.Bottom, True, 0, 0)
    Set DBuffer3.Surface = dx.MakeNormSurface(scrRect.Right, scrRect.Bottom, True, 0, 0)
    DBuffer2.Surface.BltColorFill rScreen, vbBlack
    DBuffer3.Surface.BltColorFill rScreen, vbBlack
    
    
    'load sound files
    Set Segment = Loader.LoadSegment(App.Path & "\Music\forest.mid")
    

    'Set up some variables
    WaterUp = True
  
    'Open the town map and create our map array
    Open App.Path & "\Maps\townmap.map" For Binary As #1
        Get #1, , MaxRow
        Get #1, , MaxCol

        ReDim Map(0 To MaxRow, 0 To MaxCol)

        Get #1, , Map
    Close #1
    TownMap = Map
    'Check to see if there are any animated tiles in that map
    Call CheckForAnimatedTiles
    'Set up our players position
    Player.Pushing = False
    Player.Attacking = False
    Player.Xrow = 10
    Player.Ycol = 14
    Player.Xpos = Player.Xrow * 16
    Player.Ypos = Player.Ycol * 16
    Player.CharIDX = 24
    Player.CharIDY = 64
    Player.OldXrow = 19
    Player.OldYcol = 19
    Player.Frame = 1
    'Update the collision boundaries of the player
    Call UpdateBoundaries
    'make sure we're at the first (top left) screen of the game
    ScreenX = 0
    ScreenY = 0
    
    MapCol = 10
    MapRow = 10
    
    Segment.SetLoopPoints 0, 0
    'If they are both set to 0 then it will loop the entire segment.
    Segment.SetRepeats 100
    '<Start playing music here>
    Performance.PlaySegment Segment, 0, 0
    'set the volume
    SetVolume (60)
    NumLetter = 1
    Call MainLoop
End Sub

Public Sub MainLoop()
    Dim oldtime As Long
    Dim OldTime2 As Long
    Dim FPScount As Long
    Dim FPS As Integer
    Dim FPSS As Long
    Dim Nap As Long
   
    'Store the map into a second backbuffer, so we dont have to redraw every single time
    Call RedrawTileMap(MapRow, MapCol, "All", DBuffer3.Surface)
    
    FPScount = dx.dx.TickCount
    oldtime = dx.dx.TickCount
    FlipTime = dx.dx.TickCount
    AnimTime = dx.dx.TickCount
    AnimGTime = dx.dx.TickCount
    
    dx.backbuffer.SetForeColor vbWhite
    DBuffer2.Surface.SetForeColor vbWhite
    
    Do
        'limit the FPS to around 33
        oldtime = timeGetTime() + 31
        Call CheckMovement 'Check for player movement
        Call CheckforScroll
        Call AnimateTiles 'Animate the tiles that should be animated
        Call UpdateCharFrame
        'fill the screen with green
        dx.backbuffer.BltColorFill rScreen, vbBlack
            
        dx.BltSurface DBuffer2.Surface, DBuffer3.Surface, 0, 0, 0, 0, 320, 320
        'draw the player on the buffer
        Call DrawChar
        'now draw the upper layers (leaves etc..) which will appear over the player....these steps are important in order for layering to work
        Call RedrawTileMap(MapRow, MapCol, "Upper", DBuffer2.Surface)
        'check to see if the player went on any tile that may cause an event to happen (such as loading a new map)
        Call CheckCurTile(Player.Xrow, Player.Ycol)
        dx.BltSurface dx.backbuffer, DBuffer2.Surface, PlaceScreenX, PlaceScreenY, 0, 0, 320, 320
        'display the FPS and the players X and Y
        dx.backbuffer.DrawText 500, 0, "FPS: " & FPSS & vbCrLf & "X: " & Player.Xrow & "  Y: " & Player.Ycol, False
            
        dx.primary.Flip Nothing, DDFLIP_WAIT
        'keep count of the FPS
        If dx.dx.TickCount - FPScount > 1000 Then
            FPScount = dx.dx.TickCount
            FPSS = FPS
            FPS = 0
        End If
        
        FPS = FPS + 1
        While oldtime > timeGetTime()   'delay until time runs out
            DoEvents
        Wend
    Loop
End Sub

Public Sub DrawChar()
    'draw the player to the screen
    dx.BltObject DBuffer2.Surface, surChars.Pic, Player.Xpos, Player.Ypos, Player.CharIDY, Player.CharIDY, Player.CharIDX, 24, 32
End Sub

Public Sub RedrawTileMap(ByVal Xoffset As Single, ByVal YOffset As Single, ByVal layer As String, Location As DirectDrawSurface7)
    Dim X As Integer
    Dim Y As Integer
    Dim x2, y2 As Integer
    'This is where i draw what the user can see currently. I only show a 20 by 20 area of the 200 by 200 map (town map).
    'If I know what each map tile is and all i have to do is just determine which ones the user can currently see and draw them.
    x2 = Xoffset - 11
    y2 = YOffset - 10
    For X = 0 To 19
        x2 = x2 + 1
        For Y = 0 To 19
            'Draw all the layers
            If layer = "All" Then
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(0).TileY, Map(x2, y2).layer(0).TileX, 16, 16
                Call CheckOtherLayers(x2, y2, X, Y, Location)
            ElseIf layer = "Lower" Then 'Only draw the lower layer tiles *Layer 0 and 1 are considered lower layers* *layer 2 is the tiles in which the player can go under
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(0).TileY, Map(x2, y2).layer(0).TileX, 16, 16
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(1).TileY, Map(x2, y2).layer(1).TileX, 16, 16
            Else
                If Map(x2, y2).layer(2).Used = 1 Then
                    dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(2).TileY, Map(x2, y2).layer(2).TileX, 16, 16
                End If
            End If
            y2 = y2 + 1
        Next Y
        y2 = YOffset - 10
    Next X
    
End Sub

Private Sub CheckOtherLayers(ByVal x2 As Integer, ByVal y2 As Integer, ByVal X As Integer, ByVal Y As Integer, Location As DirectDrawSurface7)
    'draw the layer1 and layer2 tiles
    If Map(x2, y2).layer(1).Used = 1 Then
        dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(1).TileY, Map(x2, y2).layer(1).TileX, 16, 16
    End If
                    
    If Map(x2, y2).layer(2).Used = 1 Then
        dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(2).TileY, Map(x2, y2).layer(2).TileX, 16, 16
    End If
End Sub

Public Sub CheckMovement()
    'checks to see if the player pressed and movement keys and at the same time will check for collisions between tiles
    If CurEvent = False Then
        If dx.CheckKey(KEY_LEFT) Then
            Player.CharIDY = 96 'change the ID of the player, so he is now facing to the left
            Player.OldIDY = 96
            
            Player.Xpos = Player.Xpos - 2 'first move him to the left....
            If CollisionDetec = True Then
                Player.Xpos = Player.Xpos + 2 'and if he collided then move him back to where he was
            End If
            Player.Moving = True
        ElseIf dx.CheckKey(KEY_RIGHT) Then
            Player.CharIDY = 32
            Player.OldIDY = 32
            
            Player.Xpos = Player.Xpos + 2
            If CollisionDetec = True Then
                Player.Xpos = Player.Xpos - 2
            End If
            Player.Moving = True
        ElseIf dx.CheckKey(KEY_UP) Then
            Player.CharIDY = 0
            Player.OldIDY = 0
            
            Player.Ypos = Player.Ypos - 2
            If CollisionDetec = True Then
                Player.Ypos = Player.Ypos + 2
            End If
            Player.Moving = True
        ElseIf dx.CheckKey(KEY_DOWN) Then
            Player.CharIDY = 64
            Player.OldIDY = 64
            
            Player.Ypos = Player.Ypos + 2
            If CollisionDetec = True Then
                Player.Ypos = Player.Ypos - 2
            End If
            Player.Moving = True
        
        Else
            Player.Moving = False
            Player.Pushing = False
            finAttack = True
           
        End If
    End If
    Call CheckCharPos 'checks his position
    Call CheckAction 'checks to see if he might have pressed an action key (such as to examine a chest etc...)
End Sub

Public Sub CheckAction()
    'If the player pressed enter check if he is near any tile that has an event added to it
    If dx.CheckKey(KEY_RETURN) Then
        Select Case Player.CharIDY
        'these "cases" check the corresponding tile infront of the player
        Case 0 'if he is facing up...
            If CurEvent = True And DoneWithMessage = True And AbletoPress = True Then 'make sure the message is complete
                AbletoPress = False
                NumLetter = 1
                CurEvent = False
            ElseIf CurEvent = False And AbletoPress = True Then 'stop the player from moving, dont let him press enter until the message is done
                Player.Moving = False
                AbletoPress = False
                DirAction = ActionUp
                Call CheckCurTile(Player.Xrow, Player.Ycol - 1) 'check the tile above the player
            End If
        Case 32 'if he is facing right...
            If CurEvent = True And DoneWithMessage = True And AbletoPress = True Then
                AbletoPress = False
                NumLetter = 1
                CurEvent = False
            ElseIf CurEvent = False And AbletoPress = True Then
                Player.Moving = False
                AbletoPress = False
                DirAction = ActionRight
                Call CheckCurTile(Player.Xrow + 1, Player.Ycol)
            End If
        Case 64 'if he is facing down..
            If CurEvent = True And DoneWithMessage = True And AbletoPress = True Then
                AbletoPress = False
                NumLetter = 1
                CurEvent = False
            ElseIf CurEvent = False And AbletoPress = True Then
                Player.Moving = False
                AbletoPress = False
                DirAction = ActionDown
                Call CheckCurTile(Player.Xrow, Player.Ycol + 1)
            End If
        Case 96 'if he is facing to the left..
            If CurEvent = True And DoneWithMessage = True And AbletoPress = True Then
                AbletoPress = False
                NumLetter = 1
                CurEvent = False
            ElseIf CurEvent = False And AbletoPress = True Then
                Player.Moving = False
                AbletoPress = False
                DirAction = ActionLeft
                Call CheckCurTile(Player.Xrow - 1, Player.Ycol)
            End If
        End Select
    Else
        AbletoPress = True
    End If
End Sub
Public Sub CheckCharPos()
    'This determines if the character has moved into another "Tile" and i update his position in the map
    Player.Xrow = (RoundPos(Player.Xpos + 12) \ 16) + ScreenX
    Player.Ycol = (RoundPos(Player.Ypos + 14) \ 16) + ScreenY
End Sub

Private Sub CheckCurTile(X As Integer, Y As Integer)
    'this sub will check if he is on a tile that will cause an event to happen
    If Map(X, Y).Linked = 1 Then
        Call LoadMap(Map(Player.Xrow, Player.Ycol).LinkInfo.MapName, Player.Xrow, Player.Ycol) 'load the map if he stepped on a tile that links or takes him to another map
    ElseIf Map(X, Y).Event.Type <> 0 Or CurEvent = True Then
        CurEvent = True
        Call MessageHandler 'call the sub that shows the message to the user
    End If
End Sub

Public Sub UpdateCharFrame()
    If dx.dx.TickCount - FlipTime > 75 Then
        FlipTime = dx.dx.TickCount
        'Draw the next frame of animation for the character moving
        If Player.Frame > 2 Then
            Player.Frame = 0
        ElseIf Player.Moving = False Then
            Player.CharIDX = 24
        Else
            Player.CharIDX = Player.Frame * 24
            Player.Frame = Player.Frame + 1
        End If
    End If
    
End Sub

Public Sub AnimateTiles()
    If UBound(AnimatedWaterTiles) > 0 And dx.dx.TickCount - AnimTime > 350 Then
        Dim X As Integer
        
        'Draw the next frame of the animated water tile
        AnimTime = dx.dx.TickCount
        
        For X = LBound(AnimatedWaterTiles) To UBound(AnimatedWaterTiles)
            'determine which tile needs to be drawn
            '*******************NEW ANIMATION CODE**********MUCH FASTER******************
            If Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX = 32 Then
                WaterUp = False
            ElseIf Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX = 0 Then
                WaterUp = True
            End If
            If WaterUp = True Then
                Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX = Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX + 16
            Else
                Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX = Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX - 16
            End If
            '******************************************************************************
                    
            dx.BltSurface DBuffer3.Surface, surTiles.Pic, (AnimatedWaterTiles(X).X - ScreenX) * 16, (AnimatedWaterTiles(X).Y - ScreenY) * 16, Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileY, Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(0).TileX, 16, 16
            
            'check if theres any tile ontop of the animated tile...if so then draw it
            If Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(1).Used = 1 Then
                dx.BltSurface DBuffer3.Surface, surTiles.Pic, (AnimatedWaterTiles(X).X - ScreenX) * 16, (AnimatedWaterTiles(X).Y - ScreenY) * 16, Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(1).TileY, Map(AnimatedWaterTiles(X).X, AnimatedWaterTiles(X).Y).layer(1).TileX, 16, 16
            End If
            
        Next X
        WaterFrameCount = WaterFrameCount + 1
    End If
    If UBound(AnimatedGrassTiles) > 0 And dx.dx.TickCount - AnimGTime > 500 Then
        AnimGTime = dx.dx.TickCount

        For X = LBound(AnimatedGrassTiles) To UBound(AnimatedGrassTiles)
            'determine which tile needs to be drawn
             '****************ANIMATION CODE*****************MUCH FASTER*************************
            If Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY = 112 Then
                GrassUp = False
            ElseIf Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY = 64 Then
                GrassUp = True
            End If
            If GrassUp = True Then
                Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY = Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY + 16
            Else
                Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY = 64
                'Map(AnimatedGrassTiles(x).x, AnimatedGrassTiles(x).y).layer(0).TileY = Map(AnimatedGrassTiles(x).x, AnimatedGrassTiles(x).y).layer(0).TileY - 16
            End If
            '*************************************************************************************

            dx.BltSurface DBuffer3.Surface, surTiles.Pic, (AnimatedGrassTiles(X).X - ScreenX) * 16, (AnimatedGrassTiles(X).Y - ScreenY) * 16, Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileY, Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(0).TileX, 16, 16

            'check if theres any tile ontop of the animated tile...if so then draw it
            If Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(1).Used = 1 Then
                dx.BltSurface DBuffer3.Surface, surTiles.Pic, (AnimatedGrassTiles(X).X - ScreenX) * 16, (AnimatedGrassTiles(X).Y - ScreenY) * 16, Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(1).TileY, Map(AnimatedGrassTiles(X).X, AnimatedGrassTiles(X).Y).layer(1).TileX, 16, 16
            End If

        Next X
        GrassFrameCount = GrassFrameCount + 1
    End If
End Sub

Public Sub CheckforScroll()
    'This determines which way we scroll the map
    If Player.Xrow - Player.OldXrow = 1 And Player.Xrow <> 200 Then
        Call ScrollMapRight
    ElseIf Player.OldXrow - Player.Xrow = 20 And Player.OldXrow <> 19 Then
        Call ScrollMapLeft
    ElseIf Player.Ycol - Player.OldYcol = 1 And Player.Ycol <> 200 Then
        Call ScrollMapDown
    ElseIf Player.OldYcol - Player.Ycol = 20 And Player.Ycol <> -1 Then
        Call ScrollMapUp
    End If
End Sub

Public Sub ScrollPixel(ByVal Xoffset As Single, ByVal YOffset As Single, ByVal layer As String, ByVal Location As DirectDrawSurface7)
    Dim X As Integer
    Dim Y As Integer
    Dim x2, y2 As Integer
    'This is where i draw what the user can see currently. I only show a 20 by 20 area of the 200 by 200 map.
    'If I know what each map tile is and all i have to do is just determine which ones the user can currently see and draw them.
    'VERY similar to the redrawtilemap sub...but this draws the next "screen" the player will see as it scrolls
    x2 = Xoffset - 20
    y2 = YOffset - 19
    For X = 0 To 39
        x2 = x2 + 1
        For Y = 0 To 39
            'Draw all the layers
            If layer = "All" Then
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(0).TileY, Map(x2, y2).layer(0).TileX, 16, 16
                Call CheckOtherLayers(x2, y2, X, Y, Location)
            ElseIf layer = "Lower" Then 'Only draw the lower layer tiles *Layer 0 and 1 are considered lower layers* *layer 2 is the tiles in which the player can go under
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(0).TileY, Map(x2, y2).layer(0).TileX, 16, 16
                dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(1).TileY, Map(x2, y2).layer(1).TileX, 16, 16
            Else
                If Map(x2, y2).layer(2).Used = 1 Then
                    dx.BltSurface Location, surTiles.Pic, X * 16, Y * 16, Map(x2, y2).layer(2).TileY, Map(x2, y2).layer(2).TileX, 16, 16
                End If
            End If
            y2 = y2 + 1
        Next Y
        y2 = YOffset - 19
    Next X
End Sub








