'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const ORIGINALSCREENWIDTH As Integer = 1680
    Private Const ORIGINALSCREENHEIGHT As Integer = 1050
    Private Const WINDOWMESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const CONTROL_MOVE As Integer = 61456
    Private Const WINDOWMESSAGE_CLICK_BUTTONDOWN As Integer = 161
    Private Const WINDOWCAPTION As Integer = 2
    Private Const WINDOWMESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3
    Private Const WIDTHSUBTRACTION As Integer = 16 'Probably the edges of the window
    Private Const HEIGHTSUBTRACTION As Integer = 38 'Probably the edges of the window
    Private Const FOGMAXSPEED As Integer = 25
    Private Const FOGMAXDELAY As Integer = 100
    Private Const ENDGAMEDELAYBLACKSCREEN As Integer = 250

    'Declare beginning necessary engine needs
    Private intScreenWidth As Integer = 800
    Private intScreenHeight As Integer = 600
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False
    Private rectFullScreen As Rectangle
    Private gGraphics As Graphics
    Private btmCanvas As New Bitmap(ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT) 'Original resolution of our game
    Private pntTopLeft As New Point(0, 0)
    Private intCanvasMode As Integer = 0 'Default menu screen
    Private intCanvasShow As Integer = 0 'Default, no animation
    Private strDirectory As String = AppDomain.CurrentDomain.BaseDirectory & "\"
    Private blnScreenChanged As Boolean = False

    'Menu necessary needs
    Private btmMenuBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\MenuBackground.jpg"))
    Private btmFogFrontPass1 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogFront.png"))
    Private btmFogBackPass1 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogBack.png"))
    Private intFogPass1X As Integer = -(ORIGINALSCREENWIDTH * 2)
    Private intFogFrontPass1Y As Integer = 650
    Private intFogBackPass1Y As Integer = 250
    Private pntFogFrontPass1 As New Point(intFogPass1X, intFogFrontPass1Y)
    Private pntFogBackPass1 As New Point(intFogPass1X, intFogBackPass1Y)
    Private btmFogFrontPass2 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogFront.png"))
    Private btmFogBackPass2 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogBack.png"))
    Private intFogPass2X As Integer = -((ORIGINALSCREENWIDTH * 2) * 2)
    Private intFogFrontPass2Y As Integer = 650
    Private intFogBackPass2Y As Integer = 250
    Private pntFogFrontPass2 As New Point(intFogPass2X, intFogFrontPass2Y)
    Private pntFogBackPass2 As New Point(intFogPass2X, intFogBackPass2Y)
    Private thrFog As System.Threading.Thread
    Private btmArcher As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Archer.png"))
    Private pntArcher As New Point(117, 0)
    Private btmLastStandText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\LastStand.png"))
    Private pntLastStandText As New Point(147, 833)
    Private udcAmbianceSound As clsSound

    'Menu buttons
    Private btmStartText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Start.png"))
    Private pntStartText As New Point(1081, 31)
    Private btmStartHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\StartHover.png"))
    Private pntStartHoverText As New Point(1059, 25)
    Private btmHighscoresText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Highscores.png"))
    Private pntHighscoresText As New Point(1198, 141)
    Private btmHighscoresHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\HighscoresHover.png"))
    Private pntHighscoresHoverText As New Point(1157, 131)
    Private btmStoryText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Story.png"))
    Private pntStoryText As New Point(1246, 283)
    Private btmStoryHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\StoryHover.png"))
    Private pntStoryHoverText As New Point(1222, 275)
    Private btmOptionsText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Options.png"))
    Private pntOptionsText As New Point(1207, 410)
    Private btmOptionsHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\OptionsHover.png"))
    Private pntOptionsHoverText As New Point(1175, 400)
    Private btmCreditsText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Credits.png"))
    Private pntCreditsText As New Point(1352, 536)
    Private btmCreditsHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\CreditsHover.png"))
    Private pntCreditsHoverText As New Point(1323, 527)
    Private btmVersusText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\Versus.png"))
    Private pntVersusText As New Point(284, 71)
    Private btmVersusHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\VersusHover.png"))
    Private pntVersusHoverText As New Point(256, 63)
    Private thrHoverSoundDelay As System.Threading.Thread

    'General, common uses
    Private btmBackText As New Bitmap(Image.FromFile(strDirectory & "Images\General\Back.png"))
    Private pntBackText As New Point(1439, 46)
    Private btmBackHoverText As New Bitmap(Image.FromFile(strDirectory & "Images\General\BackHover.png"))
    Private pntBackHoverText As New Point(1418, 35)

    'Options screen
    Private btmOptionsBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Options\OptionsBackground.jpg"))
    Private btmResolutionText As New Bitmap(Image.FromFile(strDirectory & "Images\Options\Resolution.png"))
    Private pntResolutionText As New Point(40, 41)
    Private btm800x600Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\800x600.png"))
    Private btmNot800x600Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\not800x600.png"))
    Private pnt800x600Text As New Point(85, 142)
    Private btm1024x768Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\1024x768.png"))
    Private btmNot1024x768Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\not1024x768.png"))
    Private pnt1024x768Text As New Point(85, 192)
    Private btm1280x800Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\1280x800.png"))
    Private btmNot1280x800Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\not1280x800.png"))
    Private pnt1280x800Text As New Point(85, 242)
    Private btm1280x1024Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\1280x1024.png"))
    Private btmNot1280x1024Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\not1280x1024.png"))
    Private pnt1280x1024Text As New Point(85, 293)
    Private btm1440x900Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\1440x900.png"))
    Private btmNot1440x900Text As New Bitmap(Image.FromFile(strDirectory & "Images\Options\not1440x900.png"))
    Private pnt1440x900Text As New Point(85, 342)
    Private btmFullscreenText As New Bitmap(Image.FromFile(strDirectory & "Images\Options\Fullscreen.png"))
    Private btmNotFullscreenText As New Bitmap(Image.FromFile(strDirectory & "Images\Options\notFullscreen.png"))
    Private pntFullscreenText As New Point(85, 391)
    Private btmSoundText As New Bitmap(Image.FromFile(strDirectory & "Images\Options\Sound.png"))
    Private pntSoundText As New Point(40, 447)
    Private intResolutionMode As Integer = 0 'Default 800x600
    Private btmSoundBar As New Bitmap(Image.FromFile(strDirectory & "Images\Options\SoundBar.png"))
    Private pntSoundBar As New Point(84, 547)
    Private btmSlider As New Bitmap(Image.FromFile(strDirectory & "Images\Options\Slider.png"))
    Private pntSlider As New Point(658, 533) '100% mark
    Private blnSliderWithMouseDown As Boolean = False
    Private btmSoundPercent As New Bitmap(87, 37)
    Private btmSound(100) As Bitmap '0 to 100
    Private pntSoundPercent As New Point(718, 553)

    'Loading screen
    Private btmLoadingBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingBackground.jpg"))
    Private btmLoadingBar0 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\0.png"))
    Private btmLoadingBar10 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\10.png"))
    Private btmLoadingBar20 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\20.png"))
    Private btmLoadingBar30 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\30.png"))
    Private btmLoadingBar40 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\40.png"))
    Private btmLoadingBar50 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\50.png"))
    Private btmLoadingBar60 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\60.png"))
    Private btmLoadingBar70 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\70.png"))
    Private btmLoadingBar80 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\80.png"))
    Private btmLoadingBar90 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\90.png"))
    Private btmLoadingBar99 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\99.png")) 'For trolling, Cory wanted it
    Private btmLoadingBar100 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\100.png"))
    Private btmLoadingBar As New Bitmap(1613, 134)
    Private pntLoadingBar As New Point(33, 883)
    Private btmLoadingText As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingText.png"))
    Private pntLoadingText As New Point(594, 899)
    Private btmLoadingStartText As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingStartText.png"))
    Private pntLoadingStartText As New Point(673, 909)
    Private btmLoadingParagraph25 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraph25.png"))
    Private btmLoadingParagraph50 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraph50.png"))
    Private btmLoadingParagraph75 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraph75.png"))
    Private btmLoadingParagraph100 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraph100.png"))
    Private btmLoadingParagraph As New Bitmap(1424, 472)
    Private thrParagraph As System.Threading.Thread
    Private pntLoadingParagraph As New Point(123, 261)
    Private thrLoading As System.Threading.Thread

    'Game screen
    Private btmGameBackground As Bitmap
    Private btmEndGame25 As Bitmap
    Private btmEndGame50 As Bitmap
    Private btmEndGame75 As Bitmap
    Private btmEndGame100 As Bitmap
    Private btmEndGame As Bitmap
    Private thrEndGame As System.Threading.Thread
    Private btmWordBar As New Bitmap(Image.FromFile(strDirectory & "Images\Words\WordBar.png"))
    Private pntWordBar As New Point(482, 27)
    Private udcCharacter As clsCharacter
    Private udcZombies(4) As clsZombie
    Private intZombieSpeed As Integer = 0 'Defaulted for now as nothing
    Private blnBackFromGame As Boolean = False
    Private blnGameOverFirstTime As Boolean = False
    Private udcGameAmbiance As clsSound
    Private blnEndingGameCantType As Boolean = False

    'Words
    Private astrWords() As String 'Used to fill with words
    Private intWordIndex As Integer = 0
    Private btmWord As Bitmap 'Defaulted
    Private pntWord As New Point(484, 65)
    Private strTheWord As String = ""
    Private strWord As String = ""

    'Highscores screen
    Private btmHighscoresBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Highscores\HighscoresBackground.jpg"))

    'Credits screen
    Private btmCreditsBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CreditsBackground.jpg"))

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes:'Do not allow a resize of the window if full screen, this happens if you double click the title, 
        '       but is prevented with this sub, prevents moving the window around too

        'Check if full screen
        If intResolutionMode = 5 Then
            'Prevent moving the form by control box click
            If (m.Msg = WINDOWMESSAGE_SYSTEM_COMMAND) And (m.WParam.ToInt32() = CONTROL_MOVE) Then
                Return
            End If
            'Prevent button down moving form
            If (m.Msg = WINDOWMESSAGE_CLICK_BUTTONDOWN) And (m.WParam.ToInt32() = WINDOWCAPTION) Then
                Return
            End If
        End If

        'If a double click on the title bar is triggered
        If m.Msg = WINDOWMESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
            Return
        End If

        'Still send message but without resizing
        MyBase.WndProc(m)

    End Sub

    Sub New()

        ' This call is required by the designer
        InitializeComponent() 'Do not remove

        'Load for testing the threading status
        thrRendering = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Rendering))

        'Check for multi-threading first
        If thrRendering.TrySetApartmentState(Threading.ApartmentState.MTA) Then
            'Set, multi-threading is possible
            blnThreadSupported = True
            'Stop the paint event, we will do the painting when we want to
            MyBase.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            MyBase.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            MyBase.SetStyle(ControlStyles.UserPaint, False)
            'Form buffer
            Me.DoubleBuffered = True
        End If

        'Load sound file percentages
        For intLoop As Integer = 0 To btmSound.GetUpperBound(0)
            btmSound(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\Options\Sound" & CStr(intLoop) & ".png"))
        Next

        'Set 100%
        btmSoundPercent = btmSound(100)

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Close fog
        If thrFog IsNot Nothing Then
            If thrFog.IsAlive Then
                thrFog.Abort()
            End If
        End If

        'Close rendering
        If thrRendering IsNot Nothing Then
            If thrRendering.IsAlive Then
                thrRendering.Abort()
            End If
        End If

        'Close
        If thrHoverSoundDelay IsNot Nothing Then
            If thrHoverSoundDelay.IsAlive Then
                thrHoverSoundDelay.Abort()
            End If
        End If

        'Loading
        If thrLoading IsNot Nothing Then
            If thrLoading.IsAlive Then
                thrLoading.Abort()
            End If
        End If

        'Paragraphing
        If thrParagraph IsNot Nothing Then
            If thrParagraph.IsAlive Then
                thrParagraph.Abort()
            End If
        End If

        'Ending game
        If thrEndGame IsNot Nothing Then
            If thrEndGame.IsAlive Then
                thrEndGame.Abort()
            End If
        End If

        'Stop and dispose character and zombies
        RemoveCharacterAndZombiesFromMemory()

    End Sub

    Private Sub RemoveCharacterAndZombiesFromMemory()

        'Stop and dispose character
        If udcCharacter IsNot Nothing Then
            udcCharacter.StopAndDispose()
            udcCharacter = Nothing
        End If

        'Stop and dispose zombies
        For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
            If udcZombies(intLoop) IsNot Nothing Then
                udcZombies(intLoop).StopAndDispose()
                udcZombies(intLoop) = Nothing
            End If
        Next

    End Sub

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Check if multi-threading was possible
        If Not blnThreadSupported Then
            'Display
            MessageBox.Show("This computer doesn't support multi-threading. This application will close now.", "Last Stand", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Exit
            Me.Close()
        Else
            'Force show, allows focus to happen immedately
            Me.Show()
            Me.Focus()
            'Set percentage multiplers for screen modes
            gdblScreenWidthRatio = CDbl((intScreenWidth - WIDTHSUBTRACTION) / ORIGINALSCREENWIDTH)
            gdblScreenHeightRatio = CDbl((intScreenHeight - HEIGHTSUBTRACTION) / ORIGINALSCREENHEIGHT)
            'Menu sound
            udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - WIDTHSUBTRACTION, intScreenHeight - HEIGHTSUBTRACTION) 'Full screen
            'Set for fog
            thrFog = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf KeepFogMoving))
            thrFog.Start()
            'Start rendering
            thrRendering.Start()
        End If

    End Sub

    Private Sub Rendering()

        'Loop
        While True
            'Empty variables if back from game
            If blnBackFromGame Then
                'Stop and dispose character and zombies
                RemoveCharacterAndZombiesFromMemory()
                'Set
                blnBackFromGame = False
            End If
            'Check mode before painting on canvas
            Select Case intCanvasMode
                Case 0 'Menu
                    RenderMenuScreen()
                Case 1 'Options screen
                    RenderOptionsScreen()
                Case 2 'Play game, loading screen first
                    LoadingGameScreen()
                Case 3 'Playing the game
                    StartedGameScreen()
                Case 4 'Highscores screen
                    HighscoresScreen()
                Case 5 'Credits screen
                    CreditsScreen()
            End Select
            'Select to paint on screen
            gGraphics = Me.CreateGraphics()
            'Set graphic options
            SetDefaultGraphicOptions()
            'Paint on screen
            gGraphics.DrawImage(btmCanvas, rectFullScreen)
            'If changing screen, we must change resolution in this thread or else strange things happen
            ScreenResolutionChanged()
            'Do events
            Application.DoEvents()
        End While

    End Sub

    Private Sub PaintOnCanvas()

        'Paint on canvas
        gGraphics = Graphics.FromImage(btmCanvas)

        'Set graphic options
        SetDefaultGraphicOptions()

    End Sub

    Private Sub SetDefaultGraphicOptions()

        'Set options for fastest rendering
        With gGraphics
            .CompositingMode = Drawing2D.CompositingMode.SourceOver
            .CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
            .SmoothingMode = Drawing2D.SmoothingMode.None
            .InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
            .TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixel
            .PixelOffsetMode = Drawing2D.PixelOffsetMode.HighSpeed
        End With

    End Sub

    Private Sub RenderMenuScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw menu
        gGraphics.DrawImageUnscaled(btmMenuBackground, pntTopLeft)

        'Draw fog in back
        gGraphics.DrawImageUnscaled(btmFogBackPass1, pntFogBackPass1)
        gGraphics.DrawImageUnscaled(btmFogBackPass2, pntFogBackPass2)

        'Draw Archer
        gGraphics.DrawImageUnscaled(btmArcher, pntArcher)

        'Draw fog in front
        gGraphics.DrawImageUnscaled(btmFogFrontPass1, pntFogFrontPass1)
        gGraphics.DrawImageUnscaled(btmFogFrontPass2, pntFogFrontPass2)

        'Draw start text
        If intCanvasShow = 1 Then
            gGraphics.DrawImageUnscaled(btmStartHoverText, pntStartHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmStartText, pntStartText)
        End If

        'Draw highscores text
        If intCanvasShow = 2 Then
            gGraphics.DrawImageUnscaled(btmHighscoresHoverText, pntHighscoresHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmHighscoresText, pntHighscoresText)
        End If

        'Draw story text
        If intCanvasShow = 3 Then
            gGraphics.DrawImageUnscaled(btmStoryHoverText, pntStoryHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmStoryText, pntStoryText)
        End If

        'Draw credits text
        If intCanvasShow = 4 Then
            gGraphics.DrawImageUnscaled(btmOptionsHoverText, pntOptionsHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmOptionsText, pntOptionsText)
        End If

        'Draw options text
        If intCanvasShow = 5 Then
            gGraphics.DrawImageUnscaled(btmCreditsHoverText, pntCreditsHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmCreditsText, pntCreditsText)
        End If

        'Draw versus text
        If intCanvasShow = 6 Then
            gGraphics.DrawImageUnscaled(btmVersusHoverText, pntVersusHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmVersusText, pntVersusText)
        End If

        'Draw last stand text
        gGraphics.DrawImageUnscaled(btmLastStandText, pntLastStandText)

    End Sub

    Private Sub RenderOptionsScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw options background
        gGraphics.DrawImageUnscaled(btmOptionsBackground, pntTopLeft)

        'Draw resolution text
        gGraphics.DrawImageUnscaled(btmResolutionText, pntResolutionText)

        'Draw sound text
        gGraphics.DrawImageUnscaled(btmSoundText, pntSoundText)

        'Draw last stand text
        gGraphics.DrawImageUnscaled(btmLastStandText, pntLastStandText)

        'Check which resolution
        CheckResolutionMode(0, btm800x600Text, btmNot800x600Text, pnt800x600Text)
        CheckResolutionMode(1, btm1024x768Text, btmNot1024x768Text, pnt1024x768Text)
        CheckResolutionMode(2, btm1280x800Text, btmNot1280x800Text, pnt1280x800Text)
        CheckResolutionMode(3, btm1280x1024Text, btmNot1280x1024Text, pnt1280x1024Text)
        CheckResolutionMode(4, btm1440x900Text, btmNot1440x900Text, pnt1440x900Text)
        CheckResolutionMode(5, btmFullscreenText, btmNotFullscreenText, pntFullscreenText)

        'Draw sound bar
        gGraphics.DrawImageUnscaled(btmSoundBar, pntSoundBar)

        'Draw sound percentage
        gGraphics.DrawImageUnscaled(btmSoundPercent, pntSoundPercent)

        'Draw slider
        gGraphics.DrawImageUnscaled(btmSlider, pntSlider)

        'Check
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub CheckResolutionMode(intModeSelected As Integer, btmResolutionText As Bitmap, btmNotResolutionText As Bitmap, pntResolutionText As Point)

        'Check resolution before drawing
        If intResolutionMode = intModeSelected Then
            gGraphics.DrawImageUnscaled(btmResolutionText, pntResolutionText)
        Else
            gGraphics.DrawImageUnscaled(btmNotResolutionText, pntResolutionText)
        End If

    End Sub

    Private Sub LoadingGameScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw loading background
        gGraphics.DrawImageUnscaled(btmLoadingBackground, pntTopLeft)

        'Draw loading bar
        gGraphics.DrawImageUnscaled(btmLoadingBar, pntLoadingBar)

        'Draw Loading text
        If intCanvasShow = 0 And intCanvasMode = 2 Then 'Have to check for both because at some point the rendering will glitch if not checked
            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
        Else
            gGraphics.DrawImageUnscaled(btmLoadingStartText, pntLoadingStartText)
        End If

        'Draw paragraph
        gGraphics.DrawImageUnscaled(btmLoadingParagraph, pntLoadingParagraph)

    End Sub

    Private Sub StartedGameScreen()

        'Move graphics to copy and print dead first
        gGraphics = Graphics.FromImage(btmGameBackground)
        'Set graphic options
        SetDefaultGraphicOptions()

        'Draw dead zombies permanently
        For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
            If udcZombies(intLoop) IsNot Nothing Then
                If udcZombies(intLoop).PaintOnBackgroundAfterDead Then
                    gGraphics.DrawImageUnscaled(udcZombies(intLoop).btmZombie, udcZombies(intLoop).pntZombie)
                    'Increase speed
                    intZombieSpeed += 1
                    'Create a new zombie
                    udcZombies(intLoop) = New clsZombie(Me, ORIGINALSCREENWIDTH, 325, intZombieSpeed)
                    udcZombies(intLoop).Start()
                End If
            End If
        Next

        'Paint on canvas
        PaintOnCanvas()

        'Draw the background, the armory area
        gGraphics.DrawImageUnscaled(btmGameBackground, pntTopLeft)

        'Draw word bar
        gGraphics.DrawImageUnscaled(btmWordBar, pntWordBar)

        'Draw word in the word bar
        gGraphics.DrawImageUnscaled(btmWord, pntWord)

        'Draw character
        gGraphics.DrawImageUnscaled(udcCharacter.btmCharacter, udcCharacter.pntCharacter)

        'Draw zombies pinning
        For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
            If udcZombies(intLoop) IsNot Nothing Then
                'Check distance
                If udcZombies(intLoop).pntZombie.X <= 200 And Not udcZombies(intLoop).IsPinning Then
                    'Set
                    udcZombies(intLoop).Pin() 'Pinning character
                    'Character dying sound
                    If blnGameOverFirstTime Then
                        'Play
                        Dim udcScreamSound As New clsSound(Me, strDirectory & "Sounds\CharacterDying.mp3", 3000, gintSoundVolume, False)
                        'Set
                        thrEndGame = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf EndingGame))
                        thrEndGame.Start()
                        'Set
                        blnGameOverFirstTime = False
                        'Set
                        blnEndingGameCantType = True
                    End If
                End If
                'Paint if pinning
                If udcZombies(intLoop).IsPinning And udcZombies(intLoop).IsAlive Then
                    gGraphics.DrawImageUnscaled(udcZombies(intLoop).btmZombie, udcZombies(intLoop).pntZombie)
                End If
            End If
        Next

        'Draw zombies walking
        For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
            If udcZombies(intLoop) IsNot Nothing Then
                If Not udcZombies(intLoop).PaintOnBackgroundAfterDead And Not udcZombies(intLoop).IsPinning Then
                    gGraphics.DrawImageUnscaled(udcZombies(intLoop).btmZombie, udcZombies(intLoop).pntZombie)
                End If
            End If
        Next

        'If game over
        If btmEndGame IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmEndGame, pntTopLeft)
        End If

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub EndingGame()

        'Set
        btmEndGame = btmEndGame25

        'Wait
        System.Threading.Thread.Sleep(ENDGAMEDELAYBLACKSCREEN)

        'Set
        btmEndGame = btmEndGame50

        'Wait
        System.Threading.Thread.Sleep(ENDGAMEDELAYBLACKSCREEN)

        'Set
        btmEndGame = btmEndGame75

        'Wait
        System.Threading.Thread.Sleep(ENDGAMEDELAYBLACKSCREEN)

        'Set
        btmEndGame = btmEndGame100

    End Sub

    Private Sub HighscoresScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw highscores background
        gGraphics.DrawImageUnscaled(btmHighscoresBackground, pntTopLeft)

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub CreditsScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw credits background
        gGraphics.DrawImageUnscaled(btmCreditsBackground, pntTopLeft)

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub ScreenResolutionChanged()

        'Change resolution if it does need to be changed
        If blnScreenChanged Then
            'Change window state
            If intResolutionMode <> 5 Then '5 = fullscreen
                'Normal
                Me.Invoke(Sub() Me.WindowState = FormWindowState.Normal) 'Prevent cross-threading
                'Change
                Me.Invoke(Sub() Me.Width = intScreenWidth) 'Prevent cross-threading
                Me.Invoke(Sub() Me.Height = intScreenHeight) 'Prevent cross-threading
            Else
                'Force full screen, let windows do stuff for us
                Me.Invoke(Sub() Me.WindowState = FormWindowState.Maximized) 'Prevent cross-threading
            End If
            'Set
            gdblScreenWidthRatio = CDbl((Me.Width - WIDTHSUBTRACTION) / ORIGINALSCREENWIDTH)
            gdblScreenHeightRatio = CDbl((Me.Height - HEIGHTSUBTRACTION) / ORIGINALSCREENHEIGHT)
            'Set screen rectangle
            rectFullScreen.Width = Me.Width - WIDTHSUBTRACTION
            rectFullScreen.Height = Me.Height - HEIGHTSUBTRACTION
            'Center the form
            If intResolutionMode <> 5 Then
                Me.Invoke(Sub() Me.Top = CInt((My.Computer.Screen.WorkingArea.Height / 2) - (Me.Height / 2))) 'Prevent cross-threading
                Me.Invoke(Sub() Me.Left = CInt((My.Computer.Screen.WorkingArea.Width / 2) - (Me.Width / 2))) 'Prevent cross-threading
            End If
            'Reset
            blnScreenChanged = False
        End If

    End Sub

    Private Sub KeepFogMoving()

        'Declare
        Dim rndNumber As New Random

        'Loop
        While True
            'Reset fog
            If intFogPass1X >= ORIGINALSCREENWIDTH * 2 Then
                intFogPass1X = -(ORIGINALSCREENWIDTH * 2)
            End If
            If intFogPass2X >= ORIGINALSCREENWIDTH * 2 Then
                intFogPass2X = -(ORIGINALSCREENWIDTH * 2)
            End If
            'Move fog
            intFogPass1X += rndNumber.Next(1, FOGMAXSPEED + 1)
            intFogPass2X += rndNumber.Next(1, FOGMAXSPEED + 1)
            'Move fog
            pntFogFrontPass1.X = intFogPass1X
            pntFogBackPass1.X = intFogPass1X
            pntFogFrontPass2.X = intFogPass2X
            pntFogBackPass2.X = intFogPass2X
            'Reduce CPU usage
            System.Threading.Thread.Sleep(rndNumber.Next(1, FOGMAXDELAY + 1))
        End While

    End Sub

    Private Function blnMouseInRegion(pntMouse As Point, intImageWidth As Integer, intImageHeight As Integer, pntStartingPoint As Point) As Boolean

        'Return
        If pntMouse.X >= CInt(pntStartingPoint.X * gdblScreenWidthRatio) And
        pntMouse.X <= CInt((pntStartingPoint.X + intImageWidth) * gdblScreenWidthRatio) And
        pntMouse.Y >= CInt(pntStartingPoint.Y * gdblScreenHeightRatio) And
        pntMouse.Y <= CInt((pntStartingPoint.Y + intImageHeight) * gdblScreenHeightRatio) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub HoverText(intCanvasShowToSet As Integer, ByRef blnByRefOnTopOf As Boolean)

        'Set canvas show
        intCanvasShow = intCanvasShowToSet

        'Check if mouse was on top of button once before
        If Not blnByRefOnTopOf Then
            'Set
            blnByRefOnTopOf = True
            'Hover sound
            If thrHoverSoundDelay Is Nothing Then
                'Play sound
                Dim udcButtonHoverSound As New clsSound(Me, strDirectory & "Sounds\ButtonHover.mp3", 3000, gintSoundVolume, False)
                'Start a waiting thread of 2500 ms
                thrHoverSoundDelay = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf HoverSoundWaiting))
                thrHoverSoundDelay.Start()
            End If
        End If

    End Sub

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x and y points to ensure it always works.

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)
        Static sblnOnTopOf As Boolean = False

        'Print for mouse cordinates if necessary
        'Debug.Print("Mouse point = " & CStr(pntMouse.X) & ", " & CStr(pntMouse.Y))

        'Menu stuff
        If intCanvasMode = 0 Then

            'Start has been moused over
            If blnMouseInRegion(pntMouse, 212, 69, pntStartText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Highscores has been moused over
            If blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) Then
                HoverText(2, sblnOnTopOf)
                Exit Sub
            End If

            'Story has been moused over
            If blnMouseInRegion(pntMouse, 218, 87, pntStoryText) Then
                HoverText(3, sblnOnTopOf)
                Exit Sub
            End If

            'Options has been moused over
            If blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) Then
                HoverText(4, sblnOnTopOf)
                Exit Sub
            End If

            'Credits has been moused over
            If blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) Then
                HoverText(5, sblnOnTopOf)
                Exit Sub
            End If

            'Versus has been moused over
            If blnMouseInRegion(pntMouse, 256, 74, pntVersusText) Then
                HoverText(6, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint menu background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

        'Check options
        If intCanvasMode = 1 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Sound changing
            If blnMouseInRegion(pntMouse, 600, 46, pntSoundBar) And blnSliderWithMouseDown Then 'In the bar
                If blnSliderWithMouseDown Then
                    ChangeSoundVolume(pntMouse)
                End If
                'Exit
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

        End If

        'Check game screen
        If intCanvasMode = 3 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

        End If

        'Check highscores screen
        If intCanvasMode = 4 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

        End If

        If intCanvasMode = 5 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

        End If

    End Sub

    Private Sub HoverSoundWaiting()

        'Sleep for 3000 ms
        System.Threading.Thread.Sleep(3000)
        thrHoverSoundDelay = Nothing

    End Sub

    Private Sub ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer)

        'Set
        intCanvasMode = intCanvasModeToSet

        'Set
        intCanvasShow = intCanvasShowToSet

        'Play sound
        Dim udcButtonPressedSound As New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 3000, gintSoundVolume, False)

    End Sub

    Private Sub Paragraphing()

        'Set
        btmLoadingParagraph = btmLoadingParagraph25

        'Sleep
        System.Threading.Thread.Sleep(333)

        'Set
        btmLoadingParagraph = btmLoadingParagraph50

        'Sleep
        System.Threading.Thread.Sleep(333)

        'Set
        btmLoadingParagraph = btmLoadingParagraph75

        'Sleep
        System.Threading.Thread.Sleep(333)

        'Set
        btmLoadingParagraph = btmLoadingParagraph100

    End Sub

    Private Sub LoadingGame()

        'Notes: This procedure is for loading the game to play

        'Set
        btmLoadingBar = btmLoadingBar0

        'Set
        btmLoadingBar = btmLoadingBar10

        'Set
        btmEndGame = Nothing

        'Set
        blnEndingGameCantType = False

        'Set
        blnGameOverFirstTime = True

        'Set to be fresh
        btmGameBackground = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground.jpg"))

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325)

        'Set
        btmLoadingBar = btmLoadingBar20

        'Set zombie speed
        intZombieSpeed = 25

        'Zombie 1
        udcZombies(0) = New clsZombie(Me, ORIGINALSCREENWIDTH, 325, intZombieSpeed)

        'Set
        btmLoadingBar = btmLoadingBar30

        'Zombie 2
        udcZombies(1) = New clsZombie(Me, ORIGINALSCREENWIDTH + 500, 325, intZombieSpeed)

        'Set
        btmLoadingBar = btmLoadingBar40

        'Zombie 3
        udcZombies(2) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1000, 325, intZombieSpeed)

        'Set
        btmLoadingBar = btmLoadingBar50

        'Zombie 4
        udcZombies(3) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1500, 325, intZombieSpeed)

        'Set
        btmLoadingBar = btmLoadingBar60

        'Zombie 5
        udcZombies(4) = New clsZombie(Me, ORIGINALSCREENWIDTH + 2000, 325, intZombieSpeed)

        'Set
        btmLoadingBar = btmLoadingBar70

        'Set
        btmLoadingBar = btmLoadingBar80

        'Set
        btmEndGame25 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame25.png"))
        btmEndGame50 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame50.png"))
        btmEndGame75 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame75.png"))
        btmEndGame100 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame100.png"))

        'Set
        btmLoadingBar = btmLoadingBar90

        'Set
        btmLoadingBar = btmLoadingBar99

        'Load all words
        LoadAllWords()

        'Grab a random word
        LoadARandomWord()

        'Set
        btmLoadingBar = btmLoadingBar100

        'Set
        intCanvasShow = 1 'Means completely loaded

    End Sub

    Private Sub ShowNextScreenAndExitMenu(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer)

        'Stop fog thread
        thrFog.Abort()

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(intCanvasModeToSet, intCanvasShowToSet)

        'Stop sound
        udcAmbianceSound.StopAndCloseSound()

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x and y points to ensure it always works.

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If menu
        If intCanvasMode = 0 Then

            'Start was clicked
            If blnMouseInRegion(pntMouse, 212, 69, pntStartText) Then
                'Stop fog thread
                thrFog.Abort()
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(2, 0)
                'Stop sound
                udcAmbianceSound.StopAndCloseSound()
                'Start loading game
                thrLoading = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame))
                thrLoading.Start()
                'Start paragraphing
                thrParagraph = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Paragraphing))
                thrParagraph.Start()
                'Exit
                Exit Sub
            End If

            'Highscores was clicked
            If blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) Then
                ShowNextScreenAndExitMenu(4, 0)
                'Exit
                Exit Sub
            End If

            'Options was clicked
            If blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) Then
                ShowNextScreenAndExitMenu(1, 0)
                'Exit
                Exit Sub
            End If

            'Credits was clicked
            If blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) Then
                ShowNextScreenAndExitMenu(5, 0)
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If options
        If intCanvasMode = 1 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Restart fog
                RestartFog()
                'Exit
                Exit Sub
            End If

            'Resolution change 800x600
            If blnMouseInRegion(pntMouse, 153, 32, pnt800x600Text) Then
                'Resize screen
                ResizeScreen(0, 800, 600)
                'Exit
                Exit Sub
            End If

            'Resolution change 1024x768
            If blnMouseInRegion(pntMouse, 163, 32, pnt1024x768Text) Then
                'Resize screen
                ResizeScreen(1, 1024, 768)
                'Exit
                Exit Sub
            End If

            'Resolution change 1280x800
            If blnMouseInRegion(pntMouse, 163, 32, pnt1280x800Text) Then
                'Resize screen
                ResizeScreen(2, 1280, 800)
                'Exit
                Exit Sub
            End If

            'Resolution change 1280x1024
            If blnMouseInRegion(pntMouse, 170, 32, pnt1280x1024Text) Then
                'Resize screen
                ResizeScreen(3, 1280, 1024)
                'Exit
                Exit Sub
            End If

            'Resolution change 1440x900
            If blnMouseInRegion(pntMouse, 170, 32, pnt1440x900Text) Then
                'Resize screen
                ResizeScreen(4, 1440, 900)
                'Exit
                Exit Sub
            End If

            'Resolution change full screen
            If blnMouseInRegion(pntMouse, 175, 38, pntFullscreenText) Then
                'Resize screen
                ResizeScreen(5)
                'Exit
                Exit Sub
            End If

            'Sound changing
            If blnMouseInRegion(pntMouse, 600, 46, pntSoundBar) Then 'In the bar
                ChangeSoundVolume(pntMouse)
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If loading game screen
        If intCanvasMode = 2 Then

            'Loading start text bar was clicked
            If intCanvasShow = 1 Then 'If finished loading
                If blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                    'Play ambiance sound
                    udcGameAmbiance = New clsSound(Me, strDirectory & "\Sounds\HellAmbiance.mp3", 36000, gintSoundVolume, True)
                    'Set
                    intCanvasMode = 3
                    'Set
                    intCanvasShow = 0
                    'Start character
                    udcCharacter.Start()
                    'Start zombies
                    For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                        If udcZombies(intLoop) IsNot Nothing Then
                            udcZombies(intLoop).Start()
                        End If
                    Next
                    'Exit
                    Exit Sub
                End If
            End If

            'Exit
            Exit Sub

        End If

        'If game screen
        If intCanvasMode = 3 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Stop sound
                udcGameAmbiance.StopAndCloseSound()
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Set
                blnBackFromGame = True
                'Abort
                If thrEndGame IsNot Nothing Then
                    If thrEndGame.IsAlive Then
                        thrEndGame.Abort()
                    End If
                End If
                'Restart fog
                RestartFog()
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If highscores screen
        If intCanvasMode = 4 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Restart fog
                RestartFog()
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If credits screen
        If intCanvasMode = 5 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Restart fog
                RestartFog()
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

    End Sub

    Private Sub ChangeSoundVolume(pntMouse As Point)

        'Declare
        Dim intMousePointCalculation As Integer = 0

        'Move slider
        Select Case intResolutionMode
            Case 0 'Max 39 to 319 pntMouse
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) - 12 '12 Adjustment for pixels
            Case 1
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) - 10 '10 Adjustment for pixels
            Case 2
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) - 8 '8 Adjustment for pixels
            Case 3
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) - 6 '6 Adjustment for pixels
            Case 4
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) - 4 '4 Adjustment for pixels
            Case 5
                intMousePointCalculation = CInt((pntMouse.X / gdblScreenWidthRatio) - ((53 / 2) * gdblScreenWidthRatio)) 'No adjustment for pixels screen is wider than normal
        End Select

        'If close just make it marked as that number, mouse cordinates vary by screen resolution, 57 to 659 is what we can get when tested, let's have only 58 to 657
        If intMousePointCalculation <= 61 Then
            intMousePointCalculation = 58
        End If
        If intMousePointCalculation >= 654 Then
            intMousePointCalculation = 657
        End If

        'Set point for the slider
        pntSlider.X = intMousePointCalculation

        'Set volume
        If pntSlider.X = 58 Then
            gintSoundVolume = 0
            btmSoundPercent = btmSound(0)
        ElseIf pntSlider.X = 657 Then
            gintSoundVolume = 1000
            btmSoundPercent = btmSound(100)
        Else
            gintSoundVolume = CInt((pntSlider.X - 58) * (1000 / 600)) '58 to 657 = 600 pixel bounds, inclusive of 58 as a number.
            btmSoundPercent = btmSound(CInt(gintSoundVolume / 10))
        End If

    End Sub

    Private Sub ResizeScreen(intResolutionModeToSet As Integer, Optional intWidth As Integer = 0, Optional intHeight As Integer = 0)

        'Set
        intResolutionMode = intResolutionModeToSet

        'Set
        intScreenWidth = intWidth
        intScreenHeight = intHeight

        'Set
        blnScreenChanged = True

    End Sub

    Private Sub RestartFog()

        'Reset fog
        intFogPass1X = -(ORIGINALSCREENWIDTH * 2)
        intFogPass2X = -((ORIGINALSCREENWIDTH * 2) * 2)
        pntFogBackPass1.X = -(ORIGINALSCREENWIDTH * 2)
        pntFogFrontPass1.X = -(ORIGINALSCREENWIDTH * 2)
        pntFogBackPass2.X = -((ORIGINALSCREENWIDTH * 2) * 2)
        pntFogFrontPass2.X = -((ORIGINALSCREENWIDTH * 2) * 2)

        'Set for fog
        thrFog = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf KeepFogMoving))
        thrFog.Start()

    End Sub

    Private Sub frmGame_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If options
        If intCanvasMode = 1 Then

            'Sound changing
            If blnMouseInRegion(pntMouse, 600, 46, pntSoundBar) Then 'In the bar
                blnSliderWithMouseDown = True
            End If

        End If

    End Sub

    Private Sub frmGame_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp

        'If options
        If intCanvasMode = 1 Then

            'Sound changing
            blnSliderWithMouseDown = False

        End If

    End Sub

    Private Sub frmGame_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

        'Check if playing game
        If intCanvasMode = 3 Then

            'Check the word being typed
            If Not blnEndingGameCantType Then
                If strWord.Substring(0, 1) = LCase(e.KeyChar) Or strWord.Substring(0, 1) = UCase(e.KeyChar) Then
                    'Change word by one less letter
                    strWord = strWord.Substring(1)
                    'Increase
                    intWordIndex += 1
                    'Check if word is done
                    If Len(strWord) = 0 Then
                        'Get a new word
                        LoadARandomWord()
                        'Character shoots
                        udcCharacter.CharacterShot()
                        'Kill closest zombie
                        Dim intIndex As Integer = intGetIndexOfClosestZombie() 'If returns too high, means player typed too good, shoot nothing
                        If intIndex <> udcZombies.GetUpperBound(0) + 1 Then
                            udcZombies(intIndex).Dead()
                        End If
                    Else
                        'Show to screen
                        btmWord = New Bitmap(New Bitmap(Image.FromFile(strDirectory.Substring(0, Len(strDirectory) - 1) & "Images\Words\" & UCase(strTheWord.Substring(0, 1)) &
                                  strTheWord.Substring(1) & "\" & CStr(intWordIndex) & ".png")))
                    End If
                End If
            End If

            'Exit
            Exit Sub

        End If

    End Sub

    Private Function intGetIndexOfClosestZombie() As Integer

        'Declare
        Dim intClosestX As Integer = Integer.MaxValue
        Dim intIndex As Integer = udcZombies.GetUpperBound(0) + 1

        'Loop to get closest zombie
        For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
            If udcZombies(intLoop) IsNot Nothing Then
                If udcZombies(intLoop).IsAlive Then
                    If intClosestX > udcZombies(intLoop).pntZombie.X Then
                        intClosestX = udcZombies(intLoop).pntZombie.X
                        intIndex = intLoop
                    End If
                End If
            End If
        Next

        'Return
        Return intIndex

    End Function

    Private Sub LoadAllWords()

        'Declare
        Dim intLoop As Integer = 0

        'Load each word, get all folders in the word directory
        For Each strDirectoryWord As String In System.IO.Directory.GetDirectories(strDirectory & "Images\Words")
            'Declare
            Dim sdiDirectory As New System.IO.DirectoryInfo(strDirectoryWord)
            'Redim
            ReDim Preserve astrWords(intLoop)
            astrWords(intLoop) = sdiDirectory.Name
            'Increase
            intLoop += 1
        Next

    End Sub

    Private Sub LoadARandomWord()

        'Declare
        Dim rndNumber As New Random
        strTheWord = astrWords(rndNumber.Next(0, astrWords.GetUpperBound(0) + 1))
        strWord = strTheWord

        'Set
        intWordIndex = 0

        'Set bitmap of the word
        btmWord = New Bitmap(Image.FromFile(strDirectory.Substring(0, Len(strDirectory) - 1) & "Images\Words\" & UCase(strTheWord.Substring(0, 1)) &
                  strTheWord.Substring(1) & "\0.png"))

    End Sub

End Class