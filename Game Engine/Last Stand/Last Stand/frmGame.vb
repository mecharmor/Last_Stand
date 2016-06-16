'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const GAME_VERSION As String = "1.27"
    Private Const ORIGINAL_SCREEN_WIDTH As Integer = 1680
    Private Const ORIGINAL_SCREEN_HEIGHT As Integer = 1050
    Private Const WINDOW_MESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const CONTROL_MOVE As Integer = 61456
    Private Const WINDOW_MESSAGE_CLICK_BUTTON_DOWN As Integer = 161
    Private Const WINDOW_CAPTION As Integer = 2
    Private Const WINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3
    Private Const WIDTH_SUBTRACTION As Integer = 16 'Probably the edges of the window
    Private Const HEIGHT_SUBTRACTION As Integer = 38 'Probably the edges of the window
    Private Const FOG_MAX_SPEED As Integer = 25
    Private Const FOG_MAX_DELAY As Integer = 100
    Private Const LOADING_TRANSPARENCY_DELAY As Integer = 400
    Private Const BLACK_SCREEN_DELAY As Integer = 500
    Private Const NUMBER_OF_ZOMBIES_CREATED As Integer = 150
    Private Const NUMBER_OF_ZOMBIES_AT_ONE_TIME As Integer = 5
    Private Const ZOMBIE_PINNING_X_DISTANCE As Integer = 200

    'Declare beginning necessary engine needs
    Private intScreenWidth As Integer = 800
    Private intScreenHeight As Integer = 600
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False
    Private rectFullScreen As Rectangle
    Private gGraphics As Graphics
    Private btmCanvas As New Bitmap(ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT, System.Drawing.Imaging.PixelFormat.Format32bppPArgb) 'Original resolution of our game
    Private pntTopLeft As New Point(0, 0)
    Private intCanvasMode As Integer = 0 'Default menu screen
    Private intCanvasShow As Integer = 0 'Default, no animation
    Private strDirectory As String = AppDomain.CurrentDomain.BaseDirectory & "\"
    Private blnScreenChanged As Boolean = False

    'Menu necessary needs
    Private btmMenuBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\MenuBackground.jpg"))
    Private btmFogFrontPass1 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogFront.png"))
    Private btmFogBackPass1 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogBack.png"))
    Private intFogPass1X As Integer = -(ORIGINAL_SCREEN_WIDTH * 2)
    Private intFogFrontPass1Y As Integer = 650
    Private intFogBackPass1Y As Integer = 250
    Private pntFogFrontPass1 As New Point(intFogPass1X, intFogFrontPass1Y)
    Private pntFogBackPass1 As New Point(intFogPass1X, intFogBackPass1Y)
    Private btmFogFrontPass2 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogFront.png"))
    Private btmFogBackPass2 As New Bitmap(Image.FromFile(strDirectory & "Images\Menu\FogBack.png"))
    Private intFogPass2X As Integer = -((ORIGINAL_SCREEN_WIDTH * 2) * 2)
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
    Private udcAmbianceOptionsSound As clsSound

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
    Private btmLoadingBarPicture(10) As Bitmap '0 To 10 = 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
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
    Private pntLoadingParagraph As New Point(123, 261)
    Private thrParagraph As System.Threading.Thread
    Private thrLoadingGame As System.Threading.Thread
    Private thrLoading(9) As System.Threading.Thread '0 to 9 which is 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
    Private blnLoadingGameFinished(9) As Boolean '0 to 9 which is 10, 20, 30, 40, 50, 60, 70, 80, 90, 100

    'Game screen
    Private btmGameBackground As Bitmap 'Point for this is created in the public module
    Private btmGameBackgroundCloneScreenShown As Bitmap
    Private btmDeathScreen As Bitmap
    Private btmDeathOverlay As Bitmap
    Private btmBlackScreen25 As Bitmap
    Private btmBlackScreen50 As Bitmap
    Private btmBlackScreen75 As Bitmap
    Private btmBlackScreen100 As Bitmap
    Private btmBlackScreen As Bitmap
    Private thrBlackScreen As System.Threading.Thread
    Private blnBlackScreenFinished As Boolean = False
    Private blnTriggeredBlackScreenThread As Boolean = False
    Private blnRemovedGameObjectsFromMemory As Boolean = False
    Private thrEndGame As System.Threading.Thread
    Private btmWordBar As Bitmap
    Private pntWordBar As New Point(482, 27)
    Private udcCharacter As clsCharacter
    Private intZombieSpeed As Integer = 10 ' Default
    Private blnBackFromGame As Boolean = False
    Private blnGameOverFirstTime As Boolean = False
    Private blnEndingGameCantType As Boolean = False
    Private btmAK47Magazine As Bitmap
    Private pntAK47Magazine As New Point(59, 877)
    Private btmWinOverlay As Bitmap
    Private udcScreamSound As clsSound
    Private intLevel As Integer = 1 'Starting
    Private intZombieKills As Integer = 0
    Private intReloadTimeUsed As Integer = 0
    Private intZombieKillsCombined As Integer = 0
    Private intReloadTimeUsedCombined As Integer = 0
    Private blnComparedHighscore As Boolean = False
    Private blnBeatLevel As Boolean = False

    'Sounds for the levels
    Private udcGameBackgroundSound(4) As clsSound

    'Zombies that produce game lose
    Private udcUnderwaterZombieGameLose As clsZombieGameLose
    Private blnZombieGameLoseHappened As Boolean = False

    'Helicopter
    Private udcHelicopter As clsHelicopter

    'Path choices
    Private btmPath As Bitmap
    Private btmPath1_0 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\1_0.jpg"))
    Private btmPath1_1 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\1_1.jpg"))
    Private btmPath1_2 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\1_2.jpg"))
    Private btmPath2_0 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\2_0.jpg"))
    Private btmPath2_1 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\2_1.jpg"))
    Private btmPath2_2 As New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\Paths\2_2.jpg"))
    Private blnLightZap1 As Boolean = False
    Private blnLightZap2 As Boolean = False
    Private udcZombieGrowlSound As clsSound

    'Stop watch
    Private swhStopWatch As Stopwatch

    'Words
    Private astrWords(0) As String 'Used to fill with words
    Private intWordIndex As Integer = 0
    Private strTheWord As String = ""
    Private strWord As String = ""

    'Highscores screen
    Private btmHighscoresBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Highscores\HighscoresBackground.jpg"))
    Private strHighscores As String = ""
    Private blnHighscoresIsAccess As Boolean = False

    'Credits screen
    Private btmCreditsBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CreditsBackground.jpg"))
    Private btmJohnGonzales25 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\JohnGonzales25.png"))
    Private btmJohnGonzales50 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\JohnGonzales50.png"))
    Private btmJohnGonzales75 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\JohnGonzales75.png"))
    Private btmJohnGonzales100 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\JohnGonzales.jpg"))
    Private btmJohnGonzales As Bitmap
    Private pntJohnGonzales As New Point(200, 150)
    Private btmZacharyStafford25 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\ZacharyStafford25.png"))
    Private btmZacharyStafford50 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\ZacharyStafford50.png"))
    Private btmZacharyStafford75 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\ZacharyStafford75.png"))
    Private btmZacharyStafford100 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\ZacharyStafford.jpg"))
    Private btmZacharyStafford As Bitmap
    Private pntZacharyStafford As New Point(940, 150)
    Private btmCoryLewis25 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CoryLewis25.png"))
    Private btmCoryLewis50 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CoryLewis50.png"))
    Private btmCoryLewis75 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CoryLewis75.png"))
    Private btmCoryLewis100 As New Bitmap(Image.FromFile(strDirectory & "Images\Credits\CoryLewis.jpg"))
    Private btmCoryLewis As Bitmap
    Private pntCoryLewis As New Point(570, 575)
    Private thrCredits As System.Threading.Thread

    'Versus screen
    Private strIPAddress As String = ""
    Private strNickname As String = "Player"
    Private strNicknameConnected As String = ""
    Private strIPAddressConnect As String = ""
    Private intCanvasVersusShow As Integer = 0
    Private blnFirstTimeNicknameTyping As Boolean = True 'Defaulted
    Private blnFirstTimeIPAddressTyping As Boolean = True 'Defaulted
    Private btmVersusBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusBackground.jpg"))
    Private btmVersusNickname As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusNickname.png"))
    Private pntVersusBlackOutline As New Point(100, 150)
    Private btmVersusHost As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusHost.png"))
    Private pntVersusHost As New Point(267, 605)
    Private btmVersusOr As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusOr.png"))
    Private pntVersusOr As New Point(831, 649)
    Private btmVersusJoin As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusJoin.png"))
    Private pntVersusJoin As New Point(951, 601)
    Private btmVersusIPAddress As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusIPAddress.png"))
    Private btmVersusConnect As New Bitmap(Image.FromFile(strDirectory & "Images\Versus\VersusConnect.png"))
    Private pntVersusConnect As New Point(535, 615)
    Private blnPlayPressedSoundNow As Boolean = False

    'Versus other variables
    Private blnHost As Boolean = False
    Private tcplServer As System.Net.Sockets.TcpListener
    Private thrListening As System.Threading.Thread
    Private tcpcClient As System.Net.Sockets.TcpClient
    Private thrConnecting As System.Threading.Thread
    Private srClientData As System.IO.StreamReader
    Private udcVersusConnectedThread As clsVersusConnectedThread
    Private blnConnectionCompleted As Boolean = False
    Private blnWaiting As Boolean = False
    Private blnReadyEarly As Boolean = False

    'Loading versus game variables
    Private blnGameIsVersus As Boolean = False
    Private thrLoadingVersus As System.Threading.Thread
    Private thrParagraphVersus As System.Threading.Thread
    Private btmLoadingParagraphVersus25 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraphVersus25.png"))
    Private btmLoadingParagraphVersus50 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraphVersus50.png"))
    Private btmLoadingParagraphVersus75 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraphVersus75.png"))
    Private btmLoadingParagraphVersus100 As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingParagraphVersus100.png"))
    Private btmLoadingParagraphVersus As New Bitmap(1424, 472)
    Private btmLoadingWaitingText As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingWaitingText.png"))
    Private pntLoadingWaitingText As New Point(594, 899)

    'In versus game playing variables
    Private udcCharacterOne As clsCharacter 'Host
    Private udcCharacterTwo As clsCharacter 'Join
    Private intZombieSpeedOne As Integer = 10
    Private intZombieSpeedTwo As Integer = 10
    Private btmYouWon As Bitmap
    Private pntYouWon As New Point(498, 875)
    Private btmYouLost As Bitmap
    Private pntYouLost As New Point(486, 872)
    Private btmVersusWhoWon As Bitmap
    Private intZombieKillsOne As Integer = 0
    Private intZombieKillsTwo As Integer = 0

    'Story
    Private btmStoryBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Story\StoryBackground.jpg"))
    Private btmStoryParagraph1_25 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph1_25.png"))
    Private btmStoryParagraph1_50 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph1_50.png"))
    Private btmStoryParagraph1_75 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph1_75.png"))
    Private btmStoryParagraph1_100 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph1_100.png"))
    Private btmStoryParagraph2_25 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph2_25.png"))
    Private btmStoryParagraph2_50 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph2_50.png"))
    Private btmStoryParagraph2_75 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph2_75.png"))
    Private btmStoryParagraph2_100 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph2_100.png"))
    Private btmStoryParagraph3_25 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph3_25.png"))
    Private btmStoryParagraph3_50 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph3_50.png"))
    Private btmStoryParagraph3_75 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph3_75.png"))
    Private btmStoryParagraph3_100 As New Bitmap(Image.FromFile(strDirectory & "Images\Story\Paragraph3_100.png"))
    Private btmStoryParagraph As Bitmap
    Private pntStoryParagraph As Point 'Set in the story thread
    Private thrStory As System.Threading.Thread
    Private udcStoryParagraph1Sound As clsSound
    Private udcStoryParagraph2Sound As clsSound
    Private udcStoryParagraph3Sound As clsSound

    'Game version mismatch
    Private btmGameMismatchBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Game Mismatch\GameMismatchBackground.jpg"))
    Private strGameVersionFromConnection As String = ""
    Private thrGameMismatch As System.Threading.Thread

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes:'Do not allow a resize of the window if full screen, this happens if you double click the title, 
        '       but is prevented with this sub, prevents moving the window around too

        'Check if full screen
        If intResolutionMode = 5 Then
            'Prevent moving the form by control box click
            If (m.Msg = WINDOW_MESSAGE_SYSTEM_COMMAND) And (m.WParam.ToInt32() = CONTROL_MOVE) Then
                Return
            End If
            'Prevent button down moving form
            If (m.Msg = WINDOW_MESSAGE_CLICK_BUTTON_DOWN) And (m.WParam.ToInt32() = WINDOW_CAPTION) Then
                Return
            End If
        End If

        'If a double click on the title bar is triggered
        If m.Msg = WINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
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

        'Load loading bar pictures
        LoadLoadingBarPictures()

        'Set for loading
        btmLoadingBar = btmLoadingBarPicture(0)

    End Sub

    Private Sub LoadLoadingBarPictures()

        'Declare
        Dim intPictureCounter As Integer = 0

        'Load loading bar pictures
        For intLoop As Integer = 0 To btmLoadingBarPicture.GetUpperBound(0)
            'Set
            btmLoadingBarPicture(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\" & CStr(intPictureCounter) & ".png"))
            'Increase
            intPictureCounter += 10
        Next

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Fog
        AbortThread(thrFog)

        'Set
        thrFog = Nothing

        'Rendering
        AbortThread(thrRendering)

        'Set
        thrRendering = Nothing

        'Sound delay
        AbortThread(thrHoverSoundDelay)

        'Set
        thrHoverSoundDelay = Nothing

        'Loading game
        AbortThread(thrLoadingGame)

        'Set
        thrLoadingGame = Nothing

        'Loading
        For intLoop As Integer = 0 To thrLoading.GetUpperBound(0)
            AbortThread(thrLoading(intLoop))
            'Set
            thrLoading(intLoop) = Nothing
        Next

        'Paragraphing
        AbortThread(thrParagraph)

        'Set
        thrParagraph = Nothing

        'Ending game
        AbortThread(thrEndGame)

        'Set
        thrEndGame = Nothing

        'Credits
        AbortThread(thrCredits)

        'Set
        thrCredits = Nothing

        'Story
        AbortThread(thrStory)

        'Set
        thrStory = Nothing

        'Level completed
        AbortThread(thrBlackScreen)

        'Set
        thrBlackScreen = Nothing

        'Remove game objects
        RemoveGameObjectsFromMemory()

        'Empty versus multiplayer variables
        EmptyMultiplayerVariables()

        'Stop ambient music
        If udcAmbianceSound IsNot Nothing Then
            udcAmbianceSound.StopAndCloseSound()
            udcAmbianceSound = Nothing
        End If

        'Stop ambient options music
        If udcAmbianceOptionsSound IsNot Nothing Then
            udcAmbianceOptionsSound.StopAndCloseSound()
            udcAmbianceOptionsSound = Nothing
        End If

        'Dipose graphics
        gGraphics.Dispose()

    End Sub

    Private Sub AbortThread(thrToAbort As System.Threading.Thread)

        'Check thread
        If thrToAbort IsNot Nothing Then
            If thrToAbort.IsAlive Then
                thrToAbort.Abort()
            End If
        End If

    End Sub

    Private Sub EmptyMultiplayerVariables()

        'Loading versus
        AbortThread(thrLoadingVersus)

        'Set
        thrLoadingVersus = Nothing

        'Paragraph versus
        AbortThread(thrParagraphVersus)

        'Set
        thrParagraphVersus = Nothing

        'Listening TCP/IP
        AbortThread(thrListening)

        'Set
        thrListening = Nothing

        'Connecting TCP/IP
        AbortThread(thrConnecting)

        'Set
        thrConnecting = Nothing

        'Mismatch
        AbortThread(thrGameMismatch)

        'Set
        thrGameMismatch = Nothing

        'Empty versus multiplayer variables
        If gswClientData IsNot Nothing Then
            gswClientData.Close()
            gswClientData = Nothing
        End If
        If srClientData IsNot Nothing Then
            srClientData.Close()
            srClientData = Nothing
        End If
        If tcplServer IsNot Nothing Then
            tcplServer.Stop()
            tcplServer = Nothing
        End If
        If tcpcClient IsNot Nothing Then
            tcpcClient.Close()
            tcpcClient = Nothing
        End If

        'Make this last because of try and end try can create glitch problems
        If udcVersusConnectedThread IsNot Nothing Then
            'Remove handlers
            RemoveHandler udcVersusConnectedThread.GotData, AddressOf DataArrival
            RemoveHandler udcVersusConnectedThread.ConnectionGone, AddressOf ConnectionLost
            'Abort threads
            udcVersusConnectedThread.AbortCheckConnectionThread()
            udcVersusConnectedThread.AbortReadLineThread()
            udcVersusConnectedThread = Nothing
        End If

        'Wait or else a complication will start when reconnecting after an exit of versus game
        System.Threading.Thread.Sleep(1)

    End Sub

    Private Sub RemoveGameObjectsFromMemory(Optional blnStopScreamSound As Boolean = True)

        'Stop sound effects
        If blnStopScreamSound Then
            If udcScreamSound IsNot Nothing Then
                udcScreamSound.StopAndCloseSound()
                udcScreamSound = Nothing
            End If
        End If

        'Stop game background sounds
        For intLoop As Integer = 0 To udcGameBackgroundSound.GetUpperBound(0)
            If udcGameBackgroundSound(intLoop) IsNot Nothing Then
                udcGameBackgroundSound(intLoop).StopAndCloseSound()
                udcGameBackgroundSound(intLoop) = Nothing
            End If
        Next

        'Stop and dispose helicopter
        If udcHelicopter IsNot Nothing Then
            udcHelicopter.StopRotatingBladeSound()
            'Stop and dispose
            udcHelicopter.StopAndDispose()
            udcHelicopter = Nothing
        End If

        'Stop and dispose character, remove handler
        If udcCharacter IsNot Nothing Then
            'Stop reloading sound
            udcCharacter.StopReloadingSound()
            'Stop and dispose
            udcCharacter.StopAndDispose()
            udcCharacter = Nothing
        End If

        'Stop and dispose zombies
        If gaudcZombies IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
                If gaudcZombies(intLoop) IsNot Nothing Then
                    gaudcZombies(intLoop).StopAndDispose()
                    gaudcZombies(intLoop) = Nothing
                End If
            Next
        End If

        'Stop and dispose versus characters, remove handler
        If udcCharacterOne IsNot Nothing Then
            'Stop reloading sound
            udcCharacterOne.StopReloadingSound()
            'Stop and dispose
            udcCharacterOne.StopAndDispose()
            udcCharacterOne = Nothing
        End If
        If udcCharacterTwo IsNot Nothing Then
            'Stop reloading sound
            udcCharacterTwo.StopReloadingSound()
            'Stop and dispose
            udcCharacterTwo.StopAndDispose()
            udcCharacterTwo = Nothing
        End If

        'Stop and dispose zombies
        If gaudcZombiesOne IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                If gaudcZombiesOne(intLoop) IsNot Nothing Then
                    gaudcZombiesOne(intLoop).StopAndDispose()
                    gaudcZombiesOne(intLoop) = Nothing
                End If
            Next
        End If
        If gaudcZombiesTwo IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
                If gaudcZombiesTwo(intLoop) IsNot Nothing Then
                    gaudcZombiesTwo(intLoop).StopAndDispose()
                    gaudcZombiesTwo(intLoop) = Nothing
                End If
            Next
        End If

        'Stop and dispose zombies that produce game lose
        If udcUnderwaterZombieGameLose IsNot Nothing Then
            udcUnderwaterZombieGameLose.StopAndDispose()
            udcUnderwaterZombieGameLose = Nothing
        End If

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
            'Get highscores early because grabbing information from the database access files is slow
            LoadHighscoresStringAccess()
            'Set percentage multiplers for screen modes
            gdblScreenWidthRatio = CDbl((intScreenWidth - WIDTH_SUBTRACTION) / ORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((intScreenHeight - HEIGHT_SUBTRACTION) / ORIGINAL_SCREEN_HEIGHT)
            'Menu sound
            udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - WIDTH_SUBTRACTION, intScreenHeight - HEIGHT_SUBTRACTION) 'Full screen
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
                'Set
                btmLoadingParagraphVersus = Nothing
                'Set
                blnConnectionCompleted = False
                'Set
                blnGameIsVersus = False
                'Set
                blnReadyEarly = False
                'Set
                blnWaiting = False
                'Set
                btmVersusWhoWon = Nothing
                'Set
                intCanvasVersusShow = 0
                'End game abort
                AbortThread(thrEndGame)
                'Set
                thrEndGame = Nothing
                'Level completed
                AbortThread(thrBlackScreen)
                'Set
                thrBlackScreen = Nothing
                'Stop and dispose game objects
                RemoveGameObjectsFromMemory()
                'Empty versus multiplayer variables
                EmptyMultiplayerVariables()
                'Set
                btmLoadingParagraph = Nothing
                'Set
                btmLoadingBar = btmLoadingBarPicture(0)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0, blnPlayPressedSoundNow)
                'Set
                blnPlayPressedSoundNow = False
                'Restart fog
                RestartFog()
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
                Case 6 'Versus screen
                    VersusScreen()
                Case 7 'Loading versus connected game
                    LoadingVersusConnectedScreen()
                Case 8 'Playing versus game
                    StartedVersusGameScreen()
                Case 9 'Story
                    StoryScreen()
                Case 10 'Mismatch game versions
                    GameVersionMismatch()
                Case 11, 12 'Path system
                    PathChoices()
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
        PaintOnBitmap(btmCanvas)

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
        PaintOnBitmap(btmCanvas)

        'Draw options background
        gGraphics.DrawImageUnscaled(btmOptionsBackground, pntTopLeft)

        'Draw resolution text
        gGraphics.DrawImageUnscaled(btmResolutionText, pntResolutionText)

        'Draw sound text
        gGraphics.DrawImageUnscaled(btmSoundText, pntSoundText)

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

        'Draw version
        DrawText(gGraphics, "Version " & GAME_VERSION, 50, Color.Black, 33, 953, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Version " & GAME_VERSION, 50, Color.Black, 35, 955, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Version " & GAME_VERSION, 50, Color.Red, 30, 950, 1000, 125) 'Overlay

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
        PaintOnBitmap(btmCanvas)

        'Draw loading background
        gGraphics.DrawImageUnscaled(btmLoadingBackground, pntTopLeft)

        'Draw loading bar
        gGraphics.DrawImageUnscaled(btmLoadingBar, pntLoadingBar)

        'Draw Loading text
        If intCanvasShow = 0 And intCanvasMode = 2 Then
            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
        Else
            gGraphics.DrawImageUnscaled(btmLoadingStartText, pntLoadingStartText)
        End If

        'Draw paragraph
        If btmLoadingParagraph IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmLoadingParagraph, pntLoadingParagraph)
        End If

    End Sub

    Private Sub PathChoices()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw background
        gGraphics.DrawImageUnscaled(btmPath, pntTopLeft)

        'Text
        DrawText(gGraphics, "Pick your path...", 50, Color.Black, 33, 953, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Pick your path...", 50, Color.Black, 35, 955, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Pick your path...", 50, Color.Red, 30, 950, 1000, 125) 'Overlay

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub KillZombiesMarkedToDie(audcZombiesType() As clsZombie, udcCharacterType As clsCharacter)

        'Kill zombies that were marked
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            If audcZombiesType(intLoop).MarkedToDie Then
                If Not audcZombiesType(intLoop).HasPassedMarkToDie Then
                    'Set
                    audcZombiesType(intLoop).HasPassedMarkToDie = True
                    'Set to die in frames
                    audcZombiesType(intLoop).Dying()
                    'Character shoots
                    udcCharacterType.Shoot()
                    'Check if needs to reload
                    If udcCharacterType.BulletsUsed = 30 Then
                        udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload
                    End If
                    'Exit
                    Exit Sub
                End If
            End If
        Next

    End Sub

    Private Sub PaintOnBitmap(btmToPaintOn As Bitmap)

        'What to draw on
        gGraphics = Graphics.FromImage(btmToPaintOn)

        'Set graphic options
        SetDefaultGraphicOptions()

    End Sub

    Private Sub DrawStats(gGraphicsDevice As Graphics, intZombieKillsToBe As Integer, intReloadTimeUsedToBe As Integer)

        'Draw zombie kills
        DrawText(gGraphicsDevice, CStr(intZombieKillsToBe), 85, Color.Black, 1065, 417, 1000, 125)
        DrawText(gGraphicsDevice, CStr(intZombieKillsToBe), 85, Color.White, 1060, 412, 1000, 125)

        'Declare
        Dim tsTimeSpan As TimeSpan
        Dim strElapsedTime As String

        'Set
        tsTimeSpan = swhStopWatch.Elapsed
        strElapsedTime = CStr(CInt(tsTimeSpan.TotalSeconds)) & " Sec"

        'Draw survival time
        DrawText(gGraphicsDevice, strElapsedTime, 85, Color.Black, 1065, 615, 1000, 125)
        DrawText(gGraphicsDevice, strElapsedTime, 85, Color.White, 1060, 610, 1000, 125)

        'Declare
        Dim intElapsedTime As Integer = 0
        Dim intWPM As Integer = 0

        'Get elapsed time
        intElapsedTime = CInt(Replace(strElapsedTime, " Sec", "")) - intReloadTimeUsedToBe

        'Set WPM
        intWPM = CInt((intZombieKillsToBe / (intElapsedTime / 60)))

        'Draw WPM
        DrawText(gGraphics, CStr(intWPM), 85, Color.Black, 1065, 808, 1000, 125)
        DrawText(gGraphics, CStr(intWPM), 85, Color.White, 1060, 803, 1000, 125)

    End Sub

    Private Sub StartedGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Paint on the canvas
            PaintOnBitmap(btmCanvas)

            'Check if beat level
            If blnBeatLevel Then

                'Paint black background
                gGraphics.DrawImageUnscaled(btmBlackScreen100, pntTopLeft)

                'Check which level was completed
                Select Case intLevel
                    Case 1
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Set
                        btmPath = btmPath1_0
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(11, 0, False) 'This exits the game screen to path choices
                    Case 2, 3
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Set
                        btmPath = btmPath2_0
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(12, 0, False) 'This exits the game screen to path choices
                    Case 5
                        'Show win overlay
                        gGraphics.DrawImageUnscaled(btmWinOverlay, pntTopLeft)
                        'Draw stats
                        DrawStats(gGraphics, intZombieKillsCombined, intReloadTimeUsedCombined)
                End Select

            Else

                'Show death overlay
                gGraphics.DrawImageUnscaled(btmDeathScreen, pntTopLeft)
                gGraphics.DrawImageUnscaled(btmDeathOverlay, pntTopLeft)

                'Draw stats
                DrawStats(gGraphics, intZombieKillsCombined, intReloadTimeUsedCombined)

            End If

            'Remove objects only once
            If Not blnRemovedGameObjectsFromMemory Then
                'Set
                blnRemovedGameObjectsFromMemory = True
                'Remove objects
                RemoveGameObjectsFromMemory(False) 'Don't stop the scream sound here
            End If

        Else

            'Paint on copy of the background
            PaintOnBitmap(btmGameBackground)

            'Check for helicopter
            If udcHelicopter IsNot Nothing Then
                'Draw
                gGraphics.DrawImageUnscaled(udcHelicopter.HelicopterImage, udcHelicopter.HelicopterPoint)
            End If

            'Check status of character
            CharacterMovementWithStatus()

            'Draw dead zombies permanently
            For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
                If gaudcZombies(intLoop).Spawned Then
                    If gaudcZombies(intLoop).IsDead Then
                        'Set
                        gaudcZombies(intLoop).Spawned = False
                        'Set new point
                        Dim pntTemp As New Point(gaudcZombies(intLoop).ZombiePoint.X + CInt(Math.Abs(gpntGameBackground.X)), gaudcZombies(intLoop).ZombiePoint.Y)
                        'Draw dead
                        gGraphics.DrawImageUnscaled(gaudcZombies(intLoop).ZombieImage, pntTemp)
                        'Increase count
                        intZombieKills += 1
                        'Start a new zombie
                        gaudcZombies(intZombieKills + NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
                    End If
                End If
            Next

            'Paint background and word
            PaintBackgroundAndWord()

            'Draw character
            gGraphics.DrawImageUnscaled(udcCharacter.CharacterImage, udcCharacter.CharacterPoint)

            'Check if made it to the end of the level
            If gpntGameBackground.X <= -2850 Then
                'Check if level was already beaten
                If Not blnBeatLevel Then
                    'Set
                    blnBeatLevel = True
                    'Set
                    blnEndingGameCantType = True
                    'Stop the character moving
                    udcCharacter.Stand()
                    'Exit sound
                    Select Case intLevel
                        Case 1, 2
                            'Play door
                            Dim udcOpeningMetalDoor As New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\OpeningMetalDoor.mp3", 3000, gintSoundVolume)
                    End Select
                    'Start black screen thread
                    thrBlackScreen = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf BlackScreening))
                    thrBlackScreen.Start()
                    'Stop level music
                    udcGameBackgroundSound(intLevel - 1).StopAndCloseSound()
                    udcGameBackgroundSound(intLevel - 1) = Nothing
                    'Pause the stop watch
                    swhStopWatch.Stop()
                    'Set reload time
                    intReloadTimeUsed = udcCharacter.ReloadTimes * 3
                    'Keep the reload time updated
                    intReloadTimeUsedCombined += intReloadTimeUsed
                    'Keep the zombie kills updated
                    intZombieKillsCombined += intZombieKills
                End If
            End If

            'Draw zombies
            For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
                'Check if spawned
                If gaudcZombies(intLoop).Spawned Then
                    'Check if can pin
                    If Not gaudcZombies(intLoop).MarkedToDie Then
                        'Check distance
                        If gaudcZombies(intLoop).ZombiePoint.X <= ZOMBIE_PINNING_X_DISTANCE And Not gaudcZombies(intLoop).IsPinning Then
                            'Check if level not beat
                            If Not blnBeatLevel Then
                                'Check
                                SetXYPositionOfPinningZombie(gaudcZombies, intLoop, 1)
                                'Set
                                gaudcZombies(intLoop).Pin()
                                'Check if first time game over by pin
                                If Not blnTriggeredBlackScreenThread Then
                                    'Set
                                    blnTriggeredBlackScreenThread = True
                                    'Set
                                    blnEndingGameCantType = True
                                    'Stop character from moving
                                    If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                                        udcCharacter.Stand()
                                    End If
                                    'Start black screen thread
                                    thrBlackScreen = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf BlackScreening))
                                    thrBlackScreen.Start()
                                    'Stop reloading sound
                                    udcCharacter.StopReloadingSound()
                                    'Play
                                    udcScreamSound = New clsSound(Me, strDirectory & "Sounds\CharacterDying.mp3", 3000, gintSoundVolume, False)
                                    'Stop level music
                                    udcGameBackgroundSound(intLevel - 1).StopAndCloseSound()
                                    udcGameBackgroundSound(intLevel - 1) = Nothing
                                    'Pause the stop watch
                                    swhStopWatch.Stop()
                                    'Set reload time
                                    intReloadTimeUsed = udcCharacter.ReloadTimes * 3
                                    'Keep the reload time updated
                                    intReloadTimeUsedCombined += intReloadTimeUsed
                                    'Keep the zombie kills updated
                                    intZombieKillsCombined += intZombieKills
                                End If
                            End If
                        End If
                    End If
                    'Draw zombies dying, pinning or walking
                    gGraphics.DrawImageUnscaled(gaudcZombies(intLoop).ZombieImage, gaudcZombies(intLoop).ZombiePoint)
                End If
            Next

            'Show magazine with bullet count
            gGraphics.DrawImageUnscaled(btmAK47Magazine, pntAK47Magazine)

            'Draw bullet count on magazine
            DrawText(gGraphics, CStr(30 - udcCharacter.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)

            'Check if black screen needs to be drawed
            If btmBlackScreen IsNot Nothing Then
                gGraphics.DrawImageUnscaled(btmBlackScreen, pntTopLeft)
            End If

        End If

        'Make copy if died
        If blnTriggeredBlackScreenThread Then
            'Only copy if not existent for first time
            If btmBlackScreen Is btmBlackScreen50 Then
                If btmDeathScreen Is Nothing Then
                    'Before fading the screen, copy it to show for the death overlay
                    Dim rectSource As New Rectangle(0, 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)
                    btmDeathScreen = btmCanvas.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)
                End If
            End If
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

    Private Sub CharacterMovementWithStatus()

        'Check status of character
        If Not blnTriggeredBlackScreenThread Then 'If game not over

            'Check if need to stop running
            If udcCharacter.StopCharacterFromRunning Then
                'Set
                udcCharacter.StopCharacterFromRunning = False
                'Stand, stop running
                udcCharacter.Stand()
                'Exit
                Exit Sub
            End If

            'Check the mode of character
            Select Case udcCharacter.StatusModeStartToDo

                Case clsCharacter.eintStatusMode.Stand
                    'This is neutral, but check for things that are about to be done
                    Select Case udcCharacter.StatusModeAboutToDo
                        Case clsCharacter.eintStatusMode.Reload
                            'Check frame first
                            If udcCharacter.GetPictureFrame = 1 Then
                                'Reset
                                udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                                'Reload
                                udcCharacter.Reload()
                            End If
                        Case clsCharacter.eintStatusMode.Run
                            'Check frame first
                            If udcCharacter.GetPictureFrame = 1 Then
                                'Reset
                                udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                                'Run
                                udcCharacter.Run()
                            End If
                    End Select

                Case clsCharacter.eintStatusMode.Reload
                    'Reload
                    udcCharacter.Reload()

                Case clsCharacter.eintStatusMode.Shoot
                    'Kill zombies that were marked and shoot
                    KillZombiesMarkedToDie(gaudcZombies, udcCharacter) 'Has shoot inside the sub procedure

                Case clsCharacter.eintStatusMode.Run
                    'Run
                    udcCharacter.Run()

            End Select

        End If

    End Sub

    Private Function GetNumberToPin(audcZombiesType() As clsZombie, intNumberToCheck As Integer) As Boolean

        'Loop to get the correct loop number
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            If audcZombiesType(intLoop).Spawned Then
                If Not audcZombiesType(intLoop).MarkedToDie Then
                    If audcZombiesType(intLoop).PinXYValueChanged = intNumberToCheck Then
                        Return True
                    End If
                End If
            End If
        Next

        'Else
        Return False

    End Function

    Private Sub SetXYPositionOfPinningZombie(audcZombiesType() As clsZombie, intLoop As Integer, intPinNumber As Integer)

        'Notes: Recursive sub that sets a new x, y for each zombie pinning the character currently

        'Check
        If Not GetNumberToPin(audcZombiesType, intPinNumber) Then
            'Change
            If intPinNumber <> 1 Then
                audcZombiesType(intLoop).ZombiePoint = New Point(audcZombiesType(intLoop).ZombiePoint.X - (20 * (intPinNumber - 1)),
                                                       audcZombiesType(intLoop).ZombiePoint.Y)
            End If
            'Set
            audcZombiesType(intLoop).PinXYValueChanged = intPinNumber
        Else
            SetXYPositionOfPinningZombie(audcZombiesType, intLoop, intPinNumber + 1)
        End If

    End Sub

    Private Sub CheckingForFirstTimePin(blnEndGame As Boolean, Optional strData As String = "")

        'Check if happened first time
        If blnGameOverFirstTime Then
            'Set
            blnEndingGameCantType = True
            'Set
            blnGameOverFirstTime = False
            'Stop reloading sound and send data
            If blnGameIsVersus Then
                'Send data
                gSendData(strData)
                'Stop sound
                udcCharacterOne.StopReloadingSound()
                udcCharacterTwo.StopReloadingSound()
            Else
                'Stop sound
                udcCharacter.StopReloadingSound()
            End If
            'Play
            udcScreamSound = New clsSound(Me, strDirectory & "Sounds\CharacterDying.mp3", 3000, gintSoundVolume, False)
            'Set
            thrEndGame = New System.Threading.Thread(Sub() EndingGame(blnEndGame))
            thrEndGame.Start()
        End If

    End Sub

    Private Sub PaintBackgroundAndWord()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)

        'Clone the only necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackground.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)

        'Draw the background to screen with the cloned version
        gGraphics.DrawImageUnscaled(btmGameBackgroundCloneScreenShown, pntTopLeft)

        'Dispose
        btmGameBackgroundCloneScreenShown.Dispose()
        btmGameBackgroundCloneScreenShown = Nothing

        'Draw word bar
        gGraphics.DrawImageUnscaled(btmWordBar, pntWordBar)

        'Draw text in the word bar
        DrawText(gGraphics, strTheWord, 50, Color.Black, 530, 95, 1000, 100) 'Shadow
        DrawText(gGraphics, strTheWord, 50, Color.Black, 528, 93, 1000, 100) 'Shadow
        DrawText(gGraphics, strTheWord, 50, Color.White, 525, 90, 1000, 100) 'White text

        'Word overlay
        If blnGameIsVersus Then
            If blnHost Then
                DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Red, 525, 90, 1000, 100) 'Overlay
            Else
                DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Blue, 525, 90, 1000, 100) 'Overlay
            End If
        Else
            DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Red, 525, 90, 1000, 100) 'Overlay
        End If

    End Sub

    'Private Sub DrawEndGameScreen(intZombieKillsType As Integer, udcCharacterType As clsCharacter, Optional intEscapeType As Integer = 0)

    '    'Paint on canvas
    '    PaintOnBitmap(btmCanvas)

    '    'Screen is black, show death overlay
    '    If intEscapeType = 1 Then
    '        gGraphics.DrawImageUnscaled(btmWinOverlay, pntTopLeft)
    '    Else
    '        Select Case intLevel
    '            Case 1
    '                gGraphics.DrawImageUnscaled(gbtmDeathOverlay(0), pntTopLeft)
    '            Case 2
    '                gGraphics.DrawImageUnscaled(gbtmDeathOverlay(1), pntTopLeft)
    '            Case 3
    '                gGraphics.DrawImageUnscaled(gbtmDeathOverlay(2), pntTopLeft)
    '            Case 4
    '                gGraphics.DrawImageUnscaled(gbtmDeathOverlay(3), pntTopLeft)
    '            Case 5
    '                gGraphics.DrawImageUnscaled(gbtmDeathOverlay(4), pntTopLeft)
    '        End Select
    '    End If

    '    'Draw zombie kills
    '    DrawText(gGraphics, CStr(intZombieKillsType), 85, Color.Black, 915, 437, 1000, 125)
    '    DrawText(gGraphics, CStr(intZombieKillsType), 85, Color.White, 910, 432, 1000, 125)

    '    'Stop the watch
    '    StopTheWatch()

    '    'Draw survival time
    '    DrawText(gGraphics, strElapsedTime, 85, Color.Black, 925, 605, 1000, 125)
    '    DrawText(gGraphics, strElapsedTime, 85, Color.White, 920, 600, 1000, 125)

    '    'Declare
    '    Dim intElapsedTime As Integer = 0
    '    Dim intWPM As Integer = 0

    '    'Get elapsed time
    '    If intReloadTimeUsedCombined = 0 Then
    '        intElapsedTime = CInt(Replace(strElapsedTime, " Sec", "")) - intReloadTimeUsed
    '    Else
    '        intElapsedTime = CInt(Replace(strElapsedTime, " Sec", "")) - intReloadTimeUsedCombined
    '    End If

    '    'Set WPM
    '    If Not blnGameIsVersus Then
    '        If intZombieKillsCombined = 0 Then
    '            intWPM = CInt((intZombieKillsType / (intElapsedTime / 60)))
    '        Else
    '            intWPM = CInt((intZombieKillsCombined / (intElapsedTime / 60)))
    '        End If
    '    Else
    '        intWPM = CInt((intZombieKillsType / (intElapsedTime / 60)))
    '    End If

    '    'Draw WPM
    '    DrawText(gGraphics, CStr(intWPM), 85, Color.Black, 925, 763, 1000, 125)
    '    DrawText(gGraphics, CStr(intWPM), 85, Color.White, 920, 758, 1000, 125)

    '    'Draw who won if versus game
    '    If blnGameIsVersus Then
    '        If btmVersusWhoWon IsNot Nothing Then
    '            If btmVersusWhoWon Is btmYouWon Then
    '                gGraphics.DrawImageUnscaled(btmVersusWhoWon, pntYouWon)
    '            Else
    '                gGraphics.DrawImageUnscaled(btmVersusWhoWon, pntYouLost)
    '            End If
    '        End If
    '    Else
    '        'Check if highscores was already beaten
    '        If Not blnComparedHighscore Then
    '            'Set
    '            blnComparedHighscore = True
    '            'Compare score with highscores
    '            CompareHighscoresAccess(intWPM)
    '        End If
    '    End If

    'End Sub

    Private Sub CompareHighscoresTextFile(intWPMToBe As Integer)

        'Notes: If there was an error, the database content from the text file will load after. 
        '       This is a backup to not having the correct database path, an install and registered Microsoft Provider 12.0.

        'Declare
        Dim ioSR As IO.StreamReader
        Dim strTemp As String = ""
        Dim astrTempSplit() As String
        Dim intRank As Integer = 0
        Dim blnRankBeat As Boolean = False

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Images\Highscores\Highscores.txt")

        'Use loop to get zombie kills
        While ioSR.Peek <> -1
            'Check if line is empty
            strTemp = ioSR.ReadLine
            'Check
            If strTemp = "" Then
                'Exit
                Exit While
            End If
            'Get zombie kills
            astrTempSplit = strTemp.Split(CChar(" ")) 'Element 3 is zombie kills
            'Increase
            intRank += 1
            'Check
            If intZombieKills > CInt(astrTempSplit(3)) Then
                'Set
                blnRankBeat = True
                'Exit
                Exit While
            End If
        End While

        'Close
        ioSR.Close()

        'Check if rank beat
        If blnRankBeat Then
            WriteToDatabase(intRank, intWPMToBe)
        End If

    End Sub

    Private Sub CompareHighscoresAccess(intWPMToBe As Integer)

        'Declare
        Dim strSQL As String = "SELECT Kills FROM HighscoresTable ORDER BY Rank ASC"
        Dim strConnection As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = Highscores.accdb"
        Dim dtProperties As New DataTable()
        Dim dbDataAdapter As System.Data.OleDb.OleDbDataAdapter
        Dim intRank As Integer = 0
        Dim blnRankBeat As Boolean = False

        'Set
        blnHighscoresIsAccess = False

        'Prevent errors
        Try

            'Set
            dbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(strSQL, strConnection)

            'Set
            dbDataAdapter.Fill(dtProperties)

            'Compare
            For intLoop As Integer = 0 To dtProperties.Rows.Count - 1
                If intZombieKills > CInt(dtProperties.Rows(intLoop).Item(0).ToString) Then
                    'Set
                    intRank = intLoop + 1
                    'Set
                    blnRankBeat = True
                    'Exit
                    Exit For
                End If
            Next

            'Set
            blnHighscoresIsAccess = True

        Catch ex As Exception
            'Error, possibly this computer doesn't have Microsoft Provider 12.0 installed or registered
            CompareHighscoresTextFile(intWPMToBe)

        End Try

        'Check if rank beat
        If blnHighscoresIsAccess Then
            If blnRankBeat Then
                WriteToDatabase(intRank, intWPMToBe)
            End If
        End If

    End Sub

    Private Sub WriteToDatabase(intRank As Integer, intWPMToBe As Integer)


        'Declare
        Dim strInputBox As String = ""

        'Get user info
        While strInputBox = ""
            strInputBox = InputBox("You beat a highscore, please enter your name...", "Last Stand", "")
            'Check for invalid string characters
            If CheckStringForInvalidCharacters(strInputBox) Then
                MessageBox.Show("Can only accept standard characters. Please try again...", "Last Stand", MessageBoxButtons.OK, MessageBoxIcon.Error)
                strInputBox = ""
            End If
        End While

        'Check
        If blnHighscoresIsAccess Then
            'Enter the information into the database
            EnterTheInformationIntoTheDatabaseAccess(strInputBox, intRank, intWPMToBe)
            'Reload the highscore database
            LoadHighscoresStringAccess(False)
        Else
            'Enter the information into the database
            EnterTheInformationIntoTheDatabaseTextFile(strInputBox, intRank, intWPMToBe)
            'Reload the highscore database
            LoadHighscoresStringTextFile()
        End If

    End Sub

    Private Function CheckStringForInvalidCharacters(strToCheck As String) As Boolean

        'Declare
        Dim astrInvalidCharacters(28) As String

        'Announce
        astrInvalidCharacters(0) = "'"
        astrInvalidCharacters(1) = "`"
        astrInvalidCharacters(2) = Chr(34) & Chr(34) 'Chr(34) & Chr(34) = single quote
        astrInvalidCharacters(3) = "|"
        astrInvalidCharacters(4) = ":"
        astrInvalidCharacters(5) = ";"
        astrInvalidCharacters(6) = "-"
        astrInvalidCharacters(7) = "="
        astrInvalidCharacters(8) = "*"
        astrInvalidCharacters(9) = "/"
        astrInvalidCharacters(10) = "\"
        astrInvalidCharacters(11) = ","
        astrInvalidCharacters(12) = "{"
        astrInvalidCharacters(13) = "}"
        astrInvalidCharacters(14) = "_"
        astrInvalidCharacters(15) = "("
        astrInvalidCharacters(16) = ")"
        astrInvalidCharacters(17) = "&"
        astrInvalidCharacters(18) = "+"
        astrInvalidCharacters(19) = "~"
        astrInvalidCharacters(20) = "!"
        astrInvalidCharacters(21) = "@"
        astrInvalidCharacters(22) = "#"
        astrInvalidCharacters(23) = "$"
        astrInvalidCharacters(24) = "%"
        astrInvalidCharacters(25) = "^"
        astrInvalidCharacters(26) = "?"
        astrInvalidCharacters(27) = "<"
        astrInvalidCharacters(28) = ">"

        'Check the string
        For intLoop As Integer = 0 To astrInvalidCharacters.GetUpperBound(0)
            If InStr(strToCheck, astrInvalidCharacters(intLoop)) > 0 Then
                Return True
            End If
        Next

        'If made it here
        Return False

    End Function

    Private Sub EnterTheInformationIntoTheDatabaseAccess(strName As String, intRank As Integer, intWPMToBe As Integer)

        'Declare
        Dim strSQL As String = "UPDATE HighscoresTable SET [Player Name]='" & strName & "',WPM=" & intWPMToBe & ",Kills=" & intZombieKills & " WHERE Rank=" & intRank
        Dim strConnection As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = Highscores.accdb"
        Dim dtProperties As New DataTable()
        Dim dbDataAdapter As System.Data.OleDb.OleDbDataAdapter

        'Prevent errors
        Try

            'Set
            dbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(strSQL, strConnection)

            'Set
            dbDataAdapter.Fill(dtProperties)

            'Set
            dtProperties.AcceptChanges()

        Catch ex As Exception
            'No debug, previously checked before made it here

        End Try

    End Sub

    Private Sub EnterTheInformationIntoTheDatabaseTextFile(strName As String, intRank As Integer, intWPMToBe As Integer)

        'Notes: If there was an error, the database content from the text file will load after. 
        '       This is a backup to not having the correct database path, an install and registered Microsoft Provider 12.0.

        'Declare
        Dim ioSR As IO.StreamReader
        Dim astrLine() As String = Nothing

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Images\Highscores\Highscores.txt")

        'Use loop to get scores
        While ioSR.Peek <> -1
            If astrLine Is Nothing Then
                ReDim Preserve astrLine(0)
                astrLine(0) = ioSR.ReadLine
            Else
                ReDim Preserve astrLine(astrLine.GetUpperBound(0) + 1)
                astrLine(astrLine.GetUpperBound(0)) = ioSR.ReadLine
            End If
        End While

        'Close
        ioSR.Close()

        'Change part of the array where the rank was beat
        astrLine(intRank - 1) = CStr(intRank) & ". " & intWPMToBe & " WPM " & intZombieKills & " zombie kills - " & strName

        'Delete file
        Kill(strDirectory & "Images\Highscores\Highscores.txt")

        'Write to file
        Dim ioSW As IO.StreamWriter

        'Set
        ioSW = IO.File.AppendText(strDirectory & "Images\Highscores\Highscores.txt")

        'Write array
        For intLoop As Integer = 0 To astrLine.GetUpperBound(0)
            ioSW.WriteLine(astrLine(intLoop))
        Next

        'Close
        ioSW.Close()

    End Sub

    Private Sub BlackScreening()

        'Set
        btmBlackScreen = btmBlackScreen25

        'Wait
        System.Threading.Thread.Sleep(BLACK_SCREEN_DELAY)

        'Set
        btmBlackScreen = btmBlackScreen50

        'Wait
        System.Threading.Thread.Sleep(BLACK_SCREEN_DELAY)

        'Set
        btmBlackScreen = btmBlackScreen75

        'Wait
        System.Threading.Thread.Sleep(BLACK_SCREEN_DELAY)

        'Set
        btmBlackScreen = btmBlackScreen100

        'Wait
        System.Threading.Thread.Sleep(BLACK_SCREEN_DELAY)

        'Set
        blnBlackScreenFinished = True

    End Sub

    Private Sub EndingGame(Optional blnWon As Boolean = False)

        'Set
        btmBlackScreen = btmBlackScreen25

        'Wait
        System.Threading.Thread.Sleep(500)

        'Set
        btmBlackScreen = btmBlackScreen50

        'Wait
        System.Threading.Thread.Sleep(500)

        'Set
        btmBlackScreen = btmBlackScreen75

        'Wait
        System.Threading.Thread.Sleep(500)

        'Set
        btmBlackScreen = btmBlackScreen100

        'Check if versus game
        If blnGameIsVersus Then
            'Show won or lost
            If blnWon Then
                btmVersusWhoWon = btmYouWon
            Else
                btmVersusWhoWon = btmYouLost
            End If
        End If

    End Sub

    Private Sub DrawDatabaseType(strText As String)

        'Draw
        DrawText(gGraphics, strText, 34, Color.Black, 32, 22, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(gGraphics, strText, 34, Color.Black, 33, 23, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(gGraphics, strText, 34, Color.Black, 34, 24, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(gGraphics, strText, 34, Color.Black, 35, 25, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(gGraphics, strText, 34, Color.Red, 30, 20, ORIGINAL_SCREEN_WIDTH, 100)

    End Sub

    Private Sub HighscoresScreen()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw highscores background
        gGraphics.DrawImageUnscaled(btmHighscoresBackground, pntTopLeft)

        'Draw highscores
        DrawText(gGraphics, strHighscores, 34, Color.Black, 32, 222, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 33, 223, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 34, 224, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 35, 225, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Red, 30, 220, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)

        'Check if Access
        If blnHighscoresIsAccess Then
            DrawDatabaseType("Database type is Access")
        Else
            DrawDatabaseType("Database type is Text File")
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

    Private Sub LoadHighscoresStringAccess(Optional blnRunErrorCode As Boolean = True)

        'Set
        blnHighscoresIsAccess = False

        'Reset string
        strHighscores = ""

        'Declare
        Dim strSQL As String = "SELECT * FROM HighscoresTable ORDER BY Rank ASC"
        Dim strConnection As String = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = Highscores.accdb"
        Dim dtProperties As New DataTable()
        Dim dbDataAdapter As System.Data.OleDb.OleDbDataAdapter

        'Prevent errors
        Try

            'Set
            dbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(strSQL, strConnection)

            'Set
            dbDataAdapter.Fill(dtProperties)

            'Load string
            For intLoop As Integer = 0 To dtProperties.Rows.Count - 1
                'Add to string
                strHighscores &= dtProperties.Rows(intLoop).Item(0).ToString & ". " & dtProperties.Rows(intLoop).Item(1).ToString &
                                 " WPM " & dtProperties.Rows(intLoop).Item(3).ToString & " zombie kills - " & dtProperties.Rows(intLoop).Item(2).ToString
                'Add new line if necessary
                If intLoop <> dtProperties.Rows.Count - 1 Then
                    strHighscores &= vbNewLine
                End If
            Next

            'Set
            blnHighscoresIsAccess = True

        Catch ex As Exception
            'Error, possibly this computer doesn't have Microsoft Provider 12.0 installed or registered
            If blnRunErrorCode Then
                LoadHighscoresStringTextFile()
            End If

        End Try

    End Sub

    Private Sub LoadHighscoresStringTextFile()

        'Notes: If there was an error, the database content from the text file will load after. 
        '       This is a backup to not having the correct database path, an install and registered Microsoft Provider 12.0.

        'Reset
        strHighscores = ""

        'Declare
        Dim ioSR As IO.StreamReader

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Images\Highscores\Highscores.txt")

        'Use loop to get scores
        While ioSR.Peek <> -1
            strHighscores &= ioSR.ReadLine
            If ioSR.Peek <> -1 Then
                strHighscores &= vbNewLine
            End If
        End While

        'Close
        ioSR.Close()

    End Sub

    Private Sub CreditsScreen()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw credits background
        gGraphics.DrawImageUnscaled(btmCreditsBackground, pntTopLeft)

        'Draw credit pictures
        If btmJohnGonzales IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmJohnGonzales, pntJohnGonzales)
        End If
        If btmZacharyStafford IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmZacharyStafford, pntZacharyStafford)
        End If
        If btmCoryLewis IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmCoryLewis, pntCoryLewis)
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

    Private Sub CreditsFadeIn()

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales25
        btmZacharyStafford = btmZacharyStafford25
        btmCoryLewis = btmCoryLewis25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales50
        btmZacharyStafford = btmZacharyStafford50
        btmCoryLewis = btmCoryLewis50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales75
        btmZacharyStafford = btmZacharyStafford75
        btmCoryLewis = btmCoryLewis75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales100
        btmZacharyStafford = btmZacharyStafford100
        btmCoryLewis = btmCoryLewis100

    End Sub

    Private Sub VersusScreen()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw background
        gGraphics.DrawImageUnscaled(btmVersusBackground, pntTopLeft)

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

        'Other
        Select Case intCanvasVersusShow
            Case 0 'Nickname
                'Draw host button
                gGraphics.DrawImageUnscaled(btmVersusHost, pntVersusHost)
                'Draw or
                gGraphics.DrawImageUnscaled(btmVersusOr, pntVersusOr)
                'Draw join button
                gGraphics.DrawImageUnscaled(btmVersusJoin, pntVersusJoin)
                'Draw nickname
                gGraphics.DrawImageUnscaled(btmVersusNickname, pntVersusBlackOutline)
                'Draw player name text
                DrawText(gGraphics, strNickname, 110, Color.White, 103, 350, 1500, 275)
            Case 1 'Host
                'Draw hosting text
                DrawText(gGraphics, "Hosting...", 72, Color.Red, 600, 450, 1000, 150)
            Case 2 'Join
                'Draw IP address
                gGraphics.DrawImageUnscaled(btmVersusIPAddress, pntVersusBlackOutline)
                'Draw IP address to type
                DrawText(gGraphics, strIPAddressConnect, 110, Color.White, 103, 350, 1500, 275)
                'Draw connect button
                gGraphics.DrawImageUnscaled(btmVersusConnect, pntVersusConnect)
            Case 3 'Connecting
                'Draw connecting text
                DrawText(gGraphics, "Connecting...", 72, Color.Red, 500, 450, 1000, 150)
        End Select

        'Draw ip address text
        DrawText(gGraphics, strIPAddress, 72, Color.Red, 15, 40, 1000, 150)

        'Draw port forwarding
        DrawText(gGraphics, "Router port forwarding: 10101", 50, Color.White, 300, 875, 1200, 125)

    End Sub

    Private Sub DrawText(gGraphicsDevice As Graphics, strText As String, sngFontSize As Single, colColor As Color, sngX As Single, sngY As Single,
                         sngWidth As Single, sngHeight As Single)

        'Draw text
        Dim myFont As New Font("SketchFlow Print", sngFontSize, FontStyle.Regular)
        Dim myBrush As New System.Drawing.SolidBrush(colColor)
        gGraphics.DrawString(strText, myFont, myBrush, New RectangleF(sngX, sngY, sngWidth, sngHeight))

    End Sub

    Private Sub LoadingVersusConnectedScreen()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw loading background
        gGraphics.DrawImageUnscaled(btmLoadingBackground, pntTopLeft)

        'Draw loading bar
        gGraphics.DrawImageUnscaled(btmLoadingBar, pntLoadingBar)

        'Draw loading, waiting, and start text
        If blnWaiting Then
            gGraphics.DrawImageUnscaled(btmLoadingWaitingText, pntLoadingWaitingText)
        Else
            If intCanvasShow = 0 And intCanvasMode = 7 Then
                gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
            Else
                gGraphics.DrawImageUnscaled(btmLoadingStartText, pntLoadingStartText)
            End If
        End If

        'Draw paragraph
        If btmLoadingParagraphVersus IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmLoadingParagraphVersus, pntLoadingParagraph)
        End If

    End Sub

    Private Sub StartedVersusGameScreen()

        'Check for game over death overlay screen
        If btmBlackScreen Is btmBlackScreen100 Then

            'Draw end game screen
            If blnHost Then
                'Set
                If udcCharacterOne IsNot Nothing Then
                    intReloadTimeUsed = udcCharacterOne.ReloadTimes * 3
                End If
                'Print overlay
                'DrawEndGameScreen(intZombieKillsOne, udcCharacterOne)
            Else
                'Set
                If udcCharacterTwo IsNot Nothing Then
                    intReloadTimeUsed = udcCharacterTwo.ReloadTimes * 3
                End If
                'Print overlay
                'DrawEndGameScreen(intZombieKillsTwo, udcCharacterTwo)
            End If

            'Remove objects
            RemoveGameObjectsFromMemory()

        Else

            'Move graphics to copy and print dead first
            gGraphics = Graphics.FromImage(btmGameBackground)

            'Set graphic options
            SetDefaultGraphicOptions()

            'Kill zombies that were marked
            KillZombiesMarkedToDie(gaudcZombiesOne, udcCharacterOne)

            'Kill zombies that were marked
            KillZombiesMarkedToDie(gaudcZombiesTwo, udcCharacterTwo)

            'Draw dead zombies permanently
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                If gaudcZombiesOne(intLoop) IsNot Nothing Then
                    If gaudcZombiesOne(intLoop).Spawned Then
                        If gaudcZombiesOne(intLoop).IsDead Then
                            'Set
                            gaudcZombiesOne(intLoop).Spawned = False
                            'Set new point
                            Dim pntTemp As New Point(gaudcZombiesOne(intLoop).ZombiePoint.X + CInt(Math.Abs(gpntGameBackground.X)), gaudcZombiesOne(intLoop).ZombiePoint.Y)
                            'Draw dead
                            gGraphics.DrawImageUnscaled(gaudcZombiesOne(intLoop).ZombieImage, pntTemp)
                            'Increase count
                            intZombieKillsOne += 1
                            'Start a new zombie
                            gaudcZombiesOne(intZombieKillsOne + NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
                        End If
                    End If
                End If
            Next
            For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
                If gaudcZombiesTwo(intLoop) IsNot Nothing Then
                    If gaudcZombiesTwo(intLoop).Spawned Then
                        If gaudcZombiesTwo(intLoop).IsDead Then
                            'Set
                            gaudcZombiesTwo(intLoop).Spawned = False
                            'Set new point
                            Dim pntTemp As New Point(gaudcZombiesTwo(intLoop).ZombiePoint.X + CInt(Math.Abs(gpntGameBackground.X)), gaudcZombiesTwo(intLoop).ZombiePoint.Y)
                            'Draw dead
                            gGraphics.DrawImageUnscaled(gaudcZombiesTwo(intLoop).ZombieImage, pntTemp)
                            'Increase count
                            intZombieKillsTwo += 1
                            'Start a new zombie
                            gaudcZombiesTwo(intZombieKillsTwo + NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
                        End If
                    End If
                End If
            Next

            'Paint background and word
            PaintBackgroundAndWord()

            'Draw character hoster
            gGraphics.DrawImageUnscaled(udcCharacterOne.CharacterImage, udcCharacterOne.CharacterPoint)

            'First horde of zombies for hoster
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                If gaudcZombiesOne(intLoop) IsNot Nothing Then
                    'Check if spawned
                    If gaudcZombiesOne(intLoop).Spawned Then
                        'Check if can pin
                        If Not gaudcZombiesOne(intLoop).MarkedToDie Then
                            'Check distance
                            If gaudcZombiesOne(intLoop).ZombiePoint.X <= 200 And Not gaudcZombiesOne(intLoop).IsPinning Then
                                'Set
                                gaudcZombiesOne(intLoop).Pin()
                                'Check if hosting
                                If blnHost Then
                                    'Check for first time pin
                                    CheckingForFirstTimePin(False, "6|") 'Joiner won
                                End If
                            End If
                        End If
                        'Draw zombies dying, pinning or walking
                        gGraphics.DrawImageUnscaled(gaudcZombiesOne(intLoop).ZombieImage, gaudcZombiesOne(intLoop).ZombiePoint)
                    End If
                End If
            Next

            'Draw character joiner
            gGraphics.DrawImageUnscaled(udcCharacterTwo.CharacterImage, udcCharacterTwo.CharacterPoint)

            'Second horde of zombies for joiner
            For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
                If gaudcZombiesTwo(intLoop) IsNot Nothing Then
                    'Check if spawned
                    If gaudcZombiesTwo(intLoop).Spawned Then
                        'Check if can pin
                        If Not gaudcZombiesTwo(intLoop).MarkedToDie Then
                            'Check distance
                            If gaudcZombiesTwo(intLoop).ZombiePoint.X <= 200 And Not gaudcZombiesTwo(intLoop).IsPinning Then
                                'Set
                                gaudcZombiesTwo(intLoop).Pin()
                                'Check if hosting
                                If blnHost Then
                                    'Check for first time pin
                                    CheckingForFirstTimePin(True, "7|") 'Joiner lost
                                End If
                            End If
                        End If
                        'Draw zombies dying, pinning or walking
                        gGraphics.DrawImageUnscaled(gaudcZombiesTwo(intLoop).ZombieImage, gaudcZombiesTwo(intLoop).ZombiePoint)
                    End If
                End If
            Next

            'Draw nickname
            If blnHost Then
                DrawText(gGraphics, strNickname, 36, Color.Red, 90, 205, 1000, 150) 'Host sees own name
                DrawText(gGraphics, strNicknameConnected, 36, Color.Blue, 200, 255, 1000, 150) 'Host sees joiner name
            Else
                DrawText(gGraphics, strNicknameConnected, 36, Color.Red, 90, 205, 1000, 150) 'Joiner sees host name
                DrawText(gGraphics, strNickname, 36, Color.Blue, 200, 255, 1000, 150) 'Joiner sees own name
            End If

            'Show magazine with bullet count
            gGraphics.DrawImageUnscaled(btmAK47Magazine, pntAK47Magazine)

            'Draw bullet count on magazine
            If blnHost Then
                DrawText(gGraphics, CStr(30 - udcCharacterOne.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75) 'Host
            Else
                DrawText(gGraphics, CStr(30 - udcCharacterTwo.BulletsUsed), 40, Color.Blue, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75) 'Joiner
            End If

            'If game over
            If btmBlackScreen IsNot Nothing Then
                'Fade to black
                gGraphics.DrawImageUnscaled(btmBlackScreen, pntTopLeft)
            End If

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

    Private Sub StoryScreen()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw story background
        gGraphics.DrawImageUnscaled(btmStoryBackground, pntTopLeft)

        'Draw story text
        If btmStoryParagraph IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmStoryParagraph, pntStoryParagraph)
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

    Private Sub StoryTelling()

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(83, 229)

        'Set
        btmStoryParagraph = btmStoryParagraph1_25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph1_50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph1_75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph1_100

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 1 sound
        udcStoryParagraph1Sound = New clsSound(Me, strDirectory & "Sounds\StoryParagraph1.mp3", 49000, gintSoundVolume)

        'Sleep
        System.Threading.Thread.Sleep(49000)

        'Set
        btmStoryParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(38, 168)

        'Set
        btmStoryParagraph = btmStoryParagraph2_25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph2_50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph2_75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph2_100

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 2 sound
        udcStoryParagraph2Sound = New clsSound(Me, strDirectory & "Sounds\StoryParagraph2.mp3", 64000, gintSoundVolume)

        'Sleep
        System.Threading.Thread.Sleep(64000)

        'Set
        btmStoryParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(31, 180)

        'Set
        btmStoryParagraph = btmStoryParagraph3_25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_100

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 3 sound
        udcStoryParagraph2Sound = New clsSound(Me, strDirectory & "Sounds\StoryParagraph3.mp3", 86000, gintSoundVolume)

    End Sub

    Private Sub GameVersionMismatch()

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Draw story background
        gGraphics.DrawImageUnscaled(btmGameMismatchBackground, pntTopLeft)

        'Display
        MismatchingText("Your version: " & GAME_VERSION, 878, 197)

        'Draw text
        If blnHost Then
            MismatchingText("Joiner version: " & strGameVersionFromConnection, 878, 397)
        Else
            MismatchingText("Host version: " & strGameVersionFromConnection, 878, 397)
        End If

    End Sub

    Private Sub MismatchingText(strText As String, intX As Integer, intY As Integer)

        'Draw text
        DrawText(gGraphics, strText, 50, Color.Black, intX, intY, 1000, 100) 'Shadow
        DrawText(gGraphics, strText, 50, Color.Black, intX + 2, intY + 2, 1000, 100) 'Shadow
        DrawText(gGraphics, strText, 50, Color.Red, intX - 2, intY - 2, 1000, 100)

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
            gdblScreenWidthRatio = CDbl((Me.Width - WIDTH_SUBTRACTION) / ORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((Me.Height - HEIGHT_SUBTRACTION) / ORIGINAL_SCREEN_HEIGHT)
            'Set screen rectangle
            rectFullScreen.Width = Me.Width - WIDTH_SUBTRACTION
            rectFullScreen.Height = Me.Height - HEIGHT_SUBTRACTION
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
            If intFogPass1X >= ORIGINAL_SCREEN_WIDTH * 2 Then
                intFogPass1X = -(ORIGINAL_SCREEN_WIDTH * 2)
            End If
            If intFogPass2X >= ORIGINAL_SCREEN_WIDTH * 2 Then
                intFogPass2X = -(ORIGINAL_SCREEN_WIDTH * 2)
            End If
            'Move fog
            intFogPass1X += rndNumber.Next(1, FOG_MAX_SPEED + 1)
            intFogPass2X += rndNumber.Next(1, FOG_MAX_SPEED + 1)
            'Move fog
            pntFogFrontPass1.X = intFogPass1X
            pntFogBackPass1.X = intFogPass1X
            pntFogFrontPass2.X = intFogPass2X
            pntFogBackPass2.X = intFogPass2X
            'Reduce CPU usage
            System.Threading.Thread.Sleep(rndNumber.Next(1, FOG_MAX_DELAY + 1))
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
                Dim udcButtonHoverSound As New clsSound(Me, strDirectory & "Sounds\ButtonHover.mp3", 500, gintSoundVolume)
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
            If blnSliderWithMouseDown Then
                ChangeSoundVolume(pntMouse)
                'Exit
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

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

            'Exit
            Exit Sub

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

            'Exit
            Exit Sub

        End If

        'Check credits screen
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

            'Exit
            Exit Sub

        End If

        'Check versus screen
        If intCanvasMode = 6 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

        'Check versus game screen
        If intCanvasMode = 8 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

        'Check story
        If intCanvasMode = 9 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

        'Path 1 choice
        If intCanvasMode = 11 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Paths in the sewer
            If blnMouseInRegion(pntMouse, 389, 329, New Point(230, 427)) Then
                btmPath = btmPath1_1
                'Play light switch
                If Not blnLightZap1 Then
                    Dim udcLightZapSound As New clsSound(Me, strDirectory & "Sounds\LightZap.mp3", 1000, gintSoundVolume)
                    blnLightZap1 = True
                End If
                'Exit
                Exit Sub
            Else
                If blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) Then
                    btmPath = btmPath1_2
                    'Play light switch
                    If Not blnLightZap2 Then
                        Dim udcLightZapSound As New clsSound(Me, strDirectory & "Sounds\LightZap.mp3", 1000, gintSoundVolume)
                        blnLightZap2 = True
                    End If
                    'Exit
                    Exit Sub
                Else
                    btmPath = btmPath1_0
                End If
            End If

            'Reset
            blnLightZap1 = False
            blnLightZap2 = False

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

        'Path 2 choice
        If intCanvasMode = 12 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                HoverText(1, sblnOnTopOf)
                Exit Sub
            End If

            'Paths in the sewer
            If blnMouseInRegion(pntMouse, 296, 304, New Point(643, 306)) Then
                btmPath = btmPath2_1
                'Play light switch
                If Not blnLightZap1 Then
                    Dim udcLightZapSound As New clsSound(Me, strDirectory & "Sounds\LightZap.mp3", 1000, gintSoundVolume)
                    'Zombie growl sound
                    udcZombieGrowlSound = New clsSound(Me, strDirectory & "Sounds\ZombieGrowl.mp3", 3000, gintSoundVolume)
                    blnLightZap1 = True
                End If
                'Exit
                Exit Sub
            Else
                If blnMouseInRegion(pntMouse, 297, 318, New Point(1138, 384)) Then
                    btmPath = btmPath2_2
                    'Play light switch
                    If Not blnLightZap2 Then
                        Dim udcLightZapSound As New clsSound(Me, strDirectory & "Sounds\LightZap.mp3", 1000, gintSoundVolume)
                        blnLightZap2 = True
                    End If
                    'Exit
                    Exit Sub
                Else
                    btmPath = btmPath2_0
                End If
            End If

            'Reset
            If udcZombieGrowlSound IsNot Nothing Then
                udcZombieGrowlSound.StopAndCloseSound()
                udcZombieGrowlSound = Nothing
            End If

            'Reset
            blnLightZap1 = False
            blnLightZap2 = False

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

            'Exit
            Exit Sub

        End If

    End Sub

    Private Sub HoverSoundWaiting()

        'Sleep for 500 ms
        System.Threading.Thread.Sleep(500)

        'Set
        thrHoverSoundDelay = Nothing

    End Sub

    Private Sub ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer,
                                                                      Optional blnPlayPressedSound As Boolean = True)

        'Set
        intCanvasMode = intCanvasModeToSet

        'Set
        intCanvasShow = intCanvasShowToSet

        'Play sound
        If blnPlayPressedSound Then
            Dim udcButtonPressedSound As New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 1000, gintSoundVolume, False)
        End If

    End Sub

    Private Sub Paragraphing()

        'Set
        btmLoadingParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph100

    End Sub

    Private Sub SetBlackScreenVariables()

        'Set
        btmBlackScreen = Nothing

        'Set
        blnBlackScreenFinished = False

        'Set
        blnTriggeredBlackScreenThread = False

    End Sub

    Private Sub NextLevel(btmGameBackgroundLevel As Bitmap, intLevelToBe As Integer, Optional blnLoadHelicopter As Boolean = False)

        'Set
        If btmDeathScreen IsNot Nothing Then
            btmDeathScreen.Dispose()
            btmDeathScreen = Nothing
        End If

        'Set
        blnRemovedGameObjectsFromMemory = False

        'Set
        gintBullets = 0

        'Set
        SetBlackScreenVariables()

        'Set
        gpntGameBackground.X = 0

        'Set to be fresh every game
        btmGameBackground = btmGameBackgroundLevel

        'Set
        intLevel = intLevelToBe

        'Set
        intZombieKills = 0

        'Set
        intReloadTimeUsed = 0

        'Set
        blnComparedHighscore = False

        'Set
        blnEndingGameCantType = False

        'Set
        blnBeatLevel = False

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter", intLevel)

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Check level for game losing zombies
        Select Case intLevel
            Case 2
                'SpawnUnderwaterZombieGameLose()
        End Select

        'Helicopter
        If blnLoadHelicopter Then
            udcHelicopter = New clsHelicopter(Me)
            udcHelicopter.Start()
        End If

        'Start character
        udcCharacter.Start()

        'Start zombies
        For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombies(intLoop).Start()
        Next

        'Start stop watch
        swhStopWatch.Start()

        'Play sound music
        Select Case intLevel
            Case 2
                udcGameBackgroundSound(1) = New clsSound(Me, strDirectory & "Sounds\GameBackground2.mp3", 41000, CInt(gintSoundVolume / 8), True)
            Case 3
                udcGameBackgroundSound(2) = New clsSound(Me, strDirectory & "Sounds\GameBackground3.mp3", 38000, CInt(gintSoundVolume / 8), True)
            Case 4
                udcGameBackgroundSound(3) = New clsSound(Me, strDirectory & "Sounds\GameBackground4.mp3", 31000, CInt(gintSoundVolume / 8), True)
            Case 5
                udcGameBackgroundSound(4) = New clsSound(Me, strDirectory & "Sounds\GameBackground5.mp3", 59000, CInt(gintSoundVolume / 8), True)
        End Select

    End Sub

    Private Sub SpawnUnderwaterZombieGameLose()

        'Set
        blnZombieGameLoseHappened = False

        'Exit
        Exit Sub

        ''Delcare
        'Dim rndNumber As New Random

        ''Check if need to spawn zombie
        'Select Case rndNumber.Next(0, 4)
        '    Case 1, 2, 3 '75% chance
        '        'Create
        '        udcUnderwaterZombieGameLose = New clsZombieGameLose(Me, "Underwater", 300, 900)
        'End Select

    End Sub

    Private Sub LoadingGame()

        'Notes: This procedure is for loading the game to play

        'Set
        If btmDeathScreen IsNot Nothing Then
            btmDeathScreen.Dispose()
            btmDeathScreen = Nothing
        End If

        'Set
        blnRemovedGameObjectsFromMemory = False

        'Set
        intZombieKillsCombined = 0
        intReloadTimeUsedCombined = 0

        'Set
        gpntGameBackground.X = 0

        'Set
        intLevel = 1

        'Set
        blnComparedHighscore = False

        'Set
        gintBullets = 0 'Starting bullets for key press

        'Set
        SetBlackScreenVariables()

        'Set
        intZombieKills = 0

        'Set
        intReloadTimeUsed = 0

        'Set to be fresh every game
        btmGameBackground = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground1.jpg"))

        'Set
        blnEndingGameCantType = False

        'Set
        blnGameOverFirstTime = True

        'Set
        blnBeatLevel = False

        'Check loaded game objects
        For intLoop As Integer = 0 To blnLoadingGameFinished.GetUpperBound(0) - 1
            'Create a pause waiting
            While Not blnLoadingGameFinished(intLoop)
            End While
            'Set
            btmLoadingBar = btmLoadingBarPicture(intLoop + 1)
        Next

        'Grab a random word
        LoadARandomWord()

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter", intLevel)

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Check loaded game objects
        While Not blnLoadingGameFinished(blnLoadingGameFinished.GetUpperBound(0))
        End While

        'Set
        btmLoadingBar = btmLoadingBarPicture(blnLoadingGameFinished.GetUpperBound(0) + 1)

        'Set
        intCanvasShow = 1 'Means completely loaded

    End Sub

    Private Sub LoadZombies(strGameType As String)

        'Check if multiplayer or not
        Select Case strGameType
            Case "Level 1 Single Player"
                'Zombies
                For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                    ReDim Preserve gaudcZombies(intLoop)
                    Select Case intLoop
                        Case 0
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed, "udcCharacter")
                        Case 1
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, intZombieSpeed, "udcCharacter")
                        Case 2
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, intZombieSpeed, "udcCharacter")
                        Case 3
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, intZombieSpeed, "udcCharacter")
                        Case 4
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, intZombieSpeed, "udcCharacter")
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 5, "udcCharacter")
                        Case 9, 14, 19, 24
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 25, "udcCharacter")
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 10, "udcCharacter")
                        Case 29, 34, 39, 44
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 30, "udcCharacter")
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 15, "udcCharacter")
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 35, "udcCharacter")
                        Case 65 To 74
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 5, "udcCharacter")
                        Case 75 To 84
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 10, "udcCharacter")
                        Case 85 To 94
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 15, "udcCharacter")
                        Case 95 To 104
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 20, "udcCharacter")
                        Case Else
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 45, "udcCharacter")
                    End Select
                Next
            Case "Level 1 Multiplayer"
                'Zombies
                For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                    ReDim Preserve gaudcZombiesOne(intLoop)
                    Select Case intLoop
                        Case 0
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne, "udcCharacterOne")
                        Case 1
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, intZombieSpeedOne, "udcCharacterOne")
                        Case 2
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, intZombieSpeedOne, "udcCharacterOne")
                        Case 3
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, intZombieSpeedOne, "udcCharacterOne")
                        Case 4
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, intZombieSpeedOne, "udcCharacterOne")
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 5, "udcCharacterOne")
                        Case 9, 14, 19, 24
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 25, "udcCharacterOne")
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 10, "udcCharacterOne")
                        Case 29, 34, 39, 44
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 30, "udcCharacterOne")
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 15, "udcCharacterOne")
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 35, "udcCharacterOne")
                        Case 65 To 74
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 5, "udcCharacterOne")
                        Case 75 To 84
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 10, "udcCharacterOne")
                        Case 85 To 94
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 15, "udcCharacterOne")
                        Case 95 To 104
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 20, "udcCharacterOne")
                        Case Else
                            gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 45, "udcCharacterOne")
                    End Select
                Next
                For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                    ReDim Preserve gaudcZombiesTwo(intLoop)
                    Select Case intLoop
                        Case 0
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                        Case 1
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                        Case 2
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                        Case 3
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                        Case 4
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 5, "udcCharacterTwo")
                        Case 9, 14, 19, 24
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 25, "udcCharacterTwo")
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 10, "udcCharacterTwo")
                        Case 29, 34, 39, 44
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 30, "udcCharacterTwo")
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 15, "udcCharacterTwo")
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 35, "udcCharacterTwo")
                        Case 65 To 74
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 5, "udcCharacterTwo")
                        Case 75 To 84
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 10, "udcCharacterTwo")
                        Case 85 To 94
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 15, "udcCharacterTwo")
                        Case 95 To 104
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 20, "udcCharacterTwo")
                        Case Else
                            gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 45, "udcCharacterTwo")
                    End Select
                Next
        End Select

    End Sub

    Private Sub LoadPieceOfBitmaps(ByRef abtmByRefBitmap() As Bitmap, strMainFolder As String, strSubFolder As String)

        'Load
        For intLoop As Integer = 0 To abtmByRefBitmap.GetUpperBound(0)
            abtmByRefBitmap(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\" & strMainFolder & "\" & strSubFolder & "\" & CStr(intLoop + 1) & ".png"))
        Next

    End Sub

    Private Sub LoadingGameWithMultipleThreads()

        'Notes: This only loads once, so don't put variables in here that have to keep resetting.

        'Declare
        Static sblnFirstTimeLoadingGameObjects As Boolean = True

        'Check
        If sblnFirstTimeLoadingGameObjects Then

            'Load 10%
            thrLoading(0) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame10))
            thrLoading(0).Start()

            'Load 20%
            thrLoading(1) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame20))
            thrLoading(1).Start()

            'Load 30%
            thrLoading(2) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame30))
            thrLoading(2).Start()

            'Load 40%
            thrLoading(3) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame40))
            thrLoading(3).Start()

            'Load 50%
            thrLoading(4) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame50))
            thrLoading(4).Start()

            'Load 60%
            thrLoading(5) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame60))
            thrLoading(5).Start()

            'Load 70%
            thrLoading(6) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame70))
            thrLoading(6).Start()

            'Load 80%
            thrLoading(7) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame80))
            thrLoading(7).Start()

            'Load 90%
            thrLoading(8) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame90))
            thrLoading(8).Start()

            'Load 100%
            thrLoading(9) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame100))
            thrLoading(9).Start()

            'Set
            sblnFirstTimeLoadingGameObjects = False

        End If

    End Sub

    Private Sub LoadingGame10()

        'Words
        AddWordsToArray()

        'Word bar
        btmWordBar = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\WordBar.png"))

        'Set
        blnLoadingGameFinished(0) = True

    End Sub

    Private Sub LoadingGame20()

        'Character
        LoadPieceOfBitmaps(gbtmCharacterStand, "Character\Generic", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShoot, "Character\Generic", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReload, "Character\Generic", "Reload")
        LoadPieceOfBitmaps(gbtmCharacterRunning, "Character\Generic", "Running")

        'Set
        blnLoadingGameFinished(1) = True

    End Sub

    Private Sub LoadingGame30()

        'Character red
        LoadPieceOfBitmaps(gbtmCharacterStandRed, "Character\Red", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShootRed, "Character\Red", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReloadRed, "Character\Red", "Reload")

        'Set
        blnLoadingGameFinished(2) = True

    End Sub

    Private Sub LoadingGame40()

        'Character blue
        LoadPieceOfBitmaps(gbtmCharacterStandBlue, "Character\Blue", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShootBlue, "Character\Blue", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReloadBlue, "Character\Blue", "Reload")

        'Set
        blnLoadingGameFinished(3) = True

    End Sub

    Private Sub LoadingGame50()

        'Zombies
        LoadPieceOfBitmaps(gbtmZombieWalk, "Zombies\Generic", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeath1, "Zombies\Generic", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeath2, "Zombies\Generic", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePin, "Zombies\Generic", "Pinning")

        'Set
        blnLoadingGameFinished(4) = True

    End Sub

    Private Sub LoadingGame60()

        'Zombies red
        LoadPieceOfBitmaps(gbtmZombieWalkRed, "Zombies\Red", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeathRed1, "Zombies\Red", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeathRed2, "Zombies\Red", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePinRed, "Zombies\Red", "Pinning")

        'Set
        blnLoadingGameFinished(5) = True

    End Sub

    Private Sub LoadingGame70()

        'Zombies blue
        LoadPieceOfBitmaps(gbtmZombieWalkBlue, "Zombies\Blue", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeathBlue1, "Zombies\Blue", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeathBlue2, "Zombies\Blue", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePinBlue, "Zombies\Blue", "Pinning")

        'Set
        blnLoadingGameFinished(6) = True

    End Sub

    Private Sub LoadingGame80()

        'Helicopter
        gbtmHelicopter(0) = New Bitmap(Image.FromFile(strDirectory & "Images\Helicopter\Helicopter1.jpg"))
        gbtmHelicopter(1) = New Bitmap(Image.FromFile(strDirectory & "Images\Helicopter\Helicopter2.jpg"))
        gbtmHelicopter(2) = New Bitmap(Image.FromFile(strDirectory & "Images\Helicopter\Helicopter3.jpg"))
        gbtmHelicopter(3) = New Bitmap(Image.FromFile(strDirectory & "Images\Helicopter\Helicopter4.jpg"))
        gbtmHelicopter(4) = New Bitmap(Image.FromFile(strDirectory & "Images\Helicopter\Helicopter5.jpg"))

        'Magazine
        btmAK47Magazine = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\AK47Magazine.png"))

        'Set
        blnLoadingGameFinished(7) = True

    End Sub

    Private Sub LoadingGame90()

        'Death overlay
        btmDeathOverlay = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\DeathOverlay.png"))

        'Win overlay
        btmWinOverlay = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\WinOverlay.jpg"))

        'Versus win or lose
        btmYouWon = New Bitmap(Image.FromFile(strDirectory & "Images\Versus\YouWon.png"))
        btmYouLost = New Bitmap(Image.FromFile(strDirectory & "Images\Versus\YouLost.png"))

        'Set
        blnLoadingGameFinished(8) = True

    End Sub

    Private Sub LoadingGame100()

        'End game
        btmBlackScreen25 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\BlackScreen25.png"))
        btmBlackScreen50 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\BlackScreen50.png"))
        btmBlackScreen75 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\BlackScreen75.png"))
        btmBlackScreen100 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\BlackScreen100.png"))

        'Zombies that produce game lose
        gbtmUnderwaterZombieGameLose = New Bitmap(Image.FromFile(strDirectory & "Images\ZombieGameLose\UnderwaterZombieGameLose.png"))

        'Set
        blnLoadingGameFinished(9) = True

    End Sub

    Private Sub AddWordsToArray()

        'Words
        AddWordToArray("and")
        AddWordToArray("aim")
        AddWordToArray("archer")
        AddWordToArray("arrive")
        AddWordToArray("attack")
        AddWordToArray("battle")
        AddWordToArray("begins")
        AddWordToArray("bio")
        AddWordToArray("blood")
        AddWordToArray("carve")
        AddWordToArray("causes")
        AddWordToArray("chemical")
        AddWordToArray("china")
        AddWordToArray("communications")
        AddWordToArray("compound")
        AddWordToArray("contractions")
        AddWordToArray("coughing")
        AddWordToArray("countries")
        AddWordToArray("creating")
        AddWordToArray("deadly")
        AddWordToArray("death")
        AddWordToArray("destruction")
        AddWordToArray("die")
        AddWordToArray("egypt")
        AddWordToArray("equipped")
        AddWordToArray("escaped")
        AddWordToArray("europe")
        AddWordToArray("ever")
        AddWordToArray("every")
        AddWordToArray("everyone")
        AddWordToArray("fall")
        AddWordToArray("finalize")
        AddWordToArray("find")
        AddWordToArray("france")
        AddWordToArray("has")
        AddWordToArray("hear")
        AddWordToArray("hosts")
        AddWordToArray("in")
        AddWordToArray("incursions")
        AddWordToArray("infected")
        AddWordToArray("intelligence")
        AddWordToArray("involuntary")
        AddWordToArray("iranian")
        AddWordToArray("is")
        AddWordToArray("israel")
        AddWordToArray("jordan")
        AddWordToArray("killing")
        AddWordToArray("lab")
        AddWordToArray("largest")
        AddWordToArray("leader")
        AddWordToArray("manage")
        AddWordToArray("muscles")
        AddWordToArray("not")
        AddWordToArray("of")
        AddWordToArray("one")
        AddWordToArray("painful")
        AddWordToArray("pathogen")
        AddWordToArray("pit")
        AddWordToArray("populations")
        AddWordToArray("respiratory")
        AddWordToArray("russia")
        AddWordToArray("schemed")
        AddWordToArray("simultaneous")
        AddWordToArray("site")
        AddWordToArray("summer")
        AddWordToArray("supreme")
        AddWordToArray("symptoms")
        AddWordToArray("targeting")
        AddWordToArray("terror")
        AddWordToArray("the")
        AddWordToArray("this")
        AddWordToArray("three")
        AddWordToArray("to")
        AddWordToArray("total")
        AddWordToArray("weapons")
        AddWordToArray("with")

    End Sub

    Private Sub AddWordToArray(strWord As String)

        'Declare
        Static sblnFirstTime As Boolean = True

        'Check
        If sblnFirstTime Then
            'Set
            sblnFirstTime = False
            'Set
            astrWords(0) = strWord
        Else
            ReDim Preserve astrWords(astrWords.GetUpperBound(0) + 1)
            'Set
            astrWords(astrWords.GetUpperBound(0)) = strWord
        End If

    End Sub

    Private Sub ShowNextScreenAndExitMenu(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer)

        'Stop fog thread
        AbortThread(thrFog)

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(intCanvasModeToSet, intCanvasShowToSet)

        'Stop sound
        If udcAmbianceSound IsNot Nothing Then
            udcAmbianceSound.StopAndCloseSound()
        End If

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If menu
        If intCanvasMode = 0 Then

            'Start was clicked
            If blnMouseInRegion(pntMouse, 212, 69, pntStartText) Then
                'Wait
                WaitUntilVariablesReset()
                'Stop fog thread
                AbortThread(thrFog)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(2, 0)
                'Stop sound
                If udcAmbianceSound IsNot Nothing Then
                    udcAmbianceSound.StopAndCloseSound()
                End If
                'Start paragraphing
                thrParagraph = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Paragraphing))
                thrParagraph.Start()
                'Start loading game
                LoadingGameWithMultipleThreads()
                'Setup the loading game thread
                thrLoadingGame = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame))
                thrLoadingGame.Start()
                'Reset fog
                ResetFog()
                'Exit
                Exit Sub
            End If

            'Highscores was clicked
            If blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) Then
                'Wait
                WaitUntilVariablesReset()
                'Change
                ShowNextScreenAndExitMenu(4, 0)
                'Reset fog
                ResetFog()
                'Exit
                Exit Sub
            End If

            'Story was clicked
            If blnMouseInRegion(pntMouse, 218, 87, pntStoryText) Then
                'Wait
                WaitUntilVariablesReset()
                'Change
                ShowNextScreenAndExitMenu(9, 0)
                'Reset fog
                ResetFog()
                'Set story fade in
                thrStory = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf StoryTelling))
                thrStory.Start()
                'Exit
                Exit Sub
            End If

            'Options was clicked
            If blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) Then
                'Wait
                WaitUntilVariablesReset()
                'Change
                ShowNextScreenAndExitMenu(1, 0)
                'Reset fog
                ResetFog()
                'Options sound
                udcAmbianceOptionsSound = New clsSound(Me, strDirectory & "Sounds\AmbianceOptions.mp3", 10000, gintSoundVolume, True) '38 seconds + extra
                'Exit
                Exit Sub
            End If

            'Credits was clicked
            If blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) Then
                'Wait
                WaitUntilVariablesReset()
                'Change
                ShowNextScreenAndExitMenu(5, 0)
                'Reset fog
                ResetFog()
                'Set credits fade in
                thrCredits = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf CreditsFadeIn))
                thrCredits.Start()
                'Exit
                Exit Sub
            End If

            'Versus was clicked
            If blnMouseInRegion(pntMouse, 256, 74, pntVersusText) Then
                'Wait
                WaitUntilVariablesReset()
                'Set
                strIPAddress = strGetLocalIPAddress()
                'Set
                blnFirstTimeNicknameTyping = True
                blnFirstTimeIPAddressTyping = True
                'Change
                ShowNextScreenAndExitMenu(6, 0)
                'Reset fog
                ResetFog()
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
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 21000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Restart fog
                RestartFog()
                'Stop options sound
                If udcAmbianceOptionsSound IsNot Nothing Then
                    udcAmbianceOptionsSound.StopAndCloseSound()
                End If
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

            'Exit
            Exit Sub

        End If

        'If loading game screen
        If intCanvasMode = 2 Then

            'Loading start text bar was clicked
            If intCanvasShow = 1 Then 'If finished loading
                If blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                    'Set
                    intCanvasMode = 3
                    'Set
                    intCanvasShow = 0
                    'Start character
                    udcCharacter.Start()
                    'Start zombies
                    For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
                        gaudcZombies(intLoop).Start()
                    Next
                    'Start stop watch
                    swhStopWatch = New Stopwatch
                    swhStopWatch.Start()
                    'Start music
                    udcGameBackgroundSound(0) = New clsSound(Me, strDirectory & "Sounds\GameBackground1.mp3", 40000, CInt(gintSoundVolume / 8), True)
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
                'Set
                blnPlayPressedSoundNow = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                blnBackFromGame = True
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
                'Abort thread
                AbortThread(thrCredits)
                'Reset
                btmJohnGonzales = Nothing
                btmZacharyStafford = Nothing
                btmCoryLewis = Nothing
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If versus screen
        If intCanvasMode = 6 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Check nickname
                DefaultNickname()
                'Set
                blnPlayPressedSoundNow = True
                'Go back to menu
                GoBackToMenuConnectionGone()
                'Exit
                Exit Sub
            End If

            'Check for host or join
            If intCanvasVersusShow = 0 Then
                'Host was clicked
                If blnMouseInRegion(pntMouse, 545, 176, pntVersusHost) Then
                    'Set
                    blnHost = True
                    'Check nickname
                    DefaultNickname()
                    'Set
                    intCanvasVersusShow = 1
                    'Set
                    tcplServer = New System.Net.Sockets.TcpListener(System.Net.IPAddress.Any, 10101)
                    'Start
                    tcplServer.Start()
                    'Start thread listening
                    thrListening = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Listening))
                    thrListening.Start()
                    'Exit
                    Exit Sub
                End If
                'Join was clicked
                If blnMouseInRegion(pntMouse, 441, 170, pntVersusJoin) Then
                    'Set
                    blnHost = False
                    'Set
                    strIPAddressConnect = strGetLocalIPAddress()
                    'Check nickname
                    DefaultNickname()
                    'Set
                    intCanvasVersusShow = 2
                    'Exit
                    Exit Sub
                End If
                'Exit
                Exit Sub
            End If

            'Check for connect button
            If intCanvasVersusShow = 2 Then
                'Connect was clicked
                If blnMouseInRegion(pntMouse, 605, 124, pntVersusConnect) Then
                    'Set
                    intCanvasVersusShow = 3
                    'Start thread connecting
                    thrConnecting = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Connecting))
                    thrConnecting.Start()
                    'Exit
                    Exit Sub
                End If
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If loading versus screen
        If intCanvasMode = 7 Then

            'Check if hosting
            If blnHost Then
                'Start was clicked
                If Not blnWaiting And intCanvasShow = 1 Then
                    'Start was clicked
                    If blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                        'Send playing
                        gSendData("2|" & strNickname)
                        'Start game locally
                        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(8, 0, False)
                        'Started versus game, start objects
                        StartVersusGameObjects()
                        'Exit
                        Exit Sub
                    End If
                End If
            Else
                'Start was clicked
                If Not blnWaiting And intCanvasShow = 1 Then
                    'Start was clicked
                    If blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                        'Send ready to play as joiner
                        gSendData("2|" & strNickname)
                        blnWaiting = True
                        'Exit
                        Exit Sub
                    End If
                End If
            End If

            'Exit
            Exit Sub

        End If

        'If versus game screen
        If intCanvasMode = 8 Then

            'Back was pressed
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Set
                blnPlayPressedSoundNow = True
                'Quit versus multiplayer
                GoBackToMenuConnectionGone()
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'If story screen
        If intCanvasMode = 9 Then

            'Back has been moused over
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0)
                'Restart fog
                RestartFog()
                'Abort thread
                AbortThread(thrStory)
                'Reset
                btmStoryParagraph = Nothing
                'Stop sounds
                If udcStoryParagraph1Sound IsNot Nothing Then
                    udcStoryParagraph1Sound.StopAndCloseSound()
                End If
                If udcStoryParagraph2Sound IsNot Nothing Then
                    udcStoryParagraph2Sound.StopAndCloseSound()
                End If
                If udcStoryParagraph3Sound IsNot Nothing Then
                    udcStoryParagraph3Sound.StopAndCloseSound()
                End If
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'Path choice 1
        If intCanvasMode = 11 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Set
                blnPlayPressedSoundNow = True
                'Set
                blnBackFromGame = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Exit
                Exit Sub
            End If

            'Path 1
            If blnMouseInRegion(pntMouse, 389, 329, New Point(230, 427)) Then
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground2.jpg")), 2)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(3, 0, False)
                'Exit
                Exit Sub
            End If

            'Path 2
            If blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) Then
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground3.jpg")), 3)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(3, 0, False)
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'Path choice 2
        If intCanvasMode = 12 Then

            'Back was clicked
            If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
                'Set
                blnPlayPressedSoundNow = True
                'Set
                blnBackFromGame = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
                'Exit
                Exit Sub
            End If

            'Path 1
            If blnMouseInRegion(pntMouse, 296, 304, New Point(643, 306)) Then
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground5.jpg")), 5, True)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(3, 0, False)
                'Exit
                Exit Sub
            End If

            'Path 2
            If blnMouseInRegion(pntMouse, 297, 318, New Point(1138, 384)) Then
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground4.jpg")), 4)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(3, 0, False)
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

    End Sub

    Private Sub WaitUntilVariablesReset()

        'Wait until
        While blnBackFromGame
            System.Threading.Thread.Sleep(1)
        End While

    End Sub

    Private Sub StartVersusGameObjects()

        'Start characters
        udcCharacterOne.Start()
        udcCharacterTwo.Start()

        'Start zombies
        For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombiesOne(intLoop).Start()
        Next
        For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombiesTwo(intLoop).Start()
        Next

        'Start stop watch
        swhStopWatch = New Stopwatch
        swhStopWatch.Start()

    End Sub

    Private Sub GoBackToMenuConnectionGone()

        'Menu sound
        udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra

        'Set
        blnBackFromGame = True

    End Sub

    Private Sub DefaultNickname()

        'Check nickname
        If strNickname = "" Then
            strNickname = "Player"
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

        'After setting volume, change option sound
        udcAmbianceOptionsSound.ChangeVolumeWhileSoundIsPlaying()

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

    Private Sub ResetFog()

        'Reset fog
        intFogPass1X = -(ORIGINAL_SCREEN_WIDTH * 2)
        intFogPass2X = -((ORIGINAL_SCREEN_WIDTH * 2) * 2)
        pntFogBackPass1.X = -(ORIGINAL_SCREEN_WIDTH * 2)
        pntFogFrontPass1.X = -(ORIGINAL_SCREEN_WIDTH * 2)
        pntFogBackPass2.X = -((ORIGINAL_SCREEN_WIDTH * 2) * 2)
        pntFogFrontPass2.X = -((ORIGINAL_SCREEN_WIDTH * 2) * 2)

    End Sub

    Private Sub RestartFog()

        'Set for fog
        thrFog = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf KeepFogMoving))
        thrFog.Start()

    End Sub

    Private Sub frmGame_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If options
        If intCanvasMode = 1 Then

            'Check if changing volume
            If blnMouseInRegion(pntMouse, 600, 46, pntSoundBar) Then 'In the bar
                'Change
                ChangeSoundVolume(pntMouse)
                'Set
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

        'Keys pressed
        'Debug.Print(CStr(Asc(e.KeyChar)))

        'Check if playing game
        If intCanvasMode = 3 Then

            'Exit if not finished loading or can't type because game over
            If udcCharacter Is Nothing Or blnEndingGameCantType Then
                'Exit
                Exit Sub
            End If

            'Check for max bullets
            If udcCharacter.BulletsUsed = 30 Then
                'Exit
                Exit Sub
            End If

            'Check the status of the character
            Select Case udcCharacter.StatusModeProcessing

                Case clsCharacter.eintStatusMode.Stand

                    'Check key press
                    Select Case Asc(e.KeyChar)

                        Case 32 ' = spacebar
                            'Check to make sure not maxed on bullets
                            If udcCharacter.BulletsUsed <> 0 Then
                                'Character start to reload
                                udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload
                            End If

                        Case 39 ' = '
                            'Running forward
                            udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Run

                        Case Else
                            'Check the word being typed
                            CheckTheWordBeingTyped(e, udcCharacter, gaudcZombies) 'Includes the start to do status mode

                    End Select

                Case clsCharacter.eintStatusMode.Reload

                    'Check key press
                    Select Case Asc(e.KeyChar)

                        Case 32 ' = spacebar
                            'Do nothing

                        Case 39 ' = '
                            'Running forward
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run

                        Case Else
                            'Do nothing

                    End Select

                Case clsCharacter.eintStatusMode.Run

                    'Check key press
                    Select Case Asc(e.KeyChar)

                        Case 32 ' = spacebar
                            'Check to make sure not maxed on bullets
                            If udcCharacter.BulletsUsed <> 0 Then
                                'Character start to reload
                                udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload
                            End If

                        Case 39 ' = '
                            'Do Nothing

                        Case 59 ' = ;
                            'Stop running
                            udcCharacter.StopCharacterFromRunning = True 'Have to use this because neutral will not stop running properly

                        Case Else
                            'Check the word being typed
                            CheckTheWordBeingTyped(e, udcCharacter, gaudcZombies) 'Includes the start to do status mode

                    End Select

                Case clsCharacter.eintStatusMode.Shoot

                    'Check key press
                    Select Case Asc(e.KeyChar)

                        Case 32 ' = spacebar
                            'Check to make sure not maxed on bullets
                            If udcCharacter.BulletsUsed <> 0 Then
                                'Character start to reload
                                udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload
                            End If

                        Case 39 ' = '
                            'Running forward
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run

                        Case Else
                            'Check the word being typed
                            CheckTheWordBeingTyped(e, udcCharacter, gaudcZombies) 'Includes the start to do status mode

                    End Select

            End Select

            'Exit
            Exit Sub

        End If

        'Check if changing versus nickname or IP address
        If intCanvasMode = 6 Then

            'Check show for typing nickname
            If intCanvasVersusShow = 0 Then
                'Check length
                If Len(strNickname) < 9 Then
                    'Check the keypress
                    Select Case Asc(e.KeyChar)
                        Case 8 'Backspace
                            'Check if first time key pressing
                            KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickname)
                        Case 32, 65 To 90, 97 To 122 '32 = spacebar, 65 to 90 is upper case A-Z, 97 to 122 is lower case a-z
                            'Check if first time key pressing
                            If blnFirstTimeNicknameTyping Then
                                strNickname = e.KeyChar
                                blnFirstTimeNicknameTyping = False
                            Else
                                strNickname &= e.KeyChar
                            End If
                        Case Else 'Used for debugging
                            'Debug.Print(Cstr(Asc(e.KeyChar)))
                    End Select
                Else
                    'Allow backspace
                    If Asc(e.KeyChar) = 8 Then
                        'Check if first time key pressing
                        KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickname)
                    End If
                End If
                'Exit
                Exit Sub
            End If

            'Check show for typing IP address
            If intCanvasVersusShow = 2 Then 'Joining
                'Check if pressing for first time
                If blnFirstTimeIPAddressTyping Then
                    'Check the keypress
                    Select Case Asc(e.KeyChar)
                        Case 8, 48 To 57 '8 is backspace, 48 to 57 is 0 to 9 in numbers
                            'Check if first time key pressing
                            strIPAddressConnect = e.KeyChar
                            blnFirstTimeIPAddressTyping = False
                        Case Else 'Used for debugging
                            'Debug.Print(CStr(Asc(e.KeyChar)))
                    End Select
                Else
                    'Check length
                    If Len(strIPAddressConnect) < 15 Then 'XXX.XXX.XXX.XXX
                        'Check the keypress
                        Select Case Asc(e.KeyChar)
                            Case 8 'Backspace
                                'Check if first time key pressing
                                KeyPressBackspacing(blnFirstTimeIPAddressTyping, strIPAddressConnect)
                            Case 46, 48 To 57 '46 = period, 48 to 57 is 0 to 9 in numbers
                                'Check if first time key pressing
                                If blnFirstTimeIPAddressTyping Then
                                    strIPAddressConnect = e.KeyChar
                                    blnFirstTimeIPAddressTyping = False
                                Else
                                    strIPAddressConnect &= e.KeyChar
                                End If
                            Case Else 'Used for debugging
                                'Debug.Print(CStr(Asc(e.KeyChar)))
                        End Select
                    Else
                        'Allow backspace
                        If Asc(e.KeyChar) = 8 Then
                            'Check if first time key pressing
                            KeyPressBackspacing(blnFirstTimeIPAddressTyping, strIPAddressConnect)
                        End If
                    End If
                End If
                'Exit
                Exit Sub
            End If

            'Exit
            Exit Sub

        End If

        'Check if in versus game
        If intCanvasMode = 8 Then

            'Exit if character doesn't exist
            If udcCharacterOne Is Nothing Or udcCharacterTwo Is Nothing Then
                Exit Sub
            End If

            ''Check if hoster or joiner
            'If blnHost Then
            '    'Exit if reloading, or game over
            '    If udcCharacterOne.IsReloading Or blnEndingGameCantType Or udcCharacterOne.BulletsUsed >= 30 Then
            '        Exit Sub
            '    End If
            '    'Check for spacebar and count bullets
            '    If CheckForSpacebar(e, udcCharacterOne, True) Then
            '        Exit Sub
            '    End If
            '    'Check the word being typed
            '    CheckTheWordBeingTyped(e, udcCharacterOne, gaudcZombiesOne)
            'Else
            '    'Exit if reloading, or game over
            '    If udcCharacterTwo.IsReloading Or blnEndingGameCantType Or udcCharacterTwo.BulletsUsed >= 30 Then
            '        Exit Sub
            '    End If
            '    'Check for spacebar and count bullets
            '    If CheckForSpacebar(e, udcCharacterTwo, True) Then
            '        Exit Sub
            '    End If
            '    'Check the word being typed
            '    CheckTheWordBeingTyped(e, udcCharacterTwo, gaudcZombiesTwo)
            'End If

            'Exit
            Exit Sub

        End If

    End Sub

    Private Function CheckForSpacebar(e As KeyPressEventArgs, udcCharacterType As clsCharacter, Optional blnSendData As Boolean = False) As Boolean

        'Check for spacebar
        If Asc(e.KeyChar) = 32 And blnCharacterCanReload() Then
            'Prepare to send data
            udcCharacterType.PrepareSendData = blnSendData
            'Character reloads manually
            udcCharacterType.Reload()
            'Return
            Return True
        End If

        'Return
        Return False

    End Function

    Private Function blnCharacterCanReload() As Boolean

        'Declare
        Dim blnTemp As Boolean = False

        'Check
        If blnGameIsVersus Then
            If blnHost Then
                If udcCharacterOne.BulletsUsed <> 0 Then
                    blnTemp = True
                End If
            Else
                If udcCharacterTwo.BulletsUsed <> 0 Then
                    blnTemp = True
                End If
            End If
        Else
            If udcCharacter.BulletsUsed <> 0 Then
                blnTemp = True
            End If
        End If

        'Return
        Return blnTemp

    End Function

    Private Sub CheckTheWordBeingTyped(e As KeyPressEventArgs, udcCharacterType As clsCharacter, audcZombiesType() As clsZombie)

        'Check the word being typed
        If strWord.Substring(0, 1) = LCase(e.KeyChar) Or strWord.Substring(0, 1) = UCase(e.KeyChar) Then
            'Change word by one less letter
            strWord = strWord.Substring(1)
            'Increase
            intWordIndex += 1
            'Check if word is done
            If Len(strWord) = 0 Then
                'Get a new word
                LoadARandomWord()
                'Check game type
                If blnGameIsVersus Then
                    'Check if hosting
                    If blnHost Then
                        'Declare
                        Dim intIndex As Integer = intGetIndexOfClosestZombie(audcZombiesType)
                        'Setup dying
                        audcZombiesType(intIndex).MarkedToDie = True
                        'Send data
                        gSendData("3|" & intIndex) 'Host sends data to joiner
                    Else
                        'Send data
                        gSendData("3|") 'Joiner sends data to host
                    End If
                Else
                    'Declare
                    Dim intIndex As Integer = intGetIndexOfClosestZombie(audcZombiesType)
                    'Setup dying
                    audcZombiesType(intIndex).MarkedToDie = True
                End If
                'Set to shoot
                udcCharacterType.StatusModeStartToDo = clsCharacter.eintStatusMode.Shoot
                'Increase bullet
                gintBullets += 1
                'Wait and create a sync with the class
                While gintBullets <> udcCharacterType.BulletsUsed
                End While
            End If
        End If

    End Sub

    Private Sub LoadARandomWord()

        'Delcare
        Dim rndNumber As New Random

        'Set
        strTheWord = astrWords(rndNumber.Next(0, astrWords.GetUpperBound(0) + 1))
        strWord = strTheWord

        'Set
        intWordIndex = 0

    End Sub

    Private Sub KeyPressBackspacing(ByRef blnByRefFirstTimeToCheck As Boolean, ByRef strByRefToChange As String)

        'Check if first time key pressing
        If blnByRefFirstTimeToCheck Then
            strByRefToChange = ""
            blnByRefFirstTimeToCheck = False
        Else
            If Len(strByRefToChange) <> 0 Then
                strByRefToChange = strByRefToChange.Substring(0, Len(strByRefToChange) - 1)
            End If
        End If

    End Sub

    Private Function intGetIndexOfClosestZombie(audcZombie() As clsZombie) As Integer

        'Declare
        Dim intClosestX As Integer = Integer.MaxValue
        Dim intIndex As Integer = 0

        'Loop to get closest zombie
        For intLoop As Integer = 0 To audcZombie.GetUpperBound(0)
            If audcZombie(intLoop).Spawned Then
                If Not audcZombie(intLoop).MarkedToDie Then
                    If intClosestX > audcZombie(intLoop).ZombiePoint.X Then
                        intClosestX = audcZombie(intLoop).ZombiePoint.X
                        intIndex = intLoop
                    End If
                End If
            End If
        Next

        'Return
        Return intIndex

    End Function

    Private Function strGetLocalIPAddress() As String

        'Declare
        Dim strTemp As String = ""

        'Loop
        For intLoop As Integer = 0 To System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList.GetUpperBound(0)
            If InStr(System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList(intLoop).ToString(), ":") = 0 Then
                strTemp = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList(intLoop).ToString()
                Exit For
            End If
        Next

        'Return IP
        Return strTemp

    End Function

    Private Sub Listening()

        'Loop
        While True
            'Check for pending
            Try
                If tcplServer.Pending Then
                    'Accept
                    tcpcClient = New System.Net.Sockets.TcpClient
                    tcpcClient = tcplServer.AcceptTcpClient
                    'Set streams, class, handlers, and prepare the next screen
                    PrepareTheConnectionState()
                    'Exit this thread
                    Exit While
                End If
            Catch ex As Exception
                'No debug
            End Try
        End While

    End Sub

    Private Sub Connecting()

        'Declare
        Dim blnConnected As Boolean = False

        'Loop
        While True
            'Try to connect
            Try
                tcpcClient = New System.Net.Sockets.TcpClient(strIPAddressConnect, 10101)
                blnConnected = True
            Catch ex As Exception
                'No debug
            End Try
            'Check if connected
            If blnConnected Then
                'Set streams, class, handlers, and prepare the next screen
                PrepareTheConnectionState()
                'Stop connecting
                Exit While
            End If
        End While

    End Sub

    Private Sub PrepareTheConnectionState()

        'Set streams
        gswClientData = New System.IO.StreamWriter(tcpcClient.GetStream())
        srClientData = New System.IO.StreamReader(tcpcClient.GetStream())

        'Set read line and check connection class
        udcVersusConnectedThread = New clsVersusConnectedThread(tcpcClient, srClientData)

        'Add handlers
        AddHandler udcVersusConnectedThread.GotData, AddressOf DataArrival
        AddHandler udcVersusConnectedThread.ConnectionGone, AddressOf ConnectionLost

        'Send data waiting for a connection
        If blnHost Then
            'Check
            While Not blnConnectionCompleted
                'Send
                gSendData("0|") 'Waiting for completed connection
                'Wait or else complications for exiting and reconnecting
                System.Threading.Thread.Sleep(1)
            End While
        End If

    End Sub

    Private Sub ParagraphingVersus()

        'Set
        btmLoadingParagraphVersus = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus25

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus50

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus75

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus100

    End Sub

    Private Sub LoadingVersusGame()

        'Notes: This procedure is for loading the game to play versus

        'Set
        If btmDeathScreen IsNot Nothing Then
            btmDeathScreen.Dispose()
            btmDeathScreen = Nothing
        End If

        'Set
        blnRemovedGameObjectsFromMemory = False

        'Set
        gpntGameBackground.X = 0

        'Set
        intLevel = 1

        'Set
        blnComparedHighscore = False

        'Set
        gintBullets = 0 'Starting bullets for key press

        'Set
        SetBlackScreenVariables()

        'Set
        blnGameIsVersus = True

        'Set
        intZombieKillsOne = 0
        intZombieKillsTwo = 0

        'Set
        intReloadTimeUsed = 0

        'Set to be fresh every game
        btmGameBackground = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground1.jpg"))

        'Set
        blnEndingGameCantType = False

        'Set
        blnGameOverFirstTime = True

        'Set
        blnBeatLevel = False

        'Check loaded game objects
        For intLoop As Integer = 0 To blnLoadingGameFinished.GetUpperBound(0) - 1
            'Create a pause waiting
            While Not blnLoadingGameFinished(intLoop)
            End While
            'Set
            btmLoadingBar = btmLoadingBarPicture(intLoop + 1)
        Next

        'Grab a random word
        LoadARandomWord()

        'Load in a special way
        If blnHost Then
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, False) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, True) 'Join
        Else
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, True) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, False) 'Join
        End If

        'Zombies
        LoadZombies("Level 1 Multiplayer")

        'Set if hosting
        If blnHost And Not blnReadyEarly Then
            blnWaiting = True
        End If

        'Check loaded game objects
        While Not blnLoadingGameFinished(blnLoadingGameFinished.GetUpperBound(0))
        End While

        'Set
        btmLoadingBar = btmLoadingBarPicture(blnLoadingGameFinished.GetUpperBound(0) + 1)

        'Set
        intCanvasShow = 1 'Means completely loaded

    End Sub

    Private Sub GamesMismatched(strGameVersionFromConnecter As String)

        'Set
        strGameVersionFromConnection = strGameVersionFromConnecter

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(10, 0, False)

        'Start thread
        thrGameMismatch = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Mismatching))
        thrGameMismatch.Start()

    End Sub

    Private Sub CheckingGameVersion(strData As String)

        'Check game version
        If GAME_VERSION <> strGetBlockData(strData, 1) Then
            GamesMismatched(strGetBlockData(strData, 1))
        Else
            PrepareVersusToPlayAfterMismatchPass()
        End If

    End Sub

    Private Sub PrepareVersusToPlayAfterMismatchPass()

        'Set
        intCanvasMode = 7

        'Start paragraphing
        thrParagraphVersus = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf ParagraphingVersus))
        thrParagraphVersus.Start()

        'Start loading game
        LoadingGameWithMultipleThreads()

        'Start loading game
        thrLoadingVersus = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingVersusGame))
        thrLoadingVersus.Start()

    End Sub

    Private Sub Mismatching()

        'Wait 5 seconds
        System.Threading.Thread.Sleep(5000)

        'Set
        blnPlayPressedSoundNow = False

        'Disconnect
        GoBackToMenuConnectionGone()

    End Sub

    Private Sub DataArrival(strData As String)

        'Show data
        'Debug.Print(strData)

        'Check if host
        If blnHost Then

            'Check data
            Select Case (strGetBlockData(strData))
                Case "0" 'Completely connected
                    'Send data once
                    If Not blnConnectionCompleted Then
                        'Send data of game version
                        gSendData("1|" & GAME_VERSION) 'Will exit if not matching
                        'Set
                        blnConnectionCompleted = True
                    End If
                Case "1"
                    'Check game versions
                    CheckingGameVersion(strData)
                Case "2" 'Waiting to start game
                    'Set name
                    strNicknameConnected = strGetBlockData(strData, 1)
                    'Ready to play but waiting for host to hit start
                    blnWaiting = False
                    blnReadyEarly = True
                Case "3" 'Show join has shot
                    'Declare
                    Dim intIndex As Integer = intGetIndexOfClosestZombie(gaudcZombiesTwo)
                    'Send data
                    gSendData("4|" & intIndex)
                    'Setup dying
                    gaudcZombiesTwo(intIndex).MarkedToDie = True
                Case "4" 'Not used
                Case "5" 'Show join reloading
                    udcCharacterTwo.Reload()
                Case "6" 'Not used
                Case "7" 'Not used
            End Select

        Else 'Joiner

            'Check data
            Select Case (strGetBlockData(strData))
                Case "0" 'Send data back
                    gSendData("0|") 'Waiting for host to connect
                Case "1"
                    'Send data of game version
                    gSendData("1|" & GAME_VERSION) 'Will exit if not matching
                    'Check game versions
                    CheckingGameVersion(strData)
                Case "2"
                    'Set name
                    strNicknameConnected = strGetBlockData(strData, 1)
                    'Start was hit, game started for joiner
                    ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(8, 0, False)
                    blnWaiting = False
                    'Started versus game, start objects
                    StartVersusGameObjects()
                Case "3" 'Show host has shot
                    'Setup dying
                    gaudcZombiesOne(CInt(strGetBlockData(strData, 1))).MarkedToDie = True
                Case "4" 'Kill zombie
                    'Setup dying
                    gaudcZombiesTwo(CInt(strGetBlockData(strData, 1))).MarkedToDie = True
                Case "5" 'Show host reloading
                    'Reload
                    udcCharacterOne.Reload()
                Case "6", "7" 'End game, hoster died
                    'Set
                    blnEndingGameCantType = True
                    'Set
                    blnGameOverFirstTime = False
                    'Stop reloading sound
                    udcCharacterOne.StopReloadingSound()
                    udcCharacterTwo.StopReloadingSound()
                    'Play
                    udcScreamSound = New clsSound(Me, strDirectory & "Sounds\CharacterDying.mp3", 3000, gintSoundVolume, False)
                    'Check for who won
                    If strGetBlockData(strData) = "6" Then
                        'Joiner won
                        thrEndGame = New System.Threading.Thread(Sub() EndingGame(True))
                    Else
                        'Joiner lost
                        thrEndGame = New System.Threading.Thread(Sub() EndingGame())
                    End If
                    thrEndGame.Start()
            End Select

        End If

    End Sub

    Private Function strGetBlockData(strData As String, Optional intArrayElement As Integer = 0) As String

        'Notes: Data looks like "X|" or "X|String"

        'Declare
        Dim astrTemp() As String = Split(strData, "|")

        'Return
        Return astrTemp(intArrayElement)

    End Function

    Private Sub ConnectionLost()

        'Set
        blnPlayPressedSoundNow = False

        'Go back to menu
        GoBackToMenuConnectionGone()

    End Sub

End Class