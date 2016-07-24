'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const GAME_VERSION As String = "1.35"
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
    Private Const MOUSE_OVER_SOUND_DELAY As Double = 35 '35 milliseconds

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

    'Options mouse over
    Private tmrOptionsMouseOver As New System.Timers.Timer()
    Private strOptionsButtonSpot As String = ""

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
    Private abtmSound(100) As Bitmap '0 to 100
    Private pntSoundPercent As New Point(718, 553)

    'Loading screen
    Private btmLoadingBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\LoadingBackground.jpg"))
    Private abtmLoadingBarPicture(10) As Bitmap '0 To 10 = 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
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
    Private athrLoading(9) As System.Threading.Thread '0 to 9 which is 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
    Private ablnLoadingGameFinished(9) As Boolean '0 to 9 which is 10, 20, 30, 40, 50, 60, 70, 80, 90, 100

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
    Private btmAK47Magazine As Bitmap
    Private pntAK47Magazine As New Point(59, 877)
    Private btmWinOverlay As Bitmap
    Private intLevel As Integer = 1 'Starting
    Private intZombieKills As Integer = 0
    Private intReloadTimeUsed As Integer = 0
    Private intZombieKillsCombined As Integer = 0
    Private intReloadTimeUsedCombined As Integer = 0
    Private blnComparedHighscore As Boolean = False
    Private blnBeatLevel As Boolean = False

    'Key press
    Private strKeyPressBuffer As String = ""
    Private blnPreventKeyPressEvent As Boolean = False

    'Sounds
    Private audcAmbianceSound(1) As clsSound '2 ambiance sounds so far
    Private udcButtonHoverSound As clsSound
    Private udcButtonPressedSound As clsSound
    Private audcStoryParagraphSound(2) As clsSound '3 paragraph sounds in the story area
    Private audcGameBackgroundSound(4) As clsSound '5 background sounds so far
    Private udcScreamSound As clsSound
    Private udcGunShotSound As clsSound
    Private audcZombieDeathSound(1) As clsSound '2 zombie death sounds so far
    Private udcReloadingSound As clsSound
    Private udcStepSound As clsSound
    Private udcWaterFootStepLeftSound As clsSound
    Private udcWaterFootStepRightSound As clsSound
    Private udcGravelFootStepLeftSound As clsSound
    Private udcGravelFootStepRightSound As clsSound
    Private udcOpeningMetalDoorSound As clsSound
    Private udcLightZapSound As clsSound
    Private udcZombieGrowlSound As clsSound
    Private udcRotatingBladeSound As clsSound 'Helicopter

    'Sounds multiplayer

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

    'Game version mismatch
    Private btmGameMismatchBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Game Mismatch\GameMismatchBackground.jpg"))
    Private strGameVersionFromConnection As String = ""
    Private thrGameMismatch As System.Threading.Thread

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes: Do not allow a resize of the window if full screen, this happens if you double click the title, but is prevented with this sub.

        'Check if full screen
        If intResolutionMode = 5 Then
            'Prevent moving the form by control box click
            If m.Msg = WINDOW_MESSAGE_SYSTEM_COMMAND And m.WParam.ToInt32() = CONTROL_MOVE Then
                Return
            End If
            'Prevent button down moving form
            If m.Msg = WINDOW_MESSAGE_CLICK_BUTTON_DOWN And m.WParam.ToInt32() = WINDOW_CAPTION Then
                Return
            End If
        End If

        'If a double click on the title bar is triggered
        If m.Msg = WINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
            Return
        End If

        'Still send message
        MyBase.WndProc(m) 'Must have mybase

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

        'Setup mouse over timer
        SetupMouseOverTimer()

        'Load sound file percentages
        For intLoop As Integer = 0 To abtmSound.GetUpperBound(0)
            abtmSound(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\Options\Sound" & CStr(intLoop) & ".png"))
        Next

        'Set 100%
        btmSoundPercent = abtmSound(100)

        'Load loading bar pictures
        LoadLoadingBarPictures()

        'Set for loading
        btmLoadingBar = abtmLoadingBarPicture(0)

    End Sub

    Private Sub SetupMouseOverTimer()

        'Disable timer by default, a mouse over the options will enable it later
        tmrOptionsMouseOver.Enabled = False

        'Set timer
        tmrOptionsMouseOver.AutoReset = True

        'Add handlers
        AddHandler tmrOptionsMouseOver.Elapsed, AddressOf ElapsedOptionsMouseOver

        'Set timer delay
        tmrOptionsMouseOver.Interval = MOUSE_OVER_SOUND_DELAY

    End Sub

    Private Sub ElapsedOptionsMouseOver(sender As Object, e As EventArgs)

        'Disable timer
        tmrOptionsMouseOver.Enabled = False

        'Play sound
        udcButtonHoverSound.PlaySound(gintSoundVolume)

    End Sub

    Private Sub LoadLoadingBarPictures()

        'Declare
        Dim intPictureCounter As Integer = 0

        'Load loading bar pictures
        For intLoop As Integer = 0 To abtmLoadingBarPicture.GetUpperBound(0)
            'Set
            abtmLoadingBarPicture(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\Loading Game\" & CStr(intPictureCounter) & ".png"))
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

        'Loading game
        AbortThread(thrLoadingGame)

        'Set
        thrLoadingGame = Nothing

        'Loading
        For intLoop As Integer = 0 To athrLoading.GetUpperBound(0)
            AbortThread(athrLoading(intLoop))
            'Set
            athrLoading(intLoop) = Nothing
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

        'Stop all sounds
        StopAndCloseAllSounds()

        'Dipose graphics
        gGraphics.Dispose()

    End Sub

    Private Sub StopAndCloseAllSounds()

        'Remove options mouse over timer
        RemoveOptionsMouseOverTimer()

        'Ambiance
        For intLoop As Integer = 0 To audcAmbianceSound.GetUpperBound(0)
            If audcAmbianceSound IsNot Nothing Then
                audcAmbianceSound(intLoop).StopAndCloseSound()
                audcAmbianceSound(intLoop) = Nothing
            End If
        Next

        'Button pressed
        If udcButtonPressedSound IsNot Nothing Then
            udcButtonPressedSound.StopAndCloseSound()
            udcButtonPressedSound = Nothing
        End If

        'Button hover
        If udcButtonHoverSound IsNot Nothing Then
            udcButtonHoverSound.StopAndCloseSound()
            udcButtonHoverSound = Nothing
        End If

        'Story paragraphs
        For intLoop As Integer = 0 To audcStoryParagraphSound.GetUpperBound(0)
            If audcStoryParagraphSound(intLoop) IsNot Nothing Then
                audcStoryParagraphSound(intLoop).StopAndCloseSound()
                audcStoryParagraphSound(intLoop) = Nothing
            End If
        Next

        'Game backgrounds
        For intLoop As Integer = 0 To audcGameBackgroundSound.GetUpperBound(0)
            If audcGameBackgroundSound(intLoop) IsNot Nothing Then
                audcGameBackgroundSound(intLoop).StopAndCloseSound()
                audcGameBackgroundSound(intLoop) = Nothing
            End If
        Next

        'Scream
        If udcScreamSound IsNot Nothing Then
            udcScreamSound.StopAndCloseSound()
            udcScreamSound = Nothing
        End If

        'Gun shot
        If udcGunShotSound IsNot Nothing Then
            udcGunShotSound.StopAndCloseSound()
            udcGunShotSound = Nothing
        End If

        'Zombie deaths
        For intLoop As Integer = 0 To audcZombieDeathSound.GetUpperBound(0)
            If audcZombieDeathSound(intLoop) IsNot Nothing Then
                audcZombieDeathSound(intLoop).StopAndCloseSound()
                audcZombieDeathSound(intLoop) = Nothing
            End If
        Next

        'Reloading
        If udcReloadingSound IsNot Nothing Then
            udcReloadingSound.StopAndCloseSound()
            udcReloadingSound = Nothing
        End If

        'Step
        If udcStepSound IsNot Nothing Then
            udcStepSound.StopAndCloseSound()
            udcStepSound = Nothing
        End If

        'Water foot step left
        If udcWaterFootStepLeftSound IsNot Nothing Then
            udcWaterFootStepLeftSound.StopAndCloseSound()
            udcWaterFootStepLeftSound = Nothing
        End If

        'Water foot step right
        If udcWaterFootStepRightSound IsNot Nothing Then
            udcWaterFootStepRightSound.StopAndCloseSound()
            udcWaterFootStepRightSound = Nothing
        End If

        'Gravel foot step left
        If udcGravelFootStepLeftSound IsNot Nothing Then
            udcGravelFootStepLeftSound.StopAndCloseSound()
            udcGravelFootStepLeftSound = Nothing
        End If

        'Gravel foot step right
        If udcGravelFootStepRightSound IsNot Nothing Then
            udcGravelFootStepRightSound.StopAndCloseSound()
            udcGravelFootStepRightSound = Nothing
        End If

        'Opening metal door
        If udcOpeningMetalDoorSound IsNot Nothing Then
            udcOpeningMetalDoorSound.StopAndCloseSound()
            udcOpeningMetalDoorSound = Nothing
        End If

        'Light zap
        If udcLightZapSound IsNot Nothing Then
            udcLightZapSound.StopAndCloseSound()
            udcLightZapSound = Nothing
        End If

        'Zombie growl
        If udcZombieGrowlSound IsNot Nothing Then
            udcZombieGrowlSound.StopAndCloseSound()
            udcZombieGrowlSound = Nothing
        End If

        'Rotating blade
        If udcRotatingBladeSound IsNot Nothing Then
            udcRotatingBladeSound.StopAndCloseSound()
            udcRotatingBladeSound = Nothing
        End If

    End Sub

    Private Sub RemoveOptionsMouseOverTimer()

        'Disable timers
        tmrOptionsMouseOver.Enabled = False

        'Stop and dispose timer
        tmrOptionsMouseOver.Stop()
        tmrOptionsMouseOver.Dispose()

        'Remove handlers
        RemoveHandler tmrOptionsMouseOver.Elapsed, AddressOf ElapsedOptionsMouseOver

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
                udcScreamSound.StopSound()
            End If
        End If

        'Stop game background sounds
        For intLoop As Integer = 0 To audcGameBackgroundSound.GetUpperBound(0)
            If audcGameBackgroundSound(intLoop) IsNot Nothing Then
                audcGameBackgroundSound(intLoop).StopSound()
            End If
        Next

        'Stop helicopter blade sound
        If udcRotatingBladeSound IsNot Nothing Then
            udcRotatingBladeSound.StopSound()
        End If

        'Stop and dispose helicopter
        If udcHelicopter IsNot Nothing Then
            'Stop and dispose
            udcHelicopter.StopAndDispose()
            udcHelicopter = Nothing
        End If

        'Stop reloading sound
        If udcReloadingSound IsNot Nothing Then
            udcReloadingSound.StopSound()
        End If

        'Stop and dispose character, remove handler
        If udcCharacter IsNot Nothing Then
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

        'Stop and dispose versus character (host)
        If udcCharacterOne IsNot Nothing Then
            'Stop and dispose
            udcCharacterOne.StopAndDispose()
            udcCharacterOne = Nothing
        End If

        'Stop and dispose versus character (joiner)
        If udcCharacterTwo IsNot Nothing Then
            'Stop and dispose
            udcCharacterTwo.StopAndDispose()
            udcCharacterTwo = Nothing
        End If

        'Stop and dispose zombies (host)
        If gaudcZombiesOne IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                If gaudcZombiesOne(intLoop) IsNot Nothing Then
                    gaudcZombiesOne(intLoop).StopAndDispose()
                    gaudcZombiesOne(intLoop) = Nothing
                End If
            Next
        End If

        'Stop and dispose zombies (joiner)
        If gaudcZombiesTwo IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
                If gaudcZombiesTwo(intLoop) IsNot Nothing Then
                    gaudcZombiesTwo(intLoop).StopAndDispose()
                    gaudcZombiesTwo(intLoop) = Nothing
                End If
            Next
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
            'Load necessary sounds for the beginning engine
            LoadNecessarySoundsForBeginningEngine()
            'Get highscores early because grabbing information from the database access files is slow
            LoadHighscoresStringAccess()
            'Set percentage multiplers for screen modes
            gdblScreenWidthRatio = CDbl((intScreenWidth - WIDTH_SUBTRACTION) / ORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((intScreenHeight - HEIGHT_SUBTRACTION) / ORIGINAL_SCREEN_HEIGHT)
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - WIDTH_SUBTRACTION, intScreenHeight - HEIGHT_SUBTRACTION) 'Full screen
            'Set for fog
            thrFog = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf KeepFogMoving))
            thrFog.Start()
            'Start rendering
            thrRendering.Start()
        End If

    End Sub

    Private Sub LoadNecessarySoundsForBeginningEngine()

        'Load ambiance
        For intLoop As Integer = 0 To audcAmbianceSound.GetUpperBound(0)
            audcAmbianceSound(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Ambiance" & CStr(intLoop + 1) & ".mp3", 1) 'Repeat only needs 1
        Next

        'Load button pressed
        udcButtonPressedSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonPressed.mp3", 1)

        'Load button hover
        udcButtonHoverSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonHover.mp3", 5)

        'Load story paragraphs
        For intLoop As Integer = 0 To audcStoryParagraphSound.GetUpperBound(0)
            audcStoryParagraphSound(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\StoryParagraph" & CStr(intLoop + 1) & ".mp3", 1)
        Next

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
                btmLoadingBar = abtmLoadingBarPicture(0)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0, blnPlayPressedSoundNow)
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

        'Draw options text
        If intCanvasShow = 4 Then
            gGraphics.DrawImageUnscaled(btmOptionsHoverText, pntOptionsHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmOptionsText, pntOptionsText)
        End If

        'Draw credits text
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
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(11, 0, False) 'This exits the game screen to path choices
                    Case 2, 3
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Set
                        btmPath = btmPath2_0
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(12, 0, False) 'This exits the game screen to path choices
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

            'Play single player game
            PlaySinglePlayerGame()

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

    Private Sub PlaySinglePlayerGame()

        'Paint on copy of the background
        PaintOnBitmap(btmGameBackground)

        'Check for helicopter
        If udcHelicopter IsNot Nothing Then
            'Draw
            gGraphics.DrawImageUnscaled(udcHelicopter.HelicopterImage, udcHelicopter.HelicopterPoint)
        End If

        'Draw dead zombies permanently
        For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
            If gaudcZombies(intLoop).Spawned And gaudcZombies(intLoop).IsDead Then
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
        Next

        'Paint on canvas
        PaintOnBitmap(btmCanvas)

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)

        'Clone the only necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackground.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Draw the background to screen with the cloned version
        gGraphics.DrawImageUnscaled(btmGameBackgroundCloneScreenShown, pntTopLeft)

        'Dispose because clone just makes the picture expand with more data
        btmGameBackgroundCloneScreenShown.Dispose()
        btmGameBackgroundCloneScreenShown = Nothing

        'Draw word bar
        gGraphics.DrawImageUnscaled(btmWordBar, pntWordBar)

        'Check if made it to the end of the level
        If gpntGameBackground.X <= -2750 Then
            'Check if level was already beaten or if a zombie grabbed the player
            If Not blnBeatLevel And Not blnTriggeredBlackScreenThread Then
                'Set
                blnBeatLevel = True
                'Set
                blnPreventKeyPressEvent = True
                'Make character stand because he was running
                If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                    'Stand
                    udcCharacter.Stand()
                End If
                'Door sound
                Select Case intLevel
                    Case 1, 2
                        'Play door
                        udcOpeningMetalDoorSound.PlaySound(gintSoundVolume)
                End Select
                'Start black screen thread
                thrBlackScreen = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf BlackScreening))
                thrBlackScreen.Start()
                'Stop level music
                audcGameBackgroundSound(intLevel - 1).StopSound()
                'Pause the stop watch
                swhStopWatch.Stop()
                'Set reload time
                intReloadTimeUsed = udcCharacter.ReloadTimes * 3
                'Keep the reload time updated
                intReloadTimeUsedCombined += intReloadTimeUsed
                'Keep the zombie kills updated
                intZombieKillsCombined += intZombieKills
            End If
        Else

            'Check character status if not game over
            If Not blnBeatLevel And Not blnTriggeredBlackScreenThread Then
                'Check character status
                CheckCharacterStatus()
            End If

        End If

        'Draw character
        gGraphics.DrawImageUnscaled(udcCharacter.CharacterImage, udcCharacter.CharacterPoint)

        'Draw zombies
        For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
            'Check if spawned
            If gaudcZombies(intLoop).Spawned Then
                'Check if can pin
                If Not gaudcZombies(intLoop).IsDying Then
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
                                blnPreventKeyPressEvent = True
                                'Stop character from moving
                                If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                                    udcCharacter.Stand()
                                End If
                                'Start black screen thread
                                thrBlackScreen = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf BlackScreening))
                                thrBlackScreen.Start()
                                'Stop reloading sound
                                udcReloadingSound.StopSound()
                                'Play
                                udcScreamSound.PlaySound(gintSoundVolume)
                                'Stop level music
                                audcGameBackgroundSound(intLevel - 1).StopSound()
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

        'Show magazine with bullet count
        gGraphics.DrawImageUnscaled(btmAK47Magazine, pntAK47Magazine)

        'Draw bullet count on magazine
        DrawText(gGraphics, CStr(30 - udcCharacter.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)

        'Check if black screen needs to be drawed
        If btmBlackScreen IsNot Nothing Then
            gGraphics.DrawImageUnscaled(btmBlackScreen, pntTopLeft)
        End If

        'Make copy if died
        MakeCopyOfScreenBecauseCharacterDied()

    End Sub

    Private Sub MakeCopyOfScreenBecauseCharacterDied()

        'Check for death, copy only if black screen is 50% opacity
        If blnTriggeredBlackScreenThread Then
            'Only copy if not existent for first time
            If btmBlackScreen Is btmBlackScreen50 Then
                'If death screen shot hasn't been made yet
                If btmDeathScreen Is Nothing Then
                    'Before fading the screen, copy it to show for the death overlay
                    Dim rectSource As New Rectangle(0, 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)
                    btmDeathScreen = btmCanvas.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)
                End If
            End If
        End If

    End Sub

    Private Sub CheckCharacterStatus()

        'Check what to do at this moment
        If udcCharacter.StopCharacterFromRunning Then

            'Set
            udcCharacter.StopCharacterFromRunning = False
            'Clear the key press buffer
            strKeyPressBuffer = ""
            'Make character stand
            udcCharacter.Stand()

        Else

            'Check to see if character must prepare to do something first
            Select Case udcCharacter.StatusModeAboutToDo

                Case clsCharacter.eintStatusMode.Reload
                    'Clear the key press buffer
                    strKeyPressBuffer = ""
                    'Make the character reload
                    If udcCharacter.GetPictureFrame = 1 Then
                        'Reload
                        udcCharacter.Reload()
                        'Set
                        udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                    End If

                Case clsCharacter.eintStatusMode.Run
                    'Clear the key press buffer
                    strKeyPressBuffer = ""
                    'Make the character run
                    If udcCharacter.GetPictureFrame = 1 Then
                        'Run
                        udcCharacter.Run()
                        'Set
                        udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                    End If

                Case Else 'What is the character going to currently do

                    'If not stopping, what else is the character going to do
                    Select Case udcCharacter.StatusModeStartToDo

                        Case clsCharacter.eintStatusMode.Run
                            'Clear the key press buffer
                            strKeyPressBuffer = ""
                            'Make the character run
                            udcCharacter.Run()

                        Case clsCharacter.eintStatusMode.Reload
                            'Clear the key press buffer
                            strKeyPressBuffer = ""
                            'Make the character reload
                            udcCharacter.Reload()

                        Case clsCharacter.eintStatusMode.Stand, clsCharacter.eintStatusMode.Shoot
                            'Make the character shoot with buffer
                            CheckTheKeyPressBuffer()

                    End Select

            End Select

        End If

    End Sub

    Private Sub CheckTheKeyPressBuffer()

        'Check the key press buffer
        If strKeyPressBuffer <> "" Then

            'Declare to be split
            Dim astrTemp() As String = Split(strKeyPressBuffer, ".")

            'Keep an index to remove only the necessary pieces of the key press buffer
            Dim intIndexesToRemove As Integer = 0

            'Loop through
            For intLoop As Integer = 0 To astrTemp.GetUpperBound(0)

                'Increase
                intIndexesToRemove += 1

                'Check if typing matches
                If strWord.Substring(0, 1) = LCase(astrTemp(intLoop)) Then

                    'Change the word by one less letter
                    strWord = strWord.Substring(1)

                    'Increase
                    intWordIndex += 1

                    'Check if word is done
                    If Len(strWord) = 0 Then

                        'Get a new word
                        LoadARandomWord()

                        'Shoot and kill a zombie
                        udcCharacter.Shoot()

                        'Kill closest zombie
                        gaudcZombies(intGetIndexOfClosestZombie(gaudcZombies)).Dying()

                        'Check if needs to reload
                        If udcCharacter.BulletsUsed = 30 Then

                            'Set
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload

                            'Reset
                            strKeyPressBuffer = ""

                            'Exit
                            Exit Sub

                        End If

                    End If

                End If

            Next

            'Remove only the necessary
            Dim strIndexPieces As String = ""

            'Loop
            For intLoop As Integer = 0 To intIndexesToRemove - 1
                strIndexPieces &= "." & astrTemp(intLoop)
            Next

            'Remove the piece
            strKeyPressBuffer = strKeyPressBuffer.Substring(Len(strIndexPieces) - 1)

        End If

    End Sub

    Private Function GetNumberToPin(audcZombiesType() As clsZombie, intNumberToCheck As Integer) As Boolean

        'Loop to get the correct loop number
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'If spawned, if dying, and if pin value = a number already existing in a zombie value
            If audcZombiesType(intLoop).Spawned And Not audcZombiesType(intLoop).IsDying And audcZombiesType(intLoop).PinXYValueChanged = intNumberToCheck Then
                Return True
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
            blnPreventKeyPressEvent = True
            'Set
            blnGameOverFirstTime = False
            'Stop reloading sound and send data
            If blnGameIsVersus Then
                'Send data
                gSendData(strData)
            End If
            'Stop reloading sound
            udcReloadingSound.StopSound()
            'Play
            udcScreamSound.PlaySound(gintSoundVolume)
            'Set
            thrEndGame = New System.Threading.Thread(Sub() EndingGame(blnEndGame))
            thrEndGame.Start()
        End If

    End Sub

    Private Function intGetIndexOfClosestZombie(audcZombie() As clsZombie) As Integer

        'Declare
        Dim intClosestX As Integer = Integer.MaxValue
        Dim intIndex As Integer = 0

        'Loop to get closest zombie
        For intLoop As Integer = 0 To audcZombie.GetUpperBound(0)
            'If spawned, if not dying, and if closest point is less than another point
            If audcZombie(intLoop).Spawned And Not audcZombie(intLoop).IsDying And intClosestX > audcZombie(intLoop).ZombiePoint.X Then
                'Set closest
                intClosestX = audcZombie(intLoop).ZombiePoint.X
                'Set index
                intIndex = intLoop
            End If
        Next

        'Return
        Return intIndex

    End Function

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
            'PaintBackgroundAndWord()

            'Draw character hoster
            gGraphics.DrawImageUnscaled(udcCharacterOne.CharacterImage, udcCharacterOne.CharacterPoint)

            'First horde of zombies for hoster
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                If gaudcZombiesOne(intLoop) IsNot Nothing Then
                    'Check if spawned
                    If gaudcZombiesOne(intLoop).Spawned Then
                        'Check if can pin
                        If Not gaudcZombiesOne(intLoop).IsDying Then
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
        audcStoryParagraphSound(0).PlaySound(gintSoundVolume)

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
        audcStoryParagraphSound(1).PlaySound(gintSoundVolume)

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
        audcStoryParagraphSound(2).PlaySound(gintSoundVolume)

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

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x and y points to ensure it always works.

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'Print for mouse cordinates if necessary
        'Debug.Print("Mouse point = " & CStr(pntMouse.X) & ", " & CStr(pntMouse.Y))

        'Which screen is showing
        Select Case intCanvasMode

            Case 0 'Menu
                MenuMouseOverScreen(pntMouse)

            Case 1 'Options
                OptionsMouseOverScreen(pntMouse)

            Case 3 'Game
                OneBackButtonMouseOverScreen(pntMouse, "GameBack")

            Case 4 'Highscores
                OneBackButtonMouseOverScreen(pntMouse, "HighscoresBack")

            Case 5 'Credits
                OneBackButtonMouseOverScreen(pntMouse, "CreditsBack")

            Case 6 'Versus
                OneBackButtonMouseOverScreen(pntMouse, "VersusBack")

            Case 8 'Versus game
                OneBackButtonMouseOverScreen(pntMouse, "VersusGameBack")

            Case 9 'Story
                OneBackButtonMouseOverScreen(pntMouse, "StoryBack")

            Case 11 'Path 1 choice
                Path1ChoiceMouseOverScreen(pntMouse)

            Case 12 'Path 2 choice
                Path2ChoiceMouseOverScreen(pntMouse)

        End Select

    End Sub

    Private Sub MenuMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 212, 69, pntStartText) 'Start has been moused over
                'Hover sound
                HoverText(1, "MenuStart")

            Case blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) 'Highscores has been moused over
                'Hover sound
                HoverText(2, "MenuHighscores")

            Case blnMouseInRegion(pntMouse, 218, 87, pntStoryText) 'Story has been moused over
                'Hover sound
                HoverText(3, "MenuStory")

            Case blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) 'Options has been moused over
                'Hover sound
                HoverText(4, "MenuOptions")

            Case blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) 'Credits has been moused over
                'Hover sound
                HoverText(5, "MenuCredits")

            Case blnMouseInRegion(pntMouse, 256, 74, pntVersusText) 'Versus has been moused over
                'Hover sound
                HoverText(6, "MenuVersus")

            Case Else
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

    End Sub

    Private Sub HoverText(intCanvasShowToBe As Integer, strOptionsButtonSpotToBe As String)

        'Set canvas show
        intCanvasShow = intCanvasShowToBe

        'Check if same button spot as previously
        If strOptionsButtonSpot <> strOptionsButtonSpotToBe Then
            'Set
            strOptionsButtonSpot = strOptionsButtonSpotToBe
            'Set timer on
            tmrOptionsMouseOver.Enabled = True
        End If

    End Sub

    Private Sub ResetMouseOverVariables()

        'Set timer off
        tmrOptionsMouseOver.Enabled = False

        'Reset
        strOptionsButtonSpot = ""

        'Repaint menu background
        intCanvasShow = 0

    End Sub

    Private Sub OptionsMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back has been moused over
                'Hover sound
                HoverText(1, "OptionsBack")

            Case blnSliderWithMouseDown 'Slider has been moused over
                'Change volume
                ChangeSoundVolume(pntMouse)
                'Reset mouse over variables
                ResetMouseOverVariables()

            Case Else
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

    End Sub

    Private Sub OneBackButtonMouseOverScreen(pntMouse As Point, strOptionsButtonSpotToSet As String)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText)
                'Hover sound
                HoverText(1, strOptionsButtonSpotToSet)

            Case Else
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

    End Sub

    Private Sub Path1ChoiceMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText)
                'Hover sound
                HoverText(1, "Path1Back")

            Case blnMouseInRegion(pntMouse, 389, 329, New Point(230, 427)) 'Path left
                'Set
                btmPath = btmPath1_1
                'Play light switch
                If Not blnLightZap1 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap1 = True
                End If

            Case blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) 'Path right
                'Set
                btmPath = btmPath1_2
                'Play light switch
                If Not blnLightZap2 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap2 = True
                End If

            Case Else
                'Set
                btmPath = btmPath1_0
                'Reset
                blnLightZap1 = False
                blnLightZap2 = False
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

    End Sub

    Private Sub Path2ChoiceMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText)
                'Hover sound
                HoverText(1, "Path2Back")

            Case blnMouseInRegion(pntMouse, 296, 304, New Point(643, 306)) 'Path left
                'Set
                btmPath = btmPath2_1
                'Play light switch
                If Not blnLightZap1 Then
                    'Play light zap sound
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Play zombie growl sound
                    udcZombieGrowlSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap1 = True
                End If

            Case blnMouseInRegion(pntMouse, 297, 318, New Point(1138, 384)) 'Path right
                'Set
                btmPath = btmPath2_2
                'Play light switch
                If Not blnLightZap2 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap2 = True
                End If

            Case Else
                'Stop zombie growl
                udcZombieGrowlSound.StopSound()
                'Set
                btmPath = btmPath2_0
                'Reset
                blnLightZap1 = False
                blnLightZap2 = False
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

    End Sub

    Private Sub ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer,
                                                                     Optional blnPlayPressedSound As Boolean = True)

        'Set
        intCanvasMode = intCanvasModeToSet

        'Set
        intCanvasShow = intCanvasShowToSet

        'Play sound
        If blnPlayPressedSound Then
            udcButtonPressedSound.PlaySound(gintSoundVolume)
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
        blnPreventKeyPressEvent = False

        'Set
        strKeyPressBuffer = ""

        'Set
        blnBeatLevel = False

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                        udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Helicopter
        If blnLoadHelicopter Then
            udcHelicopter = New clsHelicopter(Me, udcRotatingBladeSound)
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

        'Play background sound music
        audcGameBackgroundSound(intLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)

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
        blnPreventKeyPressEvent = False

        'Set
        strKeyPressBuffer = ""

        'Set
        blnGameOverFirstTime = True

        'Set
        blnBeatLevel = False

        'Check loaded game objects
        For intLoop As Integer = 0 To ablnLoadingGameFinished.GetUpperBound(0) - 1
            'Create a pause waiting
            While Not ablnLoadingGameFinished(intLoop)
            End While
            'Set
            btmLoadingBar = abtmLoadingBarPicture(intLoop + 1)
        Next

        'Grab a random word
        LoadARandomWord()

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                        udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Check loaded game objects
        While Not ablnLoadingGameFinished(ablnLoadingGameFinished.GetUpperBound(0))
        End While

        'Set
        btmLoadingBar = abtmLoadingBarPicture(ablnLoadingGameFinished.GetUpperBound(0) + 1)

        'Set
        intCanvasShow = 1 'Means completely loaded

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

    Private Sub LoadZombies(strGameType As String)

        'Check if multiplayer or not
        Select Case strGameType
            'What level and type
            Case "Level 1 Single Player"
                'Zombies
                For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                    'Re-dim first
                    ReDim Preserve gaudcZombies(intLoop)
                    'Check which wave number
                    Select Case intLoop
                        Case 0
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed, "udcCharacter", audcZombieDeathSound)
                        Case 1
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, intZombieSpeed, "udcCharacter", audcZombieDeathSound)
                        Case 2
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, intZombieSpeed, "udcCharacter", audcZombieDeathSound)
                        Case 3
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, intZombieSpeed, "udcCharacter", audcZombieDeathSound)
                        Case 4
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, intZombieSpeed, "udcCharacter", audcZombieDeathSound)
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 5, "udcCharacter", audcZombieDeathSound)
                        Case 9, 14, 19, 24
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 25, "udcCharacter", audcZombieDeathSound)
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 10, "udcCharacter", audcZombieDeathSound)
                        Case 29, 34, 39, 44
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 30, "udcCharacter", audcZombieDeathSound)
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 15, "udcCharacter", audcZombieDeathSound)
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 35, "udcCharacter", audcZombieDeathSound)
                        Case 65 To 74
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 5, "udcCharacter", audcZombieDeathSound)
                        Case 75 To 84
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 10, "udcCharacter", audcZombieDeathSound)
                        Case 85 To 94
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 15, "udcCharacter", audcZombieDeathSound)
                        Case 95 To 104
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 20, "udcCharacter", audcZombieDeathSound)
                        Case Else
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeed + 45, "udcCharacter", audcZombieDeathSound)
                    End Select
                Next
                '    'What level and type
                'Case "Level 1 Multiplayer"
                '    'First set of multiplayer zombies
                '    For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                '        'Re-dim first
                '        ReDim Preserve gaudcZombiesOne(intLoop)
                '        'Check which wave number
                '        Select Case intLoop
                '            Case 0
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne, "udcCharacterOne")
                '            Case 1
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, intZombieSpeedOne, "udcCharacterOne")
                '            Case 2
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, intZombieSpeedOne, "udcCharacterOne")
                '            Case 3
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, intZombieSpeedOne, "udcCharacterOne")
                '            Case 4
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, intZombieSpeedOne, "udcCharacterOne")
                '            Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 5, "udcCharacterOne")
                '            Case 9, 14, 19, 24
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 25, "udcCharacterOne")
                '            Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 10, "udcCharacterOne")
                '            Case 29, 34, 39, 44
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 30, "udcCharacterOne")
                '            Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 15, "udcCharacterOne")
                '            Case 48, 49, 53, 54, 58, 59, 63, 64
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 35, "udcCharacterOne")
                '            Case 65 To 74
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 5, "udcCharacterOne")
                '            Case 75 To 84
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 10, "udcCharacterOne")
                '            Case 85 To 94
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 15, "udcCharacterOne")
                '            Case 95 To 104
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 20, "udcCharacterOne")
                '            Case Else
                '                gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, intZombieSpeedOne + 45, "udcCharacterOne")
                '        End Select
                '    Next
                '    'Second set of multiplayer zombies
                '    For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)
                '        'Re-dim first
                '        ReDim Preserve gaudcZombiesTwo(intLoop)
                '        'Check which wave number
                '        Select Case intLoop
                '            Case 0
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                '            Case 1
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                '            Case 2
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                '            Case 3
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                '            Case 4
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
                '            Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 5, "udcCharacterTwo")
                '            Case 9, 14, 19, 24
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 25, "udcCharacterTwo")
                '            Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 10, "udcCharacterTwo")
                '            Case 29, 34, 39, 44
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 30, "udcCharacterTwo")
                '            Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 15, "udcCharacterTwo")
                '            Case 48, 49, 53, 54, 58, 59, 63, 64
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 35, "udcCharacterTwo")
                '            Case 65 To 74
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 5, "udcCharacterTwo")
                '            Case 75 To 84
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 10, "udcCharacterTwo")
                '            Case 85 To 94
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 15, "udcCharacterTwo")
                '            Case 95 To 104
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 20, "udcCharacterTwo")
                '            Case Else
                '                gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 100, 350, intZombieSpeedTwo + 45, "udcCharacterTwo")
                '        End Select
                '    Next
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
            athrLoading(0) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame10))
            athrLoading(0).Start()

            'Load 20%
            athrLoading(1) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame20))
            athrLoading(1).Start()

            'Load 30%
            athrLoading(2) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame30))
            athrLoading(2).Start()

            'Load 40%
            athrLoading(3) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame40))
            athrLoading(3).Start()

            'Load 50%
            athrLoading(4) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame50))
            athrLoading(4).Start()

            'Load 60%
            athrLoading(5) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame60))
            athrLoading(5).Start()

            'Load 70%
            athrLoading(6) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame70))
            athrLoading(6).Start()

            'Load 80%
            athrLoading(7) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame80))
            athrLoading(7).Start()

            'Load 90%
            athrLoading(8) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame90))
            athrLoading(8).Start()

            'Load 100%
            athrLoading(9) = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame100))
            athrLoading(9).Start()

            'Set
            sblnFirstTimeLoadingGameObjects = False

        End If

    End Sub

    Private Sub LoadingGame10()

        'Load sounds
        For intLoop As Integer = 0 To audcGameBackgroundSound.GetUpperBound(0)
            audcGameBackgroundSound(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GameBackground" & CStr(intLoop + 1) & ".mp3", 1)
        Next
        udcScreamSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Scream.mp3", 1)
        udcGunShotSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3", 10)
        For intLoop As Integer = 0 To audcZombieDeathSound.GetUpperBound(0)
            audcZombieDeathSound(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ZombieDeath" & CStr(intLoop + 1) & ".mp3", 10)
        Next
        udcReloadingSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reloading.mp3", 2) 'Incase multiplayer
        udcStepSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Step.mp3", 6)
        udcWaterFootStepLeftSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\WaterFootStepLeft.mp3", 3)
        udcWaterFootStepRightSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\WaterFootStepRight.mp3", 3)
        udcGravelFootStepLeftSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GravelFootStepLeft.mp3", 3)
        udcGravelFootStepRightSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GravelFootStepRight.mp3", 3)
        udcOpeningMetalDoorSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\OpeningMetalDoor.mp3", 1) 'Happens only once during gameplay
        udcLightZapSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\LightZap.mp3", 5)
        udcZombieGrowlSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ZombieGrowl.mp3", 1)
        udcRotatingBladeSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\RotatingBlade.mp3", 1)

        'Set
        ablnLoadingGameFinished(0) = True

    End Sub

    Private Sub LoadingGame20()

        'Words
        AddWordsToArray()

        'Word bar
        btmWordBar = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\WordBar.png"))

        'Character
        LoadPieceOfBitmaps(gbtmCharacterStand, "Character\Generic", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShoot, "Character\Generic", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReload, "Character\Generic", "Reload")
        LoadPieceOfBitmaps(gbtmCharacterRunning, "Character\Generic", "Running")

        'Set
        ablnLoadingGameFinished(1) = True

    End Sub

    Private Sub LoadingGame30()

        'Character red
        LoadPieceOfBitmaps(gbtmCharacterStandRed, "Character\Red", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShootRed, "Character\Red", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReloadRed, "Character\Red", "Reload")

        'Set
        ablnLoadingGameFinished(2) = True

    End Sub

    Private Sub LoadingGame40()

        'Character blue
        LoadPieceOfBitmaps(gbtmCharacterStandBlue, "Character\Blue", "Standing")
        LoadPieceOfBitmaps(gbtmCharacterShootBlue, "Character\Blue", "Shoot Once")
        LoadPieceOfBitmaps(gbtmCharacterReloadBlue, "Character\Blue", "Reload")

        'Set
        ablnLoadingGameFinished(3) = True

    End Sub

    Private Sub LoadingGame50()

        'Zombies
        LoadPieceOfBitmaps(gbtmZombieWalk, "Zombies\Generic", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeath1, "Zombies\Generic", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeath2, "Zombies\Generic", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePin, "Zombies\Generic", "Pinning")

        'Set
        ablnLoadingGameFinished(4) = True

    End Sub

    Private Sub LoadingGame60()

        'Zombies red
        LoadPieceOfBitmaps(gbtmZombieWalkRed, "Zombies\Red", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeathRed1, "Zombies\Red", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeathRed2, "Zombies\Red", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePinRed, "Zombies\Red", "Pinning")

        'Set
        ablnLoadingGameFinished(5) = True

    End Sub

    Private Sub LoadingGame70()

        'Zombies blue
        LoadPieceOfBitmaps(gbtmZombieWalkBlue, "Zombies\Blue", "Movement")
        LoadPieceOfBitmaps(gbtmZombieDeathBlue1, "Zombies\Blue", "Death 1")
        LoadPieceOfBitmaps(gbtmZombieDeathBlue2, "Zombies\Blue", "Death 2")
        LoadPieceOfBitmaps(gbtmZombiePinBlue, "Zombies\Blue", "Pinning")

        'Set
        ablnLoadingGameFinished(6) = True

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
        ablnLoadingGameFinished(7) = True

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
        ablnLoadingGameFinished(8) = True

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
        ablnLoadingGameFinished(9) = True

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

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'Check which screen is currently displayed
        Select Case intCanvasMode

            Case 0 'Menu
                MenuMouseClickScreen(pntMouse)

            Case 1 'Options
                OptionsMouseClickScreen(pntMouse)

            Case 2 'Loading game
                LoadingMouseClickScreen(pntMouse)

            Case 3 'Game
                GameMouseClickScreen(pntMouse)

            Case 4 'Highscores
                HighscoresMouseClickScreen(pntMouse)

            Case 5 'Credits
                CreditsMouseClickScreen(pntMouse)

            Case 6 'Versus
                VersusMouseClickScreen(pntMouse)

            Case 7 'Loading versus
                LoadingVersusMouseClickScreen(pntMouse)

            Case 8 'Versus game
                VersusGameMouseClickScreen(pntMouse)

            Case 9 'Story
                StoryMouseClickScreen(pntMouse)

            Case 11 'Path 1 choice
                Path1ChoiceMouseClickScreen(pntMouse)

            Case 12 'Path 2 choice
                Path2ChoiceMouseClickScreen(pntMouse)

        End Select

    End Sub

    Private Sub MenuMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 212, 69, pntStartText) 'Start button was clicked
                'Change
                ShowNextScreenAndExitMenu(2, 0)
                'Start paragraphing
                thrParagraph = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Paragraphing))
                thrParagraph.Start()
                'Start loading game
                LoadingGameWithMultipleThreads()
                'Setup the loading game thread
                thrLoadingGame = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame))
                thrLoadingGame.Start()

            Case blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) 'Highscores button was clicked
                'Change
                ShowNextScreenAndExitMenu(4, 0)

            Case blnMouseInRegion(pntMouse, 218, 87, pntStoryText) 'Story button was clicked
                'Change
                ShowNextScreenAndExitMenu(9, 0)
                'Set story fade in
                thrStory = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf StoryTelling))
                thrStory.Start()

            Case blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) 'Options button was clicked
                'Change
                ShowNextScreenAndExitMenu(1, 0)
                'Options sound
                audcAmbianceSound(1).PlaySound(gintSoundVolume, True)

            Case blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) 'Credits button was clicked
                'Change
                ShowNextScreenAndExitMenu(5, 0)
                'Set credits fade in
                thrCredits = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf CreditsFadeIn))
                thrCredits.Start()

            Case blnMouseInRegion(pntMouse, 256, 74, pntVersusText) 'Versus button was clicked
                'Set
                strIPAddress = strGetLocalIPAddress()
                'Set
                blnFirstTimeNicknameTyping = True
                blnFirstTimeIPAddressTyping = True
                'Change
                ShowNextScreenAndExitMenu(6, 0)

        End Select

    End Sub

    Private Sub ShowNextScreenAndExitMenu(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer)

        'Wait until
        While blnBackFromGame
            System.Threading.Thread.Sleep(1)
        End While

        'Stop fog thread
        AbortThread(thrFog)

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(intCanvasModeToSet, intCanvasShowToSet)

        'Stop sound
        audcAmbianceSound(0).StopSound()

        'Reset fog
        ResetFog()

    End Sub

    Private Sub OptionsMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Menu sound
                audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
                'Restart fog
                RestartFog()
                'Stop options sound
                audcAmbianceSound(1).StopSound()

            Case blnMouseInRegion(pntMouse, 153, 32, pnt800x600Text) 'Resolution change 800x600
                'Resize screen
                ResizeScreen(0, 800, 600)

            Case blnMouseInRegion(pntMouse, 163, 32, pnt1024x768Text) 'Resolution change 1024x768
                'Resize screen
                ResizeScreen(1, 1024, 768)

            Case blnMouseInRegion(pntMouse, 163, 32, pnt1280x800Text) 'Resolution change 1280x800
                'Resize screen
                ResizeScreen(2, 1280, 800)

            Case blnMouseInRegion(pntMouse, 170, 32, pnt1280x1024Text) 'Resolution change 1280x1024
                'Resize screen
                ResizeScreen(3, 1280, 1024)

            Case blnMouseInRegion(pntMouse, 170, 32, pnt1440x900Text) 'Resolution change 1440x900
                'Resize screen
                ResizeScreen(4, 1440, 900)

            Case blnMouseInRegion(pntMouse, 175, 38, pntFullscreenText) 'Resolution change full screen
                'Resize screen
                ResizeScreen(5)

        End Select

    End Sub

    Private Sub LoadingMouseClickScreen(pntMouse As Point)

        'Loading start text bar was clicked and game finished loading
        If intCanvasShow = 1 And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
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
            audcGameBackgroundSound(0).PlaySound(CInt(gintSoundVolume / 4), True)
        End If

    End Sub

    Private Sub GameMouseClickScreen(pntMouse As Point)

        'Back was clicked
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
            'Set
            blnPlayPressedSoundNow = True
            'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            blnBackFromGame = True
        End If

    End Sub

    Private Sub HighscoresMouseClickScreen(pntMouse As Point)

        'Back was clicked
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
            'Restart fog
            RestartFog()
        End If

    End Sub

    Private Sub CreditsMouseClickScreen(pntMouse As Point)

        'Back was clicked
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
            'Restart fog
            RestartFog()
            'Abort thread
            AbortThread(thrCredits)
            'Reset
            btmJohnGonzales = Nothing
            btmZacharyStafford = Nothing
            btmCoryLewis = Nothing
        End If

    End Sub

    Private Sub VersusMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Check nickname
                DefaultNickname()
                'Set
                blnPlayPressedSoundNow = True
                'Go back to menu
                GoBackToMenuConnectionGone()

            Case Else
                'Check which versus to show
                Select Case intCanvasVersusShow

                    Case 0 'Check for hoster or joiner
                        'Hoster or Joiner
                        Select Case True

                            Case blnMouseInRegion(pntMouse, 545, 176, pntVersusHost) 'Host was clicked
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

                            Case blnMouseInRegion(pntMouse, 441, 170, pntVersusJoin) 'Join was clicked
                                'Set
                                blnHost = False
                                'Set
                                strIPAddressConnect = strGetLocalIPAddress()
                                'Check nickname
                                DefaultNickname()
                                'Set
                                intCanvasVersusShow = 2

                        End Select

                    Case 2 'Connect button
                        'Connect was clicked
                        If blnMouseInRegion(pntMouse, 605, 124, pntVersusConnect) Then
                            'Set
                            intCanvasVersusShow = 3
                            'Start thread connecting
                            thrConnecting = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Connecting))
                            thrConnecting.Start()
                        End If

                End Select

        End Select

    End Sub

    Private Sub LoadingVersusMouseClickScreen(pntMouse As Point)

        'Check if hosting
        If blnHost Then
            'Start was clicked
            If Not blnWaiting And intCanvasShow = 1 And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                'Send playing
                gSendData("2|" & strNickname)
                'Start game locally
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(8, 0, False)
                'Started versus game, start objects
                StartVersusGameObjects()
            End If
        Else
            'Start was clicked
            If Not blnWaiting And intCanvasShow = 1 And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                'Send ready to play as joiner
                gSendData("2|" & strNickname)
                blnWaiting = True
            End If
        End If

    End Sub

    Private Sub VersusGameMouseClickScreen(pntMouse As Point)

        'Back was pressed
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
            'Set
            blnPlayPressedSoundNow = True
            'Quit versus multiplayer
            GoBackToMenuConnectionGone()
        End If

    End Sub

    Private Sub StoryMouseClickScreen(pntMouse As Point)

        'Back has been moused over
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Then
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
            'Restart fog
            RestartFog()
            'Abort thread
            AbortThread(thrStory)
            'Reset
            btmStoryParagraph = Nothing
            'Stop sounds
            For intLoop As Integer = 0 To audcStoryParagraphSound.GetUpperBound(0)
                audcStoryParagraphSound(intLoop).StopSound()
            Next
        End If

    End Sub

    Private Sub Path1ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Set
                blnPlayPressedSoundNow = True
                'Set
                blnBackFromGame = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                audcAmbianceSound(0).PlaySound(gintSoundVolume, True)

            Case blnMouseInRegion(pntMouse, 389, 329, New Point(230, 427)) 'Path 1 choice clicked
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground2.jpg")), 2)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)

            Case blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) 'Path 2 choice clicked
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground3.jpg")), 3)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)

        End Select

    End Sub

    Private Sub Path2ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Set
                blnPlayPressedSoundNow = True
                'Set
                blnBackFromGame = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                audcAmbianceSound(0).PlaySound(gintSoundVolume, True)

            Case blnMouseInRegion(pntMouse, 296, 304, New Point(643, 306)) 'Path 1 choice
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground5.jpg")), 5, True)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)

            Case blnMouseInRegion(pntMouse, 297, 318, New Point(1138, 384)) 'Path 2 choice
                'Setup next level
                NextLevel(New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground4.jpg")), 4)
                'Change level, reuse the mechanics
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)

        End Select

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
        audcAmbianceSound(0).PlaySound(gintSoundVolume, True)

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
            btmSoundPercent = abtmSound(0)
        ElseIf pntSlider.X = 657 Then
            gintSoundVolume = 1000
            btmSoundPercent = abtmSound(100)
        Else
            gintSoundVolume = CInt((pntSlider.X - 58) * (1000 / 600)) '58 to 657 = 600 pixel bounds, inclusive of 58 as a number.
            btmSoundPercent = abtmSound(CInt(gintSoundVolume / 10))
        End If

        'After setting volume, change option sound
        audcAmbianceSound(1).ChangeVolumeWhileSoundIsPlaying()

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

        'Notes: This sub has to be synced with the rendering thread. The problem is everything here must be captured and checked by the render thread.

        'Keys pressed
        'Debug.Print(CStr(Asc(e.KeyChar)))

        'Check if playing game
        Select Case intCanvasMode
            Case 3
                SinglePlayerKeyPress(e.KeyChar)
            Case 6
                VersusKeyPress(e.KeyChar)
            Case 8
                MultiplayerKeyPress(e.KeyChar)
        End Select

    End Sub

    Private Sub SinglePlayerKeyPress(chrKeyPressed As Char)

        'Exit if character doesn't exist
        If udcCharacter IsNot Nothing Then

            'Check the status of the character
            Select Case udcCharacter.StatusModeProcessing

                Case clsCharacter.eintStatusMode.Stand

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 32 ' = spacebar
                            'Check to make sure not maxed on bullets
                            If udcCharacter.BulletsUsed <> 0 Then
                                'Character start to reload
                                udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload
                            End If

                        Case 39 ' = '
                            'Running forward
                            udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Run

                        Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                            'Add to the buffer
                            strKeyPressBuffer &= "." & chrKeyPressed.ToString

                    End Select

                Case clsCharacter.eintStatusMode.Reload

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 39 ' = '
                            'Prepare to run forward after reload
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run

                    End Select

                Case clsCharacter.eintStatusMode.Shoot

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 32 ' = spacebar
                            'Character about to reload
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload

                        Case 39 ' = '
                            'Running forward
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run

                        Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                            'Add to the buffer
                            strKeyPressBuffer &= "." & chrKeyPressed.ToString

                    End Select

                Case clsCharacter.eintStatusMode.Run

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 32 ' = spacebar
                            'Character start to reload
                            udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload

                        Case 59 ' = ;
                            'Stop immediately
                            udcCharacter.StopCharacterFromRunning = True

                        Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                            'Add to the buffer
                            strKeyPressBuffer &= "." & chrKeyPressed.ToString

                    End Select

            End Select

        End If

    End Sub

    Private Sub VersusKeyPress(chrKeyPressed As Char)

        'Check for nickname press or ip address
        Select Case intCanvasVersusShow

            Case 0 'Nickname
                'Check length
                If Len(strNickname) < 9 Then
                    'Check the keypress
                    Select Case Asc(chrKeyPressed)
                        Case 8 'Backspace
                            'Check if first time key pressing
                            KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickname)
                        Case 32, 65 To 90, 97 To 122 '32 = spacebar, 65 to 90 is upper case A-Z, 97 to 122 is lower case a-z
                            'Check if first time key pressing
                            If blnFirstTimeNicknameTyping Then
                                'Set
                                strNickname = chrKeyPressed.ToString
                                'Set
                                blnFirstTimeNicknameTyping = False
                            Else
                                'Set
                                strNickname &= chrKeyPressed.ToString
                            End If
                    End Select
                Else
                    'Allow backspace
                    If Asc(chrKeyPressed) = 8 Then
                        'Check if first time key pressing
                        KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickname)
                    End If
                End If

            Case 2 'Ip address
                'Check if pressing for first time
                If blnFirstTimeIPAddressTyping Then
                    'Check the keypress
                    Select Case Asc(chrKeyPressed)
                        Case 8, 48 To 57 '8 is backspace, 48 to 57 is 0 to 9 in numbers
                            'Set
                            strIPAddressConnect = chrKeyPressed.ToString
                            'Set
                            blnFirstTimeIPAddressTyping = False
                    End Select
                Else
                    'Check length
                    If Len(strIPAddressConnect) < 15 Then 'XXX.XXX.XXX.XXX
                        'Check the keypress
                        Select Case Asc(chrKeyPressed)
                            Case 8 'Backspace
                                'Check if first time key pressing
                                KeyPressBackspacing(blnFirstTimeIPAddressTyping, strIPAddressConnect)
                            Case 46, 48 To 57 '46 = period, 48 to 57 is 0 to 9 in numbers
                                'Check if first time key pressing
                                If blnFirstTimeIPAddressTyping Then
                                    'Set
                                    strIPAddressConnect = chrKeyPressed.ToString
                                    'Set
                                    blnFirstTimeIPAddressTyping = False
                                Else
                                    'Set
                                    strIPAddressConnect &= chrKeyPressed.ToString
                                End If
                        End Select
                    Else
                        'Allow backspace
                        If Asc(chrKeyPressed) = 8 Then
                            'Check if first time key pressing
                            KeyPressBackspacing(blnFirstTimeIPAddressTyping, strIPAddressConnect)
                        End If
                    End If
                End If

        End Select

    End Sub

    Private Sub MultiplayerKeyPress(chrKeyPressed As Char)

        'Check if hoster or joiner
        If blnHost Then 'Hoster
            'Exit if character doesn't exist
            If udcCharacterOne IsNot Nothing Then
                'Check the status of the character
                MultiplayerKeyPressStatus(udcCharacterOne, chrKeyPressed)
            End If
        Else 'Joiner
            'Exit if character doesn't exist
            If udcCharacterTwo IsNot Nothing Then
                'Check the status of the character
                MultiplayerKeyPressStatus(udcCharacterTwo, chrKeyPressed)
            End If
        End If

    End Sub

    Private Sub MultiplayerKeyPressStatus(udcCharacterType As clsCharacter, chrKeyPressed As Char)

        'Check the status of the character
        Select Case udcCharacterType.StatusModeProcessing

            Case clsCharacter.eintStatusMode.Stand

                'Check key press
                Select Case Asc(chrKeyPressed)

                    Case 32 ' = spacebar
                        'Check to make sure not maxed on bullets
                        If udcCharacterType.BulletsUsed <> 0 Then
                            'Character start to reload
                            udcCharacterType.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload
                        End If

                    Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                        'Add to the buffer
                        strKeyPressBuffer &= "." & chrKeyPressed.ToString

                End Select

            Case clsCharacter.eintStatusMode.Shoot

                'Check key press
                Select Case Asc(chrKeyPressed)

                    Case 32 ' = spacebar
                        'Character about to reload
                        udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload

                    Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                        'Add to the buffer
                        strKeyPressBuffer &= "." & chrKeyPressed.ToString

                End Select

        End Select

    End Sub

    Private Sub KeyPressBackspacing(ByRef blnByRefFirstTimeToCheck As Boolean, ByRef strByRefToChange As String)

        'Check if first time key pressing
        If blnByRefFirstTimeToCheck Then
            'Set
            strByRefToChange = ""
            'Set
            blnByRefFirstTimeToCheck = False
        Else
            'Check length
            If Len(strByRefToChange) <> 0 Then
                'Set to subtract one from length
                strByRefToChange = strByRefToChange.Substring(0, Len(strByRefToChange) - 1)
            End If
        End If

    End Sub

    Private Function strGetLocalIPAddress() As String

        'Declare
        Dim strTemp As String = ""

        'Loop
        For intLoop As Integer = 0 To System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList.GetUpperBound(0)
            'Look for the matching correct address
            If InStr(System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList(intLoop).ToString(), ":") = 0 Then
                'Set the correct address
                strTemp = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName).AddressList(intLoop).ToString()
                'Exit
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
                'If server is waiting
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
                'Set
                tcpcClient = New System.Net.Sockets.TcpClient(strIPAddressConnect, 10101)
                'Set
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
        blnPreventKeyPressEvent = False

        'Set
        strKeyPressBuffer = ""

        'Set
        blnGameOverFirstTime = True

        'Set
        blnBeatLevel = False

        'Check loaded game objects
        For intLoop As Integer = 0 To ablnLoadingGameFinished.GetUpperBound(0) - 1
            'Create a pause waiting
            While Not ablnLoadingGameFinished(intLoop)
            End While
            'Set
            btmLoadingBar = abtmLoadingBarPicture(intLoop + 1)
        Next

        'Grab a random word
        LoadARandomWord()

        'Load in a special way
        If blnHost Then
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound,
                                               udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound,
                                               udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound, True) 'Join
        Else
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound,
                                               udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound, True) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound,
                                               udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound) 'Join
        End If

        'Zombies
        LoadZombies("Level 1 Multiplayer")

        'Set if hosting
        If blnHost And Not blnReadyEarly Then
            blnWaiting = True
        End If

        'Check loaded game objects
        While Not ablnLoadingGameFinished(ablnLoadingGameFinished.GetUpperBound(0))
        End While

        'Set
        btmLoadingBar = abtmLoadingBarPicture(ablnLoadingGameFinished.GetUpperBound(0) + 1)

        'Set
        intCanvasShow = 1 'Means completely loaded

    End Sub

    Private Sub GamesMismatched(strGameVersionFromConnecter As String)

        'Set
        strGameVersionFromConnection = strGameVersionFromConnecter

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(10, 0, False)

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
                    ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(8, 0, False)
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
                    blnPreventKeyPressEvent = True
                    'Set
                    blnGameOverFirstTime = False
                    'Stop reloading sound
                    udcReloadingSound.StopSound()
                    'Play
                    udcScreamSound.PlaySound(gintSoundVolume)
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