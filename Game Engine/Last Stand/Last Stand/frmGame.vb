'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const GAME_VERSION As String = "1.49"
    Private Const ORIGINAL_SCREEN_WIDTH As Integer = 1680
    Private Const ORIGINAL_SCREEN_HEIGHT As Integer = 1050
    Private Const WINDOW_MESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const CONTROL_MOVE As Integer = 61456
    Private Const WINDOW_MESSAGE_CLICK_BUTTON_DOWN As Integer = 161
    Private Const WINDOW_CAPTION As Integer = 2
    Private Const WINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3
    Private Const WIDTH_SUBTRACTION As Integer = 16 'Probably the edges of the window
    Private Const HEIGHT_SUBTRACTION As Integer = 38 'Probably the edges of the window
    Private Const FOG_WIDTH As Integer = 6720 'Image width
    Private Const FOG_SPEED As Integer = 40
    Private Const FOG_BACK_DISTANCE_Y As Integer = 300 'Extra added distance to shift fog down
    Private Const FOG_FRONT_DISTANCE_Y As Integer = 550 'Extra added distance to shift fog down
    Private Const LOADING_TRANSPARENCY_DELAY As Integer = 500
    Private Const BLACK_SCREEN_DEATH_DELAY As Integer = 750
    Private Const BLACK_SCREEN_BEAT_LEVEL_DELAY As Integer = 350
    Private Const MOUSE_OVER_SOUND_DELAY As Double = 35 '35 milliseconds
    Private Const CREDITS_TRANSPARENCY_DELAY As Integer = 500
    Private Const NUMBER_OF_ZOMBIES_CREATED As Integer = 150
    Private Const NUMBER_OF_ZOMBIES_AT_ONE_TIME As Integer = 5
    Private Const ZOMBIE_PINNING_X_DISTANCE As Integer = 200
    Private Const JOINER_ADDED_X_DISTANCE As Integer = 100

    'Declare beginning necessary engine needs
    Private intScreenWidth As Integer = 800
    Private intScreenHeight As Integer = 600
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False
    Private rectFullScreen As Rectangle
    Private btmCanvas As New Bitmap(ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT, Imaging.PixelFormat.Format32bppPArgb) 'Screen size here
    Private pntTopLeft As New Point(0, 0)
    Private intCanvasMode As Integer = 0 'Default menu screen
    Private intCanvasShow As Integer = 0 'Default, no animation
    Private strDirectory As String = AppDomain.CurrentDomain.BaseDirectory & "\"
    Private blnScreenChanged As Boolean = False

    'Menu necessary needs
    Private btmMenuBackgroundFile As Bitmap
    Private btmMenuBackgroundMemory As Bitmap
    Private btmFogBackFile As Bitmap
    Private btmFogBackMemory As Bitmap
    Private pntFogBack As New Point(FOG_WIDTH - ORIGINAL_SCREEN_WIDTH, 0) 'Length of picture minus length of screen to be shown
    Private intFogBackX As Integer = 0
    Private btmFogBackCloneScreenShown As Bitmap
    Private btmFogFrontFile As Bitmap
    Private btmFogFrontMemory As Bitmap
    Private pntFogFront As New Point(FOG_WIDTH - ORIGINAL_SCREEN_WIDTH, 0) 'Length of picture minus length of screen to be shown
    Private intFogFrontX As Integer = 0
    Private btmFogFrontCloneScreenShown As Bitmap
    Private btmArcherFile As Bitmap
    Private btmArcherMemory As Bitmap
    Private pntArcher As New Point(117, 0)
    Private btmLastStandTextFile As Bitmap
    Private btmLastStandTextMemory As Bitmap
    Private pntLastStandText As New Point(147, 833)
    Private btmStartTextFile As Bitmap
    Private btmStartTextMemory As Bitmap
    Private pntStartText As New Point(1081, 31)
    Private btmStartHoverTextFile As Bitmap
    Private btmStartHoverTextMemory As Bitmap
    Private pntStartHoverText As New Point(1059, 25)
    Private btmHighscoresTextFile As Bitmap
    Private btmHighscoresTextMemory As Bitmap
    Private pntHighscoresText As New Point(1198, 141)
    Private btmHighscoresHoverTextFile As Bitmap
    Private btmHighscoresHoverTextMemory As Bitmap
    Private pntHighscoresHoverText As New Point(1157, 131)
    Private btmStoryTextFile As Bitmap
    Private btmStoryTextMemory As Bitmap
    Private pntStoryText As New Point(1246, 283)
    Private btmStoryHoverTextFile As Bitmap
    Private btmStoryHoverTextMemory As Bitmap
    Private pntStoryHoverText As New Point(1222, 275)
    Private btmOptionsTextFile As Bitmap
    Private btmOptionsTextMemory As Bitmap
    Private pntOptionsText As New Point(1207, 410)
    Private btmOptionsHoverTextFile As Bitmap
    Private btmOptionsHoverTextMemory As Bitmap
    Private pntOptionsHoverText As New Point(1175, 400)
    Private btmCreditsTextFile As Bitmap
    Private btmCreditsTextMemory As Bitmap
    Private pntCreditsText As New Point(1352, 536)
    Private btmCreditsHoverTextFile As Bitmap
    Private btmCreditsHoverTextMemory As Bitmap
    Private pntCreditsHoverText As New Point(1323, 527)
    Private btmVersusTextFile As Bitmap
    Private btmVersusTextMemory As Bitmap
    Private pntVersusText As New Point(284, 71)
    Private btmVersusHoverTextFile As Bitmap
    Private btmVersusHoverTextMemory As Bitmap
    Private pntVersusHoverText As New Point(256, 63)

    'Options mouse over
    Private tmrOptionsMouseOver As New System.Timers.Timer
    Private strOptionsButtonSpot As String = ""

    'General, common uses
    Private btmBackTextFile As Bitmap
    Private btmBackTextMemory As Bitmap
    Private pntBackText As New Point(1439, 46)
    Private btmBackHoverTextFile As Bitmap
    Private btmBackHoverTextMemory As Bitmap
    Private pntBackHoverText As New Point(1418, 35)

    'Options screen
    Private btmOptionsBackgroundFile As Bitmap
    Private btmOptionsBackgroundMemory As Bitmap
    Private btmResolutionTextFile As Bitmap
    Private btmResolutionTextMemory As Bitmap
    Private pntResolutionText As New Point(40, 41)
    Private btm800x600TextFile As Bitmap
    Private btm800x600TextMemory As Bitmap
    Private btmNot800x600TextFile As Bitmap
    Private btmNot800x600TextMemory As Bitmap
    Private pnt800x600Text As New Point(85, 142)
    Private btm1024x768TextFile As Bitmap
    Private btm1024x768TextMemory As Bitmap
    Private btmNot1024x768TextFile As Bitmap
    Private btmNot1024x768TextMemory As Bitmap
    Private pnt1024x768Text As New Point(85, 192)
    Private btm1280x800TextFile As Bitmap
    Private btm1280x800TextMemory As Bitmap
    Private btmNot1280x800TextFile As Bitmap
    Private btmNot1280x800TextMemory As Bitmap
    Private pnt1280x800Text As New Point(85, 242)
    Private btm1280x1024TextFile As Bitmap
    Private btm1280x1024TextMemory As Bitmap
    Private btmNot1280x1024TextFile As Bitmap
    Private btmNot1280x1024TextMemory As Bitmap
    Private pnt1280x1024Text As New Point(85, 293)
    Private btm1440x900TextFile As Bitmap
    Private btm1440x900TextMemory As Bitmap
    Private btmNot1440x900TextFile As Bitmap
    Private btmNot1440x900TextMemory As Bitmap
    Private pnt1440x900Text As New Point(85, 342)
    Private btmFullScreenTextFile As Bitmap
    Private btmFullScreenTextMemory As Bitmap
    Private btmNotFullScreenTextFile As Bitmap
    Private btmNotFullScreenTextMemory As Bitmap
    Private pntFullscreenText As New Point(85, 391)
    Private btmSoundTextFile As Bitmap
    Private btmSoundTextMemory As Bitmap
    Private pntSoundText As New Point(40, 447)
    Private intResolutionMode As Integer = 0 'Default 800x600
    Private btmSoundBarFile As Bitmap
    Private btmSoundBarMemory As Bitmap
    Private pntSoundBar As New Point(84, 547)
    Private btmSliderFile As Bitmap
    Private btmSliderMemory As Bitmap
    Private pntSlider As New Point(658, 533) '100% mark
    Private blnSliderWithMouseDown As Boolean = False
    Private btmSoundPercent As New Bitmap(87, 37)
    Private abtmSoundFiles(100) As Bitmap '0 to 100
    Private abtmSoundMemories(100) As Bitmap '0 to 100
    Private pntSoundPercent As New Point(718, 553)

    'Loading screen
    Private btmLoadingBackgroundFile As Bitmap
    Private btmLoadingBackgroundMemory As Bitmap
    Private abtmLoadingBarPictureFiles(10) As Bitmap '0 To 10 = 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
    Private abtmLoadingBarPictureMemories(10) As Bitmap
    Private btmLoadingBar As New Bitmap(1613, 134)
    Private pntLoadingBar As New Point(33, 883)
    Private btmLoadingTextFile As Bitmap
    Private btmLoadingTextMemory As Bitmap
    Private pntLoadingText As New Point(594, 899)
    Private btmLoadingStartTextFile As Bitmap
    Private btmLoadingStartTextMemory As Bitmap
    Private pntLoadingStartText As New Point(673, 909)
    Private abtmLoadingParagraphFiles(3) As Bitmap
    Private abtmLoadingParagraphMemories(3) As Bitmap
    Private btmLoadingParagraph As New Bitmap(1424, 472)
    Private pntLoadingParagraph As New Point(123, 261)
    Private tmrParagraph As New System.Timers.Timer 'Used for paragraph delay
    Private strTypeOfParagraphWait As String = ""
    Private intParagraphWaitMode As Integer = 0
    Private thrLoadBeginningGameMaterial As System.Threading.Thread
    Private thrLoadingGame As System.Threading.Thread
    Private intMemoryLoadPosition As Integer = 0
    Private blnFinishedLoading As Boolean = False

    'Game screen
    Private btmGameBackgroundMemory As Bitmap 'Point for this is created in the public module
    Private blnGameBackgroundMemoryCopied As Boolean
    Private blnCanLoadLevelWhileRendering As Boolean = False
    Private abtmGameBackgroundFiles(4) As Bitmap
    Private abtmGameBackgroundMemories(4) As Bitmap
    Private ablnGameBackgroundMemoriesCopied(4) As Boolean
    Private btmGameBackgroundCloneScreenShown As Bitmap
    Private btmDeathScreen As Bitmap
    Private btmDeathOverlayFile As Bitmap
    Private btmDeathOverlayMemory As Bitmap
    Private blnDeathOverlayMemoryCopied As Boolean
    Private abtmBlackScreenFiles(2) As Bitmap
    Private abtmBlackScreenMemories(2) As Bitmap
    Private ablnBlackScreenMemoriesCopied(2) As Boolean
    Private btmBlackScreen As Bitmap
    Private tmrBlackScreen As New System.Timers.Timer 'Used for black screen delay
    Private intBlackScreenWaitMode As Integer = 0
    Private blnBlackScreenFinished As Boolean = False
    Private blnPlayerWasPinned As Boolean = False
    Private intZombieIncreasedPinDistance As Integer = 0
    Private blnRemovedGameObjectsFromMemory As Boolean = False
    Private btmWordBarFile As Bitmap
    Private btmWordBarMemory As Bitmap
    Private blnWordBarMemoryCopied As Boolean = False
    Private pntWordBar As New Point(482, 27)
    Private udcCharacter As clsCharacter
    Private blnBackFromGame As Boolean = False
    Private btmAK47MagazineFile As Bitmap
    Private btmAK47MagazineMemory As Bitmap
    Private blnAK47MagazineMemoryCopied As Boolean
    Private pntAK47Magazine As New Point(59, 877)
    Private btmWinOverlayFile As Bitmap
    Private btmWinOverlayMemory As Bitmap
    Private blnWinOverlayMemoryCopied As Boolean
    Private intLevel As Integer = 1 'Starting
    Private intZombieKills As Integer = 0
    Private intZombieKillsCombined As Integer = 0
    Private intReloadTimes As Integer = 0
    Private blnComparedHighscore As Boolean = False
    Private blnBeatLevel As Boolean = False

    'Key press
    Private strKeyPressBuffer As String = ""
    Private blnPreventKeyPressEvent As Boolean = False

    'Sounds
    Private audcAmbianceSound(1) As clsSound '2 ambiance sounds so far
    Private udcButtonHoverSound As clsSound
    Private udcButtonPressedSound As clsSound
    Private audcStoryParagraphSounds(2) As clsSound '3 paragraph sounds in the story area
    Private audcGameBackgroundSounds(4) As clsSound '5 background sounds so far
    Private udcScreamSound As clsSound
    Private udcGunShotSound As clsSound
    Private audcZombieDeathSounds(1) As clsSound '2 zombie death sounds so far
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

    'Helicopter
    Private udcHelicopter As clsHelicopter

    'Path choices
    Private btmPath As Bitmap
    Private abtmPath1Files(2) As Bitmap
    Private abtmPath1Memories(2) As Bitmap
    Private ablnPath1MemoriesCopied(2) As Boolean
    Private abtmPath2Files(2) As Bitmap
    Private abtmPath2Memories(2) As Bitmap
    Private ablnPath2MemoriesCopied(2) As Boolean
    Private blnLightZap1 As Boolean = False
    Private blnLightZap2 As Boolean = False

    'Stop watch
    Private swhStopWatch As Stopwatch

    'Words
    Private astrWords() As String 'Used to fill with words
    Private intWordIndex As Integer = 0
    Private strTheWord As String = ""
    Private strWord As String = ""

    'Highscores screen
    Private btmHighscoresBackgroundFile As Bitmap
    Private btmHighscoresBackgroundMemory As Bitmap
    Private strHighscores As String = ""
    Private blnHighscoresIsAccess As Boolean = False

    'Credits screen
    Private btmCreditsBackgroundFile As Bitmap
    Private btmCreditsBackgroundMemory As Bitmap
    Private abtmJohnGonzalesFiles(3) As Bitmap
    Private abtmJohnGonzalesMemories(3) As Bitmap
    Private btmJohnGonzales As Bitmap
    Private pntJohnGonzales As New Point(200, 150)
    Private abtmZacharyStaffordFiles(3) As Bitmap
    Private abtmZacharyStaffordMemories(3) As Bitmap
    Private btmZacharyStafford As Bitmap
    Private pntZacharyStafford As New Point(940, 150)
    Private abtmCoryLewisFiles(3) As Bitmap
    Private abtmCoryLewisMemories(3) As Bitmap
    Private btmCoryLewis As Bitmap
    Private pntCoryLewis As New Point(570, 575)
    Private tmrCredits As New System.Timers.Timer 'Used for credits delay
    Private intCreditsWaitMode As Integer = 0

    'Versus screen
    Private strIPAddress As String = ""
    Private strNickName As String = "Player"
    Private strNickNameConnected As String = ""
    Private strIPAddressConnect As String = ""
    Private intCanvasVersusShow As Integer = 0
    Private blnFirstTimeNicknameTyping As Boolean = True 'Defaulted
    Private blnFirstTimeIPAddressTyping As Boolean = True 'Defaulted
    Private btmVersusBackgroundFile As Bitmap
    Private btmVersusBackgroundMemory As Bitmap
    Private btmVersusNickNameTextFile As Bitmap
    Private btmVersusNickNameTextMemory As Bitmap
    Private pntVersusBlackOutline As New Point(100, 150)
    Private btmVersusHostTextFile As Bitmap
    Private btmVersusHostTextMemory As Bitmap
    Private pntVersusHost As New Point(267, 605)
    Private btmVersusOrTextFile As Bitmap
    Private btmVersusOrTextMemory As Bitmap
    Private pntVersusOr As New Point(831, 649)
    Private btmVersusJoinTextFile As Bitmap
    Private btmVersusJoinTextMemory As Bitmap
    Private pntVersusJoin As New Point(951, 601)
    Private btmVersusIPAddressTextFile As Bitmap
    Private btmVersusIPAddressTextMemory As Bitmap
    Private btmVersusConnectTextFile As Bitmap
    Private btmVersusConnectTextMemory As Bitmap
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
    Private blnConnectionCompleted As Boolean = False 'This exists because data sends way too fast
    Private blnCheckedGameMismatch As Boolean = False 'This exists because data sends way too fast
    Private blnWaiting As Boolean = False
    Private blnReadyEarly As Boolean = False

    'Loading versus game variables
    Private blnGameIsVersus As Boolean = False
    Private abtmLoadingParagraphVersusFiles(3) As Bitmap
    Private abtmLoadingParagraphVersusMemories(3) As Bitmap
    Private btmLoadingParagraphVersus As New Bitmap(1424, 472)
    Private btmLoadingWaitingTextFile As Bitmap
    Private btmLoadingWaitingTextMemory As Bitmap
    Private pntLoadingWaitingText As New Point(594, 899)

    'In versus game playing variables
    Private udcCharacterOne As clsCharacter 'Host
    Private udcCharacterTwo As clsCharacter 'Join
    Private btmYouWonFile As Bitmap
    Private btmYouWonMemory As Bitmap
    Private blnYouWonMemoryCopied As Boolean
    Private btmYouLostFile As Bitmap
    Private btmYouLostMemory As Bitmap
    Private blnYouLostMemoryCopied As Boolean
    Private btmVersusWhoWon As Bitmap
    Private intZombieKillsOne As Integer = 0
    Private intZombieKillsTwo As Integer = 0
    Private intZombieKillsWaitingTwo As Integer = 0 'Used as a buffer for joiner sending shots to prepare and kill a zombie
    Private strZombieKillBufferOne As String = ""
    Private strZombieKillBufferTwo As String = ""
    Private intZombieIncreasedPinDistanceOne As Integer = 0
    Private intZombieIncreasedPinDistanceTwo As Integer = 0

    'Story
    Private btmStoryBackgroundFile As Bitmap
    Private btmStoryBackgroundMemory As Bitmap
    Private abtmStoryParagraphFiles(11) As Bitmap
    Private abtmStoryParagraphMemories(11) As Bitmap
    Private btmStoryParagraph As Bitmap
    Private pntStoryParagraph As Point 'Set in the story thread
    Private thrStory As System.Threading.Thread

    'Game version mismatch
    Private btmGameMismatchBackgroundFile As Bitmap
    Private btmGameMismatchBackgroundMemory As Bitmap
    Private strGameVersionFromConnection As String = ""
    Private thrGameMismatch As System.Threading.Thread

    'Character bitmap load files
    Private abtmCharacterStandFiles(1) As Bitmap
    Private ablnCharacterStandMemoriesCopied(1) As Boolean
    Private abtmCharacterShootFiles(1) As Bitmap
    Private ablnCharacterShootMemoriesCopied(1) As Boolean
    Private abtmCharacterReloadFiles(21) As Bitmap
    Private ablnCharacterReloadMemoriesCopied(21) As Boolean
    Private abtmCharacterRunningFiles(16) As Bitmap
    Private ablnCharacterRunningMemoriesCopied(16) As Boolean
    Private abtmCharacterStandRedFiles(1) As Bitmap
    Private ablnCharacterStandRedMemoriesCopied(1) As Boolean
    Private abtmCharacterShootRedFiles(1) As Bitmap
    Private ablnCharacterShootRedMemoriesCopied(1) As Boolean
    Private abtmCharacterReloadRedFiles(21) As Bitmap
    Private ablnCharacterReloadRedMemoriesCopied(21) As Boolean
    Private abtmCharacterStandBlueFiles(1) As Bitmap
    Private ablnCharacterStandBlueMemoriesCopied(1) As Boolean
    Private abtmCharacterShootBlueFiles(1) As Bitmap
    Private ablnCharacterShootBlueMemoriesCopied(1) As Boolean
    Private abtmCharacterReloadBlueFiles(21) As Bitmap
    Private ablnCharacterReloadBlueMemoriesCopied(21) As Boolean

    'Zombies bitmap load files
    Private abtmZombieWalkFiles(3) As Bitmap
    Private ablnZombieWalkMemoriesCopied(3) As Boolean
    Private abtmZombieDeath1Files(5) As Bitmap
    Private ablnZombieDeath1MemoriesCopied(5) As Boolean
    Private abtmZombieDeath2Files(5) As Bitmap
    Private ablnZombieDeath2MemoriesCopied(5) As Boolean
    Private abtmZombiePinFiles(1) As Bitmap
    Private ablnZombiePinMemoriesCopied(1) As Boolean
    Private abtmZombieWalkRedFiles(3) As Bitmap
    Private ablnZombieWalkRedMemoriesCopied(3) As Boolean
    Private abtmZombieDeathRed1Files(5) As Bitmap
    Private ablnZombieDeathRed1MemoriesCopied(5) As Boolean
    Private abtmZombieDeathRed2Files(5) As Bitmap
    Private ablnZombieDeathRed2MemoriesCopied(5) As Boolean
    Private abtmZombiePinRedFiles(1) As Bitmap
    Private ablnZombiePinRedMemoriesCopied(1) As Boolean
    Private abtmZombieWalkBlueFiles(3) As Bitmap
    Private ablnZombieWalkBlueMemoriesCopied(3) As Boolean
    Private abtmZombieDeathBlue1Files(5) As Bitmap
    Private ablnZombieDeathBlue1MemoriesCopied(5) As Boolean
    Private abtmZombieDeathBlue2Files(5) As Bitmap
    Private ablnZombieDeathBlue2MemoriesCopied(5) As Boolean
    Private abtmZombiePinBlueFiles(1) As Bitmap
    Private ablnZombiePinBlueMemoriesCopied(1) As Boolean

    'Helicopter bitmap load files
    Private abtmHelicopterFiles(4) As Bitmap
    Private ablnHelicopterMemoriesCopied(4) As Boolean

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

        'Set images into memory
        SetImagesIntoMemory()

        'Set 100%
        btmSoundPercent = abtmSoundMemories(100)

        'Setup mouse over timer
        SetupMouseOverTimer()

        'Setup paragraph timer
        SetupParagraphTimer()

        'Setup credits timer
        SetupCreditsTimer()

        'Setup black screen timer
        SetupBlackScreenTimer()

    End Sub

    Private Sub SetImagesIntoMemory()

        'Menu background
        LoadPictureFileAndCopyBitmapIntoMemory(btmMenuBackgroundFile, btmMenuBackgroundMemory, "Images\Menu\MenuBackground.jpg")
        LoadPictureFileAndCopyBitmapIntoMemory(btmFogBackFile, btmFogBackMemory, "Images\Menu\FogBack.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmFogFrontFile, btmFogFrontMemory, "Images\Menu\FogFront.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmArcherFile, btmArcherMemory, "Images\Menu\Archer.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmLastStandTextFile, btmLastStandTextMemory, "Images\Menu\LastStand.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmStartTextFile, btmStartTextMemory, "Images\Menu\Start.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmStartHoverTextFile, btmStartHoverTextMemory, "Images\Menu\StartHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmHighscoresTextFile, btmHighscoresTextMemory, "Images\Menu\Highscores.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmHighscoresHoverTextFile, btmHighscoresHoverTextMemory, "Images\Menu\HighscoresHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmStoryTextFile, btmStoryTextMemory, "Images\Menu\Story.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmStoryHoverTextFile, btmStoryHoverTextMemory, "Images\Menu\StoryHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmOptionsTextFile, btmOptionsTextMemory, "Images\Menu\Options.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmOptionsHoverTextFile, btmOptionsHoverTextMemory, "Images\Menu\OptionsHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmCreditsTextFile, btmCreditsTextMemory, "Images\Menu\Credits.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmCreditsHoverTextFile, btmCreditsHoverTextMemory, "Images\Menu\CreditsHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusTextFile, btmVersusTextMemory, "Images\Menu\Versus.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusHoverTextFile, btmVersusHoverTextMemory, "Images\Menu\VersusHover.png")

        'Back button
        LoadPictureFileAndCopyBitmapIntoMemory(btmBackTextFile, btmBackTextMemory, "Images\General\Back.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmBackHoverTextFile, btmBackHoverTextMemory, "Images\General\BackHover.png")

        'Options screen
        LoadPictureFileAndCopyBitmapIntoMemory(btmOptionsBackgroundFile, btmOptionsBackgroundMemory, "Images\Options\OptionsBackground.jpg")
        LoadPictureFileAndCopyBitmapIntoMemory(btmResolutionTextFile, btmResolutionTextMemory, "Images\Options\Resolution.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btm800x600TextFile, btm800x600TextMemory, "Images\Options\800x600.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNot800x600TextFile, btmNot800x600TextMemory, "Images\Options\not800x600.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btm1024x768TextFile, btm1024x768TextMemory, "Images\Options\1024x768.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNot1024x768TextFile, btmNot1024x768TextMemory, "Images\Options\not1024x768.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btm1280x800TextFile, btm1280x800TextMemory, "Images\Options\1280x800.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNot1280x800TextFile, btmNot1280x800TextMemory, "Images\Options\not1280x800.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btm1280x1024TextFile, btm1280x1024TextMemory, "Images\Options\1280x1024.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNot1280x1024TextFile, btmNot1280x1024TextMemory, "Images\Options\not1280x1024.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btm1440x900TextFile, btm1440x900TextMemory, "Images\Options\1440x900.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNot1440x900TextFile, btmNot1440x900TextMemory, "Images\Options\not1440x900.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmFullScreenTextFile, btmFullScreenTextMemory, "Images\Options\Fullscreen.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNotFullScreenTextFile, btmNotFullScreenTextMemory, "Images\Options\notFullscreen.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSoundTextFile, btmSoundTextMemory, "Images\Options\Sound.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSoundBarFile, btmSoundBarMemory, "Images\Options\SoundBar.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSliderFile, btmSliderMemory, "Images\Options\Slider.png")

        'Loading screen
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingBackgroundFile, btmLoadingBackgroundMemory, "Images\Loading Game\LoadingBackground.jpg")

        'Loading bar pictures
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingBarPictureFiles, abtmLoadingBarPictureMemories, "Images\Loading Game\LoadingBar", ".png", 0, 10)

        'Loading screen continued
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingTextFile, btmLoadingTextMemory, "Images\Loading Game\LoadingText.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingStartTextFile, btmLoadingStartTextMemory, "Images\Loading Game\LoadingStartText.png")

        'Loading paragraphs
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingParagraphFiles, abtmLoadingParagraphMemories, "Images\Loading Game\LoadingParagraph", ".png", 1, 1)

        'Highscores background
        LoadPictureFileAndCopyBitmapIntoMemory(btmHighscoresBackgroundFile, btmHighscoresBackgroundMemory, "Images\Highscores\HighscoresBackground.jpg")

        'Credits background
        LoadPictureFileAndCopyBitmapIntoMemory(btmCreditsBackgroundFile, btmCreditsBackgroundMemory, "Images\Credits\CreditsBackground.jpg")

        'John Gonzales
        LoadCreditsFilesAndCopyBitmapsIntoMemory(abtmJohnGonzalesFiles, abtmJohnGonzalesMemories, "JohnGonzales")

        'Zachary Stafford
        LoadCreditsFilesAndCopyBitmapsIntoMemory(abtmZacharyStaffordFiles, abtmZacharyStaffordMemories, "ZacharyStafford")

        'Cory Lewis
        LoadCreditsFilesAndCopyBitmapsIntoMemory(abtmCoryLewisFiles, abtmCoryLewisMemories, "CoryLewis")

        'Versus screen to put in an IP address
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusBackgroundFile, btmVersusBackgroundMemory, "Images\Versus\VersusBackground.jpg")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusNickNameTextFile, btmVersusNickNameTextMemory, "Images\Versus\VersusNickname.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusHostTextFile, btmVersusHostTextMemory, "Images\Versus\VersusHost.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusOrTextFile, btmVersusOrTextMemory, "Images\Versus\VersusOr.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusJoinTextFile, btmVersusJoinTextMemory, "Images\Versus\VersusJoin.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusIPAddressTextFile, btmVersusIPAddressTextMemory, "Images\Versus\VersusIPAddress.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusConnectTextFile, btmVersusConnectTextMemory, "Images\Versus\VersusConnect.png")

        'Loading versus screen
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingParagraphVersusFiles, abtmLoadingParagraphVersusMemories,
                                                      "Images\Loading Game\LoadingParagraphVersus", ".png", 1, 1)

        'Loading versus screen continued
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingWaitingTextFile, btmLoadingWaitingTextMemory, "Images\Loading Game\LoadingWaitingText.png")

        'Story screen
        LoadPictureFileAndCopyBitmapIntoMemory(btmStoryBackgroundFile, btmStoryBackgroundMemory, "Images\Story\StoryBackground.jpg")

        'Story paragraphs
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmStoryParagraphFiles, abtmStoryParagraphMemories, "Images\Story\Paragraph", ".png", 1, 1)

        'Game versus mismatch
        LoadPictureFileAndCopyBitmapIntoMemory(btmGameMismatchBackgroundFile, btmGameMismatchBackgroundMemory, "Images\Game Mismatch\GameMismatchBackground.jpg")

        'Load sound file percentages
        LoadSoundFilesAndCopyBitmapsIntoMemory

    End Sub

    Private Sub LoadPictureFileAndCopyBitmapIntoMemory(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap, strImageDirectory As String)

        'Load picture file bitmap
        btmByRefBitmapFile = New Bitmap(Image.FromFile(strDirectory & strImageDirectory))

        'Declare
        Dim intWidthFile As Integer = btmByRefBitmapFile.Width
        Dim intHeightFile As Integer = btmByRefBitmapFile.Height

        'Make new memory copy
        btmByRefBitmapMemory = New Bitmap(intWidthFile, intHeightFile, Imaging.PixelFormat.Format32bppPArgb)

        'Paint onto the memory copy
        DrawGraphicsByPoint(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)

        'Dispose file lock
        btmByRefBitmapFile.Dispose()

        'Empty
        btmByRefBitmapFile = Nothing

    End Sub

    Private Sub DrawGraphicsByPoint(ByRef gByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, pntBitmapToDraw As Point)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn

        'Set options for fastest rendering
        SetGraphicOptions(gGraphics)

        'Draw
        gGraphics.DrawImageUnscaled(btmBitmapToDraw, pntBitmapToDraw)

        'Dispose graphics
        gByRefGraphicsToDrawOn.Dispose()
        gByRefGraphicsToDrawOn = Nothing

        'Dispose pointer
        gGraphics.Dispose()
        gGraphics = Nothing

    End Sub

    Private Sub SetGraphicOptions(gGraphics As Graphics)

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

    Private Sub DrawGraphicsToScreenResolution(ByRef gByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, rectBitmapToDraw As Rectangle)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn

        'Set options for fastest rendering
        SetGraphicOptions(gGraphics)

        'Draw
        gGraphics.DrawImage(btmBitmapToDraw, rectBitmapToDraw) 'Scales it down by the rectangle

        'Dispose graphics
        gByRefGraphicsToDrawOn.Dispose()
        gByRefGraphicsToDrawOn = Nothing

        'Dispose pointer
        gGraphics.Dispose()
        gGraphics = Nothing

    End Sub

    Private Sub DrawText(ByRef gByRefGraphicsToDrawOn As Graphics, strText As String, sngFontSize As Single, colColor As Color, sngX As Single, sngY As Single,
                         sngWidth As Single, sngHeight As Single)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn
        Dim fntFont As New Font("SketchFlow Print", sngFontSize, FontStyle.Regular)
        Dim bruBrush As New System.Drawing.SolidBrush(colColor)

        'Set options for fastest rendering
        SetGraphicOptions(gGraphics)

        'Draw
        gGraphics.DrawString(strText, fntFont, bruBrush, New RectangleF(sngX, sngY, sngWidth, sngHeight))

        'Dispose
        fntFont.Dispose()
        bruBrush.Dispose()

        'Empty
        fntFont = Nothing
        bruBrush = Nothing

        'Dispose graphics
        gByRefGraphicsToDrawOn.Dispose()
        gByRefGraphicsToDrawOn = Nothing

        'Dispose pointer
        gGraphics.Dispose()
        gGraphics = Nothing

    End Sub

    Private Sub LoadArrayPictureFilesAndCopyBitmapsIntoMemory(ByRef abtmByRefPictureFiles() As Bitmap, ByRef abtmByRefPictureMemories() As Bitmap,
                                                              strImageDirectoryWithoutFileType As String, strFileType As String, intIndexStarting As Integer,
                                                              intIndexIncrease As Integer)

        'Declare
        Dim intIndex As Integer = intIndexStarting

        'Loading bar pictures
        For intLoop As Integer = 0 To abtmByRefPictureFiles.GetUpperBound(0)
            'Load
            LoadPictureFileAndCopyBitmapIntoMemory(abtmByRefPictureFiles(intLoop), abtmByRefPictureMemories(intLoop),
                                                   strImageDirectoryWithoutFileType & CStr(intIndex) & strFileType)
            'Increase
            intIndex += intIndexIncrease '0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
        Next

    End Sub

    Private Sub LoadCreditsFilesAndCopyBitmapsIntoMemory(ByRef abtmByRefFiles() As Bitmap, ByRef abtmByRefMemories() As Bitmap, strNameForCredits As String)

        'Load credits of the same file type
        For intLoop As Integer = 0 To (abtmByRefFiles.GetUpperBound(0) - 1)
            LoadPictureFileAndCopyBitmapIntoMemory(abtmByRefFiles(intLoop), abtmByRefMemories(intLoop),
                                                   "Images\Credits\" & strNameForCredits & CStr(intLoop + 1) & ".png")
        Next

        'Load seperate file
        LoadPictureFileAndCopyBitmapIntoMemory(abtmByRefFiles(3), abtmByRefMemories(3), "Images\Credits\" & strNameForCredits & "4.jpg")

    End Sub

    Private Sub LoadSoundFilesAndCopyBitmapsIntoMemory()

        'Load sound file percentages
        For intLoop As Integer = 0 To abtmSoundFiles.GetUpperBound(0)
            LoadPictureFileAndCopyBitmapIntoMemory(abtmSoundFiles(intLoop), abtmSoundMemories(intLoop), "Images\Options\Sound" & CStr(intLoop) & ".png")
        Next

    End Sub

    Private Sub SetupMouseOverTimer()

        'Disable timer by default, a mouse over the options will enable it later
        tmrOptionsMouseOver.Enabled = False

        'Set timer
        tmrOptionsMouseOver.AutoReset = True

        'Add handler
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

    Private Sub SetupParagraphTimer()

        'Disable timer by default
        tmrParagraph.Enabled = False 'Enabled later

        'Set timer
        tmrParagraph.AutoReset = True

        'Add handler
        AddHandler tmrParagraph.Elapsed, AddressOf ElapsedParagraph

        'Set timer delay
        tmrParagraph.Interval = LOADING_TRANSPARENCY_DELAY

    End Sub

    Private Sub ElapsedParagraph(sender As Object, e As EventArgs)

        'Check which type of wait
        Select Case strTypeOfParagraphWait

            Case "Single"
                'Check which mode
                Select Case intParagraphWaitMode
                    Case 0
                        'Set
                        btmLoadingParagraph = abtmLoadingParagraphMemories(0)
                    Case 1
                        'Set
                        btmLoadingParagraph = abtmLoadingParagraphMemories(1)
                    Case 2
                        'Set
                        btmLoadingParagraph = abtmLoadingParagraphMemories(2)
                    Case 3
                        'Set
                        btmLoadingParagraph = abtmLoadingParagraphMemories(3)
                        'Disable timer
                        tmrParagraph.Enabled = False
                End Select

            Case "Versus"
                'Check which mode
                Select Case intParagraphWaitMode
                    Case 0
                        'Set
                        btmLoadingParagraphVersus = abtmLoadingParagraphVersusMemories(0)
                    Case 1
                        'Set
                        btmLoadingParagraphVersus = abtmLoadingParagraphVersusMemories(1)
                    Case 2
                        'Set
                        btmLoadingParagraphVersus = abtmLoadingParagraphVersusMemories(2)
                    Case 3
                        'Set
                        btmLoadingParagraphVersus = abtmLoadingParagraphVersusMemories(3)
                        'Disable timer
                        tmrParagraph.Enabled = False
                End Select

        End Select

        'Increase or reset
        If intParagraphWaitMode = 3 Then
            intParagraphWaitMode = 0
        Else
            intParagraphWaitMode += 1
        End If

    End Sub

    Private Sub SetupCreditsTimer()

        'Disable timer by default
        tmrCredits.Enabled = False 'Enabled later

        'Set timer
        tmrCredits.AutoReset = True

        'Add handler
        AddHandler tmrCredits.Elapsed, AddressOf ElapsedCredits

        'Set timer delay
        tmrCredits.Interval = CREDITS_TRANSPARENCY_DELAY

    End Sub

    Private Sub ElapsedCredits(sender As Object, e As EventArgs)

        'Check which mode
        Select Case intCreditsWaitMode
            Case 0
                'Set
                btmJohnGonzales = abtmJohnGonzalesMemories(0)
                btmZacharyStafford = abtmZacharyStaffordMemories(0)
                btmCoryLewis = abtmCoryLewisMemories(0)
            Case 1
                'Set
                btmJohnGonzales = abtmJohnGonzalesMemories(1)
                btmZacharyStafford = abtmZacharyStaffordMemories(1)
                btmCoryLewis = abtmCoryLewisMemories(1)
            Case 2
                'Set
                btmJohnGonzales = abtmJohnGonzalesMemories(2)
                btmZacharyStafford = abtmZacharyStaffordMemories(2)
                btmCoryLewis = abtmCoryLewisMemories(2)
            Case 3
                'Set
                btmJohnGonzales = abtmJohnGonzalesMemories(3)
                btmZacharyStafford = abtmZacharyStaffordMemories(3)
                btmCoryLewis = abtmCoryLewisMemories(3)
                'Disable timer
                tmrCredits.Enabled = False
        End Select

        'Increase or reset
        If intCreditsWaitMode = 3 Then
            intCreditsWaitMode = 0
        Else
            intCreditsWaitMode += 1
        End If

    End Sub

    Private Sub SetupBlackScreenTimer()

        'Disable timer by default
        tmrBlackScreen.Enabled = False 'Enabled later

        'Set timer
        tmrBlackScreen.AutoReset = True

        'Add handler
        AddHandler tmrBlackScreen.Elapsed, AddressOf ElapsedBlackScreen

    End Sub

    Private Sub ElapsedBlackScreen(sender As Object, e As EventArgs)

        'Increase
        intBlackScreenWaitMode += 1

        'Disable
        If intBlackScreenWaitMode = 2 Then
            tmrBlackScreen.Enabled = False
        End If

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Rendering
        AbortThread(thrRendering)

        'Set
        thrRendering = Nothing

        'Loading beginning game material
        AbortThread(thrLoadBeginningGameMaterial)

        'Set
        thrLoadBeginningGameMaterial = Nothing

        'Loading game
        AbortThread(thrLoadingGame)

        'Set
        thrLoadingGame = Nothing

        'Story
        AbortThread(thrStory)

        'Set
        thrStory = Nothing

        'Remove game objects
        RemoveGameObjectsFromMemory()

        'Empty versus multiplayer variables
        EmptyMultiplayerVariables()

        'Stop all sounds
        StopAndCloseAllSounds()

    End Sub

    Private Sub StopAndCloseAllSounds()

        'Remove options mouse over timer
        RemoveOptionsMouseOverTimer()

        'Remove paragraph timer
        RemoveParagraphTimer()

        'Remove credits timer
        RemoveCreditsTimer()

        'Remove black screen timer
        RemoveBlackScreenTimer()

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
        For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
            If audcStoryParagraphSounds(intLoop) IsNot Nothing Then
                audcStoryParagraphSounds(intLoop).StopAndCloseSound()
                audcStoryParagraphSounds(intLoop) = Nothing
            End If
        Next

        'Game backgrounds
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            If audcGameBackgroundSounds(intLoop) IsNot Nothing Then
                audcGameBackgroundSounds(intLoop).StopAndCloseSound()
                audcGameBackgroundSounds(intLoop) = Nothing
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
        For intLoop As Integer = 0 To audcZombieDeathSounds.GetUpperBound(0)
            If audcZombieDeathSounds(intLoop) IsNot Nothing Then
                audcZombieDeathSounds(intLoop).StopAndCloseSound()
                audcZombieDeathSounds(intLoop) = Nothing
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

        'Disable timer
        tmrOptionsMouseOver.Enabled = False

        'Stop and dispose timer
        tmrOptionsMouseOver.Stop()
        tmrOptionsMouseOver.Dispose()

        'Remove handler
        RemoveHandler tmrOptionsMouseOver.Elapsed, AddressOf ElapsedOptionsMouseOver

    End Sub

    Private Sub RemoveParagraphTimer()

        'Disable timer
        tmrParagraph.Enabled = False

        'Stop and dispose timer
        tmrParagraph.Stop()
        tmrParagraph.Dispose()

        'Remove handler
        RemoveHandler tmrParagraph.Elapsed, AddressOf ElapsedParagraph

    End Sub

    Private Sub RemoveCreditsTimer()

        'Disable timer
        tmrCredits.Enabled = False

        'Stop and dispose timer
        tmrCredits.Stop()
        tmrCredits.Dispose()

        'Remove handler
        RemoveHandler tmrCredits.Elapsed, AddressOf ElapsedCredits

    End Sub

    Private Sub RemoveBlackScreenTimer()

        'Disable timer
        tmrBlackScreen.Enabled = False

        'Stop and dispose timer
        tmrBlackScreen.Stop()
        tmrBlackScreen.Dispose()

        'Remove handler
        RemoveHandler tmrBlackScreen.Elapsed, AddressOf ElapsedBlackScreen

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

        'Loading game material
        AbortThread(thrLoadBeginningGameMaterial)

        'Set
        thrLoadBeginningGameMaterial = Nothing

        'Loading versus
        AbortThread(thrLoadingGame)

        'Set
        thrLoadingGame = Nothing

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
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            If audcGameBackgroundSounds(intLoop) IsNot Nothing Then
                audcGameBackgroundSounds(intLoop).StopSound()
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
            'Hide the ugly gray screen
            Me.Hide()
            'Force focus
            Me.Focus()
            'Load necessary sounds for the beginning engine
            LoadNecessarySoundsForBeginningEngine() 'This needs a handle from the form window
            'Get highscores early because grabbing information from the database access files is slow
            LoadHighscoresStringAccess()
            'Set percentage multiplers for screen modes
            gdblScreenWidthRatio = CDbl((intScreenWidth - WIDTH_SUBTRACTION) / ORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((intScreenHeight - HEIGHT_SUBTRACTION) / ORIGINAL_SCREEN_HEIGHT)
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - WIDTH_SUBTRACTION, intScreenHeight - HEIGHT_SUBTRACTION) 'Full screen
            'Start rendering
            thrRendering.Start()
            'Un-hide
            Me.Show()
        End If

    End Sub

    Private Sub LoadNecessarySoundsForBeginningEngine()

        'Load ambiance
        For intLoop As Integer = 0 To audcAmbianceSound.GetUpperBound(0)
            audcAmbianceSound(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Ambiance" & CStr(intLoop + 1) & ".mp3", 1) 'Repeat only needs 1
        Next

        'Load button pressed
        udcButtonPressedSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonPressed.mp3", 3)

        'Load button hover
        udcButtonHoverSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonHover.mp3", 5)

        'Load story paragraphs
        For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
            audcStoryParagraphSounds(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\StoryParagraph" & CStr(intLoop + 1) & ".mp3", 1)
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
                blnCheckedGameMismatch = False
                'Set
                blnGameIsVersus = False
                'Set
                blnReadyEarly = False
                'Set
                blnWaiting = False
                'Set
                intCanvasVersusShow = 0
                'Stop and dispose game objects
                RemoveGameObjectsFromMemory()
                'Empty versus multiplayer variables
                EmptyMultiplayerVariables()
                'Set
                btmLoadingParagraph = Nothing
                '0% loaded
                btmLoadingBar = abtmLoadingBarPictureMemories(0)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0, blnPlayPressedSoundNow)
                'Set
                blnPlayPressedSoundNow = False
                'Reset fog x positions
                ResetFogXPositions()
                'Disable timer
                tmrBlackScreen.Enabled = False
                'Set
                blnFinishedLoading = False
                'Disable timer
                tmrCredits.Enabled = False
                'Reset
                btmJohnGonzales = Nothing
                btmZacharyStafford = Nothing
                btmCoryLewis = Nothing
                'Abort thread
                AbortThread(thrStory)
                'Reset
                btmStoryParagraph = Nothing
                'Stop sounds
                For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
                    audcStoryParagraphSounds(intLoop).StopSound()
                Next
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
            'Paint on screen
            DrawGraphicsToScreenResolution(Me.CreateGraphics(), btmCanvas, rectFullScreen)
            'If changing screen, we must change resolution in this thread or else strange things happen
            ScreenResolutionChanged()
            'Load parts of the game here after a screen has been shown, slowly load, this only happens during the loading screens
            Select Case intCanvasMode
                Case 2, 7
                    LoadingBitmapsIntoMemory()
            End Select
            'Check if need to load new level
            If blnCanLoadLevelWhileRendering Then
                'Check which level
                Select Case intLevel
                    Case 2
                        'Level underground tunnel
                        NextLevel(abtmGameBackgroundMemories(1))
                        'Change level, reuse the mechanics
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)
                    Case 3
                        'Level underground lab
                        NextLevel(abtmGameBackgroundMemories(2))
                        'Change level, reuse the mechanics
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)
                    Case 4
                        'Level outdoors swamp
                        NextLevel(abtmGameBackgroundMemories(3))
                        'Change level, reuse the mechanics
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)
                    Case 5
                        'Level helicopter escape
                        NextLevel(abtmGameBackgroundMemories(4), True)
                        'Change level, reuse the mechanics
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)
                End Select
                'Set
                blnCanLoadLevelWhileRendering = False
            End If
            'Do events
            Application.DoEvents()
        End While

    End Sub

    Private Sub NextLevel(btmGameBackgroundLevel As Bitmap, Optional blnLoadHelicopter As Boolean = False)

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

        'Memory copy level
        CopyLevelBitmap(btmGameBackgroundMemory, btmGameBackgroundLevel)

        'Set
        intZombieKills = 0

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
        audcGameBackgroundSounds(intLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)

    End Sub

    Private Sub RenderMenuScreen()

        'Declare
        Dim rectSourceBack As New Rectangle(pntFogBack.X + intFogBackX, 0, ORIGINAL_SCREEN_WIDTH, btmFogBackMemory.Height)
        Dim rectSourceFront As New Rectangle(pntFogFront.X + intFogFrontX, 0, ORIGINAL_SCREEN_WIDTH, btmFogFrontMemory.Height)

        'Change fog x positions
        intFogBackX -= FOG_SPEED
        intFogFrontX -= FOG_SPEED

        'Max out the limit of the x positions if necessary
        If intFogBackX + pntFogBack.X < 0 Then
            intFogBackX = 0
        End If
        If intFogFrontX + pntFogFront.X < 0 Then
            intFogFrontX = 0
        End If

        'Clone the only necessary spots
        btmFogBackCloneScreenShown = btmFogBackMemory.Clone(rectSourceBack, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short
        btmFogFrontCloneScreenShown = btmFogFrontMemory.Clone(rectSourceFront, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Paint onto canvas the menu background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmMenuBackgroundMemory, pntTopLeft)

        'Draw fog in back
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmFogBackCloneScreenShown, New Point(pntTopLeft.X, FOG_BACK_DISTANCE_Y))

        'Draw Archer
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmArcherMemory, pntArcher)

        'Draw fog in front
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmFogFrontCloneScreenShown, New Point(pntTopLeft.X, FOG_FRONT_DISTANCE_Y))

        'Dispose of fog clones because clones expand paint data, will cause memory leak
        btmFogBackCloneScreenShown.Dispose()
        btmFogBackCloneScreenShown = Nothing
        btmFogFrontCloneScreenShown.Dispose()
        btmFogFrontCloneScreenShown = Nothing

        'Draw start text
        If intCanvasShow = 1 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStartHoverTextMemory, pntStartHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStartTextMemory, pntStartText)
        End If

        'Draw highscores text
        If intCanvasShow = 2 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmHighscoresHoverTextMemory, pntHighscoresHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmHighscoresTextMemory, pntHighscoresText)
        End If

        'Draw story text
        If intCanvasShow = 3 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStoryHoverTextMemory, pntStoryHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStoryTextMemory, pntStoryText)
        End If

        'Draw options text
        If intCanvasShow = 4 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmOptionsHoverTextMemory, pntOptionsHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmOptionsTextMemory, pntOptionsText)
        End If

        'Draw credits text
        If intCanvasShow = 5 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmCreditsHoverTextMemory, pntCreditsHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmCreditsTextMemory, pntCreditsText)
        End If

        'Draw versus text
        If intCanvasShow = 6 Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusHoverTextMemory, pntVersusHoverText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusTextMemory, pntVersusText)
        End If

        'Draw last stand text
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLastStandTextMemory, pntLastStandText)

    End Sub

    Private Sub RenderOptionsScreen()

        'Draw options background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmOptionsBackgroundMemory, pntTopLeft)

        'Draw resolution text
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmResolutionTextMemory, pntResolutionText)

        'Draw sound text
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmSoundTextMemory, pntSoundText)

        'Check which resolution
        CheckResolutionMode(0, btm800x600TextMemory, btmNot800x600TextMemory, pnt800x600Text)
        CheckResolutionMode(1, btm1024x768TextMemory, btmNot1024x768TextMemory, pnt1024x768Text)
        CheckResolutionMode(2, btm1280x800TextMemory, btmNot1280x800TextMemory, pnt1280x800Text)
        CheckResolutionMode(3, btm1280x1024TextMemory, btmNot1280x1024TextMemory, pnt1280x1024Text)
        CheckResolutionMode(4, btm1440x900TextMemory, btmNot1440x900TextMemory, pnt1440x900Text)
        CheckResolutionMode(5, btmFullScreenTextMemory, btmNotFullScreenTextMemory, pntFullscreenText)

        'Draw sound bar
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmSoundBarMemory, pntSoundBar)

        'Draw sound percentage
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmSoundPercent, pntSoundPercent)

        'Draw slider
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmSliderMemory, pntSlider)

        'Draw version
        DrawText(Graphics.FromImage(btmCanvas), "Version " & GAME_VERSION, 50, Color.Black, 33, 953, 1000, 125) 'Black shadow
        DrawText(Graphics.FromImage(btmCanvas), "Version " & GAME_VERSION, 50, Color.Black, 35, 955, 1000, 125) 'Black shadow
        DrawText(Graphics.FromImage(btmCanvas), "Version " & GAME_VERSION, 50, Color.Red, 30, 950, 1000, 125) 'Overlay

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub ShowBackButtonOrHoverBackButton()

        'Show back button or hover back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmBackHoverTextMemory, pntBackHoverText)
        Else
            'Draw back text
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmBackTextMemory, pntBackText)
        End If

    End Sub

    Private Sub CheckResolutionMode(intModeSelected As Integer, btmResolutionText As Bitmap, btmNotResolutionText As Bitmap, pntResolutionText As Point)

        'Check resolution before drawing
        If intResolutionMode = intModeSelected Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmResolutionText, pntResolutionText)
        Else
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmNotResolutionText, pntResolutionText)
        End If

    End Sub

    Private Sub LoadingGameScreen()

        'Draw loading background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw Loading text
        If blnFinishedLoading Then
            'Start
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
        Else
            'Loading
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
        End If

        'Draw paragraph
        If btmLoadingParagraph IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingParagraph, pntLoadingParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub PathChoices()

        'Draw background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmPath, pntTopLeft)

        'Text
        DrawText(Graphics.FromImage(btmCanvas), "Pick your path...", 50, Color.Black, 33, 953, 1000, 125) 'Black shadow
        DrawText(Graphics.FromImage(btmCanvas), "Pick your path...", 50, Color.Black, 35, 955, 1000, 125) 'Black shadow
        DrawText(Graphics.FromImage(btmCanvas), "Pick your path...", 50, Color.Red, 30, 950, 1000, 125) 'Overlay

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub DrawStats(intZombieKillsToBe As Integer, intReloadTimeUsedToBe As Integer)

        'Draw zombie kills
        DrawText(Graphics.FromImage(btmCanvas), CStr(intZombieKillsToBe), 85, Color.Black, 1065, 417, 1000, 125)
        DrawText(Graphics.FromImage(btmCanvas), CStr(intZombieKillsToBe), 85, Color.White, 1060, 412, 1000, 125)

        'Declare
        Dim tsTimeSpan As TimeSpan
        Dim strElapsedTime As String

        'Set
        tsTimeSpan = swhStopWatch.Elapsed
        strElapsedTime = CStr(CInt(tsTimeSpan.TotalSeconds)) & " Sec"

        'Draw survival time
        DrawText(Graphics.FromImage(btmCanvas), strElapsedTime, 85, Color.Black, 1065, 615, 1000, 125)
        DrawText(Graphics.FromImage(btmCanvas), strElapsedTime, 85, Color.White, 1060, 610, 1000, 125)

        'Declare
        Dim intElapsedTime As Integer = 0
        Dim intWPM As Integer = 0

        'Get elapsed time
        intElapsedTime = CInt(Replace(strElapsedTime, " Sec", "")) - (intReloadTimeUsedToBe * 3)

        'Set WPM
        intWPM = CInt((intZombieKillsToBe / (intElapsedTime / 60)))

        'Draw WPM
        DrawText(Graphics.FromImage(btmCanvas), CStr(intWPM), 85, Color.Black, 1065, 808, 1000, 125)
        DrawText(Graphics.FromImage(btmCanvas), CStr(intWPM), 85, Color.White, 1060, 803, 1000, 125)

    End Sub

    Private Sub StartedGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Check if beat level
            If Not blnPlayerWasPinned Then

                'Paint black background
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), abtmBlackScreenMemories(2), pntTopLeft)

                'Check which level was completed
                Select Case intLevel
                    Case 1
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Set
                        btmPath = abtmPath1Memories(0)
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(11, 0, False) 'This exits the game screen to path choices
                    Case 2, 3
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Set
                        btmPath = abtmPath2Memories(0)
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(12, 0, False) 'This exits the game screen to path choices
                    Case 5
                        'Show win overlay
                        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmWinOverlayMemory, pntTopLeft)
                        'Draw stats
                        DrawStats(intZombieKillsCombined, intReloadTimes)
                End Select

            Else

                'Show death screen fading
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)

                'Show death overlay
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmDeathOverlayMemory, pntTopLeft)

                'Draw stats
                DrawStats(intZombieKillsCombined, intReloadTimes)

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

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub PlaySinglePlayerGame()

        'Check for helicopter
        If udcHelicopter IsNot Nothing Then
            'Draw
            DrawGraphicsByPoint(Graphics.FromImage(btmGameBackgroundMemory), udcHelicopter.HelicopterImage, udcHelicopter.HelicopterPoint)
        End If

        'Draw dead zombies permanently
        DrawDeadZombiesPermanently(gaudcZombies, intZombieKills)

        'Paint on canvas, clone the only necessary spot of the background, and draw word bar
        PaintOnCanvasCloneScreenAndDrawWordBar()

        'Check if made it to the end of the level
        If gpntGameBackground.X <= -2750 Then
            'Check if level was already beaten or if a zombie grabbed the player
            If Not blnBeatLevel And Not blnPlayerWasPinned Then
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
                'Start black screen
                BlackScreening(BLACK_SCREEN_BEAT_LEVEL_DELAY)
                'Stop level music
                audcGameBackgroundSounds(intLevel - 1).StopSound()
                'Pause the stop watch
                swhStopWatch.Stop()
                'Keep the reload times updated because object will be removed by memory
                intReloadTimes = udcCharacter.ReloadTimes
                'Keep the zombie kills updated
                intZombieKillsCombined += intZombieKills
            End If
        Else

            'Check character status if not game over
            If Not blnBeatLevel And Not blnPlayerWasPinned Then
                'Check character status
                CheckCharacterStatus()
            End If

        End If

        'Draw character
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), udcCharacter.CharacterImage, udcCharacter.CharacterPoint)

        'Draw zombies
        For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
            'Check if spawned
            If gaudcZombies(intLoop).Spawned Then
                'Check if can pin
                If Not gaudcZombies(intLoop).IsDying Then
                    'Check distance
                    If gaudcZombies(intLoop).ZombiePoint.X <= ZOMBIE_PINNING_X_DISTANCE - intZombieIncreasedPinDistance And Not gaudcZombies(intLoop).IsPinning Then
                        'Check if level not beat
                        If Not blnBeatLevel Then
                            'Increase distance
                            intZombieIncreasedPinDistance += 25
                            'Make zombie pin
                            gaudcZombies(intLoop).Pin()
                            'Check if first time game over by pin
                            If Not blnPlayerWasPinned Then
                                'Set
                                blnPlayerWasPinned = True
                                'Set
                                blnPreventKeyPressEvent = True
                                'Stop character from moving
                                If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                                    udcCharacter.Stand()
                                End If
                                'Start black screen
                                BlackScreening(BLACK_SCREEN_DEATH_DELAY)
                                'Stop reloading sound
                                udcReloadingSound.StopSound()
                                'Play
                                udcScreamSound.PlaySound(gintSoundVolume)
                                'Stop level music
                                audcGameBackgroundSounds(intLevel - 1).StopSound()
                                'Pause the stop watch
                                swhStopWatch.Stop()
                                'Keep the reload times updated because object will be removed by memory
                                intReloadTimes = udcCharacter.ReloadTimes
                                'Keep the zombie kills updated
                                intZombieKillsCombined += intZombieKills
                            End If
                        End If
                    End If
                End If
                'Draw zombies dying, pinning or walking
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), gaudcZombies(intLoop).ZombieImage, gaudcZombies(intLoop).ZombiePoint)
            End If
        Next

        'Draw text in the word bar
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.Black, 530, 95, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.Black, 528, 93, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.White, 525, 90, 1000, 100) 'White text

        'Word overlay
        DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(0, intWordIndex), 50, Color.Red, 525, 90, 1000, 100) 'Overlay

        'Show magazine with bullet count
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmAK47MagazineMemory, pntAK47Magazine)

        'Draw bullet count on magazine
        DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacter.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)

        'Check if need to black screen
        If blnBeatLevel Or blnPlayerWasPinned Then
            'Check which mode
            Select Case intBlackScreenWaitMode
                Case 0
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(0) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(0)
                    End If
                Case 1
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(1) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(1)
                    End If
                Case 2
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(2) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(2)
                    End If
            End Select
            'Paint black overlay
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmBlackScreen, pntTopLeft)
            'Check
            If btmBlackScreen Is abtmBlackScreenMemories(2) And blnBeatLevel Then
                'Set
                blnBlackScreenFinished = True
            End If
        End If

        'Make copy if died
        MakeCopyOfScreenBecauseCharacterDied()

    End Sub

    Private Sub DrawDeadZombiesPermanently(audcZombiesType() As clsZombie, ByRef intByRefZombieKills As Integer)

        'Declare
        Dim pntTemp As New Point() 'Default

        'Draw dead zombies permanently
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            If audcZombiesType(intLoop).Spawned And audcZombiesType(intLoop).IsDead Then
                'Set
                audcZombiesType(intLoop).Spawned = False
                'Set
                pntTemp = New Point(audcZombiesType(intLoop).ZombiePoint.X + CInt(Math.Abs(gpntGameBackground.X)), audcZombiesType(intLoop).ZombiePoint.Y)
                'Draw dead
                DrawGraphicsByPoint(Graphics.FromImage(btmGameBackgroundMemory), audcZombiesType(intLoop).ZombieImage, pntTemp)
                'Increase count
                intByRefZombieKills += 1
                'Start a new zombie
                audcZombiesType(intByRefZombieKills + NUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
            End If
        Next

    End Sub

    Private Sub PaintOnCanvasCloneScreenAndDrawWordBar()

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)

        'Clone the only necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackgroundMemory.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Draw the background to screen with the cloned version
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmGameBackgroundCloneScreenShown, pntTopLeft)

        'Dispose because clone just makes the picture expand with more data
        btmGameBackgroundCloneScreenShown.Dispose()
        btmGameBackgroundCloneScreenShown = Nothing

        'Draw word bar
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmWordBarMemory, pntWordBar)

    End Sub

    Private Sub MakeCopyOfScreenBecauseCharacterDied()

        'Check if need to copy screen
        If blnPlayerWasPinned Then
            'Check
            If btmBlackScreen Is abtmBlackScreenMemories(2) Then
                'Check if already happened
                If btmDeathScreen Is Nothing Then
                    'Before fading the screen, copy it to show for the death overlay
                    Dim rectSource As New Rectangle(0, 0, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)
                    'After paint make a copy
                    btmDeathScreen = btmCanvas.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)
                    'Set
                    blnBlackScreenFinished = True
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
                            CheckTheKeyPressBuffer(udcCharacter, gaudcZombies)

                    End Select

            End Select

        End If

    End Sub

    Private Sub CheckTheKeyPressBuffer(udcCharacterType As clsCharacter, audcZombiesType() As clsZombie)

        'Check the key press buffer
        If strKeyPressBuffer <> "" Then

            'Declare to be split
            Dim astrTemp() As String = Split(strKeyPressBuffer, ".") 'Each element is used because delimiter came first instead of last

            'Keep an index to remove only the necessary pieces of the key press buffer
            Dim intIndexesToRemove As Integer = 0

            'Prepare to send data if host
            Dim astrZombiesToKill(0) As String
            Dim strZombieDeathData As String = ""

            'Loop through
            For intLoop As Integer = 0 To (astrTemp.GetUpperBound(0) - 1)

                'Check length of word
                If Len(strWord) = 1 Then
                    'Check to make sure there is a spawned zombie to kill and isn't dying
                    If Not blnZombieToKillExists(audcZombiesType) Then
                        'Exit
                        Exit For
                    End If
                End If

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

                        'Declare
                        Dim intIndexOfClosestZombie As Integer = intGetIndexOfClosestZombie(audcZombiesType)

                        'Shoot
                        udcCharacterType.Shoot()

                        'Prepare to kill closest zombie
                        If blnGameIsVersus Then
                            'Check if hosting
                            If blnHost Then
                                'Kill zombie
                                audcZombiesType(intIndexOfClosestZombie).Dying()
                                'Add zombie to array string
                                If astrZombiesToKill.GetUpperBound(0) = 0 Then
                                    'Set
                                    astrZombiesToKill(0) = CStr(intIndexOfClosestZombie)
                                Else
                                    'Re-dim
                                    ReDim Preserve astrZombiesToKill(astrZombiesToKill.GetUpperBound(0) + 1)
                                    'Set
                                    astrZombiesToKill(astrZombiesToKill.GetUpperBound(0)) = CStr(intIndexOfClosestZombie)
                                End If
                            Else 'Joiner
                                'Send data
                                gSendData(5, "") 'This makes zombie waiting kills count up
                            End If
                        Else
                            'Kill zombie
                            audcZombiesType(intIndexOfClosestZombie).Dying()
                        End If

                        'Check if needs to reload
                        If udcCharacterType.BulletsUsed = 30 Then
                            'Set
                            udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload
                            'Reset
                            strKeyPressBuffer = ""
                            'Exit
                            Exit Sub
                        End If

                    End If

                End If

            Next

            'Remove if a zombie kill happened
            If intIndexesToRemove <> 0 Then

                'Remove only the necessary
                Dim strIndexPieces As String = ""

                'Loop
                For intLoop As Integer = 0 To intIndexesToRemove - 1
                    strIndexPieces &= astrTemp(intLoop) & "."
                Next

                'Remove the piece
                strKeyPressBuffer = strKeyPressBuffer.Substring(Len(strIndexPieces))

                'If hosting send data
                If blnHost Then
                    'Check if zombie kill data
                    If astrZombiesToKill(0) <> "" Then
                        'Prepare data to be complete
                        For intLoop As Integer = 0 To astrZombiesToKill.GetUpperBound(0)
                            strZombieDeathData &= astrZombiesToKill(intLoop) & "," 'Becomes zombie kill buffer one later for joiner
                        Next
                        'Send data
                        gSendData(4, strZombieDeathData)
                    End If
                End If

            End If

        End If

    End Sub

    Private Function blnZombieToKillExists(audcZombiesType() As clsZombie) As Boolean

        'Check through all zombies and make sure one exists to kill
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check for spawned
            If audcZombiesType(intLoop).Spawned Then
                'Check for not dying
                If Not audcZombiesType(intLoop).IsDying Then
                    'Return
                    Return True
                End If
            End If
        Next

        'Return
        Return False

    End Function

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

    Private Sub BlackScreening(intScreenDelay As Integer)

        'Set
        intBlackScreenWaitMode = 0

        'Set timer delay
        tmrBlackScreen.Interval = intScreenDelay

        'Enable
        tmrBlackScreen.Enabled = True

    End Sub

    Private Sub DrawDatabaseType(strText As String)

        'Draw
        DrawText(Graphics.FromImage(btmCanvas), strText, 34, Color.Black, 32, 22, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 34, Color.Black, 33, 23, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 34, Color.Black, 34, 24, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 34, Color.Black, 35, 25, ORIGINAL_SCREEN_WIDTH, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 34, Color.Red, 30, 20, ORIGINAL_SCREEN_WIDTH, 100)

    End Sub

    Private Sub HighscoresScreen()

        'Draw highscores background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmHighscoresBackgroundMemory, pntTopLeft)

        'Draw highscores
        DrawText(Graphics.FromImage(btmCanvas), strHighscores, 34, Color.Black, 32, 222, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strHighscores, 34, Color.Black, 33, 223, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strHighscores, 34, Color.Black, 34, 224, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strHighscores, 34, Color.Black, 35, 225, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strHighscores, 34, Color.Red, 30, 220, ORIGINAL_SCREEN_WIDTH, ORIGINAL_SCREEN_HEIGHT)

        'Check if Access
        If blnHighscoresIsAccess Then
            DrawDatabaseType("Database type is Access")
        Else
            DrawDatabaseType("Database type is Text File")
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

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

        'Draw credits background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmCreditsBackgroundMemory, pntTopLeft)

        'Draw credit pictures
        If btmJohnGonzales IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmJohnGonzales, pntJohnGonzales)
        End If
        If btmZacharyStafford IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmZacharyStafford, pntZacharyStafford)
        End If
        If btmCoryLewis IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmCoryLewis, pntCoryLewis)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub VersusScreen()

        'Draw background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusBackgroundMemory, pntTopLeft)

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

        'Other
        Select Case intCanvasVersusShow
            Case 0 'Nickname
                'Draw host button
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusHostTextMemory, pntVersusHost)
                'Draw or
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusOrTextMemory, pntVersusOr)
                'Draw join button
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusJoinTextMemory, pntVersusJoin)
                'Draw nickname
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusNickNameTextMemory, pntVersusBlackOutline)
                'Draw player name text
                DrawText(Graphics.FromImage(btmCanvas), strNickName, 110, Color.White, 103, 350, 1500, 275)
            Case 1 'Host
                'Draw hosting text
                DrawText(Graphics.FromImage(btmCanvas), "Hosting...", 72, Color.Red, 600, 450, 1000, 150)
            Case 2 'Join
                'Draw IP address
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusIPAddressTextMemory, pntVersusBlackOutline)
                'Draw IP address to type
                DrawText(Graphics.FromImage(btmCanvas), strIPAddressConnect, 110, Color.White, 103, 350, 1500, 275)
                'Draw connect button
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusConnectTextMemory, pntVersusConnect)
            Case 3 'Connecting
                'Draw connecting text
                DrawText(Graphics.FromImage(btmCanvas), "Connecting...", 72, Color.Red, 500, 450, 1000, 150)
        End Select

        'Draw ip address text
        DrawText(Graphics.FromImage(btmCanvas), strIPAddress, 72, Color.Red, 15, 25, 1000, 150)

        'Draw port forwarding
        DrawText(Graphics.FromImage(btmCanvas), "Router port forwarding: 10101", 50, Color.White, 375, 875, 1200, 125)

    End Sub

    Private Sub LoadingVersusConnectedScreen()

        'Draw loading background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw loading, waiting, and start text
        If blnWaiting Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingWaitingTextMemory, pntLoadingWaitingText)
        Else
            'Check if finished loading
            If blnFinishedLoading Then
                'Start
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
            Else
                'Loading
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
            End If
        End If

        'Draw paragraph
        If btmLoadingParagraphVersus IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmLoadingParagraphVersus, pntLoadingParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StartedVersusGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Show death overlay
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmVersusWhoWon, pntTopLeft)

            'Draw stats
            If blnHost Then
                DrawStats(intZombieKillsOne, intReloadTimes)
            Else
                DrawStats(intZombieKillsTwo, intReloadTimes)
            End If

            'Remove objects only once
            If Not blnRemovedGameObjectsFromMemory Then
                'Set
                blnRemovedGameObjectsFromMemory = True
                'Remove objects
                RemoveGameObjectsFromMemory(False) 'Don't stop the scream sound here
            End If

        Else

            'Play multiplayer game
            PlayMultiplayerGame()

        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub PlayMultiplayerGame()

        'Draw dead zombies permanently for hoster
        DrawDeadZombiesPermanently(gaudcZombiesOne, intZombieKillsOne)

        'Draw dead zombies permanently for joiner
        DrawDeadZombiesPermanently(gaudcZombiesTwo, intZombieKillsTwo)

        'Paint on canvas, clone the only necessary spot of the background, and draw word bar
        PaintOnCanvasCloneScreenAndDrawWordBar()

        'Check character status if not game over
        If Not blnPlayerWasPinned Then
            'Check character status
            If blnHost Then
                CheckCharacterMultiplayerStatus(udcCharacterOne, gaudcZombiesOne)
            Else
                CheckCharacterMultiplayerStatus(udcCharacterTwo, gaudcZombiesTwo)
            End If
        End If

        'Check the waiting kills, this is a different buffer on hoster's screen
        If blnHost And intZombieKillsWaitingTwo > 0 Then 'This only happens for hoster
            'Declare
            Dim strZombieKillData As String = ""
            'Loop
            For intLoop As Integer = 1 To intZombieKillsWaitingTwo
                'Check to make sure there is a spawned zombie to kill and isn't dying
                If Not blnZombieToKillExists(gaudcZombiesTwo) Then
                    'Exit
                    Exit For
                End If
                'Declare
                Dim intIndexOfClosestZombie As Integer = intGetIndexOfClosestZombie(gaudcZombiesTwo)
                'Show joiner shooting
                udcCharacterTwo.Shoot()
                'Kill zombie
                gaudcZombiesTwo(intIndexOfClosestZombie).Dying()
                'Add to zombie kill buffer
                strZombieKillData &= CStr(intIndexOfClosestZombie) & ","
                'Remove from zombie waiting kills
                intZombieKillsWaitingTwo -= 1
            Next
            'Send zombie death data
            If strZombieKillData <> "" Then
                gSendData(5, strZombieKillData)
            End If
        End If

        'Check zombie kill buffers for joiner's screen
        If Not blnHost Then
            'Zombies one
            CheckMultiplayerZombieKillBuffer(strZombieKillBufferOne, udcCharacterOne, gaudcZombiesOne, True)
            'Zombies two
            CheckMultiplayerZombieKillBuffer(strZombieKillBufferTwo, udcCharacterTwo, gaudcZombiesTwo) 'Joiner on the join screen doesn't need to shoot
        End If

        'Draw character hoster
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), udcCharacterOne.CharacterImage, udcCharacterOne.CharacterPoint)

        'Draw first zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesOne, ZOMBIE_PINNING_X_DISTANCE, intZombieIncreasedPinDistanceOne, 7)

        'Draw character joiner
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), udcCharacterTwo.CharacterImage, udcCharacterTwo.CharacterPoint)

        'Draw second zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesTwo, ZOMBIE_PINNING_X_DISTANCE + JOINER_ADDED_X_DISTANCE, intZombieIncreasedPinDistanceTwo, 8)

        'Send X Positions
        gSendData(3, strGetXPositionsOfZombies())

        'Draw text in the word bar
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.Black, 530, 95, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.Black, 528, 93, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 50, Color.White, 525, 90, 1000, 100) 'White text

        'Word overlay
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(0, intWordIndex), 50, Color.Red, 525, 90, 1000, 100) 'Overlay
        Else
            DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(0, intWordIndex), 50, Color.Blue, 525, 90, 1000, 100) 'Overlay
        End If

        'Show magazine with bullet count
        Graphics.FromImage(btmCanvas).DrawImageUnscaled(btmAK47MagazineMemory, pntAK47Magazine)

        'Draw nickname
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), strNickName, 36, Color.Red, 90, 205, 1000, 150) 'Host sees own name
            DrawText(Graphics.FromImage(btmCanvas), strNickNameConnected, 36, Color.Blue, 200, 255, 1000, 150) 'Host sees joiner name
        Else
            DrawText(Graphics.FromImage(btmCanvas), strNickNameConnected, 36, Color.Red, 90, 205, 1000, 150) 'Joiner sees host name
            DrawText(Graphics.FromImage(btmCanvas), strNickName, 36, Color.Blue, 200, 255, 1000, 150) 'Joiner sees own name
        End If

        'Draw bullet count on magazine
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacterOne.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)
        Else
            DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacterTwo.BulletsUsed), 40, Color.Blue, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)
        End If

        'Check if need to black screen
        If blnPlayerWasPinned Then
            'Check which mode
            Select Case intBlackScreenWaitMode
                Case 0
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(0) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(0)
                    End If
                Case 1
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(1) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(1)
                    End If
                Case 2
                    'Make sure this only happens once
                    If btmBlackScreen IsNot abtmBlackScreenMemories(2) Then
                        'Set
                        btmBlackScreen = abtmBlackScreenMemories(2)
                    End If
            End Select
            'Paint black overlay
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmBlackScreen, pntTopLeft)
        End If

        'Make copy if died
        MakeCopyOfScreenBecauseCharacterDied()

    End Sub

    Private Sub CheckMultiplayerZombieKillBuffer(ByRef strByRefZombieKillBufferType As String, udcCharacterType As clsCharacter, audcZombiesType() As clsZombie,
                                                 Optional blnShoot As Boolean = False)

        'Check string
        If strByRefZombieKillBufferType <> "" Then
            'Kill zombies in the buffer
            Dim astrTemp() As String = Split(strByRefZombieKillBufferType, ",")
            Dim intIndexesToRemove As Integer = 0
            'Loop to kill zombies
            For intLoop As Integer = 0 To (astrTemp.GetUpperBound(0) - 1)
                'Check to make sure there is a spawned zombie to kill
                If Not audcZombiesType(CInt(astrTemp(intLoop))).Spawned Then
                    'Exit
                    Exit For
                End If
                'Increase
                intIndexesToRemove += 1
                'Player shoots
                If blnShoot Then
                    udcCharacterType.Shoot()
                End If
                'Kill zombie
                audcZombiesType(CInt(astrTemp(intLoop))).Dying()
            Next
            'Remove if a zombie kill happened
            If intIndexesToRemove <> 0 Then
                'Remove only the necessary
                Dim strIndexPieces As String = ""
                'Loop
                For intLoop As Integer = 0 To (intIndexesToRemove - 1)
                    strIndexPieces &= astrTemp(intLoop) & ","
                Next
                'Remove the piece
                strByRefZombieKillBufferType = strByRefZombieKillBufferType.Substring(Len(strIndexPieces))
            End If
        End If

    End Sub

    Private Function strGetXPositionsOfZombies() As String

        'Notes: Looks like "0:1600,1:1800,|0:1600,1:1800," as final version

        'Declare
        Dim strReturn As String = ""

        'Get first set of zombie positions
        For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
            'Check if spawned and alive
            If gaudcZombiesOne(intLoop).Spawned And Not gaudcZombiesOne(intLoop).IsDying Then
                'Add onto the string
                strReturn &= CStr(intLoop) & ":" & CStr(gaudcZombiesOne(intLoop).ZombiePoint.X) & ","
            End If
        Next

        'Add onto the string
        strReturn &= "|"

        'Get second set of zombie positions
        For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
            'Check if spawned and alive
            If gaudcZombiesTwo(intLoop).Spawned And Not gaudcZombiesTwo(intLoop).IsDying Then
                'Add onto the string
                strReturn &= CStr(intLoop) & ":" & CStr(gaudcZombiesTwo(intLoop).ZombiePoint.X) & ","
            End If
        Next

        'Return
        Return strReturn

    End Function

    Private Sub CheckCharacterMultiplayerStatus(udcCharacterType As clsCharacter, audcZombiesType() As clsZombie)

        'Check to see if character must prepare to do something first
        Select Case udcCharacterType.StatusModeAboutToDo

            Case clsCharacter.eintStatusMode.Reload
                'Clear the key press buffer
                strKeyPressBuffer = ""
                'Make the character reload
                If udcCharacterType.GetPictureFrame = 1 Then
                    'Prepare to send data
                    udcCharacterType.PrepareSendData = True 'This is for multiplayer reloading
                    'Reload
                    udcCharacterType.Reload()
                    'Set
                    udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                End If

            Case Else 'What is the character going to currently do

                'If not stopping, what else is the character going to do
                Select Case udcCharacterType.StatusModeStartToDo

                    Case clsCharacter.eintStatusMode.Reload
                        'Clear the key press buffer
                        strKeyPressBuffer = ""
                        'Prepare to send data
                        udcCharacterType.PrepareSendData = True 'This is for multiplayer reloading
                        'Make the character reload
                        udcCharacterType.Reload()

                    Case clsCharacter.eintStatusMode.Stand, clsCharacter.eintStatusMode.Shoot
                        'Make the character shoot with buffer
                        CheckTheKeyPressBuffer(udcCharacterType, audcZombiesType)

                End Select

        End Select

    End Sub

    Private Sub DrawMultiplayerZombiesAndSendData(audcZombiesType() As clsZombie, intZombiePinningXDistance As Integer,
                                                  ByRef intByRefZombieIncreasedPinDistance As Integer, intDataCase As Integer)

        'Draw zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check if spawned
            If audcZombiesType(intLoop).Spawned Then
                'Check if can pin
                If Not audcZombiesType(intLoop).IsDying Then
                    'Check distance
                    If audcZombiesType(intLoop).ZombiePoint.X <= intZombiePinningXDistance - intByRefZombieIncreasedPinDistance And
                    Not audcZombiesType(intLoop).IsPinning Then
                        'Increase pin distance
                        intByRefZombieIncreasedPinDistance += 25
                        'Make zombie pin
                        audcZombiesType(intLoop).Pin()
                        'Check if first time game over by pin
                        If blnHost And Not blnPlayerWasPinned Then
                            'Set
                            blnPlayerWasPinned = True
                            'Set
                            blnPreventKeyPressEvent = True
                            'Start black screen
                            BlackScreening(BLACK_SCREEN_DEATH_DELAY)
                            'Stop reloading sound
                            udcReloadingSound.StopSound()
                            'Play
                            udcScreamSound.PlaySound(gintSoundVolume)
                            'Stop level music
                            audcGameBackgroundSounds(intLevel - 1).StopSound()
                            'Pause the stop watch
                            swhStopWatch.Stop()
                            'Keep the reload times updated because object will be removed by memory
                            intReloadTimes = udcCharacterOne.ReloadTimes 'This will always be player one
                            'Send data
                            gSendData(intDataCase, "") 'Joiner won if 7 and 8 joiner lost
                            'Show who won to host
                            Select Case intDataCase
                                Case 7
                                    btmVersusWhoWon = btmYouLostMemory
                                Case 8
                                    btmVersusWhoWon = btmYouWonMemory
                            End Select
                        End If
                    End If
                End If
                'Draw zombies dying, pinning or walking
                DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), audcZombiesType(intLoop).ZombieImage, audcZombiesType(intLoop).ZombiePoint)
            End If
        Next

    End Sub

    Private Sub StoryScreen()

        'Draw story background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStoryBackgroundMemory, pntTopLeft)

        'Draw story text
        If btmStoryParagraph IsNot Nothing Then
            DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmStoryParagraph, pntStoryParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StoryTelling()

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(83, 229)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(0)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(1)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(2)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(3)

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 1 sound
        audcStoryParagraphSounds(0).PlaySound(gintSoundVolume)

        'Sleep
        System.Threading.Thread.Sleep(49000)

        'Set
        btmStoryParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(38, 168)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(4)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(5)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(6)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(7)

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 2 sound
        audcStoryParagraphSounds(1).PlaySound(gintSoundVolume)

        'Sleep
        System.Threading.Thread.Sleep(64000)

        'Set
        btmStoryParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        pntStoryParagraph = New Point(31, 180)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(8)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(9)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(10)

        'Sleep
        System.Threading.Thread.Sleep(LOADING_TRANSPARENCY_DELAY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(11)

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 3 sound
        audcStoryParagraphSounds(2).PlaySound(gintSoundVolume)

    End Sub

    Private Sub GameVersionMismatch()

        'Draw story background
        DrawGraphicsByPoint(Graphics.FromImage(btmCanvas), btmGameMismatchBackgroundMemory, pntTopLeft)

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
        DrawText(Graphics.FromImage(btmCanvas), strText, 50, Color.Black, intX, intY, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 50, Color.Black, intX + 2, intY + 2, 1000, 100) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strText, 50, Color.Red, intX - 2, intY - 2, 1000, 100)

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
            'Sleep to prevent window resizing in a strange way, doesn't show up all the time, must have this
            System.Threading.Thread.Sleep(1)
        End If

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

            Case 2 'Loading game
                OneBackButtonMouseOverScreen(pntMouse, "LoadingGameBack")

            Case 3 'Game
                OneBackButtonMouseOverScreen(pntMouse, "GameBack")

            Case 4 'Highscores
                OneBackButtonMouseOverScreen(pntMouse, "HighscoresBack")

            Case 5 'Credits
                OneBackButtonMouseOverScreen(pntMouse, "CreditsBack")

            Case 6 'Versus
                OneBackButtonMouseOverScreen(pntMouse, "VersusBack")

            Case 7 'Loading versus game
                OneBackButtonMouseOverScreen(pntMouse, "LoadingVersusGameBack")

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
                btmPath = abtmPath1Memories(1)
                'Play light switch
                If Not blnLightZap1 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap1 = True
                End If

            Case blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) 'Path right
                'Set
                btmPath = abtmPath1Memories(2)
                'Play light switch
                If Not blnLightZap2 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap2 = True
                End If

            Case Else
                'Set
                btmPath = abtmPath1Memories(0)
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
                btmPath = abtmPath2Memories(1)
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
                btmPath = abtmPath2Memories(2)
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
                btmPath = abtmPath2Memories(0)
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

    Private Sub LoadBeginningGameMaterial()

        'Load the beginning game material thread
        thrLoadBeginningGameMaterial = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadBeginningGameMaterialThread))
        thrLoadBeginningGameMaterial.Start()

    End Sub

    Private Sub LoadBeginningGameMaterialThread()

        '0% loaded
        btmLoadingBar = abtmLoadingBarPictureMemories(0)

        'Set
        blnGameBackgroundMemoryCopied = False

        'Load variables
        LoadingGameVariables()

        'Words
        AddWordsToArray()

        'Grab a random word
        LoadARandomWord()

        'Load audio
        LoadGameAudio()

        '10% loaded
        btmLoadingBar = abtmLoadingBarPictureMemories(1)

        'Check if previously loaded to the end
        If intMemoryLoadPosition = 175 Then
            intMemoryLoadPosition -= 1
        End If

        'Continue loading
        RestartLoadingGame(False)

    End Sub

    Private Sub LoadingGameVariables()

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
        intZombieKillsCombined = 0

        'Set
        blnPreventKeyPressEvent = False

        'Set
        strKeyPressBuffer = ""

        'Set
        blnBeatLevel = False

        'Check if multiplayer game
        If blnGameIsVersus Then
            'Set
            intZombieKillsOne = 0
            intZombieKillsTwo = 0
            'Set
            intZombieKillsWaitingTwo = 0
            'Set
            strZombieKillBufferOne = ""
            strZombieKillBufferTwo = ""
            'Set
            btmVersusWhoWon = Nothing
            'Set
            intZombieIncreasedPinDistanceOne = 0
            intZombieIncreasedPinDistanceTwo = 0
        Else
            'Set
            intZombieKills = 0
            'Set
            intZombieIncreasedPinDistance = 0
        End If

    End Sub

    Private Sub SetBlackScreenVariables()

        'Set
        btmBlackScreen = Nothing

        'Set
        blnBlackScreenFinished = False

        'Set
        blnPlayerWasPinned = False

    End Sub

    Private Sub AddWordsToArray()

        'Note: This load has been designed this way incase we must rewind, load forward again

        'First time to declare an array of 0, 1 element
        If astrWords Is Nothing Then
            'Re-dim
            ReDim Preserve astrWords(0)
            'Set
            astrWords(0) = "and"
        End If

        'Wait until completed
        While astrWords.GetUpperBound(0) <> 74
            'Check which case to add a word
            Select Case astrWords.GetUpperBound(0)
                Case 0
                    ReDimPreserveAndAddWord(astrWords, "and", "aim")
                Case 1
                    ReDimPreserveAndAddWord(astrWords, "aim", "archer")
                Case 2
                    ReDimPreserveAndAddWord(astrWords, "archer", "arrive")
                Case 3
                    ReDimPreserveAndAddWord(astrWords, "arrive", "attack")
                Case 4
                    ReDimPreserveAndAddWord(astrWords, "attack", "battle")
                Case 5
                    ReDimPreserveAndAddWord(astrWords, "battle", "begins")
                Case 6
                    ReDimPreserveAndAddWord(astrWords, "begins", "bio")
                Case 7
                    ReDimPreserveAndAddWord(astrWords, "bio", "blood")
                Case 8
                    ReDimPreserveAndAddWord(astrWords, "blood", "carve")
                Case 9
                    ReDimPreserveAndAddWord(astrWords, "carve", "causes")
                Case 10
                    ReDimPreserveAndAddWord(astrWords, "causes", "chemical")
                Case 11
                    ReDimPreserveAndAddWord(astrWords, "chemical", "china")
                Case 12
                    ReDimPreserveAndAddWord(astrWords, "china", "communications")
                Case 13
                    ReDimPreserveAndAddWord(astrWords, "communications", "compound")
                Case 14
                    ReDimPreserveAndAddWord(astrWords, "compound", "contractions")
                Case 15
                    ReDimPreserveAndAddWord(astrWords, "contractions", "coughing")
                Case 16
                    ReDimPreserveAndAddWord(astrWords, "coughing", "countries")
                Case 17
                    ReDimPreserveAndAddWord(astrWords, "countries", "creating")
                Case 18
                    ReDimPreserveAndAddWord(astrWords, "creating", "deadly")
                Case 19
                    ReDimPreserveAndAddWord(astrWords, "deadly", "death")
                Case 20
                    ReDimPreserveAndAddWord(astrWords, "death", "destruction")
                Case 21
                    ReDimPreserveAndAddWord(astrWords, "destruction", "die")
                Case 22
                    ReDimPreserveAndAddWord(astrWords, "die", "egypt")
                Case 23
                    ReDimPreserveAndAddWord(astrWords, "egypt", "equipped")
                Case 24
                    ReDimPreserveAndAddWord(astrWords, "equipped", "escaped")
                Case 25
                    ReDimPreserveAndAddWord(astrWords, "escaped", "europe")
                Case 26
                    ReDimPreserveAndAddWord(astrWords, "europe", "ever")
                Case 27
                    ReDimPreserveAndAddWord(astrWords, "ever", "every")
                Case 28
                    ReDimPreserveAndAddWord(astrWords, "every", "everyone")
                Case 29
                    ReDimPreserveAndAddWord(astrWords, "everyone", "fall")
                Case 30
                    ReDimPreserveAndAddWord(astrWords, "fall", "finalize")
                Case 31
                    ReDimPreserveAndAddWord(astrWords, "finalize", "find")
                Case 32
                    ReDimPreserveAndAddWord(astrWords, "find", "france")
                Case 33
                    ReDimPreserveAndAddWord(astrWords, "france", "has")
                Case 34
                    ReDimPreserveAndAddWord(astrWords, "has", "hear")
                Case 35
                    ReDimPreserveAndAddWord(astrWords, "hear", "hosts")
                Case 36
                    ReDimPreserveAndAddWord(astrWords, "hosts", "in")
                Case 37
                    ReDimPreserveAndAddWord(astrWords, "in", "incursions")
                Case 38
                    ReDimPreserveAndAddWord(astrWords, "incursions", "infected")
                Case 39
                    ReDimPreserveAndAddWord(astrWords, "infected", "intelligence")
                Case 40
                    ReDimPreserveAndAddWord(astrWords, "intelligence", "involuntary")
                Case 41
                    ReDimPreserveAndAddWord(astrWords, "involuntary", "iranian")
                Case 42
                    ReDimPreserveAndAddWord(astrWords, "iranian", "is")
                Case 43
                    ReDimPreserveAndAddWord(astrWords, "is", "israel")
                Case 44
                    ReDimPreserveAndAddWord(astrWords, "israel", "jordan")
                Case 45
                    ReDimPreserveAndAddWord(astrWords, "jordan", "killing")
                Case 46
                    ReDimPreserveAndAddWord(astrWords, "killing", "lab")
                Case 47
                    ReDimPreserveAndAddWord(astrWords, "lab", "largest")
                Case 48
                    ReDimPreserveAndAddWord(astrWords, "largest", "leader")
                Case 49
                    ReDimPreserveAndAddWord(astrWords, "leader", "manage")
                Case 50
                    ReDimPreserveAndAddWord(astrWords, "manage", "muscles")
                Case 51
                    ReDimPreserveAndAddWord(astrWords, "muscles", "not")
                Case 52
                    ReDimPreserveAndAddWord(astrWords, "not", "of")
                Case 53
                    ReDimPreserveAndAddWord(astrWords, "of", "one")
                Case 54
                    ReDimPreserveAndAddWord(astrWords, "one", "painful")
                Case 55
                    ReDimPreserveAndAddWord(astrWords, "painful", "pathogen")
                Case 56
                    ReDimPreserveAndAddWord(astrWords, "pathogen", "pit")
                Case 57
                    ReDimPreserveAndAddWord(astrWords, "pit", "populations")
                Case 58
                    ReDimPreserveAndAddWord(astrWords, "populations", "respiratory")
                Case 59
                    ReDimPreserveAndAddWord(astrWords, "respiratory", "russia")
                Case 60
                    ReDimPreserveAndAddWord(astrWords, "russia", "schemed")
                Case 61
                    ReDimPreserveAndAddWord(astrWords, "schemed", "simultaneous")
                Case 62
                    ReDimPreserveAndAddWord(astrWords, "simultaneous", "site")
                Case 63
                    ReDimPreserveAndAddWord(astrWords, "site", "summer")
                Case 64
                    ReDimPreserveAndAddWord(astrWords, "summer", "supreme")
                Case 65
                    ReDimPreserveAndAddWord(astrWords, "supreme", "symptoms")
                Case 66
                    ReDimPreserveAndAddWord(astrWords, "symptoms", "targeting")
                Case 67
                    ReDimPreserveAndAddWord(astrWords, "targeting", "terror")
                Case 68
                    ReDimPreserveAndAddWord(astrWords, "terror", "the")
                Case 69
                    ReDimPreserveAndAddWord(astrWords, "the", "this")
                Case 70
                    ReDimPreserveAndAddWord(astrWords, "three", "to")
                Case 71
                    ReDimPreserveAndAddWord(astrWords, "to", "total")
                Case 72
                    ReDimPreserveAndAddWord(astrWords, "total", "weapons")
                Case 73
                    ReDimPreserveAndAddWord(astrWords, "weapons", "with")
            End Select
        End While

        'Check last word
        If astrWords(astrWords.GetUpperBound(0)) = "" Then
            astrWords(astrWords.GetUpperBound(0)) = "with" 'Must check last word
        End If

    End Sub

    Private Sub ReDimPreserveAndAddWord(ByRef astrByRefWord() As String, strWordPreviouslyDeclared As String, strWord As String)

        'Check previous first
        If astrByRefWord(astrByRefWord.GetUpperBound(0)) = "" Then
            'Set
            astrByRefWord(astrByRefWord.GetUpperBound(0)) = strWordPreviouslyDeclared
        End If

        'Re-dim preserve
        ReDim Preserve astrByRefWord(astrByRefWord.GetUpperBound(0) + 1)

        'Add word
        astrByRefWord(astrByRefWord.GetUpperBound(0)) = strWord

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

    Private Sub LoadGameAudio()

        'Load game background sounds
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            If audcGameBackgroundSounds(intLoop) Is Nothing Then
                audcGameBackgroundSounds(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GameBackground" & CStr(intLoop + 1) & ".mp3", 1)
            End If
        Next

        'Load scream sound
        If udcScreamSound Is Nothing Then
            udcScreamSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Scream.mp3", 1)
        End If

        'Load gun shot sound
        If udcGunShotSound Is Nothing Then
            udcGunShotSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3", 10)
        End If

        'Load zombie death sounds
        For intLoop As Integer = 0 To audcZombieDeathSounds.GetUpperBound(0)
            If audcZombieDeathSounds(intLoop) Is Nothing Then
                audcZombieDeathSounds(intLoop) = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ZombieDeath" & CStr(intLoop + 1) & ".mp3", 10)
            End If
        Next

        'Load reloading sound
        If udcReloadingSound Is Nothing Then
            udcReloadingSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Reloading.mp3", 2) 'Incase multiplayer
        End If

        'Load step sound
        If udcStepSound Is Nothing Then
            udcStepSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\Step.mp3", 6)
        End If

        'Load water foot left sound
        If udcWaterFootStepLeftSound Is Nothing Then
            udcWaterFootStepLeftSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\WaterFootStepLeft.mp3", 3)
        End If

        'Load water foot right sound
        If udcWaterFootStepRightSound Is Nothing Then
            udcWaterFootStepRightSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\WaterFootStepRight.mp3", 3)
        End If

        'Load gravel foot left sound
        If udcGravelFootStepLeftSound Is Nothing Then
            udcGravelFootStepLeftSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GravelFootStepLeft.mp3", 3)
        End If

        'Load gravel foot right sound
        If udcGravelFootStepRightSound Is Nothing Then
            udcGravelFootStepRightSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\GravelFootStepRight.mp3", 3)
        End If

        'Load opening metal door sound
        If udcOpeningMetalDoorSound Is Nothing Then
            udcOpeningMetalDoorSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\OpeningMetalDoor.mp3", 1) 'Happens only once during gameplay
        End If

        'Load light zap sound
        If udcLightZapSound Is Nothing Then
            udcLightZapSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\LightZap.mp3", 5)
        End If

        'Load zombie growl sound
        If udcZombieGrowlSound Is Nothing Then
            udcZombieGrowlSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ZombieGrowl.mp3", 1)
        End If

        'Load rotating blade sound of the helicopter
        If udcRotatingBladeSound Is Nothing Then
            udcRotatingBladeSound = New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\RotatingBlade.mp3", 1)
        End If

    End Sub

    Private Sub RestartLoadingGame(blnIncreaseMemoryPosition As Boolean)

        'Increase loading position
        If blnIncreaseMemoryPosition Then
            intMemoryLoadPosition += 1
        End If

        'Load the loading game thread
        thrLoadingGame = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGameThread))
        thrLoadingGame.Start()

    End Sub

    Private Sub LoadingGameThread()

        'Note: This will load but wait until memory has completed before doing another load, each file is loaded individually, waiting for a response from memory if it finished copying that file

        'Load continually
        Select Case intMemoryLoadPosition
            Case 0 To 4
                'File load array
                LoadGameFileWithIndex(0, abtmGameBackgroundFiles, ablnGameBackgroundMemoriesCopied, "Images\Game Play\GameBackground", ".jpg")
            Case 5
                'File load
                LoadGameFile(btmWordBarFile, blnWordBarMemoryCopied, "Images\Game Play\WordBar.png")
            Case 6, 7
                'File load array
                LoadGameFileWithIndex(6, abtmCharacterStandFiles, ablnCharacterStandMemoriesCopied, "Images\Character\Generic\Standing\", ".png")
            Case 8, 9
                'File load array
                LoadGameFileWithIndex(8, abtmCharacterShootFiles, ablnCharacterShootMemoriesCopied, "Images\Character\Generic\Shoot Once\", ".png")
            Case 10 To 31
                'File load array
                LoadGameFileWithIndex(10, abtmCharacterReloadFiles, ablnCharacterReloadMemoriesCopied, "Images\Character\Generic\Reload\", ".png")
            Case 32 To 48
                'File load array
                LoadGameFileWithIndex(32, abtmCharacterRunningFiles, ablnCharacterRunningMemoriesCopied, "Images\Character\Generic\Running\", ".png")
            Case 49, 50
                'File load array
                LoadGameFileWithIndex(49, abtmCharacterStandRedFiles, ablnCharacterStandRedMemoriesCopied, "Images\Character\Red\Standing\", ".png")
            Case 51, 52
                'File load array
                LoadGameFileWithIndex(51, abtmCharacterShootRedFiles, ablnCharacterShootRedMemoriesCopied, "Images\Character\Red\Shoot Once\", ".png")
            Case 53 To 74
                'File load array
                LoadGameFileWithIndex(53, abtmCharacterReloadRedFiles, ablnCharacterReloadRedMemoriesCopied, "Images\Character\Red\Reload\", ".png")
            Case 75, 76
                'File load array
                LoadGameFileWithIndex(75, abtmCharacterStandBlueFiles, ablnCharacterStandBlueMemoriesCopied, "Images\Character\Blue\Standing\", ".png")
            Case 77, 78
                'File load array
                LoadGameFileWithIndex(77, abtmCharacterShootBlueFiles, ablnCharacterShootBlueMemoriesCopied, "Images\Character\Blue\Shoot Once\", ".png")
            Case 79 To 100
                'File load array
                LoadGameFileWithIndex(79, abtmCharacterReloadBlueFiles, ablnCharacterReloadBlueMemoriesCopied, "Images\Character\Blue\Reload\", ".png")
            Case 101 To 104
                'File load array
                LoadGameFileWithIndex(101, abtmZombieWalkFiles, ablnZombieWalkMemoriesCopied, "Images\Zombies\Generic\Movement\", ".png")
            Case 105 To 110
                'File load array
                LoadGameFileWithIndex(105, abtmZombieDeath1Files, ablnZombieDeath1MemoriesCopied, "Images\Zombies\Generic\Death 1\", ".png")
            Case 111 To 116
                'File load array
                LoadGameFileWithIndex(111, abtmZombieDeath2Files, ablnZombieDeath2MemoriesCopied, "Images\Zombies\Generic\Death 2\", ".png")
            Case 117, 118
                'File load array
                LoadGameFileWithIndex(117, abtmZombiePinFiles, ablnZombiePinMemoriesCopied, "Images\Zombies\Generic\Pinning\", ".png")
            Case 119 To 122
                'File load array
                LoadGameFileWithIndex(119, abtmZombieWalkRedFiles, ablnZombieWalkRedMemoriesCopied, "Images\Zombies\Red\Movement\", ".png")
            Case 123 To 128
                'File load array
                LoadGameFileWithIndex(123, abtmZombieDeathRed1Files, ablnZombieDeathRed1MemoriesCopied, "Images\Zombies\Red\Death 1\", ".png")
            Case 129 To 134
                'File load array
                LoadGameFileWithIndex(129, abtmZombieDeathRed2Files, ablnZombieDeathRed2MemoriesCopied, "Images\Zombies\Red\Death 2\", ".png")
            Case 135, 136
                'File load array
                LoadGameFileWithIndex(135, abtmZombiePinRedFiles, ablnZombiePinRedMemoriesCopied, "Images\Zombies\Red\Pinning\", ".png")
            Case 137 To 140
                'File load array
                LoadGameFileWithIndex(137, abtmZombieWalkBlueFiles, ablnZombieWalkBlueMemoriesCopied, "Images\Zombies\Blue\Movement\", ".png")
            Case 141 To 146
                'File load array
                LoadGameFileWithIndex(141, abtmZombieDeathBlue1Files, ablnZombieDeathBlue1MemoriesCopied, "Images\Zombies\Blue\Death 1\", ".png")
            Case 147 To 152
                'File load array
                LoadGameFileWithIndex(147, abtmZombieDeathBlue2Files, ablnZombieDeathBlue2MemoriesCopied, "Images\Zombies\Blue\Death 2\", ".png")
            Case 153, 154
                'File load array
                LoadGameFileWithIndex(153, abtmZombiePinBlueFiles, ablnZombiePinBlueMemoriesCopied, "Images\Zombies\Blue\Pinning\", ".png")
            Case 155 To 159
                'File load array
                LoadGameFileWithIndex(155, abtmHelicopterFiles, ablnHelicopterMemoriesCopied, "Images\Helicopters\HelicopterGonzales\", ".jpg")
            Case 160
                'File load
                LoadGameFile(btmAK47MagazineFile, blnAK47MagazineMemoryCopied, "Images\Game Play\AK47Magazine.png")
            Case 161
                'File load
                LoadGameFile(btmDeathOverlayFile, blnDeathOverlayMemoryCopied, "Images\Game Play\DeathOverlay.png")
            Case 162
                'File load
                LoadGameFile(btmWinOverlayFile, blnWinOverlayMemoryCopied, "Images\Game Play\WinOverlay.jpg")
            Case 163
                'File load
                LoadGameFile(btmYouWonFile, blnYouWonMemoryCopied, "Images\Versus\YouWon.png")
            Case 164
                'File load
                LoadGameFile(btmYouLostFile, blnYouLostMemoryCopied, "Images\Versus\YouLost.png")
            Case 165 To 167
                'File load array
                LoadGameFileWithIndex(165, abtmBlackScreenFiles, ablnBlackScreenMemoriesCopied, "Images\Game Play\Black Screen\", ".png")
            Case 168 To 170
                'File load array
                LoadGameFileWithIndex(168, abtmPath1Files, ablnPath1MemoriesCopied, "Images\Game Play\Paths\Path 1\", ".jpg")
            Case 171 To 173
                'File load array
                LoadGameFileWithIndex(171, abtmPath2Files, ablnPath2MemoriesCopied, "Images\Game Play\Paths\Path 2\", ".jpg")
            Case 174
                'Not used here, but below after graphics renders, skipped in the thread loading
            Case 175 'Be careful changing this number, find = check if previously loaded to the end
                'Check if single player
                If Not blnGameIsVersus Then
                    'Character
                    udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                                    udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)
                    'Zombies
                    LoadZombies("Level 1 Single Player")
                Else
                    'Load in a special way
                    If blnHost Then
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound, True) 'Join
                    Else
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound, True) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", intLevel, udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound) 'Join
                    End If
                    'Zombies
                    LoadZombies("Level 1 Multiplayer")
                    'Set if hosting
                    If blnHost And Not blnReadyEarly Then
                        blnWaiting = True
                    End If
                End If
                '100%
                btmLoadingBar = abtmLoadingBarPictureMemories(10)
                'Set
                blnFinishedLoading = True 'Means completely loaded
        End Select

    End Sub

    Private Sub LoadGameFileWithIndex(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ablnMemoryCopied() As Boolean, strImageSubDirectory As String, strFileType As String)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract

        'File load
        LoadGameFile(abtmByRefFile(intIndexChanger), ablnMemoryCopied(intIndexChanger), strImageSubDirectory & CStr(intIndexChanger + 1) & strFileType)

    End Sub

    Private Sub LoadGameFile(ByRef btmByRefFile As Bitmap, blnMemoryCopied As Boolean, strImageDirectory As String)

        'Load file
        If btmByRefFile Is Nothing And Not blnMemoryCopied Then
            'Load
            btmByRefFile = New Bitmap(Image.FromFile(strDirectory & strImageDirectory))
        End If

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
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 10, "udcCharacter", audcZombieDeathSounds)
                        Case 1
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, 10, "udcCharacter", audcZombieDeathSounds)
                        Case 2
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, 10, "udcCharacter", audcZombieDeathSounds)
                        Case 3
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, 10, "udcCharacter", audcZombieDeathSounds)
                        Case 4
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, 10, "udcCharacter", audcZombieDeathSounds)
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 15, "udcCharacter", audcZombieDeathSounds)
                        Case 9, 14, 19, 24
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 35, "udcCharacter", audcZombieDeathSounds)
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacter", audcZombieDeathSounds)
                        Case 29, 34, 39, 44
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 40, "udcCharacter", audcZombieDeathSounds)
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacter", audcZombieDeathSounds)
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 45, "udcCharacter", audcZombieDeathSounds)
                        Case 65 To 74
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 15, "udcCharacter", audcZombieDeathSounds)
                        Case 75 To 84
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacter", audcZombieDeathSounds)
                        Case 85 To 94
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacter", audcZombieDeathSounds)
                        Case 95 To 104
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 30, "udcCharacter", audcZombieDeathSounds)
                        Case Else
                            gaudcZombies(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 55, "udcCharacter", audcZombieDeathSounds)
                    End Select
                Next


            Case "Level 1 Multiplayer"
                'Check if hosting, if not hosting then ghost like property the zombies
                If blnHost Then 'Hoster
                    LoadMultiplayerZombies(False, True) 'True walking zombies, not ghost like
                Else 'Joiner
                    LoadMultiplayerZombies(True, False) 'Ghost non-walking zombies, x position must be updated with get data
                End If

        End Select

    End Sub

    Private Sub LoadMultiplayerZombies(blnImitation As Boolean, blnSendXPositionData As Boolean)

        'Zombies for player
        For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)

            'Re-dim first
            ReDim Preserve gaudcZombiesOne(intLoop)

            'Check which wave number
            Select Case intLoop

                Case 0
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 10, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 1
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500, 325, 10, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 2
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000, 325, 10, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 3
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500, 325, 10, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 4
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000, 325, 10, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 15, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 9, 14, 19, 24
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 35, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 29, 34, 39, 44
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 40, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 48, 49, 53, 54, 58, 59, 63, 64
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 45, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 65 To 74
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 15, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 75 To 84
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 85 To 94
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 25, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case 95 To 104
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 30, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

                Case Else
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH, 325, 55, "udcCharacterOne", audcZombieDeathSounds, blnImitation)

            End Select

        Next

        'Zombies for joiner
        For intLoop As Integer = 0 To (NUMBER_OF_ZOMBIES_CREATED - 1)

            'Re-dim first
            ReDim Preserve gaudcZombiesTwo(intLoop)

            'Check which wave number
            Select Case intLoop

                Case 0
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 10, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 1
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 500 + JOINER_ADDED_X_DISTANCE, 350, 10, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 2
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1000 + JOINER_ADDED_X_DISTANCE, 350, 10, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 3
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 1500 + JOINER_ADDED_X_DISTANCE, 350, 10, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 4
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + 2000 + JOINER_ADDED_X_DISTANCE, 350, 10, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 15, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 9, 14, 19, 24
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 35, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 25, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 29, 34, 39, 44
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 40, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 25, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 48, 49, 53, 54, 58, 59, 63, 64
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 45, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 65 To 74
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 15, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 75 To 84
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 25, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 85 To 94
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 25, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case 95 To 104
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 30, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

                Case Else
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, ORIGINAL_SCREEN_WIDTH + JOINER_ADDED_X_DISTANCE, 350, 55, "udcCharacterTwo", audcZombieDeathSounds,
                                               blnImitation)

            End Select

        Next

    End Sub

    Private Sub LoadingBitmapsIntoMemory()

        'Note: This waits for files to be completed and then loads it into memory

        'Load into memory
        Select Case intMemoryLoadPosition
            Case 0 To 4
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(0, abtmGameBackgroundFiles, abtmGameBackgroundMemories, ablnGameBackgroundMemoriesCopied)
            Case 5
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmWordBarFile, btmWordBarMemory, blnWordBarMemoryCopied)
            Case 6, 7
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(6, abtmCharacterStandFiles, gabtmCharacterStandMemories, ablnCharacterStandMemoriesCopied)
            Case 8, 9
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(8, abtmCharacterShootFiles, gabtmCharacterShootMemories, ablnCharacterShootMemoriesCopied)
            Case 10 To 31
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(10, abtmCharacterReloadFiles, gabtmCharacterReloadMemories, ablnCharacterReloadMemoriesCopied)
            Case 32 To 48
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(32, abtmCharacterRunningFiles, gabtmCharacterRunningMemories, ablnCharacterRunningMemoriesCopied)
                '20% loaded
                SetLoadingBarPercentage(48, 2)
            Case 49, 50
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(49, abtmCharacterStandRedFiles, gabtmCharacterStandRedMemories, ablnCharacterStandRedMemoriesCopied)
            Case 51, 52
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(51, abtmCharacterShootRedFiles, gabtmCharacterShootRedMemories, ablnCharacterShootRedMemoriesCopied)
            Case 53 To 74
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(53, abtmCharacterReloadRedFiles, gabtmCharacterReloadRedMemories, ablnCharacterReloadRedMemoriesCopied)
                '30% loaded
                SetLoadingBarPercentage(74, 3)
            Case 75, 76
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(75, abtmCharacterStandBlueFiles, gabtmCharacterStandBlueMemories, ablnCharacterStandBlueMemoriesCopied)
            Case 77, 78
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(77, abtmCharacterShootBlueFiles, gabtmCharacterShootBlueMemories, ablnCharacterShootBlueMemoriesCopied)
            Case 79 To 100
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(79, abtmCharacterReloadBlueFiles, gabtmCharacterReloadBlueMemories, ablnCharacterReloadBlueMemoriesCopied)
                '40% loaded
                SetLoadingBarPercentage(100, 4)
            Case 101 To 104
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(101, abtmZombieWalkFiles, gabtmZombieWalkMemories, ablnZombieWalkMemoriesCopied)
            Case 105 To 110
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(105, abtmZombieDeath1Files, gabtmZombieDeath1Memories, ablnZombieDeath1MemoriesCopied)
            Case 111 To 116
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(111, abtmZombieDeath2Files, gabtmZombieDeath2Memories, ablnZombieDeath2MemoriesCopied)
            Case 117, 118
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(117, abtmZombiePinFiles, gabtmZombiePinMemories, ablnZombiePinMemoriesCopied)
                '50% loaded
                SetLoadingBarPercentage(118, 5)
            Case 119 To 122
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(119, abtmZombieWalkRedFiles, gabtmZombieWalkRedMemories, ablnZombieWalkRedMemoriesCopied)
            Case 123 To 128
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(123, abtmZombieDeathRed1Files, gabtmZombieDeathRed1Memories, ablnZombieDeathRed1MemoriesCopied)
            Case 129 To 134
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(129, abtmZombieDeathRed2Files, gabtmZombieDeathRed2Memories, ablnZombieDeathRed2MemoriesCopied)
            Case 135, 136
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(135, abtmZombiePinRedFiles, gabtmZombiePinRedMemories, ablnZombiePinRedMemoriesCopied)
                '60% loaded
                SetLoadingBarPercentage(136, 6)
            Case 137 To 140
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(137, abtmZombieWalkBlueFiles, gabtmZombieWalkBlueMemories, ablnZombieWalkBlueMemoriesCopied)
            Case 141 To 146
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(141, abtmZombieDeathBlue1Files, gabtmZombieDeathBlue1Memories, ablnZombieDeathBlue1MemoriesCopied)
            Case 147 To 152
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(147, abtmZombieDeathBlue2Files, gabtmZombieDeathBlue2Memories, ablnZombieDeathBlue2MemoriesCopied)
            Case 153, 154
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(153, abtmZombiePinBlueFiles, gabtmZombiePinBlueMemories, ablnZombiePinBlueMemoriesCopied)
                '70% loaded
                SetLoadingBarPercentage(154, 7)
            Case 155 To 159
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(155, abtmHelicopterFiles, gabtmHelicopterMemories, ablnHelicopterMemoriesCopied)
            Case 160
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmAK47MagazineFile, btmAK47MagazineMemory, blnAK47MagazineMemoryCopied)
                '80% loaded
                SetLoadingBarPercentage(160, 8)
            Case 161
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmDeathOverlayFile, btmDeathOverlayMemory, blnDeathOverlayMemoryCopied)
            Case 162
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmWinOverlayFile, btmWinOverlayMemory, blnWinOverlayMemoryCopied)
            Case 163
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmYouWonFile, btmYouWonMemory, blnYouWonMemoryCopied)
            Case 164
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmYouLostFile, btmYouLostMemory, blnYouLostMemoryCopied)
            Case 165 To 167
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(165, abtmBlackScreenFiles, abtmBlackScreenMemories, ablnBlackScreenMemoriesCopied)
            Case 168 To 170
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(168, abtmPath1Files, abtmPath1Memories, ablnPath1MemoriesCopied)
            Case 171 To 173
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(171, abtmPath2Files, abtmPath2Memories, ablnPath2MemoriesCopied)
            Case 174 'Be careful changing this number, find = check if previously loaded to the end
                'Memory copy level
                CopyLevelBitmap(btmGameBackgroundMemory, abtmGameBackgroundMemories(0))
                '90% loaded
                SetLoadingBarPercentage(174, 9)
                'Restart the thread, should do the last 100%
                RestartLoadingGame(True)
        End Select

    End Sub

    Private Sub CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ByRef abtmByRefMemory() As Bitmap,
                                                                ByRef ablnByRefMemoryCopied() As Boolean)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract

        'File load
        CopyBitmapIntoMemoryAfterDrawingScreen(abtmByRefFile(intIndexChanger), abtmByRefMemory(intIndexChanger), ablnByRefMemoryCopied(intIndexChanger))

    End Sub

    Private Sub CopyBitmapIntoMemoryAfterDrawingScreen(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap,
                                                       ByRef blnByRefBitmapMemoryCopied As Boolean)

        'Load bitmap into memory, see if file was loaded and if haven't already loaded into memory
        If btmByRefBitmapFile IsNot Nothing And Not blnByRefBitmapMemoryCopied Then

            'Declare
            Dim intWidthFile As Integer = btmByRefBitmapFile.Width
            Dim intHeightFile As Integer = btmByRefBitmapFile.Height

            'Make new memory picture
            btmByRefBitmapMemory = New Bitmap(intWidthFile, intHeightFile, Imaging.PixelFormat.Format32bppPArgb)

            'Memory copy
            DrawGraphicsByPoint(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)

            'Dispose file lock
            btmByRefBitmapFile.Dispose()

            'Set
            blnByRefBitmapMemoryCopied = True

            'Restart loading game thread
            RestartLoadingGame(True)

        Else

            'Check to see if it was already copied
            If blnByRefBitmapMemoryCopied Then

                'Restart loading game thread
                RestartLoadingGame(True)

            End If

        End If

    End Sub

    Private Sub SetLoadingBarPercentage(intMemoryLoadPositionToCheck As Integer, intPictureMemoryIndex As Integer)

        'Check if need to set loading bar percentage
        If intMemoryLoadPosition = intMemoryLoadPositionToCheck Then
            'Set
            btmLoadingBar = abtmLoadingBarPictureMemories(intPictureMemoryIndex)
        End If

    End Sub

    Private Sub CopyLevelBitmap(ByRef btmByRefBitmapMemory As Bitmap, btmBitmapMemoryToCopy As Bitmap)

        'Dispose old
        If btmByRefBitmapMemory IsNot Nothing Then
            btmByRefBitmapMemory.Dispose()
            btmByRefBitmapMemory = Nothing
        End If

        'Declare
        Dim intWidthToCopy As Integer = btmBitmapMemoryToCopy.Width
        Dim intHeightToCopy As Integer = btmBitmapMemoryToCopy.Height

        'Make new memory picture
        btmByRefBitmapMemory = New Bitmap(intWidthToCopy, intHeightToCopy, Imaging.PixelFormat.Format32bppPArgb)

        'Memory copy
        DrawGraphicsByPoint(Graphics.FromImage(btmByRefBitmapMemory), btmBitmapMemoryToCopy, pntTopLeft)

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
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case 4 'Highscores
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case 5 'Credits
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case 6 'Versus
                VersusMouseClickScreen(pntMouse)

            Case 7 'Loading versus
                LoadingVersusMouseClickScreen(pntMouse)

            Case 8 'Versus game
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case 9 'Story
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

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
                'Set
                btmLoadingParagraph = Nothing
                'Set
                strTypeOfParagraphWait = "Single"
                'Set
                intParagraphWaitMode = 0
                'Use the paragraph timer
                tmrParagraph.Enabled = True
                'Set
                blnGameIsVersus = False
                'Load game
                LoadBeginningGameMaterial()

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
                'Set
                intCreditsWaitMode = 0
                'Enable timer
                tmrCredits.Enabled = True

            Case blnMouseInRegion(pntMouse, 256, 74, pntVersusText) 'Versus button was clicked
                'Set
                gblnSendingDataCleared = True
                gblnReceivingDataCleared = True
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

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(intCanvasModeToSet, intCanvasShowToSet)

        'Stop sound
        audcAmbianceSound(0).StopSound()

        'Reset fog x positions
        ResetFogXPositions()

    End Sub

    Private Sub OptionsMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Menu sound
                audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
                'Reset fog x positions
                ResetFogXPositions()
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
        If blnFinishedLoading And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
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
            'Play background sound music
            audcGameBackgroundSounds(intLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)
        Else
            'Back was clicked
            GeneralBackButtonClick(pntMouse, True)
        End If

    End Sub

    Private Sub GeneralBackButtonClick(pntMouse As Point, blnPlayPressedSoundNowToBe As Boolean, Optional blnForceExecution As Boolean = False)

        'Back was clicked
        If blnMouseInRegion(pntMouse, 190, 74, pntBackText) Or blnForceExecution Then
            'Set
            blnPlayPressedSoundNow = blnPlayPressedSoundNowToBe
            'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            blnBackFromGame = True
        End If

    End Sub

    Private Sub VersusMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Check nickname
                DefaultNickName()
                'Go back to menu
                GeneralBackButtonClick(New Point(-1, -1), True, True) 'Point doesn't matter here, forcing back button activity

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
                                DefaultNickName()
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
                                DefaultNickName()
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
            If Not blnWaiting And blnFinishedLoading And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                'Send playing
                gSendData(2, strNickName)
                'Start game locally
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(8, 0, False)
                'Started versus game, start objects
                StartVersusGameObjects()
            Else
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)
            End If
        Else
            'Start was clicked
            If Not blnWaiting And blnFinishedLoading And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                'Send ready to play as joiner
                gSendData(2, strNickName)
                'Set
                blnWaiting = True
            Else
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)
            End If
        End If

    End Sub

    Private Sub Path1ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case blnMouseInRegion(pntMouse, 389, 329, New Point(230, 427)) 'Path 1 choice clicked
                'Setup next level
                PrepareToNextLevel(2)

            Case blnMouseInRegion(pntMouse, 368, 306, New Point(1094, 481)) 'Path 2 choice clicked
                'Setup next level
                PrepareToNextLevel(3)

        End Select

    End Sub

    Private Sub PrepareToNextLevel(intLevelToBe As Integer)

        'Setup next level
        intLevel = intLevelToBe

        'Set
        blnCanLoadLevelWhileRendering = True

    End Sub

    Private Sub Path2ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 190, 74, pntBackText) 'Back button was clicked
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case blnMouseInRegion(pntMouse, 296, 304, New Point(643, 306)) 'Path 1 choice
                'Setup next level
                PrepareToNextLevel(5)

            Case blnMouseInRegion(pntMouse, 297, 318, New Point(1138, 384)) 'Path 2 choice
                'Setup next level
                PrepareToNextLevel(4)

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

        'Play background sound music
        audcGameBackgroundSounds(intLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)

    End Sub

    Private Sub DefaultNickName()

        'Check nick name
        If strNickName = "" Then
            strNickName = "Player"
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
            btmSoundPercent = abtmSoundMemories(0)
        ElseIf pntSlider.X = 657 Then
            gintSoundVolume = 1000
            btmSoundPercent = abtmSoundMemories(100)
        Else
            gintSoundVolume = CInt((pntSlider.X - 58) * (1000 / 600)) '58 to 657 = 600 pixel bounds, inclusive of 58 as a number.
            btmSoundPercent = abtmSoundMemories(CInt(gintSoundVolume / 10))
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

    Private Sub ResetFogXPositions()

        'Reset fog
        intFogBackX = 0
        intFogFrontX = 0

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
                            strKeyPressBuffer &= chrKeyPressed.ToString & "."

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
                            strKeyPressBuffer &= chrKeyPressed.ToString & "."

                    End Select

                Case clsCharacter.eintStatusMode.Run

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 32 ' = spacebar
                            'Check to make sure not maxed on bullets
                            If udcCharacter.BulletsUsed <> 0 Then
                                'Character start to reload
                                udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Reload
                            End If

                        Case 59 ' = ;
                            'Stop immediately
                            udcCharacter.StopCharacterFromRunning = True

                        Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                            'Add to the buffer
                            strKeyPressBuffer &= chrKeyPressed.ToString & "."

                    End Select

            End Select

        End If

    End Sub

    Private Sub VersusKeyPress(chrKeyPressed As Char)

        'Check for nickname press or ip address
        Select Case intCanvasVersusShow

            Case 0 'Nickname
                'Check length
                If Len(strNickName) < 9 Then
                    'Check the keypress
                    Select Case Asc(chrKeyPressed)
                        Case 8 'Backspace
                            'Check if first time key pressing
                            KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickName)
                        Case 32, 65 To 90, 97 To 122 '32 = spacebar, 65 to 90 is upper case A-Z, 97 to 122 is lower case a-z
                            'Check if first time key pressing
                            If blnFirstTimeNicknameTyping Then
                                'Set
                                strNickName = chrKeyPressed.ToString
                                'Set
                                blnFirstTimeNicknameTyping = False
                            Else
                                'Set
                                strNickName &= chrKeyPressed.ToString
                            End If
                    End Select
                Else
                    'Allow backspace
                    If Asc(chrKeyPressed) = 8 Then
                        'Check if first time key pressing
                        KeyPressBackspacing(blnFirstTimeNicknameTyping, strNickName)
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
                        strKeyPressBuffer &= chrKeyPressed.ToString & "."

                End Select

            Case clsCharacter.eintStatusMode.Shoot

                'Check key press
                Select Case Asc(chrKeyPressed)

                    Case 32 ' = spacebar
                        'Character about to reload
                        udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload

                    Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                        'Add to the buffer
                        strKeyPressBuffer &= chrKeyPressed.ToString & "."

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
                gSendData(0, "") 'Waiting for completed connection
                'Wait or else complications for exiting and reconnecting
                System.Threading.Thread.Sleep(1)
            End While
        End If

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

    Private Sub CheckingGameVersion(strGameVersion As String)

        'Check game version
        If GAME_VERSION <> strGameVersion Then
            'Check game for mismatch
            GamesMismatched(strGameVersion)
        Else
            PrepareVersusToPlayAfterMismatchPass()
        End If

    End Sub

    Private Sub PrepareVersusToPlayAfterMismatchPass()

        'Set
        intCanvasMode = 7

        'Set
        btmLoadingParagraphVersus = Nothing

        'Set
        strTypeOfParagraphWait = "Versus"

        'Set
        intParagraphWaitMode = 0

        'Use the paragraph timer
        tmrParagraph.Enabled = True

        'Set
        blnGameIsVersus = True

        'Load game
        LoadBeginningGameMaterial()

    End Sub

    Private Sub Mismatching()

        'Wait 5 seconds
        System.Threading.Thread.Sleep(5000)

        'Disconnect
        GeneralBackButtonClick(New Point(-1, -1), False, True) 'Point doesn't matter here, forcing back button activity

    End Sub

    Private Sub DataArrival(strData As String)

        'Notes: Data with TCP/IP is network streamed and not lined up perfectly, there must be delimiters, especially ending check, buffer is an absolute must

        'Wait here
        While Not gblnReceivingDataCleared
        End While

        'Set
        gblnReceivingDataCleared = False

        'Show partial data
        'Debug.Print(strData)

        'Add to the buffer
        gstrData &= strData

        'Show complete data
        'Debug.Print(gstrData)

        'Check the buffer for ending delimiter character
        If InStr(gstrData, "~") > 0 Then
            'Set the data to be split
            Dim astrData() As String = Split(gstrData, "~")
            'Continue to use the data received
            For intLoop As Integer = 0 To astrData.GetUpperBound(0) - 1
                'Use the data received
                UseTheDataReceived(astrData(intLoop))
            Next
            'Reassemble the data
            gstrData = astrData(astrData.GetUpperBound(0))
        End If

        'Set
        gblnReceivingDataCleared = True

    End Sub

    Private Sub UseTheDataReceived(strData As String)

        'Check if host
        If blnHost Then

            'Check data
            Select Case (strGetBlockData(strData))

                Case "0" 'Completely connected
                    'Check to be connected message only sends once
                    If Not blnConnectionCompleted Then
                        'Set
                        blnConnectionCompleted = True
                        'Send data
                        gSendData(1, GAME_VERSION)
                        'Check game versions
                        CheckGameVerisonData(strData)
                    End If

                Case "1"
                    'Not used

                Case "2" 'Waiting to start game
                    'Set name
                    strNickNameConnected = strGetBlockData(strData, 1)
                    'Ready to play but waiting for host to hit start
                    blnWaiting = False
                    blnReadyEarly = True

                Case "3" 'Not used by host

                Case "4" 'Not used by host

                Case "5" 'Join has shot
                    'Increase
                    intZombieKillsWaitingTwo += 1

                Case "6" 'Joiner is reloading
                    'Show reloading
                    udcCharacterTwo.Reload()

                Case "7" 'Not used by host

                Case "8" 'Not used by host

            End Select

        Else 'Joiner

            'Check data
            Select Case (strGetBlockData(strData))

                Case "0" 'Send data back
                    'Send data once
                    If Not blnConnectionCompleted Then
                        'Set
                        blnConnectionCompleted = True
                        'Send data
                        gSendData(0, GAME_VERSION)
                    End If

                Case "1"
                    'Check game versions
                    CheckGameVerisonData(strData)

                Case "2"
                    'Set name
                    strNickNameConnected = strGetBlockData(strData, 1)
                    'Start was hit, game started for joiner
                    ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(8, 0, False)
                    blnWaiting = False
                    'Started versus game, start objects
                    StartVersusGameObjects()

                Case "3" 'Get x positions of zombies and update them, looks like "3|0:1600,1:1800,|0:1600,1:1800,~", this update happens during each render of hoster
                    'Remove the ending delimiter
                    strData = Replace(strData, "~", "")
                    'Declare
                    Dim astrTemp() As String = Split(strData, "|") '0, 1, 2 elements
                    Dim astrDataZombiesOne() As String = Split(astrTemp(1), ",") '0, 1, 2, 3, 4, 5, 6 elements need to -1 for index count
                    Dim astrDataZombiesTwo() As String = Split(astrTemp(2), ",") '0, 1, 2, 3, 4, 5, 6 elements need to -1 for index count
                    'Loop to change x positions of zombies one
                    ChangeXPositionOfZombies(astrDataZombiesOne, gaudcZombiesOne)
                    'Loop to change x positions of zombies two
                    ChangeXPositionOfZombies(astrDataZombiesTwo, gaudcZombiesTwo)

                Case "4" 'Host has shot
                    'Add to zombie kill buffer
                    strZombieKillBufferOne &= strGetBlockData(strData, 1)

                Case "5" 'Join has shot
                    'Add to zombie kill buffer
                    strZombieKillBufferTwo &= strGetBlockData(strData, 1)

                Case "6" 'Show hoster is reloading
                    'Show reloading
                    udcCharacterOne.Reload()

                Case "7", "8" 'End game for joiner
                    'Set
                    blnPlayerWasPinned = True
                    'Set
                    blnPreventKeyPressEvent = True
                    'Start black screen
                    BlackScreening(BLACK_SCREEN_DEATH_DELAY)
                    'Stop reloading sound
                    udcReloadingSound.StopSound()
                    'Play
                    udcScreamSound.PlaySound(gintSoundVolume)
                    'Stop level music
                    audcGameBackgroundSounds(intLevel - 1).StopSound()
                    'Pause the stop watch
                    swhStopWatch.Stop()
                    'Keep the reload times updated because object will be removed by memory
                    intReloadTimes = udcCharacterTwo.ReloadTimes 'This will always be player two
                    'Check the data
                    If strData = "7|" Then
                        btmVersusWhoWon = btmYouWonMemory 'Host lost
                    Else
                        btmVersusWhoWon = btmYouLostMemory 'Host won
                    End If

            End Select

        End If

    End Sub

    Private Sub ChangeXPositionOfZombies(astrDataZombiesType() As String, audcZombiesType() As clsZombie)

        'Declare
        Dim astrDataOfZombies() As String 'Elements later are 0, 1
        Dim intIndexOfZombies As Integer = 0
        Dim intXPositionOfZombies As Integer = 0

        'Loop to change x positions of zombies one
        For intLoop As Integer = 0 To astrDataZombiesType.GetUpperBound(0) - 1
            'Set temporary data
            astrDataOfZombies = Split(astrDataZombiesType(intLoop), ":") '0 = index of zombie, 1 = x position
            'Set index
            intIndexOfZombies = CInt(astrDataOfZombies(0))
            'Set x position
            intXPositionOfZombies = CInt(astrDataOfZombies(1))
            'Set x position of zombie
            If audcZombiesType IsNot Nothing Then
                If audcZombiesType(intIndexOfZombies) IsNot Nothing Then
                    audcZombiesType(intIndexOfZombies).ZombiePoint = New Point(intXPositionOfZombies,
                                                                          audcZombiesType(intIndexOfZombies).ZombiePoint.Y)
                End If
            End If
        Next

    End Sub

    Private Sub CheckGameVerisonData(strData As String)

        'Check game versions
        If Not blnCheckedGameMismatch Then
            'Set
            blnCheckedGameMismatch = True
            'Check game version
            CheckingGameVersion(strGetBlockData(strData, 1))
        End If

    End Sub

    Private Function strGetBlockData(strData As String, Optional intArrayElement As Integer = 0) As String

        'Notes: Data looks like "X|~" or "X|String~" which equals "" or "String"

        'Remove ending delimiter
        strData = strData.Replace("~", "")

        'Declare data to be split
        Dim astrTemp() As String = Split(strData, "|")

        'Return
        Return astrTemp(intArrayElement)

    End Function

    Private Sub ConnectionLost()

        'Go back to menu
        GeneralBackButtonClick(New Point(-1, -1), False, True) 'Point doesn't matter here, forcing back button activity

    End Sub

End Class