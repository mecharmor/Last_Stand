'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const strGAME_VERSION As String = "1.55"
    Private Const intORIGINAL_SCREEN_WIDTH As Integer = 840
    Private Const intORIGINAL_SCREEN_HEIGHT As Integer = 525
    Private Const intWINDOW_MESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const intCONTROL_MOVE As Integer = 61456
    Private Const intWINDOW_MESSAGE_CLICK_BUTTON_DOWN As Integer = 161
    Private Const intWINDOW_CAPTION As Integer = 2
    Private Const intWINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3
    Private Const dblONE_SECOND_DELAY As Double = 1000 'This is one second delay in milliseconds
    Private Const dblMOUSE_OVER_SOUND_DELAY As Double = 35 '35 milliseconds
    Private Const dblLOADING_TRANSPARENCY_DELAY As Double = 500
    Private Const dblSTORY_SOUND_DELAY As Double = 1000 'Story delay before playing voice
    Private Const dblSTORY_PARAGRAPH2_DELAY As Double = 49000
    Private Const dblSTORY_PARAGRAPH3_DELAY As Double = 64000
    Private Const dblCREDITS_TRANSPARENCY_DELAY As Double = 500
    Private Const dblGAME_MISMATCH_DELAY As Double = 5000 'Delay to show the game mismatch screen
    Private Const intWIDTH_SUBTRACTION As Integer = 16 'Probably the edges of the window
    Private Const intHEIGHT_SUBTRACTION As Integer = 38 'Probably the edges of the window
    Private Const intSLIDER_WIDTH As Integer = 27
    Private Const strDRAW_STATS_SECONDS_WORD As String = " Sec"
    Private Const intFOG_BACK_MEMORY_WIDTH As Integer = 1648 'Picture actual width
    Private Const intFOG_BACK_MEMORY_HEIGHT As Integer = 300 'Picture actual height
    Private Const intFOG_FRONT_MEMORY_WIDTH As Integer = 1660 'Picture actual width
    Private Const intFOG_FRONT_MEMORY_HEIGHT As Integer = 300 'Picture actual height
    Private Const dblFOG_FRAME_WAIT_TO_START As Double = 250 'Delay for fog before showing it
    Private Const dblFOG_FRAME_DELAY As Double = 35
    Private Const intFOG_SPEED As Integer = 3
    Private Const intFOG_BACK_DISTANCE_Y As Integer = 280 'Extra added distance to shift fog down
    Private Const intFOG_FRONT_DISTANCE_Y As Integer = 340 'Extra added distance to shift fog down
    Private Const intFOG_BACK_ADJUSTED_HEIGHT As Integer = intFOG_BACK_MEMORY_HEIGHT + (intORIGINAL_SCREEN_HEIGHT -
                                                           (intFOG_BACK_DISTANCE_Y + intFOG_BACK_MEMORY_HEIGHT)) 'This is the bottom cut off
    Private Const intFOG_FRONT_ADJUSTED_HEIGHT As Integer = intFOG_FRONT_MEMORY_HEIGHT + (intORIGINAL_SCREEN_HEIGHT -
                                                            (intFOG_FRONT_DISTANCE_Y + intFOG_FRONT_MEMORY_HEIGHT)) 'Bottom cut off
    Private Const intNUMBER_OF_ZOMBIES_CREATED As Integer = 150
    Private Const intNUMBER_OF_ZOMBIES_AT_ONE_TIME As Integer = 5
    Private Const intZOMBIE_PINNING_X_DISTANCE As Integer = 100
    Private Const intJOINER_ADDED_X_DISTANCE As Integer = 50
    Private Const intBLACK_SCREEN_BEAT_LEVEL_DELAY As Integer = 350
    Private Const intBLACK_SCREEN_DEATH_DELAY As Integer = 750

    'Declare beginning necessary engine needs
    Private intScreenWidth As Integer = 800
    Private intScreenHeight As Integer = 600
    Private thrRendering As System.Threading.Thread
    Private blnRendering As Boolean = False
    Private blnThreadSupported As Boolean = False
    Private rectFullScreen As Rectangle
    Private btmCanvas As New Bitmap(intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT, Imaging.PixelFormat.Format32bppPArgb) 'Screen size here
    Private pntTopLeft As New Point(0, 0)
    Private intCanvasMode As Integer = 0 'Default menu screen
    Private intCanvasShow As Integer = 0 'Default, no animation
    Private strDirectory As String = AppDomain.CurrentDomain.BaseDirectory
    Private blnScreenChanged As Boolean = False

    'Menu necessary needs
    Private btmMenuBackgroundFile As Bitmap
    Private btmMenuBackgroundMemory As Bitmap
    Private btmFogBackFile As Bitmap
    Private btmFogBackMemory As Bitmap
    Private tmrFog As New System.Timers.Timer
    Private blnProcessBackFog As Boolean = False
    Private blnProcessFrontFog As Boolean = False
    Private intFogBackRectangleMove As Integer = 0
    Private intFogBackX As Integer = intORIGINAL_SCREEN_WIDTH
    Private btmFogBackCloneScreenShown As Bitmap
    Private btmFogFrontFile As Bitmap
    Private btmFogFrontMemory As Bitmap
    Private btmFogFrontCloneScreenShown As Bitmap
    Private intFogFrontRectangleMove As Integer = 0
    Private intFogFrontX As Integer = intORIGINAL_SCREEN_WIDTH
    Private btmArcherFile As Bitmap
    Private btmArcherMemory As Bitmap
    Private pntArcher As New Point(58, 0)
    Private btmLastStandTextFile As Bitmap
    Private btmLastStandTextMemory As Bitmap
    Private pntLastStandText As New Point(73, 416)
    Private btmStartTextFile As Bitmap
    Private btmStartTextMemory As Bitmap
    Private pntStartText As New Point(540, 15)
    Private btmStartHoverTextFile As Bitmap
    Private btmStartHoverTextMemory As Bitmap
    Private pntStartHoverText As New Point(529, 12)
    Private btmHighscoresTextFile As Bitmap
    Private btmHighscoresTextMemory As Bitmap
    Private pntHighscoresText As New Point(599, 70)
    Private btmHighscoresHoverTextFile As Bitmap
    Private btmHighscoresHoverTextMemory As Bitmap
    Private pntHighscoresHoverText As New Point(578, 65)
    Private btmStoryTextFile As Bitmap
    Private btmStoryTextMemory As Bitmap
    Private pntStoryText As New Point(623, 141)
    Private btmStoryHoverTextFile As Bitmap
    Private btmStoryHoverTextMemory As Bitmap
    Private pntStoryHoverText As New Point(611, 137)
    Private btmOptionsTextFile As Bitmap
    Private btmOptionsTextMemory As Bitmap
    Private pntOptionsText As New Point(603, 205)
    Private btmOptionsHoverTextFile As Bitmap
    Private btmOptionsHoverTextMemory As Bitmap
    Private pntOptionsHoverText As New Point(587, 200)
    Private btmCreditsTextFile As Bitmap
    Private btmCreditsTextMemory As Bitmap
    Private pntCreditsText As New Point(676, 268)
    Private btmCreditsHoverTextFile As Bitmap
    Private btmCreditsHoverTextMemory As Bitmap
    Private pntCreditsHoverText As New Point(661, 263)
    Private btmVersusTextFile As Bitmap
    Private btmVersusTextMemory As Bitmap
    Private pntVersusText As New Point(142, 35)
    Private btmVersusHoverTextFile As Bitmap
    Private btmVersusHoverTextMemory As Bitmap
    Private pntVersusHoverText As New Point(128, 31)
    Private btmBackTextFile As Bitmap
    Private btmBackTextMemory As Bitmap
    Private pntBackText As New Point(719, 23)
    Private btmBackHoverTextFile As Bitmap
    Private btmBackHoverTextMemory As Bitmap
    Private pntBackHoverText As New Point(709, 17)

    'Options screen
    Private btmOptionsBackgroundFile As Bitmap
    Private btmOptionsBackgroundMemory As Bitmap
    Private btmResolutionTextFile As Bitmap
    Private btmResolutionTextMemory As Bitmap
    Private pntResolutionText As New Point(20, 20)
    Private btm800x600TextFile As Bitmap
    Private btm800x600TextMemory As Bitmap
    Private btmNot800x600TextFile As Bitmap
    Private btmNot800x600TextMemory As Bitmap
    Private pnt800x600Text As New Point(42, 71)
    Private btm1024x768TextFile As Bitmap
    Private btm1024x768TextMemory As Bitmap
    Private btmNot1024x768TextFile As Bitmap
    Private btmNot1024x768TextMemory As Bitmap
    Private pnt1024x768Text As New Point(42, 96)
    Private btm1280x800TextFile As Bitmap
    Private btm1280x800TextMemory As Bitmap
    Private btmNot1280x800TextFile As Bitmap
    Private btmNot1280x800TextMemory As Bitmap
    Private pnt1280x800Text As New Point(42, 121)
    Private btm1280x1024TextFile As Bitmap
    Private btm1280x1024TextMemory As Bitmap
    Private btmNot1280x1024TextFile As Bitmap
    Private btmNot1280x1024TextMemory As Bitmap
    Private pnt1280x1024Text As New Point(42, 146)
    Private btm1440x900TextFile As Bitmap
    Private btm1440x900TextMemory As Bitmap
    Private btmNot1440x900TextFile As Bitmap
    Private btmNot1440x900TextMemory As Bitmap
    Private pnt1440x900Text As New Point(42, 171)
    Private btmFullScreenTextFile As Bitmap
    Private btmFullScreenTextMemory As Bitmap
    Private btmNotFullScreenTextFile As Bitmap
    Private btmNotFullScreenTextMemory As Bitmap
    Private pntFullscreenText As New Point(42, 195)
    Private btmSoundTextFile As Bitmap
    Private btmSoundTextMemory As Bitmap
    Private pntSoundText As New Point(20, 223)
    Private intResolutionMode As Integer = 0 'Default 800x600
    Private btmSoundBarFile As Bitmap
    Private btmSoundBarMemory As Bitmap
    Private pntSoundBar As New Point(42, 273)
    Private btmSliderFile As Bitmap
    Private btmSliderMemory As Bitmap
    Private pntSlider As New Point(329, 266) '100% mark
    Private blnSliderWithMouseDown As Boolean = False
    Private btmSoundPercent As Bitmap
    Private abtmSoundFiles(100) As Bitmap '0 to 100
    Private abtmSoundMemories(100) As Bitmap '0 to 100
    Private pntSoundPercent As New Point(359, 276)

    'Loading screen
    Private btmLoadingBackgroundFile As Bitmap
    Private btmLoadingBackgroundMemory As Bitmap
    Private abtmLoadingBarPictureFiles(10) As Bitmap '0 To 10 = 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
    Private abtmLoadingBarPictureMemories(10) As Bitmap
    Private btmLoadingBar As Bitmap
    Private pntLoadingBar As New Point(16, 441)
    Private btmLoadingTextFile As Bitmap
    Private btmLoadingTextMemory As Bitmap
    Private pntLoadingText As New Point(297, 449)
    Private btmLoadingStartTextFile As Bitmap
    Private btmLoadingStartTextMemory As Bitmap
    Private pntLoadingStartText As New Point(336, 454)
    Private abtmLoadingParagraphFiles(3) As Bitmap
    Private abtmLoadingParagraphMemories(3) As Bitmap
    Private btmLoadingParagraph As Bitmap
    Private pntLoadingParagraph As New Point(61, 130)
    Private tmrParagraph As New System.Timers.Timer 'Used for paragraph delay
    Private strTypeOfParagraphWait As String = ""
    Private intParagraphWaitMode As Integer = 0
    Private thrLoadBeginningGameMaterial As System.Threading.Thread
    Private thrLoadingGame As System.Threading.Thread
    Private intMemoryLoadPosition As Integer = 0
    Private blnFinishedLoading As Boolean = False

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
    Private pntJohnGonzales As New Point(100, 75)
    Private abtmZacharyStaffordFiles(3) As Bitmap
    Private abtmZacharyStaffordMemories(3) As Bitmap
    Private btmZacharyStafford As Bitmap
    Private pntZacharyStafford As New Point(470, 75)
    Private abtmCoryLewisFiles(3) As Bitmap
    Private abtmCoryLewisMemories(3) As Bitmap
    Private btmCoryLewis As Bitmap
    Private pntCoryLewis As New Point(285, 287)
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
    Private pntVersusBlackOutline As New Point(50, 75)
    Private btmVersusHostTextFile As Bitmap
    Private btmVersusHostTextMemory As Bitmap
    Private pntVersusHost As New Point(133, 302)
    Private btmVersusOrTextFile As Bitmap
    Private btmVersusOrTextMemory As Bitmap
    Private pntVersusOr As New Point(415, 324)
    Private btmVersusJoinTextFile As Bitmap
    Private btmVersusJoinTextMemory As Bitmap
    Private pntVersusJoin As New Point(475, 300)
    Private btmVersusIPAddressTextFile As Bitmap
    Private btmVersusIPAddressTextMemory As Bitmap
    Private btmVersusConnectTextFile As Bitmap
    Private btmVersusConnectTextMemory As Bitmap
    Private pntVersusConnect As New Point(267, 307)
    Private blnPlayPressedSoundNow As Boolean = False

    'Loading versus game variables
    Private blnGameIsVersus As Boolean = False
    Private abtmLoadingParagraphVersusFiles(3) As Bitmap
    Private abtmLoadingParagraphVersusMemories(3) As Bitmap
    Private btmLoadingParagraphVersus As Bitmap
    Private btmLoadingWaitingTextFile As Bitmap
    Private btmLoadingWaitingTextMemory As Bitmap
    Private pntLoadingWaitingText As New Point(297, 449)

    'Story
    Private btmStoryBackgroundFile As Bitmap
    Private btmStoryBackgroundMemory As Bitmap
    Private abtmStoryParagraphFiles(11) As Bitmap
    Private abtmStoryParagraphMemories(11) As Bitmap
    Private btmStoryParagraph As Bitmap
    Private pntStoryParagraph As Point 'Set in the story thread
    Private tmrStory As New System.Timers.Timer 'Used for story delay
    Private intStoryWaitMode As Integer = 0

    'Game version mismatch
    Private btmGameMismatchBackgroundFile As Bitmap
    Private btmGameMismatchBackgroundMemory As Bitmap
    Private strGameVersionFromConnection As String = ""
    Private tmrGameMismatch As New System.Timers.Timer 'Used for mismatch screen

    'FPS reading
    Private tmrFPS As New System.Timers.Timer
    Private intFPSDisplay As Integer = 0
    Private intFPSCalculated As Integer = 0

    'Options mouse over
    Private tmrOptionsMouseOver As New System.Timers.Timer
    Private strOptionsButtonSpot As String = ""

    'Sounds
    Private audcAmbianceSound(1) As clsSound '2 ambiance sounds so far
    Private udcButtonHoverSound As clsSound
    Private udcButtonPressedSound As clsSound
    Private audcStoryParagraphSounds(2) As clsSound '3 paragraph sounds in the story area
    Private udcFinishedLoading100PercentSound As clsSound 'When game = 100% loaded
    Private udcGameLoadedPressedSound As clsSound 'When game is loaded and pressing start
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
    Private audcSmallChainGagSounds(1) As clsSound 'Chained zombie gags
    Private udcWaterSplashSound As clsSound 'Happens after zombie hits water
    Private udcFaceZombieEyesOpenSound As clsSound 'Happens when the zombie face opens eyes

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
    Private pntWordBar As New Point(241, 13)
    Private udcCharacter As clsCharacter
    Private blnBackFromGame As Boolean = False
    Private btmAK47MagazineFile As Bitmap
    Private btmAK47MagazineMemory As Bitmap
    Private blnAK47MagazineMemoryCopied As Boolean
    Private pntAK47Magazine As New Point(29, 438)
    Private btmWinOverlayFile As Bitmap
    Private btmWinOverlayMemory As Bitmap
    Private blnWinOverlayMemoryCopied As Boolean
    Private blnBeatLevel As Boolean = False

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

    'Words
    Private astrWords(298) As String 'Used to fill with words
    Private intWordIndex As Integer = 0
    Private strTheWord As String = ""
    Private strWord As String = ""
    Private strNextWord As String = ""

    'Key press
    Private strKeyPressBuffer As String = ""
    Private blnPreventKeyPressEvent As Boolean = False

    'Highscores for the game
    Private intZombieKills As Integer = 0
    Private intZombieKillsCombined As Integer = 0
    Private intReloadTimes As Integer = 0
    Private intRunTimes As Integer = 0
    Private tsTimeSpan As TimeSpan
    Private strElapsedTime As String = ""
    Private intElapsedTime As Integer = 0
    Private intWPM As Integer = 0
    Private blnSetStats As Boolean = False
    Private blnComparedHighscore As Boolean = False

    'Stop watch
    Private swhStopWatch As Stopwatch

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

    'Chained zombie bitmap load files
    Private abtmChainedZombieFiles(2) As Bitmap
    Private ablnChainedZombieMemoriesCopied(2) As Boolean

    'Helicopter bitmap load files
    Private abtmHelicopterFiles(4) As Bitmap
    Private ablnHelicopterMemoriesCopied(4) As Boolean

    'Zombie face
    Private abtmFaceZombieFiles(1) As Bitmap
    Private ablnFaceZombieMemoriesCopied(1) As Boolean
    Private blnOpenedEyes As Boolean = False

    'Pipe
    Private btmPipeValveFile As Bitmap
    Private btmPipeValveMemory As Bitmap
    Private blnPipeValveMemoryCopied As Boolean

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

    'Special effects
    Private btmGameBackground2WaterFile As Bitmap
    Private btmGameBackground2WaterMemory As Bitmap
    Private blnGameBackground2WaterMemoryCopied As Boolean = False
    Private intWaterHeight As Integer = 0 'Set later after memory load
    Private pntGameBackground2Water As Point

    'Helicopter
    Private udcHelicopterGonzales As clsHelicopter

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes: Do not allow a resize of the window if full screen, this happens if you double click the title, but is prevented with this sub.

        'Check if full screen
        If intResolutionMode = 5 Then
            'Prevent moving the form by control box click
            If m.Msg = intWINDOW_MESSAGE_SYSTEM_COMMAND And m.WParam.ToInt32() = intCONTROL_MOVE Then
                Return
            End If
            'Prevent button down moving form
            If m.Msg = intWINDOW_MESSAGE_CLICK_BUTTON_DOWN And m.WParam.ToInt32() = intWINDOW_CAPTION Then
                Return
            End If
        End If

        'If a double click on the title bar is triggered
        If m.Msg = intWINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
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

        '0% loaded
        btmLoadingBar = abtmLoadingBarPictureMemories(0)

        'Set 100%
        btmSoundPercent = abtmSoundMemories(100)

        'Setup frames per second timer
        SetupTimer(tmrFPS, AddressOf ElapsedFPS, dblONE_SECOND_DELAY)

        'Setup mouse over timer
        SetupTimer(tmrOptionsMouseOver, AddressOf ElapsedOptionsMouseOver, dblMOUSE_OVER_SOUND_DELAY)

        'Setup paragraph timer
        SetupTimer(tmrParagraph, AddressOf ElapsedParagraph, dblLOADING_TRANSPARENCY_DELAY)

        'Setup story timer
        SetupTimer(tmrStory, AddressOf ElapsedStory, dblLOADING_TRANSPARENCY_DELAY)

        'Setup credits timer
        SetupTimer(tmrCredits, AddressOf ElapsedCredits, dblCREDITS_TRANSPARENCY_DELAY)

        'Setup black screen timer
        SetupTimer(tmrBlackScreen, AddressOf ElapsedBlackScreen, dblONE_SECOND_DELAY)

        'Setup mismatch timer
        SetupTimer(tmrGameMismatch, AddressOf ElapsedGameMismatch, dblGAME_MISMATCH_DELAY)

        'Setup fog timer
        SetupTimer(tmrFog, AddressOf ElapsedFog, dblFOG_FRAME_WAIT_TO_START)

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
        LoadPictureFileAndCopyBitmapIntoMemory(btmBackTextFile, btmBackTextMemory, "Images\Menu\Back.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmBackHoverTextFile, btmBackHoverTextMemory, "Images\Menu\BackHover.png")

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
        For intLoop As Integer = 0 To abtmSoundFiles.GetUpperBound(0)
            LoadPictureFileAndCopyBitmapIntoMemory(abtmSoundFiles(intLoop), abtmSoundMemories(intLoop), "Images\Options\Sound" & CStr(intLoop) & ".png")
        Next

    End Sub

    Private Sub LoadPictureFileAndCopyBitmapIntoMemory(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap, strImageDirectory As String)

        'Load picture file bitmap
        btmByRefBitmapFile = New Bitmap(Image.FromFile(strDirectory & strImageDirectory))

        'Prepare to create new memory blank bitmap
        If IO.File.Exists(strDirectory & strImageDirectory) Then
            'Attempt to load file
            Try
                'Create new blank image
                btmByRefBitmapMemory = New Bitmap(btmByRefBitmapFile.Width, btmByRefBitmapFile.Height, Imaging.PixelFormat.Format32bppPArgb)
            Catch ex As Exception
                'Display and close
                CloseApplicationWithErrorMessage("Virtual memory is full or set very low. Please increase memory size. This application will close now.")
            End Try
        Else
            'Display and close
            CloseApplicationWithErrorMessage("The " & strDirectory & strImageDirectory & " file is missing. This application will close now.")
        End If

        'Draw the file bitmap data to memory blank sheet, this also changes the pixelformat making rendering much faster with the memory
        DrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)

        'Dispose of the file as memory will only be used now
        btmByRefBitmapFile.Dispose()
        btmByRefBitmapFile = Nothing

    End Sub

    Private Sub CloseApplicationWithErrorMessage(strErrorMessage As String)

        'Display
        MessageBox.Show(strErrorMessage, "Last Stand", MessageBoxButtons.OK,
                        MessageBoxIcon.Error)

        'Exit
        Me.Close()

    End Sub

    Private Sub DrawGraphics(ByRef gByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, pntPoint As Point)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn

        'Set options for fastest rendering
        SetGraphicOptions(gGraphics)

        'Draw
        gGraphics.DrawImageUnscaled(btmBitmapToDraw, pntPoint)

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

    Private Sub DrawGraphics(ByRef gByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, rectRectangle As Rectangle)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn

        'Set options for fastest rendering
        SetGraphicOptions(gGraphics)

        'Draw
        gGraphics.DrawImage(btmBitmapToDraw, rectRectangle) 'Scales it down by the rectangle

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

    Private Sub SetupTimer(tmrToSetup As System.Timers.Timer, steEventHandler As System.Timers.ElapsedEventHandler, dblInterval As Double)

        'Disable timer by default
        tmrToSetup.Enabled = False 'Enabled later

        'Set timer
        tmrToSetup.AutoReset = True

        'Add handler
        AddHandler tmrToSetup.Elapsed, steEventHandler

        'Set timer delay
        tmrToSetup.Interval = dblInterval

    End Sub

    Private Sub ElapsedFPS(sender As Object, e As EventArgs)

        'Set frames to be stored
        intFPSDisplay = intFPSCalculated

        'Reset
        intFPSCalculated = 0

    End Sub

    Private Sub ElapsedOptionsMouseOver(sender As Object, e As EventArgs)

        'Disable timer
        tmrOptionsMouseOver.Enabled = False

        'Play sound
        udcButtonHoverSound.PlaySound(gintSoundVolume)

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

    Private Sub ElapsedFog(sender As Object, e As EventArgs)

        'Check to change interval
        If tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START Then
            'Change interval
            tmrFog.Interval = dblFOG_FRAME_DELAY
            'Reset fog variables
            ResetFogVariables(intFogBackRectangleMove, intFogBackX, blnProcessBackFog)
            ResetFogVariables(intFogFrontRectangleMove, intFogFrontX, blnProcessFrontFog)
        End If

        'Increase fog back
        IncreaseFogVariables(intFogBackRectangleMove, intFOG_BACK_MEMORY_WIDTH, intFogBackX, blnProcessBackFog)

        'Increase fog front
        IncreaseFogVariables(intFogFrontRectangleMove, intFOG_FRONT_MEMORY_WIDTH, intFogFrontX, blnProcessFrontFog)

        'Check to reset
        If Not blnProcessBackFog And Not blnProcessFrontFog Then
            tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START
        End If

    End Sub

    Private Sub ResetFogVariables(ByRef intByRefFogRectangleMove As Integer, ByRef intByRefFogX As Integer, ByRef blnByRefProcessFog As Boolean)

        'Set
        intByRefFogRectangleMove = 0

        'Set
        intByRefFogX = intORIGINAL_SCREEN_WIDTH

        'Set
        blnByRefProcessFog = True

    End Sub

    Private Sub IncreaseFogVariables(ByRef intByRefBackRectangleMove As Integer, intFogMemoryWidth As Integer,
                                     ByRef intByRefFogX As Integer, ByRef blnByRefProcessFog As Boolean)

        'Increase
        If intByRefBackRectangleMove > intORIGINAL_SCREEN_WIDTH Then
            'Check for last end of picture on the left side
            If intFogMemoryWidth - intByRefBackRectangleMove < 0 Then
                'Decrease
                intByRefFogX -= intFOG_SPEED
                'Check
                If intORIGINAL_SCREEN_WIDTH - intByRefFogX >= intORIGINAL_SCREEN_WIDTH Then
                    'Disable
                    blnByRefProcessFog = False
                End If
            Else
                'Increase
                intByRefBackRectangleMove += intFOG_SPEED
            End If
        Else
            'Increase
            intByRefBackRectangleMove += intFOG_SPEED
        End If

    End Sub

    Private Sub ElapsedStory(sender As Object, e As EventArgs)

        'Check which mode
        Select Case intStoryWaitMode
            Case 0
                'Set
                SetStoryPointerAndBitmap(41, 114, 0)
            Case 1
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(1)
            Case 2
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(2)
            Case 3
                'Set
                ChangeStoryParagraphBitmapAndInterval(3, dblSTORY_SOUND_DELAY)
            Case 4
                'Play story paragraph 1 sound
                PlayStoryParagraphSoundAndChangeInterval(0, dblSTORY_PARAGRAPH2_DELAY)
            Case 5
                'Set
                ClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
            Case 6
                'Set
                SetStoryPointerAndBitmap(19, 84, 4)
            Case 7
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(5)
            Case 8
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(6)
            Case 9
                'Set
                ChangeStoryParagraphBitmapAndInterval(7, dblSTORY_SOUND_DELAY)
            Case 10
                'Play story paragraph 2 sound
                PlayStoryParagraphSoundAndChangeInterval(1, dblSTORY_PARAGRAPH3_DELAY)
            Case 11
                'Set
                ClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
            Case 12
                'Set
                SetStoryPointerAndBitmap(15, 90, 8)
            Case 13
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(9)
            Case 14
                'Set
                btmStoryParagraph = abtmStoryParagraphMemories(10)
            Case 15
                'Set
                ChangeStoryParagraphBitmapAndInterval(11, dblSTORY_SOUND_DELAY)
            Case 16
                'Play story paragraph 3 sound
                audcStoryParagraphSounds(2).PlaySound(gintSoundVolume)
                'Disable timer
                tmrStory.Enabled = False
        End Select

        'Increase
        intStoryWaitMode += 1

    End Sub

    Private Sub SetStoryPointerAndBitmap(intX As Integer, intY As Integer, intIndex As Integer)

        'Set
        pntStoryParagraph = New Point(intX, intY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(intIndex)

    End Sub

    Private Sub ChangeStoryParagraphBitmapAndInterval(intIndex As Integer, dblInterval As Double)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(intIndex)

        'Set
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub PlayStoryParagraphSoundAndChangeInterval(intIndex As Integer, dblInterval As Double)

        'Play story paragraph sound
        audcStoryParagraphSounds(intIndex).PlaySound(gintSoundVolume)

        'Change interval
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub ClearStoryParagraphBitmapAndChangeInterval(dblInterval As Double)

        'Set
        btmStoryParagraph = Nothing

        'Set
        tmrStory.Interval = dblInterval

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

    Private Sub ElapsedBlackScreen(sender As Object, e As EventArgs)

        'Increase
        intBlackScreenWaitMode += 1

        'Disable
        If intBlackScreenWaitMode = 2 Then
            tmrBlackScreen.Enabled = False
        End If

    End Sub

    Private Sub ElapsedGameMismatch(sender As Object, e As EventArgs)

        'Disconnect
        GeneralBackButtonClick(New Point(-1, -1), False, True) 'Point doesn't matter here, forcing back button activity

    End Sub

    Private Sub GeneralBackButtonClick(pntMouse As Point, blnPlayPressedSoundNowToBe As Boolean, Optional blnForceExecution As Boolean = False)

        'Back was clicked
        If blnMouseInRegion(pntMouse, 95, 37, pntBackText) Or blnForceExecution Then
            'Set
            blnPlayPressedSoundNow = blnPlayPressedSoundNowToBe
            'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set
            blnBackFromGame = True
        End If

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Set
        blnRendering = False

        'Wait
        While thrRendering.IsAlive
            System.Threading.Thread.Sleep(1)
        End While

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

        'Remove game objects
        RemoveGameObjectsFromMemory()

        'Empty versus multiplayer variables
        EmptyMultiplayerVariables()

        'Remove FPS timer
        RemoveTimer(tmrFPS, AddressOf ElapsedFPS)

        'Remove options mouse over timer
        RemoveTimer(tmrOptionsMouseOver, AddressOf ElapsedOptionsMouseOver)

        'Remove paragraph timer
        RemoveTimer(tmrParagraph, AddressOf ElapsedParagraph)

        'Remove story timer
        RemoveTimer(tmrStory, AddressOf ElapsedStory)

        'Remove credits timer
        RemoveTimer(tmrCredits, AddressOf ElapsedCredits)

        'Remove black screen timer
        RemoveTimer(tmrBlackScreen, AddressOf ElapsedBlackScreen)

        'Remove mismatch timer
        RemoveTimer(tmrGameMismatch, AddressOf ElapsedGameMismatch)

        'Remove fog timer
        RemoveTimer(tmrFog, AddressOf ElapsedFog)

        'Stop all sounds
        StopAndCloseAllSounds()

    End Sub

    Private Sub AbortThread(thrToAbort As System.Threading.Thread)

        'Check thread
        If thrToAbort IsNot Nothing Then
            If thrToAbort.IsAlive Then
                thrToAbort.Abort()
            End If
        End If

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
        If udcHelicopterGonzales IsNot Nothing Then
            'Stop and dispose
            udcHelicopterGonzales.StopAndDispose()
            udcHelicopterGonzales = Nothing
        End If

        'Stop gagging sound
        For intLoop As Integer = 0 To audcSmallChainGagSounds.GetUpperBound(0)
            If audcSmallChainGagSounds(intLoop) IsNot Nothing Then
                audcSmallChainGagSounds(intLoop).StopSound()
            End If
        Next

        'Stop and dispose chained zombie
        If gudcChainedZombie IsNot Nothing Then
            'Stop and dispose
            gudcChainedZombie.StopAndDispose()
            gudcChainedZombie = Nothing
        End If

        'Stop water splash
        If udcWaterSplashSound IsNot Nothing Then
            udcWaterSplashSound.StopSound()
        End If

        'Stop face zombie eyes open sound
        If udcFaceZombieEyesOpenSound IsNot Nothing Then
            udcFaceZombieEyesOpenSound.StopSound()
        End If

        'Dispose face zombie
        If gudcFaceZombie IsNot Nothing Then
            gudcFaceZombie = Nothing
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
            RemoveHandler udcVersusConnectedThread.gGotData, AddressOf DataArrival
            RemoveHandler udcVersusConnectedThread.gConnectionGone, AddressOf ConnectionLost
            'Abort threads
            udcVersusConnectedThread.AbortCheckConnectionThread()
            udcVersusConnectedThread.AbortReadLineThread()
            udcVersusConnectedThread = Nothing
        End If

        'Wait or else a complication will start when reconnecting after an exit of versus game
        System.Threading.Thread.Sleep(1)

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
                        gSendData(1, strGAME_VERSION)
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
                        gSendData(0, strGAME_VERSION)
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
                    BlackScreening(intBLACK_SCREEN_DEATH_DELAY)
                    'Stop reloading sound
                    udcReloadingSound.StopSound()
                    'Play
                    udcScreamSound.PlaySound(gintSoundVolume)
                    'Stop level music
                    audcGameBackgroundSounds(gintLevel - 1).StopSound()
                    'Pause the stop watch
                    swhStopWatch.Stop()
                    'Keep the reload times updated because object will be removed by memory
                    intReloadTimes = udcCharacterTwo.ReloadTimes 'This will always be player two
                    'Keep the running times updated because object will be removed by memory
                    intRunTimes = udcCharacterTwo.RunTimes
                    'Check the data
                    If strData = "7|" Then
                        btmVersusWhoWon = btmYouWonMemory 'Host lost
                    Else
                        btmVersusWhoWon = btmYouLostMemory 'Host won
                    End If

            End Select

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

    Private Sub CheckGameVerisonData(strData As String)

        'Check game versions
        If Not blnCheckedGameMismatch Then
            'Set
            blnCheckedGameMismatch = True
            'Check game version
            CheckingGameVersion(strGetBlockData(strData, 1))
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
                    audcZombiesType(intIndexOfZombies).Point = New Point(intXPositionOfZombies, audcZombiesType(intIndexOfZombies).Point.Y)
                End If
            End If
        Next

    End Sub

    Private Sub ConnectionLost()

        'Go back to menu
        GeneralBackButtonClick(New Point(-1, -1), False, True) 'Point doesn't matter here, forcing back button activity

    End Sub

    Private Sub RemoveTimer(tmrToRemove As System.Timers.Timer, steEventHandler As System.Timers.ElapsedEventHandler)

        'Disable timer
        tmrToRemove.Enabled = False

        'Stop and dispose timer
        tmrToRemove.Stop()
        tmrToRemove.Dispose()

        'Remove handler
        RemoveHandler tmrToRemove.Elapsed, steEventHandler

    End Sub

    Private Sub StopAndCloseAllSounds()

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

        'Loaded game 100%
        If udcFinishedLoading100PercentSound IsNot Nothing Then
            udcFinishedLoading100PercentSound.StopAndCloseSound()
            udcFinishedLoading100PercentSound = Nothing
        End If

        'Loaded game press
        If udcGameLoadedPressedSound IsNot Nothing Then
            udcGameLoadedPressedSound.StopAndCloseSound()
            udcGameLoadedPressedSound = Nothing
        End If

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

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Check if multi-threading was possible
        If Not blnThreadSupported Then
            'Display
            MessageBox.Show("This computer doesn't support multi-threading. This application will close now.", "Last Stand", MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            'Exit
            Me.Close()
        Else
            'Hide the ugly gray screen
            Me.Hide()
            'Force focus
            Me.Focus()
            'Enable FPS timer, this is the only place to enable it first, if enabled sooner, it won't work
            tmrFPS.Enabled = True
            'Load necessary sounds for the beginning engine
            LoadNecessarySoundsForBeginningEngine() 'This needs a handle from the form window
            'Get highscores early because grabbing information from the database access files is slow
            LoadHighscoresStringAccess()
            'Set percentage multiplers for screen modes
            gdblScreenWidthRatio = CDbl((intScreenWidth - intWIDTH_SUBTRACTION) / intORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((intScreenHeight - intHEIGHT_SUBTRACTION) / intORIGINAL_SCREEN_HEIGHT)
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - intWIDTH_SUBTRACTION, intScreenHeight - intHEIGHT_SUBTRACTION) 'Full screen
            'Set
            blnRendering = True
            'Start rendering
            thrRendering.Start()
            'Un-hide
            Me.Show()
            'Enable fog
            tmrFog.Enabled = True
        End If

    End Sub

    Private Sub LoadNecessarySoundsForBeginningEngine()

        'Load ambiance
        For intLoop As Integer = 0 To audcAmbianceSound.GetUpperBound(0)
            audcAmbianceSound(intLoop) = New clsSound(Me, strDirectory & "Sounds\Ambiance" & CStr(intLoop + 1) & ".mp3", 1) 'Repeat only needs 1
        Next

        'Load button pressed
        udcButtonPressedSound = New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 3)

        'Load button hover
        udcButtonHoverSound = New clsSound(Me, strDirectory & "Sounds\ButtonHover.mp3", 5)

        'Load story paragraphs
        For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
            audcStoryParagraphSounds(intLoop) = New clsSound(Me, strDirectory & "Sounds\StoryParagraph" & CStr(intLoop + 1) & ".mp3", 1)
        Next

        'Loaded game 100%
        udcFinishedLoading100PercentSound = New clsSound(Me, strDirectory & "Sounds\FinishedLoading100Percent.mp3", 2)

        'Loaded game press
        udcGameLoadedPressedSound = New clsSound(Me, strDirectory & "Sounds\GameLoadedPressed.mp3", 2)

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
                strHighscores &= dtProperties.Rows(intLoop).Item(0).ToString & ". " & dtProperties.Rows(intLoop).Item(1).ToString & strDRAW_STATS_SECONDS_WORD & " " &
                              dtProperties.Rows(intLoop).Item(2).ToString & " WPM " & dtProperties.Rows(intLoop).Item(3).ToString & " zombie kills - " &
                              dtProperties.Rows(intLoop).Item(4).ToString
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
                If Not blnHighscoresIsAccess Then
                    LoadHighscoresStringTextFile()
                End If
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
        Dim intRank As Integer = 1

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Highscores.txt")

        'Use loop to get scores
        While ioSR.Peek <> -1
            'Set, but add rank in front too
            strHighscores &= CStr(intRank) & ". " & ioSR.ReadLine
            'Check if not last line
            If ioSR.Peek <> -1 Then
                'Include line
                strHighscores &= vbNewLine
                'Increase
                intRank += 1
            End If
        End While

        'Close
        ioSR.Close()

    End Sub

    Private Sub Rendering()

        'Loop
        While blnRendering
            'Empty variables if back from game
            If blnBackFromGame Then
                'Disable timer
                tmrParagraph.Enabled = False
                'Set
                btmLoadingParagraph = Nothing
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
                '0% loaded
                btmLoadingBar = abtmLoadingBarPictureMemories(0)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0, blnPlayPressedSoundNow)
                'Set
                blnPlayPressedSoundNow = False
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
                'Disable story timer
                tmrStory.Enabled = False
                'Set interval
                tmrStory.Interval = dblLOADING_TRANSPARENCY_DELAY
                'Set
                intStoryWaitMode = 0
                'Reset
                btmStoryParagraph = Nothing
                'Disable mismatch timer
                tmrGameMismatch.Enabled = False
                'Stop sounds
                For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
                    audcStoryParagraphSounds(intLoop).StopSound()
                Next
                'Start fog
                tmrFog.Enabled = True
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
            'Paint on screen with rectangle
            DrawGraphics(Me.CreateGraphics(), btmCanvas, rectFullScreen)
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
                Select Case gintLevel
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
                        NextLevel(abtmGameBackgroundMemories(4))
                        'Change level, reuse the mechanics
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(3, 0, False)
                End Select
                'Set
                blnCanLoadLevelWhileRendering = False
            End If
            'Increase frames per second
            intFPSCalculated += 1
            'Do events at the end, is this necessary? These computers are multiple threaded...
            'Application.DoEvents()
        End While

        'Dispose of canvas bitmap
        btmCanvas.Dispose()
        btmCanvas = Nothing

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

    Private Sub RenderMenuScreen()

        'Declare
        Dim rectSourceBack As Rectangle
        Dim rectSourceFront As Rectangle

        'Prepare clone of fog back
        PrepareFog(blnProcessBackFog, intFogBackRectangleMove, intFOG_BACK_MEMORY_WIDTH, rectSourceBack, intFogBackX, intFOG_BACK_ADJUSTED_HEIGHT,
                   btmFogBackCloneScreenShown, btmFogBackMemory)

        'Prepare clone of fog front
        PrepareFog(blnProcessFrontFog, intFogFrontRectangleMove, intFOG_FRONT_MEMORY_WIDTH, rectSourceFront, intFogFrontX, intFOG_FRONT_ADJUSTED_HEIGHT,
                   btmFogFrontCloneScreenShown, btmFogFrontMemory)

        'Paint onto canvas the menu background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmMenuBackgroundMemory, pntTopLeft)

        'Draw fog in back
        DrawFog(btmFogBackCloneScreenShown, intFOG_BACK_MEMORY_WIDTH, intFogBackRectangleMove, intFogBackX, intFOG_BACK_DISTANCE_Y)

        'Draw Archer
        DrawGraphics(Graphics.FromImage(btmCanvas), btmArcherMemory, pntArcher)

        'Draw fog in front
        DrawFog(btmFogFrontCloneScreenShown, intFOG_FRONT_MEMORY_WIDTH, intFogFrontRectangleMove, intFogFrontX, intFOG_FRONT_DISTANCE_Y)

        'Draw start text
        If intCanvasShow = 1 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmStartHoverTextMemory, pntStartHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmStartTextMemory, pntStartText)
        End If

        'Draw highscores text
        If intCanvasShow = 2 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmHighscoresHoverTextMemory, pntHighscoresHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmHighscoresTextMemory, pntHighscoresText)
        End If

        'Draw story text
        If intCanvasShow = 3 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmStoryHoverTextMemory, pntStoryHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmStoryTextMemory, pntStoryText)
        End If

        'Draw options text
        If intCanvasShow = 4 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmOptionsHoverTextMemory, pntOptionsHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmOptionsTextMemory, pntOptionsText)
        End If

        'Draw credits text
        If intCanvasShow = 5 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmCreditsHoverTextMemory, pntCreditsHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmCreditsTextMemory, pntCreditsText)
        End If

        'Draw versus text
        If intCanvasShow = 6 Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusHoverTextMemory, pntVersusHoverText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusTextMemory, pntVersusText)
        End If

        'Draw last stand text
        DrawGraphics(Graphics.FromImage(btmCanvas), btmLastStandTextMemory, pntLastStandText)

    End Sub

    Private Sub PrepareFog(blnProcessFog As Boolean, intFogRectangleMove As Integer, intFogMemoryWidth As Integer, ByRef rectByRefSource As Rectangle, intFogX As Integer,
                           intFogMemoryHeight As Integer, ByRef btmByRefFogCloneScreenShown As Bitmap, btmFogMemory As Bitmap)

        'Prepare clone of fog
        If blnProcessFog Then
            'Check
            If intFogRectangleMove <> 0 Then
                'Make rectangle
                If intFogRectangleMove > intORIGINAL_SCREEN_WIDTH Then
                    'Check for last end of the picture on the left side
                    If intFogMemoryWidth - intFogRectangleMove < 0 Then
                        'Set
                        rectByRefSource = New Rectangle(0, 0, intFogX, intFogMemoryHeight)
                    Else
                        'Set
                        rectByRefSource = New Rectangle(intFogMemoryWidth - intFogRectangleMove, 0, intORIGINAL_SCREEN_WIDTH, intFogMemoryHeight)
                    End If
                Else
                    'Set
                    rectByRefSource = New Rectangle(intFogMemoryWidth - intFogRectangleMove, 0, intFogRectangleMove, intFogMemoryHeight)
                End If
                'Clone
                btmByRefFogCloneScreenShown = btmFogMemory.Clone(rectByRefSource,
                                              Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short
            End If
        End If

    End Sub

    Private Sub DrawFog(ByRef btmByRefFogCloneScreenShown As Bitmap, intFogMemoryWidth As Integer, intFogRectangleMove As Integer, intFogX As Integer,
                        intFogDistanceY As Integer)

        'Draw fog
        If btmByRefFogCloneScreenShown IsNot Nothing Then
            'Draw depending
            If intFogMemoryWidth - intFogRectangleMove < 0 Then
                'Draw
                DrawGraphics(Graphics.FromImage(btmCanvas), btmByRefFogCloneScreenShown, New Point(intORIGINAL_SCREEN_WIDTH - intFogX,
                             intFogDistanceY))
            Else
                'Draw
                DrawGraphics(Graphics.FromImage(btmCanvas), btmByRefFogCloneScreenShown, New Point(0, intFogDistanceY))
            End If
            'Dispose of fog clone
            btmByRefFogCloneScreenShown.Dispose()
            btmByRefFogCloneScreenShown = Nothing
        End If

    End Sub

    Private Sub RenderOptionsScreen()

        'Draw options background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmOptionsBackgroundMemory, pntTopLeft)

        'Draw resolution text
        DrawGraphics(Graphics.FromImage(btmCanvas), btmResolutionTextMemory, pntResolutionText)

        'Draw sound text
        DrawGraphics(Graphics.FromImage(btmCanvas), btmSoundTextMemory, pntSoundText)

        'Check which resolution
        CheckResolutionMode(0, btm800x600TextMemory, btmNot800x600TextMemory, pnt800x600Text)
        CheckResolutionMode(1, btm1024x768TextMemory, btmNot1024x768TextMemory, pnt1024x768Text)
        CheckResolutionMode(2, btm1280x800TextMemory, btmNot1280x800TextMemory, pnt1280x800Text)
        CheckResolutionMode(3, btm1280x1024TextMemory, btmNot1280x1024TextMemory, pnt1280x1024Text)
        CheckResolutionMode(4, btm1440x900TextMemory, btmNot1440x900TextMemory, pnt1440x900Text)
        CheckResolutionMode(5, btmFullScreenTextMemory, btmNotFullScreenTextMemory, pntFullscreenText)

        'Draw sound bar
        DrawGraphics(Graphics.FromImage(btmCanvas), btmSoundBarMemory, pntSoundBar)

        'Draw sound percentage
        DrawGraphics(Graphics.FromImage(btmCanvas), btmSoundPercent, pntSoundPercent)

        'Draw slider
        DrawGraphics(Graphics.FromImage(btmCanvas), btmSliderMemory, pntSlider)

        'Draw version
        DrawTextWithShadows("Version", 25, Color.Red, 15, 475, 500, 125)
        DrawTextWithShadows(strGAME_VERSION, 25, Color.White, 132, 475, 500, 62)

        'Draw frames per second
        DrawTextWithShadows("Frames Per Second", 25, Color.Red, 500, 475, 500, 62)
        DrawTextWithShadows(CStr(intFPSDisplay), 25, Color.White, 772, 475, 500, 62)

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub LoadingGameScreen()

        'Draw loading background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw Loading text
        If blnFinishedLoading Then
            'Start
            DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
        Else
            'Loading
            DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
        End If

        'Draw paragraph
        If btmLoadingParagraph IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingParagraph, pntLoadingParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StartedGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Check if beat level
            If Not blnPlayerWasPinned Then

                'Paint black background
                DrawGraphics(Graphics.FromImage(btmCanvas), abtmBlackScreenMemories(2), pntTopLeft)

                'Check which level was completed
                Select Case gintLevel
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
                        DrawGraphics(Graphics.FromImage(btmCanvas), btmWinOverlayMemory, pntTopLeft)
                        'Draw stats
                        DrawStats(intZombieKillsCombined)
                        'Check if highscores was already beaten
                        If Not blnComparedHighscore Then
                            'Set
                            blnComparedHighscore = True
                            'Compare score with highscores
                            CompareHighscoresAccess()
                        End If
                End Select

            Else

                'Show death screen fading
                DrawGraphics(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)

                'Show death overlay
                DrawGraphics(Graphics.FromImage(btmCanvas), btmDeathOverlayMemory, pntTopLeft)

                'Draw stats
                DrawStats(intZombieKillsCombined)

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

    Private Sub DrawStats(intZombieKillsToBe As Integer)

        'Check if stats has been created
        If Not blnSetStats Then
            'Set
            blnSetStats = True
            'Set timespan
            tsTimeSpan = swhStopWatch.Elapsed
            'Set seconds string
            strElapsedTime = CStr(CInt(tsTimeSpan.TotalSeconds)) & strDRAW_STATS_SECONDS_WORD
            'Get elapsed time
            intElapsedTime = CInt(Replace(strElapsedTime, strDRAW_STATS_SECONDS_WORD, "")) - (intReloadTimes * 3) - intRunTimes 'This is minus seconds wasted
            'Set WPM
            If intZombieKillsToBe = 0 Then
                intWPM = 0
            Else
                intWPM = CInt((intZombieKillsToBe / (intElapsedTime / 60)))
            End If
        End If

        'Draw zombie kills
        DrawTextWithShadows(CStr(intZombieKillsToBe), 42, Color.White, 530, 206, 500, 62)

        'Draw survival time
        DrawTextWithShadows(strElapsedTime, 42, Color.White, 530, 305, 500, 62)

        'Draw WPM
        DrawTextWithShadows(CStr(intWPM), 42, Color.White, 530, 401, 500, 62)

    End Sub

    Private Sub CompareHighscoresAccess()

        'Declare
        Dim strSQL As String = "SELECT Time FROM HighscoresTable ORDER BY Rank ASC"
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

            'Compare time
            For intLoop As Integer = 0 To dtProperties.Rows.Count - 1
                If CInt(tsTimeSpan.TotalSeconds) <= CInt(dtProperties.Rows(intLoop).Item(0).ToString) Then
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
            CompareHighscoresTextFile()

        End Try

        'Check if rank beat
        If blnHighscoresIsAccess Then
            If blnRankBeat Then
                WriteToDatabase(intRank)
            End If
        End If

    End Sub

    Private Sub CompareHighscoresTextFile()

        'Notes: If there was an error, the database content from the text file will load after. 
        '       This is a backup to not having the correct database path, an install and registered Microsoft Provider 12.0.

        'Declare
        Dim ioSR As IO.StreamReader
        Dim strTemp As String = ""
        Dim astrTempSplit() As String
        Dim intRank As Integer = 0
        Dim blnRankBeat As Boolean = False

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Highscores.txt")

        'Use loop to get times
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
            If CInt(tsTimeSpan.TotalSeconds) <= CInt(astrTempSplit(0)) Then
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
            WriteToDatabase(intRank)
        End If

    End Sub

    Private Sub WriteToDatabase(intRank As Integer)


        'Declare
        Dim strInputBox As String = ""

        'Get user info
        While strInputBox = ""
            'Display
            strInputBox = InputBox("You beat a highscore, please enter your name" & vbNewLine & "(15 characters)...", "Last Stand", "")
            'Check for invalid string characters
            If blnCheckStringForInvalidCharacters(strInputBox) Or Len(strInputBox) > 15 Or Trim(strInputBox) = "" Then
                'Show
                MessageBox.Show("Can only accept standard characters, length is too long, and/or no characters is entered. Please try again...", "Last Stand",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Reset
                strInputBox = ""
            End If
        End While

        'Check
        If blnHighscoresIsAccess Then 'Access
            'Enter the information into the database
            EnterTheInformationIntoTheDatabaseAccess(strInputBox, intRank)
            'Reload the highscore database
            LoadHighscoresStringAccess(False)
        Else 'Text document
            'Enter the information into the database
            EnterTheInformationIntoTheDatabaseTextFile(strInputBox, intRank)
            'Reload the highscore database
            LoadHighscoresStringTextFile()
        End If

    End Sub

    Private Function blnCheckStringForInvalidCharacters(strToCheck As String) As Boolean

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

    Private Sub EnterTheInformationIntoTheDatabaseAccess(strName As String, intRank As Integer)

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
            'Loop to copy all data downward
            If intRank <> 10 Then
                For intLoop As Integer = dtProperties.Rows.Count To intRank Step -1
                    dtProperties.Rows(intLoop).Item(1) = dtProperties.Rows(intLoop - 1).Item(1) 'Time
                    dtProperties.Rows(intLoop).Item(2) = dtProperties.Rows(intLoop - 1).Item(2) 'WPM
                    dtProperties.Rows(intLoop).Item(3) = dtProperties.Rows(intLoop - 1).Item(3) 'Zombie kills
                    dtProperties.Rows(intLoop).Item(4) = dtProperties.Rows(intLoop - 1).Item(4) 'Name
                Next
            End If
            'Replace rank spot
            dtProperties.Rows(intRank - 1).Item(1) = CStr(CInt(tsTimeSpan.TotalSeconds))
            dtProperties.Rows(intRank - 1).Item(2) = CStr(intWPM)
            dtProperties.Rows(intRank - 1).Item(3) = CStr(intZombieKillsCombined)
            dtProperties.Rows(intRank - 1).Item(4) = strName
            'Set
            dtProperties.AcceptChanges()
        Catch
            'No debug
        End Try

    End Sub

    Private Sub EnterTheInformationIntoTheDatabaseTextFile(strName As String, intRank As Integer)

        'Notes: If there was an error, the database content from the text file will load after. 
        '       This is a backup to not having the correct database path, an install and registered Microsoft Provider 12.0.

        'Declare
        Dim ioSR As IO.StreamReader
        Dim astrLine(0) As String
        Dim blnPassedFirstElement As Boolean = False 'Must be used because astrLine(0) could be ""

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Highscores.txt")

        'Use loop to get scores
        While ioSR.Peek <> -1
            'Check if an element exists
            If astrLine.GetUpperBound(0) = 0 And Not blnPassedFirstElement Then
                'Set
                astrLine(0) = ioSR.ReadLine
                'Set
                blnPassedFirstElement = True
            Else
                'Re-dim
                ReDim Preserve astrLine(astrLine.GetUpperBound(0) + 1)
                'Set
                astrLine(astrLine.GetUpperBound(0)) = ioSR.ReadLine
            End If
        End While

        'Close
        ioSR.Close()

        'Copy downward
        If intRank <> 10 Then
            For intLoop As Integer = astrLine.GetUpperBound(0) To intRank Step -1
                astrLine(intLoop) = astrLine(intLoop - 1)
            Next
        End If

        'Change part of the array where the rank was beat
        astrLine(intRank - 1) = CStr(CInt(tsTimeSpan.TotalSeconds)) & strDRAW_STATS_SECONDS_WORD & " " & intWPM & " WPM " & intZombieKillsCombined & " zombie kills - " &
                                strName

        'Delete file
        Kill(strDirectory & "Highscores.txt")

        'Write to file
        Dim ioSW As IO.StreamWriter

        'Set
        ioSW = IO.File.AppendText(strDirectory & "Highscores.txt")

        'Write array
        For intLoop As Integer = 0 To astrLine.GetUpperBound(0)
            ioSW.WriteLine(astrLine(intLoop))
        Next

        'Close
        ioSW.Close()

    End Sub

    Private Sub PlaySinglePlayerGame()

        'Check for special effects in background
        Select Case gintLevel
            Case 5
                'Check for helicopter
                If gpntGameBackground.X <= -600 Then
                    'Draw
                    DrawGraphics(Graphics.FromImage(btmGameBackgroundMemory), udcHelicopterGonzales.Image, udcHelicopterGonzales.Point)
                End If
        End Select

        'Draw dead zombies permanently
        DrawDeadZombiesPermanently(gaudcZombies, intZombieKills)

        'Paint on canvas and clone the only necessary spot of the background
        PaintOnCanvasAndCloneScreen()

        'Check if made it to the end of the level
        Dim intEscapeDistance As Integer = 0

        'Set escape distance
        Select Case gintLevel
            Case 1, 3, 4, 5
                intEscapeDistance = -1375
            Case 2
                intEscapeDistance = -3750
        End Select

        'Check distance for level end
        If gpntGameBackground.X <= intEscapeDistance Then
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
                Select Case gintLevel
                    Case 1, 2
                        'Play door
                        udcOpeningMetalDoorSound.PlaySound(gintSoundVolume)
                End Select
                'Start black screen
                BlackScreening(intBLACK_SCREEN_BEAT_LEVEL_DELAY)
                'Stop sounds
                StopSoundsGameDone()
                'Pause the stop watch
                swhStopWatch.Stop()
                'Keep the reload times updated because object will be removed by memory
                intReloadTimes = udcCharacter.ReloadTimes
                'Keep the running times updated because object will be removed by memory
                intRunTimes = udcCharacter.RunTimes
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
        DrawGraphics(Graphics.FromImage(btmCanvas), udcCharacter.Image, udcCharacter.Point)

        'Draw zombies
        For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
            'Check if spawned
            If gaudcZombies(intLoop).Spawned Then
                'Check if can pin
                If Not gaudcZombies(intLoop).IsDying And Not gaudcZombies(intLoop).IsPinning Then
                    'Declare
                    Dim intTempDistance As Integer = gaudcZombies(intLoop).Point.X
                    'Check distance
                    If intZombieIncreasedPinDistance = 0 Then
                        'Check distance
                        If intTempDistance <= intZOMBIE_PINNING_X_DISTANCE Then
                            'Check if level not beat
                            If Not blnBeatLevel Then
                                'Set distance for future pin
                                intZombieIncreasedPinDistance = intTempDistance - 12
                                'Make zombie pin
                                gaudcZombies(intLoop).Pin()
                                'Set
                                blnPlayerWasPinned = True
                                'Set
                                blnPreventKeyPressEvent = True
                                'Stop character from moving
                                If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                                    'Stand
                                    udcCharacter.Stand()
                                End If
                                'Start black screen
                                BlackScreening(intBLACK_SCREEN_DEATH_DELAY)
                                'Stop sounds
                                StopSoundsGameDone()
                                'Play
                                udcScreamSound.PlaySound(gintSoundVolume)
                                'Pause the stop watch
                                swhStopWatch.Stop()
                                'Keep the reload times updated because object will be removed by memory
                                intReloadTimes = udcCharacter.ReloadTimes
                                'Keep the running times updated because object will be removed by memory
                                intRunTimes = udcCharacter.RunTimes
                                'Keep the zombie kills updated
                                intZombieKillsCombined += intZombieKills
                            End If
                        End If
                    Else
                        'Check distance
                        If intTempDistance <= intZombieIncreasedPinDistance Then
                            'Increase distance
                            intZombieIncreasedPinDistance -= 12
                            'Make zombie pin
                            gaudcZombies(intLoop).Pin()
                        End If
                    End If
                End If
                'Draw zombies dying, pinning or walking
                DrawGraphics(Graphics.FromImage(btmCanvas), gaudcZombies(intLoop).Image, gaudcZombies(intLoop).Point)
            End If
        Next

        'Check level for special effect
        Select Case gintLevel
            Case 2
                'Chained zombie
                DrawGraphics(Graphics.FromImage(btmCanvas), gudcChainedZombie.Image, gudcChainedZombie.Point)
                'Declare
                Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH, intWaterHeight)
                'Clone only the necessary spot
                btmGameBackgroundCloneScreenShown = btmGameBackground2WaterMemory.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short
                'Draw the background water to screen
                DrawGraphics(Graphics.FromImage(btmCanvas), btmGameBackgroundCloneScreenShown, New Point(0, intORIGINAL_SCREEN_HEIGHT - intWaterHeight)) 'Disposed of later
                'Dipose because clone just makes the picture expand with more data
                btmGameBackgroundCloneScreenShown.Dispose()
                btmGameBackgroundCloneScreenShown = Nothing
                'Pipe valve
                If gpntGameBackground.X <= -550 And gpntGameBackground.X >= -2200 Then
                    DrawGraphics(Graphics.FromImage(btmCanvas), btmPipeValveMemory, gpntPipeValve)
                End If
                'Draw zombie face
                If gpntGameBackground.X <= -1550 And gpntGameBackground.X >= -3200 Then
                    'Check if need to open eyes
                    If gpntGameBackground.X <= -2550 Then
                        'Check
                        If Not blnOpenedEyes Then
                            'Open eyes
                            gudcFaceZombie.OpenEyes()
                            'Set
                            blnOpenedEyes = True
                        End If
                    End If
                    'Draw
                    DrawGraphics(Graphics.FromImage(btmCanvas), gudcFaceZombie.Image, gudcFaceZombie.Point)
                End If
        End Select

        'Draw word bar
        DrawGraphics(Graphics.FromImage(btmCanvas), btmWordBarMemory, pntWordBar)

        'Declare
        Dim strWordToShow As String = strTheWord.Substring(intWordIndex, Len(strTheWord) - intWordIndex)

        'Draw text in the word bar
        DrawText(Graphics.FromImage(btmCanvas), strWordToShow & " " & strNextWord, 25, Color.Black, 265, 47, 500, 50) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strWordToShow & " " & strNextWord, 25, Color.Black, 264, 46, 500, 50) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strWordToShow & " " & strNextWord, 25, Color.White, 262, 45, 500, 50) 'White text

        'Word overlay
        DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(intWordIndex, 1), 25, Color.Red, 262, 45, 500, 50) 'Overlay

        'Show magazine with bullet count
        DrawGraphics(Graphics.FromImage(btmCanvas), btmAK47MagazineMemory, pntAK47Magazine)

        'Draw bullet count on magazine
        DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacter.BulletsUsed), 20, Color.Red, pntAK47Magazine.X - 7, pntAK47Magazine.Y + 25, 50, 37)

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
            DrawGraphics(Graphics.FromImage(btmCanvas), btmBlackScreen, pntTopLeft)
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
                pntTemp = New Point(audcZombiesType(intLoop).Point.X + CInt(Math.Abs(gpntGameBackground.X)), audcZombiesType(intLoop).Point.Y)
                'Draw dead
                DrawGraphics(Graphics.FromImage(btmGameBackgroundMemory), audcZombiesType(intLoop).Image, pntTemp)
                'Increase count
                intByRefZombieKills += 1
                'Start a new zombie
                audcZombiesType(intByRefZombieKills + intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
            End If
        Next

    End Sub

    Private Sub PaintOnCanvasAndCloneScreen()

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT)

        'Clone the only necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackgroundMemory.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Draw the background to screen with the cloned version
        DrawGraphics(Graphics.FromImage(btmCanvas), btmGameBackgroundCloneScreenShown, pntTopLeft)

        'Dispose because clone just makes the picture expand with more data
        btmGameBackgroundCloneScreenShown.Dispose()
        btmGameBackgroundCloneScreenShown = Nothing

    End Sub

    Private Sub BlackScreening(intScreenDelay As Integer)

        'Set
        intBlackScreenWaitMode = 0

        'Set timer delay
        tmrBlackScreen.Interval = intScreenDelay

        'Enable
        tmrBlackScreen.Enabled = True

    End Sub

    Private Sub StopSoundsGameDone()

        'Stop gagging sound
        For intLoop As Integer = 0 To audcSmallChainGagSounds.GetUpperBound(0)
            audcSmallChainGagSounds(intLoop).StopSound()
        Next

        'Stop splash sound
        udcWaterSplashSound.StopSound() 'Happens after zombie hits water

        'Face zombie
        udcFaceZombieEyesOpenSound.StopSound() 'Happens when the zombie face opens eyes

        'Stop reloading sound
        udcReloadingSound.StopSound()

        'Stop level music
        audcGameBackgroundSounds(gintLevel - 1).StopSound()

    End Sub

    Private Sub CheckCharacterStatus()

        'Check what to do at this moment
        If udcCharacter.StopCharacterFromRunning Then

            'Set
            udcCharacter.StopCharacterFromRunning = False

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
                            'Make the character run
                            udcCharacter.Run() 'Run now

                        Case clsCharacter.eintStatusMode.Reload
                            'Clear the key press buffer
                            strKeyPressBuffer = ""
                            'Make the character reload
                            If udcCharacter.StatusModeProcessing = clsCharacter.eintStatusMode.Run Then
                                'While running, stop to reload immediately
                                udcCharacter.Reload()
                            Else
                                If udcCharacter.GetPictureFrame = 1 Then
                                    'Only reload after finishing a gun shot frame
                                    udcCharacter.Reload()
                                End If
                            End If

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
            Dim astrZombiesToKill(0) As String 'Can be used to detect ""
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
                                If astrZombiesToKill.GetUpperBound(0) = 0 And astrZombiesToKill(0) = "" Then
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
                            'Becomes zombie kill buffer one later for joiner
                            strZombieDeathData &= astrZombiesToKill(intLoop) & ","
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
            If audcZombie(intLoop).Spawned And Not audcZombie(intLoop).IsDying And intClosestX > audcZombie(intLoop).Point.X Then
                'Set closest
                intClosestX = audcZombie(intLoop).Point.X
                'Set index
                intIndex = intLoop
            End If
        Next

        'Return
        Return intIndex

    End Function

    Private Sub MakeCopyOfScreenBecauseCharacterDied()

        'Check if need to copy screen
        If blnPlayerWasPinned Then
            'Check
            If btmBlackScreen Is abtmBlackScreenMemories(2) Then
                'Check if already happened
                If btmDeathScreen Is Nothing Then
                    'Before fading the screen, copy it to show for the death overlay
                    Dim rectSource As New Rectangle(0, 0, intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT)
                    'After paint make a copy
                    btmDeathScreen = btmCanvas.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)
                    'Set
                    blnBlackScreenFinished = True
                End If
            End If
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
                            If udcCharacter.BulletsUsed <> 30 Then
                                udcCharacter.StatusModeStartToDo = clsCharacter.eintStatusMode.Run
                            End If

                        Case 65 To 90, 97 To 122 'Lower case and upper case string characters
                            'Add to the buffer
                            strKeyPressBuffer &= chrKeyPressed.ToString & "."

                    End Select

                Case clsCharacter.eintStatusMode.Reload

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 39 ' = '
                            'Prepare to run forward after reload
                            If udcCharacter.BulletsUsed <> 30 Then
                                udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run
                            End If

                    End Select

                Case clsCharacter.eintStatusMode.Shoot

                    'Check key press
                    Select Case Asc(chrKeyPressed)

                        Case 32 ' = spacebar
                            'Character about to reload
                            udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload

                        Case 39 ' = '
                            'Running forward
                            If udcCharacter.BulletsUsed <> 30 Then
                                udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Run
                            End If

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

    Private Sub HighscoresScreen()

        'Draw highscores background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmHighscoresBackgroundMemory, pntTopLeft)

        'Declare
        Static strRankNumbers As String = ""

        'Set rank numbers
        If strRankNumbers = "" Then
            'Loop
            For intLoop As Integer = 1 To 10
                'Add a delimiter of "."
                strRankNumbers &= CStr(intLoop) & "."
                'Check if need to add a new line
                If intLoop <> 10 Then
                    'Add a new line
                    strRankNumbers &= vbNewLine
                End If
            Next
        End If

        'Draw highscores
        DrawTextWithShadows(strHighscores, 17, Color.Red, 15, 110, intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT)
        DrawTextWithShadows(strRankNumbers, 17, Color.White, 15, 110, 100, intORIGINAL_SCREEN_HEIGHT)

        'Check if Access
        If blnHighscoresIsAccess Then
            'Access
            DrawTextWithShadows("Database type is", 17, Color.Red, 15, 10, intORIGINAL_SCREEN_WIDTH, 50)
            DrawTextWithShadows("Access", 17, Color.White, 172, 10, intORIGINAL_SCREEN_WIDTH, 50)
        Else
            'Text file
            DrawTextWithShadows("Database type is", 17, Color.Red, 15, 10, intORIGINAL_SCREEN_WIDTH, 50)
            DrawTextWithShadows("Text File", 17, Color.White, 172, 10, intORIGINAL_SCREEN_WIDTH, 50)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub CreditsScreen()

        'Draw credits background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmCreditsBackgroundMemory, pntTopLeft)

        'Draw John Gonzales
        If btmJohnGonzales IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmJohnGonzales, pntJohnGonzales)
        End If

        'Draw Zachary Stafford
        If btmZacharyStafford IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmZacharyStafford, pntZacharyStafford)
        End If

        'Draw Cory Lewis
        If btmCoryLewis IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmCoryLewis, pntCoryLewis)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub VersusScreen()

        'Draw background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusBackgroundMemory, pntTopLeft)

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

        'Other
        Select Case intCanvasVersusShow
            Case 0 'Nickname
                'Draw host button
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusHostTextMemory, pntVersusHost)
                'Draw or
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusOrTextMemory, pntVersusOr)
                'Draw join button
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusJoinTextMemory, pntVersusJoin)
                'Draw nickname
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusNickNameTextMemory, pntVersusBlackOutline)
                'Draw player name text
                DrawText(Graphics.FromImage(btmCanvas), strNickName, 55, Color.White, 51, 175, 750, 137)
            Case 1 'Host
                'Draw hosting text
                DrawText(Graphics.FromImage(btmCanvas), "Hosting...", 36, Color.Red, 300, 225, 500, 75)
            Case 2 'Join
                'Draw IP address
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusIPAddressTextMemory, pntVersusBlackOutline)
                'Draw IP address to type
                DrawText(Graphics.FromImage(btmCanvas), strIPAddressConnect, 55, Color.White, 51, 175, 750, 137)
                'Draw connect button
                DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusConnectTextMemory, pntVersusConnect)
            Case 3 'Connecting
                'Draw connecting text
                DrawText(Graphics.FromImage(btmCanvas), "Connecting...", 36, Color.Red, 250, 225, 500, 75)
        End Select

        'Draw ip address text
        DrawText(Graphics.FromImage(btmCanvas), strIPAddress, 36, Color.Red, 7, 12, 500, 75)

        'Draw port forwarding
        DrawText(Graphics.FromImage(btmCanvas), "Router port forwarding: 10101", 25, Color.White, 187, 437, 600, 62)

    End Sub

    Private Sub LoadingVersusConnectedScreen()

        'Draw loading background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw loading, waiting, and start text
        If blnWaiting Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingWaitingTextMemory, pntLoadingWaitingText)
        Else
            'Check if finished loading
            If blnFinishedLoading Then
                'Start
                DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
            Else
                'Loading
                DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
            End If
        End If

        'Draw paragraph
        If btmLoadingParagraphVersus IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingParagraphVersus, pntLoadingParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StartedVersusGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Show death overlay
            DrawGraphics(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)
            DrawGraphics(Graphics.FromImage(btmCanvas), btmVersusWhoWon, pntTopLeft)

            'Draw stats
            If blnHost Then
                DrawStats(intZombieKillsOne)
            Else
                DrawStats(intZombieKillsTwo)
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

        'Paint on canvas and clone the only necessary spot of the background
        PaintOnCanvasAndCloneScreen()

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
        DrawGraphics(Graphics.FromImage(btmCanvas), udcCharacterOne.Image, udcCharacterOne.Point)

        'Draw first zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesOne, intZOMBIE_PINNING_X_DISTANCE, intZombieIncreasedPinDistanceOne, 7)

        'Draw character joiner
        DrawGraphics(Graphics.FromImage(btmCanvas), udcCharacterTwo.Image, udcCharacterTwo.Point)

        'Draw second zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesTwo, intZOMBIE_PINNING_X_DISTANCE + intJOINER_ADDED_X_DISTANCE, intZombieIncreasedPinDistanceTwo, 8)

        'Send X Positions
        gSendData(3, strGetXPositionsOfZombies())

        'Draw word bar
        DrawGraphics(Graphics.FromImage(btmCanvas), btmWordBarMemory, pntWordBar)

        'Draw text in the word bar
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 25, Color.Black, 265, 47, 500, 50) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 25, Color.Black, 264, 46, 500, 50) 'Shadow
        DrawText(Graphics.FromImage(btmCanvas), strTheWord, 25, Color.White, 262, 45, 500, 50) 'White text

        'Word overlay
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(0, intWordIndex), 25, Color.Red, 262, 45, 500, 50) 'Overlay
        Else
            DrawText(Graphics.FromImage(btmCanvas), strTheWord.Substring(0, intWordIndex), 25, Color.Blue, 262, 45, 500, 50) 'Overlay
        End If

        'Show magazine with bullet count
        Graphics.FromImage(btmCanvas).DrawImageUnscaled(btmAK47MagazineMemory, pntAK47Magazine)

        'Draw nickname
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), strNickName, 18, Color.Red, 45, 102, 500, 75) 'Host sees own name
            DrawText(Graphics.FromImage(btmCanvas), strNickNameConnected, 18, Color.Blue, 100, 127, 500, 75) 'Host sees joiner name
        Else
            DrawText(Graphics.FromImage(btmCanvas), strNickNameConnected, 18, Color.Red, 45, 102, 500, 75) 'Joiner sees host name
            DrawText(Graphics.FromImage(btmCanvas), strNickName, 18, Color.Blue, 100, 127, 500, 75) 'Joiner sees own name
        End If

        'Draw bullet count on magazine
        If blnHost Then
            DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacterOne.BulletsUsed), 20, Color.Red, pntAK47Magazine.X - 7, pntAK47Magazine.Y + 25, 50, 37)
        Else
            DrawText(Graphics.FromImage(btmCanvas), CStr(30 - udcCharacterTwo.BulletsUsed), 20, Color.Blue, pntAK47Magazine.X - 7, pntAK47Magazine.Y + 25, 50, 37)
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
            DrawGraphics(Graphics.FromImage(btmCanvas), btmBlackScreen, pntTopLeft)
        End If

        'Make copy if died
        MakeCopyOfScreenBecauseCharacterDied()

    End Sub

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
                        'Make character reload
                        If udcCharacterType.GetPictureFrame = 1 Then
                            'Prepare to send data
                            udcCharacterType.PrepareSendData = True 'This is for multiplayer reloading
                            'Make the character reload
                            udcCharacterType.Reload() 'Only reload after finishing a gun shot frame
                        End If

                    Case clsCharacter.eintStatusMode.Stand, clsCharacter.eintStatusMode.Shoot
                        'Make the character shoot with buffer
                        CheckTheKeyPressBuffer(udcCharacterType, audcZombiesType)

                End Select

        End Select

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

    Private Sub DrawMultiplayerZombiesAndSendData(audcZombiesType() As clsZombie, intZombiePinningXDistance As Integer,
                                                  ByRef intByRefZombieIncreasedPinDistance As Integer, intDataCase As Integer)

        'Draw zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'Check if spawned
            If audcZombiesType(intLoop).Spawned Then
                'Check if can pin
                If Not audcZombiesType(intLoop).IsDying And Not audcZombiesType(intLoop).IsPinning Then
                    'Declare
                    Dim intTempDistance As Integer = audcZombiesType(intLoop).Point.X
                    'Check distance
                    If intByRefZombieIncreasedPinDistance = 0 Then
                        'Check distance
                        If intTempDistance <= intZombiePinningXDistance Then
                            'Make zombie pin
                            audcZombiesType(intLoop).Pin() 'For hoster and joiner
                            'Check if first time game over by pin
                            If blnHost And Not blnPlayerWasPinned Then
                                'Set distance for future pin
                                intByRefZombieIncreasedPinDistance = intTempDistance - 25
                                'Set
                                blnPlayerWasPinned = True
                                'Set
                                blnPreventKeyPressEvent = True
                                'Start black screen
                                BlackScreening(intBLACK_SCREEN_DEATH_DELAY)
                                'Stop reloading sound
                                udcReloadingSound.StopSound()
                                'Play
                                udcScreamSound.PlaySound(gintSoundVolume)
                                'Stop level music
                                audcGameBackgroundSounds(gintLevel - 1).StopSound()
                                'Pause the stop watch
                                swhStopWatch.Stop()
                                'Keep the reload times updated because object will be removed by memory
                                intReloadTimes = udcCharacterOne.ReloadTimes 'This will always be player one
                                'Keep the running times updated because object will be removed by memory
                                intRunTimes = udcCharacterOne.RunTimes
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
                    Else
                        'Check distance
                        If intTempDistance <= intByRefZombieIncreasedPinDistance Then
                            'Increase distance
                            intByRefZombieIncreasedPinDistance -= 25
                            'Make zombie pin
                            gaudcZombies(intLoop).Pin()
                        End If
                    End If
                End If
                'Draw zombies dying, pinning or walking
                DrawGraphics(Graphics.FromImage(btmCanvas), audcZombiesType(intLoop).Image, audcZombiesType(intLoop).Point)
            End If
        Next

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
                strReturn &= CStr(intLoop) & ":" & CStr(gaudcZombiesOne(intLoop).Point.X) & ","
            End If
        Next

        'Add onto the string
        strReturn &= "|"

        'Get second set of zombie positions
        For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
            'Check if spawned and alive
            If gaudcZombiesTwo(intLoop).Spawned And Not gaudcZombiesTwo(intLoop).IsDying Then
                'Add onto the string
                strReturn &= CStr(intLoop) & ":" & CStr(gaudcZombiesTwo(intLoop).Point.X) & ","
            End If
        Next

        'Return
        Return strReturn

    End Function

    Private Sub StoryScreen()

        'Draw story background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmStoryBackgroundMemory, pntTopLeft)

        'Draw story text
        If btmStoryParagraph IsNot Nothing Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmStoryParagraph, pntStoryParagraph)
        End If

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub GameVersionMismatch()

        'Draw story background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmGameMismatchBackgroundMemory, pntTopLeft)

        'Display
        DrawTextWithShadows("Your version: " & strGAME_VERSION, 25, Color.Red, 439, 98, 500, 50)

        'Draw text
        If blnHost Then
            DrawTextWithShadows("Joiner version: " & strGameVersionFromConnection, 25, Color.Red, 439, 198, 500, 50)
        Else
            DrawTextWithShadows("Host version: " & strGameVersionFromConnection, 25, Color.Red, 439, 198, 500, 50)
        End If

    End Sub

    Private Sub NextLevel(btmGameBackgroundLevel As Bitmap)

        'Load default variables
        LoadDefaultVariables()

        'Memory copy level
        CopyLevelBitmap(btmGameBackgroundMemory, btmGameBackgroundLevel)

        'Set
        intZombieKills = 0 'Must be reset to set the beginning number of zombies

        'Character
        udcCharacter = New clsCharacter(Me, 50, 162, "udcCharacter", udcReloadingSound, udcGunShotSound, udcStepSound,
                                        udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Load into levels
        Select Case gintLevel
            Case 2
                'Face zombie
                gudcFaceZombie = New clsFaceZombie(Me, 2000, 0, udcFaceZombieEyesOpenSound)
                'Chained zombie
                gudcChainedZombie = New clsChainedZombie(Me, 750, 0, audcSmallChainGagSounds) 'Near the door
                gudcChainedZombie.Start()
            Case 5
                'Helicopter
                udcHelicopterGonzales = New clsHelicopter(Me, udcRotatingBladeSound)
                udcHelicopterGonzales.Start()
        End Select

        'Start character
        udcCharacter.Start()

        'Start zombies
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombies(intLoop).Start()
        Next

        'Start stop watch
        swhStopWatch.Start()

        'Play background sound music
        audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)

    End Sub

    Private Sub PathChoices()

        'Draw background
        DrawGraphics(Graphics.FromImage(btmCanvas), btmPath, pntTopLeft)

        'Text
        DrawTextWithShadows("Pick your path...", 25, Color.Red, 15, 475, 500, 62)

        'Show back button or hover back button
        ShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub CheckResolutionMode(intModeSelected As Integer, btmResolutionText As Bitmap, btmNotResolutionText As Bitmap, pntResolutionText As Point)

        'Check resolution before drawing
        If intResolutionMode = intModeSelected Then
            DrawGraphics(Graphics.FromImage(btmCanvas), btmResolutionText, pntResolutionText)
        Else
            DrawGraphics(Graphics.FromImage(btmCanvas), btmNotResolutionText, pntResolutionText)
        End If

    End Sub

    Private Sub DrawTextWithShadows(strText As String, intFontSize As Integer, colColor As Color, sngX As Single, sngY As Single, sngWidth As Single,
                                    sngHeight As Single)

        'Draw shadow first
        For intLoop As Integer = 1 To 5
            DrawText(Graphics.FromImage(btmCanvas), strText, intFontSize, Color.Black, sngX + CSng(intLoop), sngY + CSng(intLoop), sngWidth, sngHeight) 'Black shadow
        Next

        'Draw overlay
        DrawText(Graphics.FromImage(btmCanvas), strText, intFontSize, colColor, sngX, sngY, sngWidth, sngHeight) 'Overlay

    End Sub

    Private Sub DrawText(ByRef gByRefGraphicsToDrawOn As Graphics, strText As String, sngFontSize As Single, colColor As Color, sngX As Single, sngY As Single,
                         sngWidth As Single, sngHeight As Single)

        'Declare
        Dim gGraphics As Graphics = gByRefGraphicsToDrawOn
        Dim fntFont As New Font("Times New Roman", sngFontSize, FontStyle.Regular)
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

    Private Sub ShowBackButtonOrHoverBackButton()

        'Show back button or hover back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            DrawGraphics(Graphics.FromImage(btmCanvas), btmBackHoverTextMemory, pntBackHoverText)
        Else
            'Draw back text
            DrawGraphics(Graphics.FromImage(btmCanvas), btmBackTextMemory, pntBackText)
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
            gdblScreenWidthRatio = CDbl((Me.Width - intWIDTH_SUBTRACTION) / intORIGINAL_SCREEN_WIDTH)
            gdblScreenHeightRatio = CDbl((Me.Height - intHEIGHT_SUBTRACTION) / intORIGINAL_SCREEN_HEIGHT)
            'Set screen rectangle
            rectFullScreen.Width = Me.Width - intWIDTH_SUBTRACTION
            rectFullScreen.Height = Me.Height - intHEIGHT_SUBTRACTION
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

    Private Sub frmGame_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If options
        If intCanvasMode = 1 Then

            'Check if changing volume
            If blnMouseInRegion(pntMouse, 300, 23, pntSoundBar) Then 'In the bar
                'Change
                ChangeSoundVolume(pntMouse)
                'Set
                blnSliderWithMouseDown = True
            End If

        End If

    End Sub

    Private Sub ChangeSoundVolume(pntMouse As Point)

        'Use formulas created with excel, points were x = 39 to 319 y = 0 to 1000, with y=mx+b, slider was 39 to 319 with 29 to 329 but now factor ratio of mouse x
        Select Case pntMouse.X
            Case Is <= CInt(39 * (intScreenWidth / 800)) '0%
                'Picture change
                btmSoundPercent = abtmSoundMemories(0)
                'Volume change
                gintSoundVolume = 0
                'Move slider
                pntSlider.X = 29
            Case Is >= CInt(319 * (intScreenWidth / 800)) '100%
                'Picture change
                btmSoundPercent = abtmSoundMemories(100)
                'Volume change
                gintSoundVolume = 1000
                'Move slider
                pntSlider.X = 329
            Case Else '1% to 99%
                'Picture change
                btmSoundPercent = abtmSoundMemories(CInt(((3.5714 * (pntMouse.X / (intScreenWidth / 800))) - 139.29) / 10))
                'Volume change
                gintSoundVolume = CInt((3.5714 * (pntMouse.X / (intScreenWidth / 800))) - 139.29)
                'Move slider
                pntSlider.X = CInt((1.0714 * (pntMouse.X / (intScreenWidth / 800))) - 12.786)
        End Select

        'After setting volume, change option sound
        audcAmbianceSound(1).ChangeVolumeWhileSoundIsPlaying()

    End Sub

    Private Sub frmGame_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp

        'If options
        If intCanvasMode = 1 Then

            'Sound changing
            blnSliderWithMouseDown = False

        End If

    End Sub

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

            Case blnMouseInRegion(pntMouse, 106, 34, pntStartText) 'Start has been moused over
                'Hover sound
                HoverText(1, "MenuStart")

            Case blnMouseInRegion(pntMouse, 206, 49, pntHighscoresText) 'Highscores has been moused over
                'Hover sound
                HoverText(2, "MenuHighscores")

            Case blnMouseInRegion(pntMouse, 109, 43, pntStoryText) 'Story has been moused over
                'Hover sound
                HoverText(3, "MenuStory")

            Case blnMouseInRegion(pntMouse, 144, 44, pntOptionsText) 'Options has been moused over
                'Hover sound
                HoverText(4, "MenuOptions")

            Case blnMouseInRegion(pntMouse, 142, 39, pntCreditsText) 'Credits has been moused over
                'Hover sound
                HoverText(5, "MenuCredits")

            Case blnMouseInRegion(pntMouse, 128, 37, pntVersusText) 'Versus has been moused over
                'Hover sound
                HoverText(6, "MenuVersus")

            Case Else
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

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

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back has been moused over
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

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
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

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
                'Hover sound
                HoverText(1, "Path1Back")

            Case blnMouseInRegion(pntMouse, 194, 164, New Point(115, 213)) 'Path left
                'Set
                btmPath = abtmPath1Memories(1)
                'Play light switch
                If Not blnLightZap1 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap1 = True
                End If

            Case blnMouseInRegion(pntMouse, 184, 153, New Point(547, 240)) 'Path right
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

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
                'Hover sound
                HoverText(1, "Path2Back")

            Case blnMouseInRegion(pntMouse, 148, 152, New Point(321, 153)) 'Path left
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

            Case blnMouseInRegion(pntMouse, 148, 159, New Point(569, 192)) 'Path right
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

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'Check which screen is currently displayed
        Select Case intCanvasMode

            Case 0 'Menu
                MenuMouseClickScreen(pntMouse)

            Case 1 'Options
                OptionsMouseClickScreen(pntMouse)

            Case 2 'Start, loading game
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

            Case blnMouseInRegion(pntMouse, 106, 35, pntStartText) 'Start button was clicked
                'Show paragraph and set variables
                ShowParagraphAndSetVariables(2, 0, "Single", False)

            Case blnMouseInRegion(pntMouse, 207, 50, pntHighscoresText) 'Highscores button was clicked
                'Change
                ShowNextScreenAndExitMenu(4, 0)

            Case blnMouseInRegion(pntMouse, 109, 44, pntStoryText) 'Story button was clicked
                'Change
                ShowNextScreenAndExitMenu(9, 0)
                'Reset
                intStoryWaitMode = 0
                'Enable story timer
                tmrStory.Enabled = True

            Case blnMouseInRegion(pntMouse, 145, 45, pntOptionsText) 'Options button was clicked
                'Change
                ShowNextScreenAndExitMenu(1, 0)
                'Options sound
                audcAmbianceSound(1).PlaySound(gintSoundVolume, True)

            Case blnMouseInRegion(pntMouse, 143, 39, pntCreditsText) 'Credits button was clicked
                'Change
                ShowNextScreenAndExitMenu(5, 0)
                'Set
                intCreditsWaitMode = 0
                'Enable timer
                tmrCredits.Enabled = True

            Case blnMouseInRegion(pntMouse, 128, 37, pntVersusText) 'Versus button was clicked
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

    Private Sub ShowParagraphAndSetVariables(intCanvasModeToBe As Integer, intCanvasShowToBe As Integer, strTypeOfParagraphWaitType As String,
                                             blnGameIsVersusToBe As Boolean)

        'Change
        ShowNextScreenAndExitMenu(intCanvasModeToBe, intCanvasShowToBe)

        'Empty paragraph first
        If blnGameIsVersusToBe Then
            'Set
            btmLoadingParagraphVersus = Nothing
        Else
            'Set
            btmLoadingParagraph = Nothing
        End If

        'Set
        strTypeOfParagraphWait = strTypeOfParagraphWaitType

        'Set
        intParagraphWaitMode = 0

        'Use the paragraph timer
        tmrParagraph.Enabled = True

        'Set
        blnGameIsVersus = blnGameIsVersusToBe

        'Load game
        LoadBeginningGameMaterial()

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

        'Reset and disable fog
        ResetAndDisableFog()

    End Sub

    Private Sub ResetAndDisableFog()

        'Reset
        blnProcessBackFog = False
        blnProcessFrontFog = False
        intFogBackRectangleMove = 0
        intFogFrontRectangleMove = 0
        intFogBackX = intORIGINAL_SCREEN_WIDTH
        intFogFrontX = intORIGINAL_SCREEN_WIDTH

        'Disable fog
        tmrFog.Enabled = False

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

    Private Sub OptionsMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                'Menu sound
                audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0)
                'Stop options sound
                audcAmbianceSound(1).StopSound()
                'Enable fog
                tmrFog.Enabled = True

            Case blnMouseInRegion(pntMouse, 77, 16, pnt800x600Text) 'Resolution change 800x600
                'Resize screen
                ResizeScreen(0, 800, 600)

            Case blnMouseInRegion(pntMouse, 82, 16, pnt1024x768Text) 'Resolution change 1024x768
                'Resize screen
                ResizeScreen(1, 1024, 768)

            Case blnMouseInRegion(pntMouse, 82, 16, pnt1280x800Text) 'Resolution change 1280x800
                'Resize screen
                ResizeScreen(2, 1280, 800)

            Case blnMouseInRegion(pntMouse, 85, 16, pnt1280x1024Text) 'Resolution change 1280x1024
                'Resize screen
                ResizeScreen(3, 1280, 1024)

            Case blnMouseInRegion(pntMouse, 85, 16, pnt1440x900Text) 'Resolution change 1440x900
                'Resize screen
                ResizeScreen(4, 1440, 900)

            Case blnMouseInRegion(pntMouse, 88, 19, pntFullscreenText) 'Resolution change full screen
                'Resize screen
                ResizeScreen(5, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)

        End Select

    End Sub

    Private Sub ResizeScreen(intResolutionModeToSet As Integer, intWidth As Integer, intHeight As Integer)

        'Set
        intResolutionMode = intResolutionModeToSet

        'Set
        intScreenWidth = intWidth
        intScreenHeight = intHeight

        'Set
        blnScreenChanged = True

    End Sub

    Private Sub LoadingMouseClickScreen(pntMouse As Point)

        'Loading start text bar was clicked and game finished loading
        If blnFinishedLoading And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
            'Disable timer
            tmrParagraph.Enabled = False 'Bitmap is emptied later after leaving the game
            'Play game loaded press sound
            udcGameLoadedPressedSound.PlaySound(gintSoundVolume)
            'Set
            intCanvasMode = 3
            'Set
            intCanvasShow = 0
            'Start character
            udcCharacter.Start()
            'Start zombies
            For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
                gaudcZombies(intLoop).Start()
            Next
            'Start stop watch
            swhStopWatch = New Stopwatch
            swhStopWatch.Start()
            'Play background sound music
            audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)
        Else
            'Back was clicked
            GeneralBackButtonClick(pntMouse, True)
        End If

    End Sub

    Private Sub VersusMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
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

                            Case blnMouseInRegion(pntMouse, 272, 88, pntVersusHost) 'Host was clicked
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

                            Case blnMouseInRegion(pntMouse, 220, 85, pntVersusJoin) 'Join was clicked
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
                        If blnMouseInRegion(pntMouse, 302, 62, pntVersusConnect) Then
                            'Set
                            intCanvasVersusShow = 3
                            'Start thread connecting
                            thrConnecting = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Connecting))
                            thrConnecting.Start()
                        End If

                End Select

        End Select

    End Sub

    Private Sub DefaultNickName()

        'Check nick name
        If strNickName = "" Then
            strNickName = "Player"
        End If

    End Sub

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
        AddHandler udcVersusConnectedThread.gGotData, AddressOf DataArrival
        AddHandler udcVersusConnectedThread.gConnectionGone, AddressOf ConnectionLost

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

        'Start mismatch timer
        tmrGameMismatch.Enabled = True

    End Sub

    Private Sub CheckingGameVersion(strGameVersion As String)

        'Check game version
        If strGAME_VERSION <> strGameVersion Then
            'Check game for mismatch
            GamesMismatched(strGameVersion)
        Else
            'Prepare versus to play because mismatch has passed
            ShowParagraphAndSetVariables(7, 0, "Versus", True)
        End If

    End Sub

    Private Sub LoadingVersusMouseClickScreen(pntMouse As Point)

        'Check if hosting
        If blnHost Then
            'Start was clicked
            If Not blnWaiting And blnFinishedLoading And blnMouseInRegion(pntMouse, 1613, 134, pntLoadingBar) Then
                'Play game loaded press sound
                udcGameLoadedPressedSound.PlaySound(gintSoundVolume)
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
                'Play game loaded press sound
                udcGameLoadedPressedSound.PlaySound(gintSoundVolume)
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

    Private Sub StartVersusGameObjects()

        'Start characters
        udcCharacterOne.Start()
        udcCharacterTwo.Start()

        'Start zombies
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombiesOne(intLoop).Start()
        Next
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1)
            gaudcZombiesTwo(intLoop).Start()
        Next

        'Start stop watch
        swhStopWatch = New Stopwatch
        swhStopWatch.Start()

        'Play background sound music
        audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(gintSoundVolume / 4), True)

    End Sub

    Private Sub Path1ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case blnMouseInRegion(pntMouse, 194, 164, New Point(115, 213)) 'Path 1 choice clicked
                'Setup next level
                PrepareToNextLevel(2)

            Case blnMouseInRegion(pntMouse, 184, 153, New Point(547, 240)) 'Path 2 choice clicked
                'Setup next level
                PrepareToNextLevel(3)

        End Select

    End Sub

    Private Sub PrepareToNextLevel(intLevelToBe As Integer)

        'Setup next level
        gintLevel = intLevelToBe

        'Set
        blnCanLoadLevelWhileRendering = True

    End Sub

    Private Sub Path2ChoiceMouseClickScreen(pntMouse As Point)

        'Check which region is being clicked
        Select Case True

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                'Back was clicked
                GeneralBackButtonClick(pntMouse, True)

            Case blnMouseInRegion(pntMouse, 148, 152, New Point(321, 153)) 'Path 1 choice
                'Setup next level
                PrepareToNextLevel(5)

            Case blnMouseInRegion(pntMouse, 148, 159, New Point(569, 192)) 'Path 2 choice
                'Setup next level
                PrepareToNextLevel(4)

        End Select

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
        If astrWords(298) = "" Then
            AddWordsToArray()
        End If

        'Grab a random word
        LoadARandomWord()

        'Load audio
        LoadGameAudio()

        '10% loaded
        btmLoadingBar = abtmLoadingBarPictureMemories(1)

        'Check if previously loaded to the end
        If intMemoryLoadPosition = 182 Then
            intMemoryLoadPosition -= 1
        End If

        'Continue loading
        RestartLoadingGame(False)

    End Sub

    Private Sub LoadingGameVariables()

        'Load default varibles
        LoadDefaultVariables()

        'Set
        gintLevel = 1

        'Set
        blnOpenedEyes = False

        'Set
        gpntPipeValve.X = 1250

        'Set
        strTheWord = ""

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

        'Set defaults for highscores and records used with the stats screen
        intZombieKillsCombined = 0
        intReloadTimes = 0
        tsTimeSpan = New TimeSpan(0)
        strElapsedTime = ""
        intElapsedTime = 0
        intWPM = 0
        blnSetStats = False
        blnComparedHighscore = False

    End Sub

    Private Sub LoadDefaultVariables()

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
        btmBlackScreen = Nothing

        'Set
        blnBlackScreenFinished = False

        'Set
        blnPlayerWasPinned = False

        'Set
        gintBullets = 0 'Starting bullets for key press

        'Set
        blnPreventKeyPressEvent = False

        'Set
        strKeyPressBuffer = ""

        'Set
        blnBeatLevel = False

    End Sub

    Private Sub AddWordsToArray()

        astrWords(0) = "a"
        astrWords(1) = "ab"
        astrWords(2) = "able"
        astrWords(3) = "ace"
        astrWords(4) = "aim"
        astrWords(5) = "alt"
        astrWords(6) = "and"
        astrWords(7) = "archer"
        astrWords(8) = "arm"
        astrWords(9) = "arrive"
        astrWords(10) = "as"
        astrWords(11) = "at"
        astrWords(12) = "attack"
        astrWords(13) = "ate"
        astrWords(14) = "axe"
        astrWords(15) = "back"
        astrWords(16) = "bag"
        astrWords(17) = "bait"
        astrWords(18) = "bake"
        astrWords(19) = "ball"
        astrWords(20) = "bang"
        astrWords(21) = "bar"
        astrWords(22) = "bat"
        astrWords(23) = "battle"
        astrWords(24) = "be"
        astrWords(25) = "bee"
        astrWords(26) = "best"
        astrWords(27) = "beast"
        astrWords(28) = "beat"
        astrWords(29) = "beer"
        astrWords(30) = "begins"
        astrWords(31) = "bird"
        astrWords(32) = "bio"
        astrWords(33) = "blood"
        astrWords(34) = "boat"
        astrWords(35) = "bot"
        astrWords(36) = "broad"
        astrWords(37) = "bullet"
        astrWords(38) = "cage"
        astrWords(39) = "call"
        astrWords(40) = "can"
        astrWords(41) = "cap"
        astrWords(42) = "car"
        astrWords(43) = "carve"
        astrWords(44) = "causes"
        astrWords(45) = "chemical"
        astrWords(46) = "china"
        astrWords(47) = "cog"
        astrWords(48) = "col"
        astrWords(49) = "communications"
        astrWords(50) = "compound"
        astrWords(51) = "contractions"
        astrWords(52) = "coughing"
        astrWords(53) = "countries"
        astrWords(54) = "create"
        astrWords(55) = "creating"
        astrWords(56) = "daily"
        astrWords(57) = "dam"
        astrWords(58) = "day"
        astrWords(59) = "dead"
        astrWords(60) = "deadly"
        astrWords(61) = "deal"
        astrWords(62) = "death"
        astrWords(63) = "destruction"
        astrWords(64) = "die"
        astrWords(65) = "dog"
        astrWords(66) = "ear"
        astrWords(67) = "early"
        astrWords(68) = "ease"
        astrWords(69) = "easy"
        astrWords(70) = "eat"
        astrWords(71) = "edge"
        astrWords(72) = "edit"
        astrWords(73) = "egypt"
        astrWords(74) = "equipped"
        astrWords(75) = "escaped"
        astrWords(76) = "europe"
        astrWords(77) = "ever"
        astrWords(78) = "every"
        astrWords(79) = "everyone"
        astrWords(80) = "face"
        astrWords(81) = "fact"
        astrWords(82) = "fail"
        astrWords(83) = "fake"
        astrWords(84) = "fall"
        astrWords(85) = "false"
        astrWords(86) = "fan"
        astrWords(87) = "far"
        astrWords(88) = "finalize"
        astrWords(89) = "find"
        astrWords(90) = "france"
        astrWords(91) = "gain"
        astrWords(92) = "gap"
        astrWords(93) = "gas"
        astrWords(94) = "gate"
        astrWords(95) = "gene"
        astrWords(96) = "gentle"
        astrWords(97) = "giant"
        astrWords(98) = "hail"
        astrWords(99) = "half"
        astrWords(100) = "ham"
        astrWords(101) = "hand"
        astrWords(102) = "hang"
        astrWords(103) = "hard"
        astrWords(104) = "has"
        astrWords(105) = "hat"
        astrWords(106) = "hear"
        astrWords(107) = "hosts"
        astrWords(108) = "i"
        astrWords(109) = "ice"
        astrWords(110) = "idea"
        astrWords(111) = "ideal"
        astrWords(112) = "if"
        astrWords(113) = "ill"
        astrWords(114) = "image"
        astrWords(115) = "import"
        astrWords(116) = "in"
        astrWords(117) = "incursions"
        astrWords(118) = "infected"
        astrWords(119) = "intelligence"
        astrWords(120) = "involuntary"
        astrWords(121) = "iranian"
        astrWords(122) = "is"
        astrWords(123) = "israel"
        astrWords(124) = "jail"
        astrWords(125) = "jam"
        astrWords(126) = "jar"
        astrWords(127) = "jaw"
        astrWords(128) = "jet"
        astrWords(129) = "job"
        astrWords(130) = "join"
        astrWords(131) = "joke"
        astrWords(132) = "jordan"
        astrWords(133) = "joy"
        astrWords(134) = "keen"
        astrWords(135) = "key"
        astrWords(136) = "kick"
        astrWords(137) = "kid"
        astrWords(138) = "kill"
        astrWords(139) = "killing"
        astrWords(140) = "kind"
        astrWords(141) = "king"
        astrWords(142) = "kiss"
        astrWords(143) = "knee"
        astrWords(144) = "knot"
        astrWords(145) = "know"
        astrWords(146) = "lab"
        astrWords(147) = "lack"
        astrWords(148) = "lake"
        astrWords(149) = "lamp"
        astrWords(150) = "land"
        astrWords(151) = "largest"
        astrWords(152) = "last"
        astrWords(153) = "late"
        astrWords(154) = "leader"
        astrWords(155) = "mad"
        astrWords(156) = "mail"
        astrWords(157) = "main"
        astrWords(158) = "make"
        astrWords(159) = "man"
        astrWords(160) = "manage"
        astrWords(161) = "many"
        astrWords(162) = "map"
        astrWords(163) = "mark"
        astrWords(164) = "muscles"
        astrWords(165) = "nail"
        astrWords(166) = "name"
        astrWords(167) = "native"
        astrWords(168) = "navy"
        astrWords(169) = "near"
        astrWords(170) = "neat"
        astrWords(171) = "neck"
        astrWords(172) = "need"
        astrWords(173) = "not"
        astrWords(174) = "oath"
        astrWords(175) = "obey"
        astrWords(176) = "odd"
        astrWords(177) = "odor"
        astrWords(178) = "of"
        astrWords(179) = "off"
        astrWords(180) = "offense"
        astrWords(181) = "offer"
        astrWords(182) = "pace"
        astrWords(183) = "pack"
        astrWords(184) = "pad"
        astrWords(185) = "page"
        astrWords(186) = "pain"
        astrWords(187) = "painful"
        astrWords(188) = "pair"
        astrWords(189) = "pale"
        astrWords(190) = "pan"
        astrWords(191) = "panel"
        astrWords(192) = "panic"
        astrWords(193) = "park"
        astrWords(194) = "pathogen"
        astrWords(195) = "pit"
        astrWords(196) = "populations"
        astrWords(197) = "quick"
        astrWords(198) = "quit"
        astrWords(199) = "quiz"
        astrWords(200) = "quote"
        astrWords(201) = "race"
        astrWords(202) = "rage"
        astrWords(203) = "rail"
        astrWords(204) = "rain"
        astrWords(205) = "raise"
        astrWords(206) = "range"
        astrWords(207) = "rapid"
        astrWords(208) = "rare"
        astrWords(209) = "rarely"
        astrWords(210) = "rash"
        astrWords(211) = "rate"
        astrWords(212) = "raw"
        astrWords(213) = "react"
        astrWords(214) = "respiratory"
        astrWords(215) = "russia"
        astrWords(216) = "sad"
        astrWords(217) = "safe"
        astrWords(218) = "safety"
        astrWords(219) = "sail"
        astrWords(220) = "sake"
        astrWords(221) = "sale"
        astrWords(222) = "salt"
        astrWords(223) = "same"
        astrWords(224) = "sand"
        astrWords(225) = "save"
        astrWords(226) = "say"
        astrWords(227) = "schemed"
        astrWords(228) = "simultaneous"
        astrWords(229) = "site"
        astrWords(230) = "summer"
        astrWords(231) = "supreme"
        astrWords(232) = "symptoms"
        astrWords(233) = "tail"
        astrWords(234) = "take"
        astrWords(235) = "talk"
        astrWords(236) = "tall"
        astrWords(237) = "tap"
        astrWords(238) = "tape"
        astrWords(239) = "target"
        astrWords(240) = "targeting"
        astrWords(241) = "task"
        astrWords(242) = "taste"
        astrWords(243) = "tax"
        astrWords(244) = "tea"
        astrWords(245) = "teach"
        astrWords(246) = "tear"
        astrWords(247) = "terror"
        astrWords(248) = "the"
        astrWords(249) = "this"
        astrWords(250) = "three"
        astrWords(251) = "to"
        astrWords(252) = "total"
        astrWords(253) = "ugly"
        astrWords(254) = "uncle"
        astrWords(255) = "underdog"
        astrWords(256) = "undo"
        astrWords(257) = "union"
        astrWords(258) = "unit"
        astrWords(259) = "unite"
        astrWords(260) = "valley"
        astrWords(261) = "value"
        astrWords(262) = "vary"
        astrWords(263) = "vast"
        astrWords(264) = "verb"
        astrWords(265) = "verdict"
        astrWords(266) = "very"
        astrWords(267) = "victory"
        astrWords(268) = "video"
        astrWords(269) = "view"
        astrWords(270) = "wage"
        astrWords(271) = "waist"
        astrWords(272) = "wait"
        astrWords(273) = "wake"
        astrWords(274) = "walk"
        astrWords(275) = "wall"
        astrWords(276) = "want"
        astrWords(277) = "war"
        astrWords(278) = "warm"
        astrWords(279) = "warn"
        astrWords(280) = "wash"
        astrWords(281) = "waste"
        astrWords(282) = "watch"
        astrWords(283) = "water"
        astrWords(284) = "wave"
        astrWords(285) = "wax"
        astrWords(286) = "weapons"
        astrWords(287) = "with"
        astrWords(288) = "yard"
        astrWords(289) = "year"
        astrWords(290) = "yell"
        astrWords(291) = "yes"
        astrWords(292) = "yet"
        astrWords(293) = "you"
        astrWords(294) = "your"
        astrWords(295) = "youth"
        astrWords(296) = "zero"
        astrWords(297) = "zone"
        astrWords(298) = "zoo"

    End Sub

    Private Sub LoadARandomWord()

        'Delcare
        Dim rndNumber As New Random
        Dim intRandomNumber1 As Integer = rndNumber.Next(0, astrWords.GetUpperBound(0) + 1)
        Dim intRandomNumber2 As Integer = rndNumber.Next(0, astrWords.GetUpperBound(0) + 1)

        'Check
        If strTheWord = "" Then
            'Set
            strTheWord = astrWords(intRandomNumber1)
            'Check
            While intRandomNumber1 = intRandomNumber2
                intRandomNumber2 = rndNumber.Next(0, astrWords.GetUpperBound(0) + 1)
            End While
            'Set
            strNextWord = astrWords(intRandomNumber2)
        Else
            'Set
            strTheWord = strNextWord
            'Check
            While strNextWord = astrWords(intRandomNumber1)
                intRandomNumber1 = rndNumber.Next(0, astrWords.GetUpperBound(0) + 1)
            End While
            'Set
            strNextWord = astrWords(intRandomNumber1)
        End If

        'Set
        strWord = strTheWord

        'Set
        intWordIndex = 0

    End Sub

    Private Sub LoadGameAudio()

        'Load game background sounds
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            If audcGameBackgroundSounds(intLoop) Is Nothing Then
                audcGameBackgroundSounds(intLoop) = New clsSound(Me, strDirectory & "Sounds\GameBackground" & CStr(intLoop + 1) & ".mp3", 1)
            End If
        Next

        'Load scream sound
        If udcScreamSound Is Nothing Then
            udcScreamSound = New clsSound(Me, strDirectory & "Sounds\Scream.mp3", 1)
        End If

        'Load gun shot sound
        If udcGunShotSound Is Nothing Then
            udcGunShotSound = New clsSound(Me, strDirectory & "Sounds\GunShot.mp3", 10)
        End If

        'Load zombie death sounds
        For intLoop As Integer = 0 To audcZombieDeathSounds.GetUpperBound(0)
            If audcZombieDeathSounds(intLoop) Is Nothing Then
                audcZombieDeathSounds(intLoop) = New clsSound(Me, strDirectory & "Sounds\ZombieDeath" & CStr(intLoop + 1) & ".mp3", 10)
            End If
        Next

        'Load reloading sound
        If udcReloadingSound Is Nothing Then
            udcReloadingSound = New clsSound(Me, strDirectory & "Sounds\Reloading.mp3", 2) 'Incase multiplayer
        End If

        'Load step sound
        If udcStepSound Is Nothing Then
            udcStepSound = New clsSound(Me, strDirectory & "Sounds\Step.mp3", 6)
        End If

        'Load water foot left sound
        If udcWaterFootStepLeftSound Is Nothing Then
            udcWaterFootStepLeftSound = New clsSound(Me, strDirectory & "Sounds\WaterFootStepLeft.mp3", 3)
        End If

        'Load water foot right sound
        If udcWaterFootStepRightSound Is Nothing Then
            udcWaterFootStepRightSound = New clsSound(Me, strDirectory & "Sounds\WaterFootStepRight.mp3", 3)
        End If

        'Load gravel foot left sound
        If udcGravelFootStepLeftSound Is Nothing Then
            udcGravelFootStepLeftSound = New clsSound(Me, strDirectory & "Sounds\GravelFootStepLeft.mp3", 3)
        End If

        'Load gravel foot right sound
        If udcGravelFootStepRightSound Is Nothing Then
            udcGravelFootStepRightSound = New clsSound(Me, strDirectory & "Sounds\GravelFootStepRight.mp3", 3)
        End If

        'Load opening metal door sound
        If udcOpeningMetalDoorSound Is Nothing Then
            udcOpeningMetalDoorSound = New clsSound(Me, strDirectory & "Sounds\OpeningMetalDoor.mp3", 1) 'Happens only once during gameplay
        End If

        'Load light zap sound
        If udcLightZapSound Is Nothing Then
            udcLightZapSound = New clsSound(Me, strDirectory & "Sounds\LightZap.mp3", 5)
        End If

        'Load zombie growl sound
        If udcZombieGrowlSound Is Nothing Then
            udcZombieGrowlSound = New clsSound(Me, strDirectory & "Sounds\ZombieGrowl.mp3", 1)
        End If

        'Load rotating blade sound of the helicopter
        If udcRotatingBladeSound Is Nothing Then
            udcRotatingBladeSound = New clsSound(Me, strDirectory & "Sounds\RotatingBlade.mp3", 1)
        End If

        'Load chained zombie gag sound
        For intLoop As Integer = 0 To audcSmallChainGagSounds.GetUpperBound(0)
            If audcSmallChainGagSounds(intLoop) Is Nothing Then
                audcSmallChainGagSounds(intLoop) = New clsSound(Me, strDirectory & "Sounds\SmallChainGag" & CStr(intLoop + 1) & ".mp3", 2)
            End If
        Next

        'Load water splash sound
        If udcWaterSplashSound Is Nothing Then
            udcWaterSplashSound = New clsSound(Me, strDirectory & "Sounds\WaterSplash.mp3", 5)
        End If

        'Load face zombie eyes open sound
        If udcFaceZombieEyesOpenSound Is Nothing Then
            udcFaceZombieEyesOpenSound = New clsSound(Me, strDirectory & "Sounds\FaceZombieEyesOpen.mp3", 1)
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

        'Note: This will load but wait until memory has completed before doing another load, each file is loaded individually, waiting for a response from memory as well

        'Load continually
        Select Case intMemoryLoadPosition
            Case 0 To 4
                'File load array
                LoadGameFileWithIndexByDirectory(0, abtmGameBackgroundFiles, ablnGameBackgroundMemoriesCopied, "Images\Game Play\GameBackground", "GameBackground", ".jpg")
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
            Case 174 To 176
                'File load array
                LoadGameFileWithIndex(174, abtmChainedZombieFiles, ablnChainedZombieMemoriesCopied, "Images\ChainedZombie\", ".png")
            Case 177
                'File load
                LoadGameFile(btmGameBackground2WaterFile, blnGameBackground2WaterMemoryCopied, "Images\Game Play\GameBackground2\Water.png")
            Case 178, 179
                'File load array
                LoadGameFileWithIndex(178, abtmFaceZombieFiles, ablnFaceZombieMemoriesCopied, "Images\Face Zombies\", ".png")
            Case 180
                'File load
                LoadGameFile(btmPipeValveFile, blnPipeValveMemoryCopied, "Images\Game Play\GameBackground2\PipeValve.png")
            Case 181
                'Not used here, but below after graphics renders, skipped in the thread loading
            Case 182 'Be careful changing this number, find = check if previously loaded to the end
                'Check if single player
                If Not blnGameIsVersus Then
                    'Character
                    udcCharacter = New clsCharacter(Me, 50, 162, "udcCharacter", udcReloadingSound, udcGunShotSound, udcStepSound,
                                                    udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)
                    'Zombies
                    LoadZombies("Level 1 Single Player")
                Else
                    'Load in a special way
                    If blnHost Then
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, 50, 150, "udcCharacterOne", udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, 100, 175, "udcCharacterTwo", udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound, True) 'Join
                    Else
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, 50, 150, "udcCharacterOne", udcReloadingSound, udcGunShotSound, udcStepSound,
                                                           udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound,
                                                           udcGravelFootStepRightSound, True) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, 100, 175, "udcCharacterTwo", udcReloadingSound, udcGunShotSound, udcStepSound,
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
                'Play sound
                udcFinishedLoading100PercentSound.PlaySound(gintSoundVolume)
        End Select

    End Sub

    Private Sub LoadGameFileWithIndexByDirectory(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ablnMemoryCopied() As Boolean,
                                                 strImageDirectoryWithoutFolderNumber As String, strFileNameWithoutNumber As String, strFileType As String)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract

        'File load
        LoadGameFile(abtmByRefFile(intIndexChanger), ablnMemoryCopied(intIndexChanger), strImageDirectoryWithoutFolderNumber & CStr(intIndexChanger + 1) & "\" &
                     strFileNameWithoutNumber & CStr(intIndexChanger + 1) & strFileType)

    End Sub

    Private Sub LoadGameFileWithIndex(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ablnMemoryCopied() As Boolean, strImageSubDirectory As String,
                                      strFileType As String)

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
                For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_CREATED - 1)
                    'Re-dim first
                    ReDim Preserve gaudcZombies(intLoop)
                    'Check which wave number
                    Select Case intLoop
                        Case 0
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 5, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 1
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 250, 162, 5, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 2
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 500, 162, 5, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 3
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 750, 162, 5, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 4
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 1000, 162, 5, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 7, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 9, 14, 19, 24
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 17, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 29, 34, 39, 44
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 20, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 48, 49, 53, 54, 58, 59, 63, 64
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 22, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 65 To 74
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 7, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 75 To 84
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 85 To 94
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case 95 To 104
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 15, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                        Case Else
                            gaudcZombies(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 27, "udcCharacter", audcZombieDeathSounds,
                                                    udcWaterSplashSound)
                    End Select
                Next


            Case "Level 1 Multiplayer"
                'Check if hosting, if not hosting then ghost like property the zombies
                If blnHost Then 'Hoster
                    LoadMultiplayerZombies(False) 'True walking zombies, not ghost like
                Else 'Joiner
                    LoadMultiplayerZombies(True) 'Ghost non-walking zombies, x position must be updated with get data
                End If

        End Select

    End Sub

    Private Sub LoadMultiplayerZombies(blnImitation As Boolean)

        'Zombies for player
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_CREATED - 1)

            'Re-dim first
            ReDim Preserve gaudcZombiesOne(intLoop)

            'Check which wave number
            Select Case intLoop

                Case 0
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 5, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 1
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 500, 162, 5, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 2
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 1000, 162, 5, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 3
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 1500, 162, 5, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 4
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 2000, 162, 5, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 7, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 9, 14, 19, 24
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 17, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 29, 34, 39, 44
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 20, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 48, 49, 53, 54, 58, 59, 63, 64
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 22, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 65 To 74
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 7, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 75 To 84
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 85 To 94
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 12, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case 95 To 104
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 15, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

                Case Else
                    gaudcZombiesOne(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH, 162, 27, "udcCharacterOne", audcZombieDeathSounds,
                                               udcWaterSplashSound, blnImitation)

            End Select

        Next

        'Zombies for joiner
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_CREATED - 1)

            'Re-dim first
            ReDim Preserve gaudcZombiesTwo(intLoop)

            'Check which wave number
            Select Case intLoop

                Case 0
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 5, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 1
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 500 + intJOINER_ADDED_X_DISTANCE, 175, 5, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 2
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 1000 + intJOINER_ADDED_X_DISTANCE, 175, 5, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 3
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 1500 + intJOINER_ADDED_X_DISTANCE, 175, 5, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 4
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + 2000 + intJOINER_ADDED_X_DISTANCE, 175, 5, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 5 To 8, 10 To 13, 15 To 18, 20 To 23
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 7, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 9, 14, 19, 24
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 17, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 25 To 28, 30 To 33, 35 To 38, 40 To 43
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 12, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 29, 34, 39, 44
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 20, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 45 To 47, 50 To 52, 55 To 57, 60 To 62
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 12, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 48, 49, 53, 54, 58, 59, 63, 64
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 22, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 65 To 74
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 7, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 75 To 84
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 12, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 85 To 94
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 12, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case 95 To 104
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 15, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

                Case Else
                    gaudcZombiesTwo(intLoop) = New clsZombie(Me, intORIGINAL_SCREEN_WIDTH + intJOINER_ADDED_X_DISTANCE, 175, 27, "udcCharacterTwo",
                                               audcZombieDeathSounds, udcWaterSplashSound, blnImitation)

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
            Case 174 To 176
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(174, abtmChainedZombieFiles, gabtmChainedZombieMemories, ablnChainedZombieMemoriesCopied)
            Case 177
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmGameBackground2WaterFile, btmGameBackground2WaterMemory, blnGameBackground2WaterMemoryCopied)
                'Set
                If btmGameBackground2WaterMemory IsNot Nothing Then
                    intWaterHeight = btmGameBackground2WaterMemory.Height
                End If
            Case 178, 179
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(178, abtmFaceZombieFiles, gabtmFaceZombieMemories, ablnFaceZombieMemoriesCopied)
            Case 180
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmPipeValveFile, btmPipeValveMemory, blnPipeValveMemoryCopied)
            Case 181 'Be careful changing this number, find = check if previously loaded to the end
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

            'New image into memory
            btmByRefBitmapMemory = New Bitmap(btmByRefBitmapFile.Width, btmByRefBitmapFile.Height, Imaging.PixelFormat.Format32bppPArgb)

            'Memory copy
            DrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)

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

        'Make new image into memory
        btmByRefBitmapMemory = New Bitmap(btmBitmapMemoryToCopy.Width, btmBitmapMemoryToCopy.Height, Imaging.PixelFormat.Format32bppPArgb)

        'Memory copy
        DrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmBitmapMemoryToCopy, pntTopLeft)

    End Sub

End Class