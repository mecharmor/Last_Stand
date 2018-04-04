'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants for outside of in-game
    Private Const strGAME_VERSION As String = "1.75"
    Private Const intORIGINAL_SCREEN_WIDTH As Integer = 840
    Private Const intORIGINAL_SCREEN_HEIGHT As Integer = 525
    Private Const intWINDOW_MESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const intCONTROL_MOVE As Integer = 61456
    Private Const intWINDOW_MESSAGE_CLICK_BUTTON_DOWN As Integer = 161
    Private Const intWINDOW_CAPTION As Integer = 2
    Private Const intWINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3
    Private Const intWM_POSCHANGED As Integer = &H46 'Block this to allow Window to full-screen properly
    Private Const dblONE_SECOND_DELAY As Double = 1000 'This is one second delay in milliseconds
    Private Const dblMOUSE_OVER_SOUND_DELAY As Double = 35 '35 milliseconds
    Private Const dblLOADING_TRANSPARENCY_DELAY As Double = 125
    Private Const dblSTORY_SOUND_DELAY As Double = 1000 'Story delay before playing voice
    Private Const dblSTORY_PARAGRAPH2_DELAY As Double = 49000
    Private Const dblSTORY_PARAGRAPH3_DELAY As Double = 64000
    Private Const dblCREDITS_TRANSPARENCY_DELAY As Double = 125
    Private Const dblGAME_MISMATCH_DELAY As Double = 5000 'Delay to show the game mismatch screen
    Private Const intSLIDER_WIDTH As Integer = 27
    Private Const strDRAW_STATS_SECONDS_WORD As String = " Sec"
    Private Const intFOG_BACK_MEMORY_WIDTH As Integer = 1648 'Picture actual width
    Private Const intFOG_BACK_MEMORY_HEIGHT As Integer = 300 'Picture actual height
    Private Const intFOG_FRONT_MEMORY_WIDTH As Integer = 1660 'Picture actual width
    Private Const intFOG_FRONT_MEMORY_HEIGHT As Integer = 300 'Picture actual height
    Private Const dblFOG_FRAME_WAIT_TO_START As Double = 250 'Delay for fog before showing it
    Private Const dblFOG_FRAME_DELAY As Double = 15
    Private Const intFOG_SPEED As Integer = 1
    Private Const intFOG_BACK_Y As Integer = 280 'Extra added distance to shift fog down
    Private Const intFOG_FRONT_Y As Integer = 340 'Extra added distance to shift fog down
    Private Const intFOG_BACK_ADJUSTED_HEIGHT As Integer = intFOG_BACK_MEMORY_HEIGHT + (intORIGINAL_SCREEN_HEIGHT -
                                                           (intFOG_BACK_Y + intFOG_BACK_MEMORY_HEIGHT)) 'This is the bottom cut off
    Private Const intFOG_FRONT_ADJUSTED_HEIGHT As Integer = intFOG_FRONT_MEMORY_HEIGHT + (intORIGINAL_SCREEN_HEIGHT -
                                                            (intFOG_FRONT_Y + intFOG_FRONT_MEMORY_HEIGHT)) 'Bottom cut off
    Private Const intSOUND_START_PIXEL As Integer = 42
    Private Const intSOUND_END_PIXEL As Integer = 341
    Private Const intSOUND_LEFT_END_OF_SLIDER As Integer = 32
    Private Const intSOUND_RIGHT_END_OF_SLIDER As Integer = 328
    Private Const intSOUND_SLIDER_HALF_PIXEL As Integer = 13
    Private Const intSOUND_DIVIDER_PERCENT As Integer = 3 '299 / 100 = 2.9 per sound %
    Private Const intSOUND_MULTIPLER_PERCENT As Integer = 10 'Used for converting 0 to 100 with 0 to 1000
    Private Const intSOUND_OPENGL_ARRAY_OFFSET As Integer = 24
    'Controls text
    Private Const strCONTROLS_DETAIL As String = "(Spacebar to reload), (' to run), (; to stop)"
    'Black screen death delay
    Private Const intBLACK_SCREEN_BEAT_LEVEL_DELAY As Integer = 350
    Private Const intBLACK_SCREEN_DEATH_DELAY As Integer = 750
    'Zombies starting
    Private Const intNUMBER_OF_ZOMBIES_CREATED As Integer = 500
    Private Const intNUMBER_OF_ZOMBIES_AT_ONE_TIME As Integer = 5
    'Character
    Private Const intCHARACTER_X As Integer = 50
    Private Const intCHARACTER_Y As Integer = 162
    'Character multiplayer
    Private Const intJOINER_ADDED_X As Integer = 50
    Private Const intJOINER_ADDED_Y As Integer = 25
    Private Const intCHARACTER_HOSTER_X As Integer = 50
    Private Const intCHARACTER_HOSTER_Y As Integer = 162
    Private Const intCHARACTER_JOINER_X As Integer = intCHARACTER_HOSTER_X + intJOINER_ADDED_X
    Private Const intCHARACTER_JOINER_Y As Integer = intCHARACTER_HOSTER_Y + intJOINER_ADDED_Y
    'Zombies for single
    Private Const intZOMBIE_X As Integer = 350
    Private Const intZOMBIE_Y As Integer = intCHARACTER_Y
    'Zombies for multiplayer
    Private Const intZOMBIE_HOSTER_X As Integer = 350
    Private Const intZOMBIE_HOSTER_Y As Integer = intCHARACTER_HOSTER_Y
    Private Const intZOMBIE_JOINER_X As Integer = intZOMBIE_HOSTER_X + intJOINER_ADDED_X
    Private Const intZOMBIE_JOINER_Y As Integer = intCHARACTER_JOINER_Y
    Private Const intZOMBIE_PINNING_X As Integer = 100
    Private Const intINCREASE_ZOMBIE_SPEED_AFTER_DEATHS As Integer = 15
    Private Const intZOMBIE_BETWEEN_ZOMBIE_WIDTH As Integer = 100
    'On screen words
    Private Const intON_SCREEN_WORD_MISSED_RED_X As Integer = 275
    Private Const intON_SCREEN_WORD_MISSED_RED_Y As Integer = 125
    Private Const intON_SCREEN_WORD_MISSED_BLUE_X As Integer = intON_SCREEN_WORD_MISSED_RED_X + intJOINER_ADDED_X
    Private Const intON_SCREEN_WORD_MISSED_BLUE_Y As Integer = intON_SCREEN_WORD_MISSED_RED_Y + intJOINER_ADDED_Y

    'Delare PIXELFORMATDESCRIPTOR constants
    Private Const intPFD_VERSION As Integer = 1
    Private Const intPFD_DRAW_TO_WINDOW As Integer = &H4
    Private Const intPFD_SUPPORT_OPENGL As Integer = &H20
    Private Const intPFD_DOUBLEBUFFER As Integer = &H1
    Private Const intPFD_TYPE_RGBA As Integer = &H0
    Private Const intPFD_BITS As Integer = 32 'Do not use 16 or else aero theme could disable
    Private Const intPFD_SIXTEEN_BIT_Z_BUFFER As Integer = 16
    Private Const intPFD_MAIN_PLANE As Integer = &H0

    'Declare standard GL constants
    Private Const intGL_PROJECTION As Integer = &H1701
    Private Const intGL_MODELVIEW As Integer = &H1700
    Private Const intGL_SMOOTH As Integer = &H1D01 '7425
    Private Const intGL_DEPTH_TEST As Integer = &HB71 '2929
    Private Const intGL_LEQUAL As Integer = &H203 '515
    Private Const intGL_PERSPECTIVE_CORRECTION_HINT As Integer = &HC50 '3152
    Private Const intGL_NICEST As Integer = &H1102 '4354
    Private Const intGL_COLOR_BUFFER_BIT As Integer = &H4000 '16384
    Private Const intGL_DEPTH_BUFFER_BIT As Integer = &H100 '256
    Private Const intGL_VERSION As Integer = &H1F02

    'Declare texture GL constants
    Private Const intGL_TEXTURE_2D As Integer = &HDE1
    Private Const intINTERNAL_FORMAT_RGB As Integer = 3
    Private Const intINTERNAL_FORMAT_RGBA As Integer = 4
    Private Const intGL_RGBA8 As Integer = &H8058
    Private Const intGL_RGB As Integer = &H1907
    Private Const intGL_BGRA_EXT As Integer = &H80E1
    Private Const intGL_RGBA As Integer = &H1908
    Private Const intGL_UNSIGNED_BYTE As Integer = &H1401
    Private Const intGL_BLEND As Integer = &HBE2
    Private Const intGL_TRIANGLES As Integer = &H4
    Private Const intGL_CLAMP_TO_EDGE As Integer = &H812F
    Private Const intGL_TEXTURE_WRAP_S As Integer = &H2802
    Private Const intGL_TEXTURE_WRAP_T As Integer = &H2803
    Private Const intGL_TEXTURE_MIN_FILTER As Integer = 10241
    Private Const intGL_TEXTURE_MAG_FILTER As Integer = 10240
    Private Const intGL_LINEAR As Integer = 9729
    Private Const intGL_NEAREST As Integer = &H2600
    Private Const intGL_SRC_ALPHA As Integer = &H302
    Private Const intGL_ONE_MINUS_SRC_ALPHA As Integer = &H303
    Private Const intGL_VIEWPORT As Integer = &HBA2
    Private Const intGL_MAX_TEXTURE_MAX_ANISOTROPY_EXT As Integer = &H84FF

    'Declare error GL constants
    Private Const intGL_INVALID_ENUM As Integer = &H500
    Private Const intGL_INVALID_VALUE As Integer = &H501
    Private Const intGL_INVALID_OPERATION As Integer = &H502
    Private Const intGL_STACK_OVERFLOW As Integer = &H503
    Private Const intGL_STACK_UNDERFLOW As Integer = &H504
    Private Const intGL_OUT_OF_MEMORY As Integer = &H505
    Private Const intGL_INVALID_FRAMEBUFFER_OPERATION As Integer = &H506
    Private Const intGL_CONTEXT_LOST As Integer = &H507
    Private Const intGL_TABLE_TOO_LARGE As Integer = &H8031

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
    Private strDirectory As String = Replace(Replace(AppDomain.CurrentDomain.BaseDirectory, "bin\Debug\", ""), "bin\Release\", "")
    Private blnScreenChanged As Boolean = False

    'Text for the engine
    Private abtmTextSize25TNRFiles(287) As Bitmap 'TNR = Times New Roman
    Private abtmTextSize25TNRMemories(287) As Bitmap
    Private abtmTextSize36TNRFiles(287) As Bitmap
    Private abtmTextSize36TNRMemories(287) As Bitmap
    Private abtmTextSize42TNRFiles(287) As Bitmap
    Private abtmTextSize42TNRMemories(287) As Bitmap
    Private abtmTextSize55TNRFiles(287) As Bitmap
    Private abtmTextSize55TNRMemories(287) As Bitmap
    Private abtmTextSize72TNRFiles(287) As Bitmap
    Private abtmTextSize72TNRMemories(287) As Bitmap

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
    Private btmGraphicsTextFile As Bitmap
    Private btmGraphicsTextMemory As Bitmap
    Private pntGraphicsText As New Point(20, 20)
    Private btmGDIPlusSoftwareTextFile As Bitmap
    Private btmGDIPlusSoftwareTextMemory As Bitmap
    Private btmNotGDIPlusSoftwareTextFile As Bitmap
    Private btmNotGDIPlusSoftwareTextMemory As Bitmap
    Private pntGDIPlusSoftwareText As New Point(42, 71)
    Private btmOpenGLHardwareTextFile As Bitmap
    Private btmOpenGLHardwareTextMemory As Bitmap
    Private btmNotOpenGLHardwareTextFile As Bitmap
    Private btmNotOpenGLHardwareTextMemory As Bitmap
    Private pntOpenGLHardwareText As New Point(200, 71)
    Private btmSoundTextFile As Bitmap
    Private btmSoundTextMemory As Bitmap
    Private pntSoundText As New Point(20, 100)
    Private btmMCISSTextFile As Bitmap
    Private btmMCISSTextMemory As Bitmap
    Private pntMCISSText As New Point(42, 141)
    Private btmResolutionTextFile As Bitmap
    Private btmResolutionTextMemory As Bitmap
    Private pntResolutionText As New Point(20, 165)
    Private btm800x600TextFile As Bitmap
    Private btm800x600TextMemory As Bitmap
    Private btmNot800x600TextFile As Bitmap
    Private btmNot800x600TextMemory As Bitmap
    Private pnt800x600Text As New Point(42, 214)
    Private btm1024x768TextFile As Bitmap
    Private btm1024x768TextMemory As Bitmap
    Private btmNot1024x768TextFile As Bitmap
    Private btmNot1024x768TextMemory As Bitmap
    Private pnt1024x768Text As New Point(42, 239)
    Private btm1280x800TextFile As Bitmap
    Private btm1280x800TextMemory As Bitmap
    Private btmNot1280x800TextFile As Bitmap
    Private btmNot1280x800TextMemory As Bitmap
    Private pnt1280x800Text As New Point(42, 264)
    Private btm1280x1024TextFile As Bitmap
    Private btm1280x1024TextMemory As Bitmap
    Private btmNot1280x1024TextFile As Bitmap
    Private btmNot1280x1024TextMemory As Bitmap
    Private pnt1280x1024Text As New Point(42, 289)
    Private btm1440x900TextFile As Bitmap
    Private btm1440x900TextMemory As Bitmap
    Private btmNot1440x900TextFile As Bitmap
    Private btmNot1440x900TextMemory As Bitmap
    Private pnt1440x900Text As New Point(42, 314)
    Private btmFullScreenTextFile As Bitmap
    Private btmFullScreenTextMemory As Bitmap
    Private btmNotFullScreenTextFile As Bitmap
    Private btmNotFullScreenTextMemory As Bitmap
    Private pntFullscreenText As New Point(42, 339)
    Private intResolutionMode As Integer = 0 'Default 800x600
    Private btmVolumeTextFile As Bitmap
    Private btmVolumeTextMemory As Bitmap
    Private pntVolumeText As New Point(20, 370)
    Private btmSoundBarFile As Bitmap
    Private btmSoundBarMemory As Bitmap
    Private pntSoundBar As New Point(42, 423)
    Private btmSliderFile As Bitmap
    Private btmSliderMemory As Bitmap
    Private pntSlider As New Point(328, 416) '100% mark
    Private blnSliderWithMouseDown As Boolean = False
    Private btmSoundPercent As Bitmap
    Private abtmSoundFiles(100) As Bitmap '0 to 100
    Private abtmSoundMemories(100) As Bitmap '0 to 100
    Private pntSoundPercent As New Point(359, 426)

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
    Private astrHighscores(9) As String

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
    Private btmVersusHostTextHoverFile As Bitmap
    Private btmVersusHostTextHoverMemory As Bitmap
    Private pntVersusHost As New Point(169, 302)
    Private pntVersusHostHover As New Point(151, 294)
    Private btmVersusOrTextFile As Bitmap
    Private btmVersusOrTextMemory As Bitmap
    Private pntVersusOr As New Point(410, 324)
    Private btmVersusJoinTextFile As Bitmap
    Private btmVersusJoinTextMemory As Bitmap
    Private btmVersusJoinTextHoverFile As Bitmap
    Private btmVersusJoinTextHoverMemory As Bitmap
    Private pntVersusJoin As New Point(480, 300)
    Private pntVersusJoinHover As New Point(464, 294)
    Private btmVersusIPAddressTextFile As Bitmap
    Private btmVersusIPAddressTextMemory As Bitmap
    Private btmVersusConnectTextFile As Bitmap
    Private btmVersusConnectTextMemory As Bitmap
    Private btmVersusConnectTextHoverFile As Bitmap
    Private btmVersusConnectTextHoverMemory As Bitmap
    Private pntVersusConnect As New Point(267, 307)
    Private pntVersusConnectHover As New Point(249, 301)
    Private blnPlayPressedSoundNow As Boolean = False
    Private blnDoNotPlayPressedSoundAfterPressingBack As Boolean = False

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
    Private blnDrawFPS As Boolean = False 'Turns on to show over the screen

    'System FPS reading
    Private tmrSystemFPS As New Diagnostics.Stopwatch() 'Use this before and after code to read the execution time used

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
    Private whateverelse As System.Threading.Volatile
    Private btmAK47MagazineFile As Bitmap
    Private btmAK47MagazineMemory As Bitmap
    Private blnAK47MagazineMemoryCopied As Boolean
    Private pntAK47Magazine As New Point(29, 438)
    Private btmWinOverlayFile As Bitmap
    Private btmWinOverlayMemory As Bitmap
    Private blnWinOverlayMemoryCopied As Boolean
    Private blnBeatLevel As Boolean = False

    'Extra for game screen
    Private udcOnScreenWordOne As clsOnScreenWord 'Single player or hoster
    Private udcOnScreenWordTwo As clsOnScreenWord 'Joiner

    'In versus game playing variables
    Private udcCharacterOne As clsCharacter 'Host
    Private udcCharacterTwo As clsCharacter 'Join
    Private btmYouWonFile As Bitmap
    Private btmYouWonMemory As Bitmap
    Private blnYouWonMemoryCopied As Boolean
    Private btmYouLostFile As Bitmap
    Private btmYouLostMemory As Bitmap
    Private blnYouLostMemoryCopied As Boolean
    Private intVersusWhoWonMode As Integer = 0
    Private intZombiesKilledIncreaseSpawnOne As Integer = 0
    Private intZombiesKilledIncreaseSpawnTwo As Integer = 0
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
    Private blnListening As Boolean = False
    Private tcpcClient As System.Net.Sockets.TcpClient
    Private thrConnecting As System.Threading.Thread
    Private blnConnecting As Boolean = False
    Private srClientData As System.IO.StreamReader
    Private udcVersusConnectedThread As clsVersusConnectedThread
    Private blnConnectionCompleted As Boolean = False 'This exists because data sends way too fast
    Private blnCheckedGameMismatch As Boolean = False 'This exists because data sends way too fast
    Private blnWaiting As Boolean = False
    Private blnReadyEarly As Boolean = False

    'Words
    Private astrWords(298) As String 'Used to fill with words
    Private strEntireLengthOfWords As String = ""
    Private astrWordParts(11) As String '12 words such as "i a i a i a i a i a i a" > 20 characters

    'Key press
    Private strKeyPressBuffer As String = ""
    Private blnWordWrong As Boolean = False
    Private intTypedWords As Integer = 0

    'Highscores for the game
    Private intZombieKills As Integer = 0
    Private intZombiesKilledIncreaseSpawn As Integer = 0
    Private intZombieKillsCombined As Integer = 0
    Private blnStartedTimeElapsed As Boolean = False
    Private swhTimeElapsed As Stopwatch
    Private tsTimeSpan As TimeSpan 'Used for converting stop watch
    Private strElapsedTime As String = ""
    Private intElapsedTime As Integer = 0
    Private intWPM As Integer = 0
    Private blnSetStats As Boolean = False
    Private blnComparedHighscore As Boolean = False

    'Word bitmap load files
    Private abtmOnScreenWordMissedRedFiles(1) As Bitmap
    Private ablnOnScreenWordMissedRedMemoriesCopied(1) As Boolean
    Private abtmOnScreenWordMissedBlueFiles(1) As Bitmap
    Private ablnOnScreenWordMissedBlueMemoriesCopied(1) As Boolean

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

    'Destroyed brick wall
    Private btmDestroyedBrickWallFile As Bitmap
    Private btmDestroyedBrickWallMemory As Bitmap
    Private blnDestroyedBrickWallMemoryCopied As Boolean

    'Pipe
    Private btmPipeValveFile As Bitmap
    Private btmPipeValveMemory As Bitmap
    Private blnPipeValveMemoryCopied As Boolean

    'Zombie face
    Private abtmFaceZombieFiles(1) As Bitmap
    Private ablnFaceZombieMemoriesCopied(1) As Boolean
    Private blnOpenedEyes As Boolean = False

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

    'Graphic system
    Private blnOpenGL As Boolean = False

    'Window OpenGL variables
    Private iptGDIContext As IntPtr
    Private iptPermanentContext As IntPtr = Nothing
    Private uintPixelFormatTemp As UInt32
    Private strucPixelFormatDescriptor As PIXELFORMATDESCRIPTOR
    Private blnActivatedRenderingContext As Boolean = False

    'Functions and subs for User32
    Private Declare Function GetDC Lib "user32" (iptHandleToBe As IntPtr) As IntPtr
    Private Declare Function ReleaseDC Lib "user32" (iptWindowHandle As IntPtr, iptHandleDC As IntPtr) As Boolean

    'Functions and subs for GDI32
    Private Declare Function ChoosePixelFormat Lib "gdi32" (iptHandleToBe As IntPtr,
                                                            ByRef strucByRefPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As UInt32
    Private Declare Function SetPixelFormat Lib "gdi32" (iptHandleToBe As IntPtr, uintPixelFormat As UInt32,
                                                         ByRef strucByRefPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Boolean
    Private Declare Function SwapBuffers Lib "gdi32" (iptGDIContextToBe As IntPtr) As Boolean
    Private Declare Function DeleteDC Lib "gdi32" (iptHandleToBe As IntPtr) As Boolean

    'Functions and subs for OpenGL
    Private Declare Function wglCreateContext Lib "opengl32" (iptHandleToBe As IntPtr) As IntPtr
    Private Declare Function wglDeleteContext Lib "opengl32" (iptHandleToBe As IntPtr) As IntPtr
    Private Declare Function wglMakeCurrent Lib "opengl32" (iptGDIContextToBe As IntPtr, iptPermanentContextToBe As IntPtr) As Boolean
    Private Declare Sub glViewport Lib "opengl32" (intX As Integer, intY As Integer, intWidth As Integer, intHeight As Integer)
    Private Declare Sub glMatrixMode Lib "opengl32" (intMode As Integer)
    Private Declare Sub glLoadIdentity Lib "opengl32" ()
    Private Declare Sub glShadeModel Lib "opengl32" (intShade As Integer)
    Private Declare Sub glClearDepth Lib "opengl32" (dblDepth As Double)
    Private Declare Sub glEnable Lib "opengl32" (intCapability As Integer)
    Private Declare Sub glDisable Lib "opengl32" (intCapability As Integer)
    Private Declare Sub glDepthFunc Lib "opengl32" (intFunction As Integer)
    Private Declare Sub glHint Lib "opengl32" (intTarget As Integer, intMode As Integer)
    Private Declare Function glGetError Lib "opengl32" () As Integer 'Returns enumeration error codes from previous mistake
    Private Declare Sub glClear Lib "opengl32" (intMask As Integer)
    Private Declare Sub glGenTextures Lib "opengl32" (intOnlyOneNumberTexture As Integer, ByRef iptByRefTexture As IntPtr)
    Private Declare Sub glGenTextures Lib "opengl32" (intNumberOfTextures As Integer, ByRef aiptByRefTextures() As IntPtr)
    Private Declare Sub glBindTexture Lib "opengl32" (intTarget As Integer, iptTexture As IntPtr)
    Private Declare Sub glTexImage2D Lib "opengl32" (intTarger As Integer, intLevel As Integer, intInternalFormat As Integer, intWidth As Integer,
                                                     intHeight As Integer, intBorder As Integer, intFormat As Integer, intType As Integer,
                                                     iptBitmapData As IntPtr)
    Private Declare Sub glTexCoord2f Lib "opengl32" (sngS As Single, sngT As Single)
    Private Declare Sub glVertex3f Lib "opengl32" (sngX As Single, sngY As Single, sngZ As Single)
    Private Declare Sub glTexParameteri Lib "opengl32" (intTarget As Integer, intPName As Integer, intParam As Integer)
    Private Declare Sub glBlendFunc Lib "opengl32" (intSFactor As Integer, intDFactor As Integer)
    Private Declare Sub glBegin Lib "opengl32" (intMode As Integer)
    Private Declare Sub glEnd Lib "opengl32" ()
    Private Declare Function glGetString Lib "opengl32" (intName As Integer) As IntPtr
    Private Declare Sub glFinish Lib "opengl32" ()

    'OpenGL necessary engine needs
    Private blnOpenGLLoaded As Boolean = False
    Private strOpenGLVersion As String = ""
    Private aiptOptionsTextures(125) As IntPtr 'Includes back button
    Private intSoundTextureIndex As Integer = 0 '24 to 124
    Private aiptMenuTextures(16) As IntPtr
    Private intOpenGLFrontFogX As Integer = 0
    Private intOpenGLBackFogX As Integer = 0
    Private aiptLoadingTextures(17) As IntPtr
    Private intLoadingBarTextureIndex As Integer = 0 '1 to 10
    Private intLoadingParagraphTextureIndex As Integer = 0 '1 to 4
    Private aiptHighscoreTextures(0) As IntPtr
    Private aiptStoryTextures(12) As IntPtr
    Private intStoryParagraphTextureIndex As Integer = 0 '1 to 12
    Private aiptCreditTextures(16) As IntPtr
    Private intJohnGonzlesTextureIndex As Integer = 0
    Private intZacharyStaffordTextureIndex As Integer = 0
    Private intCoryLewisTextureIndex As Integer = 0
    Private aiptVersusScreenTextures(9) As IntPtr
    Private aiptGameMismatchTextures(0) As IntPtr
    Private aiptVersusLoadingTextures(5) As IntPtr
    Private intVersusLoadingParagraphTextureIndex As Integer = 0 '2 to 5

    'OpenGL texture index notes
    '=======================================================
    'aiptLoadingTextures with 
    'intLoadingBarTextureIndex and (1 to 11)
    'intLoadingParagraphTextureIndex (14, 15, 16, 17)
    '=======================================================
    '0 = btmLoadingBackgroundMemory
    '1 = abtmLoadingBarPictureMemories(0)
    '2 = abtmLoadingBarPictureMemories(1)
    '3 = abtmLoadingBarPictureMemories(2)
    '4 = abtmLoadingBarPictureMemories(3)
    '5 = abtmLoadingBarPictureMemories(4)
    '6 = abtmLoadingBarPictureMemories(5)
    '7 = abtmLoadingBarPictureMemories(6)
    '8 = abtmLoadingBarPictureMemories(7)
    '9 = abtmLoadingBarPictureMemories(8)
    '10 = abtmLoadingBarPictureMemories(9)
    '11 = abtmLoadingBarPictureMemories(10)
    '12 = btmLoadingTextMemory
    '13 = btmLoadingStartTextMemory
    '14 = abtmLoadingParagraphMemories(0)
    '15 = abtmLoadingParagraphMemories(1)
    '16 = abtmLoadingParagraphMemories(2)
    '17 = abtmLoadingParagraphMemories(3)
    '=======================================================
    'intVersusLoadingParagraphTextureIndex
    '=======================================================
    '2 = abtmLoadingParagraphVersusMemories(0)
    '3 = abtmLoadingParagraphVersusMemories(1)
    '4 = abtmLoadingParagraphVersusMemories(2)
    '5 = abtmLoadingParagraphVersusMemories(3)

    'Text for OpenGL
    Private aiptTextSize25Textures(287) As IntPtr
    Private aiptTextSize36Textures(287) As IntPtr
    Private aiptTextSize42Textures(287) As IntPtr
    Private aiptTextSize55Textures(287) As IntPtr
    Private aiptTextSize72Textures(287) As IntPtr

    'Loading OpenGL game
    Private aiptGameBackgroundTextures(4) As IntPtr
    Private iptWordBarTexture As IntPtr
    Private aiptCharacterTextures(94) As IntPtr
    Private aiptZombieTextures(53) As IntPtr
    Private aiptHelicopterTextures(4) As IntPtr
    Private iptAK47MagazineTexture As IntPtr
    Private iptDeathOverlayTexture As IntPtr
    Private iptWinOverlayTexture As IntPtr
    Private iptYouWonTexture As IntPtr
    Private iptYouLostTexture As IntPtr
    Private aiptBlackScreenTextures(2) As IntPtr
    Private aiptPathTextures(5) As IntPtr
    Private intPathTextureIndex As Integer = 0
    Private aiptChainedZombieTextures(2) As IntPtr
    Private aiptFaceZombieTextures(1) As IntPtr
    Private iptPipeValveTexture As IntPtr
    Private iptDestroyedBrickWallTexture As IntPtr
    Private aiptOnScreenWordMissedTextures(3) As IntPtr
    Private iptGameBackgroundTexture As IntPtr 'Temporary created on the spot during game
    Private iptDeathScreenTexture As IntPtr 'Temporary created on the spot during game
    Private iptGameBackground2WaterTexture As IntPtr 'Temporary created on the spot during game

    'OpenGL texture notes
    '=======================================================
    'aiptCharacterTextures
    '=======================================================
    '0 = gabtmCharacterStandMemories(0)
    '1 = gabtmCharacterStandMemories(1)
    '2 = gabtmCharacterShootMemories(0)
    '3 = gabtmCharacterShootMemories(1)
    '4 = gabtmCharacterReloadMemories(0)
    '5 = gabtmCharacterReloadMemories(1)
    '6 = gabtmCharacterReloadMemories(2)
    '7 = gabtmCharacterReloadMemories(3)
    '8 = gabtmCharacterReloadMemories(4)
    '9 = gabtmCharacterReloadMemories(5)
    '10 = gabtmCharacterReloadMemories(6)
    '11 = gabtmCharacterReloadMemories(7)
    '12 = gabtmCharacterReloadMemories(8)
    '13 = gabtmCharacterReloadMemories(9)
    '14 = gabtmCharacterReloadMemories(10)
    '15 = gabtmCharacterReloadMemories(11)
    '16 = gabtmCharacterReloadMemories(12)
    '17 = gabtmCharacterReloadMemories(13)
    '18 = gabtmCharacterReloadMemories(14)
    '19 = gabtmCharacterReloadMemories(15)
    '20 = gabtmCharacterReloadMemories(16)
    '21 = gabtmCharacterReloadMemories(17)
    '22 = gabtmCharacterReloadMemories(18)
    '23 = gabtmCharacterReloadMemories(19)
    '24 = gabtmCharacterReloadMemories(20)
    '25 = gabtmCharacterReloadMemories(21)
    '26 = gabtmCharacterRunningMemories(0)
    '27 = gabtmCharacterRunningMemories(1)
    '28 = gabtmCharacterRunningMemories(2)
    '29 = gabtmCharacterRunningMemories(3)
    '30 = gabtmCharacterRunningMemories(4)
    '31 = gabtmCharacterRunningMemories(5)
    '32 = gabtmCharacterRunningMemories(6)
    '33 = gabtmCharacterRunningMemories(7)
    '34 = gabtmCharacterRunningMemories(8)
    '35 = gabtmCharacterRunningMemories(9)
    '36 = gabtmCharacterRunningMemories(10)
    '37 = gabtmCharacterRunningMemories(11)
    '38 = gabtmCharacterRunningMemories(12)
    '39 = gabtmCharacterRunningMemories(13)
    '40 = gabtmCharacterRunningMemories(14)
    '41 = gabtmCharacterRunningMemories(15)
    '42 = gabtmCharacterRunningMemories(16)
    '43 = gabtmCharacterStandRedMemories(0)
    '44 = gabtmCharacterStandRedMemories(1)
    '45 = gabtmCharacterShootRedMemories(0)
    '46 = gabtmCharacterShootRedMemories(1)
    '47 = gabtmCharacterReloadRedMemories(0)
    '48 = gabtmCharacterReloadRedMemories(1)
    '49 = gabtmCharacterReloadRedMemories(2)
    '50 = gabtmCharacterReloadRedMemories(3)
    '51 = gabtmCharacterReloadRedMemories(4)
    '52 = gabtmCharacterReloadRedMemories(5)
    '53 = gabtmCharacterReloadRedMemories(6)
    '54 = gabtmCharacterReloadRedMemories(7)
    '55 = gabtmCharacterReloadRedMemories(8)
    '56 = gabtmCharacterReloadRedMemories(9)
    '57 = gabtmCharacterReloadRedMemories(10)
    '58 = gabtmCharacterReloadRedMemories(11)
    '59 = gabtmCharacterReloadRedMemories(12)
    '60 = gabtmCharacterReloadRedMemories(13)
    '61 = gabtmCharacterReloadRedMemories(14)
    '62 = gabtmCharacterReloadRedMemories(15)
    '63 = gabtmCharacterReloadRedMemories(16)
    '64 = gabtmCharacterReloadRedMemories(17)
    '65 = gabtmCharacterReloadRedMemories(18)
    '66 = gabtmCharacterReloadRedMemories(19)
    '67 = gabtmCharacterReloadRedMemories(20)
    '68 = gabtmCharacterReloadRedMemories(21)
    '69 = gabtmCharacterStandBlueMemories(0)
    '70 = gabtmCharacterStandBlueMemories(1)
    '71 = gabtmCharacterShootBlueMemories(0)
    '72 = gabtmCharacterShootBlueMemories(1)
    '73 = gabtmCharacterReloadBlueMemories(0)
    '74 = gabtmCharacterReloadBlueMemories(1)
    '75 = gabtmCharacterReloadBlueMemories(2)
    '76 = gabtmCharacterReloadBlueMemories(3)
    '77 = gabtmCharacterReloadBlueMemories(4)
    '78 = gabtmCharacterReloadBlueMemories(5)
    '79 = gabtmCharacterReloadBlueMemories(6)
    '80 = gabtmCharacterReloadBlueMemories(7)
    '81 = gabtmCharacterReloadBlueMemories(8)
    '82 = gabtmCharacterReloadBlueMemories(9)
    '83 = gabtmCharacterReloadBlueMemories(10)
    '84 = gabtmCharacterReloadBlueMemories(11)
    '85 = gabtmCharacterReloadBlueMemories(12)
    '86 = gabtmCharacterReloadBlueMemories(13)
    '87 = gabtmCharacterReloadBlueMemories(14)
    '88 = gabtmCharacterReloadBlueMemories(15)
    '89 = gabtmCharacterReloadBlueMemories(16)
    '90 = gabtmCharacterReloadBlueMemories(17)
    '91 = gabtmCharacterReloadBlueMemories(18)
    '92 = gabtmCharacterReloadBlueMemories(19)
    '93 = gabtmCharacterReloadBlueMemories(20)
    '94 = gabtmCharacterReloadBlueMemories(21)
    '=======================================================
    'aiptZombieTextures
    '=======================================================
    '0 = gabtmZombieWalkMemories(0)
    '1 = gabtmZombieWalkMemories(1)
    '2 = gabtmZombieWalkMemories(2)
    '3 = gabtmZombieWalkMemories(3)
    '4 = gabtmZombieDeath1Memories(0)
    '5 = gabtmZombieDeath1Memories(1)
    '6 = gabtmZombieDeath1Memories(2)
    '7 = gabtmZombieDeath1Memories(3)
    '8 = gabtmZombieDeath1Memories(4)
    '9 = gabtmZombieDeath1Memories(5)
    '10 = gabtmZombieDeath2Memories(0)
    '11 = gabtmZombieDeath2Memories(1)
    '12 = gabtmZombieDeath2Memories(2)
    '13 = gabtmZombieDeath2Memories(3)
    '14 = gabtmZombieDeath2Memories(4)
    '15 = gabtmZombieDeath2Memories(5)
    '16 = gabtmZombiePinMemories(0)
    '17 = gabtmZombiePinMemories(1)
    '18 = gabtmZombieWalkRedMemories(0)
    '19 = gabtmZombieWalkRedMemories(1)
    '20 = gabtmZombieWalkRedMemories(2)
    '21 = gabtmZombieWalkRedMemories(3)
    '22 = gabtmZombieDeathRed1Memories(0)
    '23 = gabtmZombieDeathRed1Memories(1)
    '24 = gabtmZombieDeathRed1Memories(2)
    '25 = gabtmZombieDeathRed1Memories(3)
    '26 = gabtmZombieDeathRed1Memories(4)
    '27 = gabtmZombieDeathRed1Memories(5)
    '28 = gabtmZombieDeathRed2Memories(0)
    '29 = gabtmZombieDeathRed2Memories(1)
    '30 = gabtmZombieDeathRed2Memories(2)
    '31 = gabtmZombieDeathRed2Memories(3)
    '32 = gabtmZombieDeathRed2Memories(4)
    '33 = gabtmZombieDeathRed2Memories(5)
    '34 = gabtmZombiePinRedMemories(0)
    '35 = gabtmZombiePinRedMemories(1)
    '36 = gabtmZombieWalkBlueMemories(0)
    '37 = gabtmZombieWalkBlueMemories(1)
    '38 = gabtmZombieWalkBlueMemories(2)
    '39 = gabtmZombieWalkBlueMemories(3)
    '40 = gabtmZombieDeathBlue1Memories(0)
    '41 = gabtmZombieDeathBlue1Memories(1)
    '42 = gabtmZombieDeathBlue1Memories(2)
    '43 = gabtmZombieDeathBlue1Memories(3)
    '44 = gabtmZombieDeathBlue1Memories(4)
    '45 = gabtmZombieDeathBlue1Memories(5)
    '46 = gabtmZombieDeathBlue2Memories(0)
    '47 = gabtmZombieDeathBlue2Memories(1)
    '48 = gabtmZombieDeathBlue2Memories(2)
    '49 = gabtmZombieDeathBlue2Memories(3)
    '50 = gabtmZombieDeathBlue2Memories(4)
    '51 = gabtmZombieDeathBlue2Memories(5)
    '52 = gabtmZombiePinBlueMemories(0)
    '53 = gabtmZombiePinBlueMemories(1)
    '=======================================================
    'aiptHelicopterTextures
    '=======================================================
    '0 = gabtmHelicopterMemories(0)
    '1 = gabtmHelicopterMemories(1)
    '2 = gabtmHelicopterMemories(2)
    '3 = gabtmHelicopterMemories(3)
    '4 = gabtmHelicopterMemories(4)
    '=======================================================
    'aiptBlackScreenTextures
    '=======================================================
    '0 = abtmBlackScreenMemories(0)
    '1 = abtmBlackScreenMemories(1)
    '2 = abtmBlackScreenMemories(2)
    '=======================================================
    'aiptPathTextures with intPathTextureIndex
    '=======================================================
    '0 = abtmPath1Memories(0)
    '1 = abtmPath1Memories(1)
    '2 = abtmPath1Memories(2)
    '3 = abtmPath2Memories(0)
    '4 = abtmPath2Memories(1)
    '5 = abtmPath2Memories(2)
    '=======================================================
    'aiptChainedZombieTextures
    '=======================================================
    '0 = gabtmChainedZombieMemories(0)
    '1 = gabtmChainedZombieMemories(1)
    '2 = gabtmChainedZombieMemories(2)
    '=======================================================
    'aiptFaceZombieTextures
    '=======================================================
    '0 = gabtmFaceZombieMemories(0)
    '1 = gabtmFaceZombieMemories(1)
    '=======================================================
    'aiptOnScreenWordMissedTextures
    '=======================================================
    '0 = gabtmOnScreenWordMissedRedMemories(0)
    '1 = gabtmOnScreenWordMissedRedMemories(1)
    '2 = gabtmOnScreenWordMissedBlueMemories(0)
    '3 = gabtmOnScreenWordMissedBlueMemories(1)

    'Declare cheats
    Private blnGodMode As Boolean = False

    'Delegate sub procedure
    Private Delegate Sub SubDelegate()

    Private Structure VIEWPORT
        Public intX As Integer
        Public intY As Integer
        Public intWidth As Integer
        Public intHeight As Integer
    End Structure

    Private Structure PIXELFORMATDESCRIPTOR
        Public intSize As Integer 'Set before using this structure
        Public intVersion As Integer
        Public intFormatSupport As Integer 'Support window, opengl, and double buffering
        Public intTypeRGBA As Integer 'Request an RGBA format
        Public intBits As Integer 'Select our color depth
        Public intColorBit1, intColorBit2, intColorBit3, intColorBit4, intColorBit5, intColorBit6 As Integer
        Public intAlphaBuffer As Integer
        Public intShiftBit As Integer
        Public intAccumulationBuffer As Integer
        Public intAccumulationBit1, intAccumulationBit2, intAccumulationBit3, intAccumulationBit4 As Integer
        Public intSixteenBitZBuffer As Integer '16 bit z-buffer (depth buffer)
        Public intStencilBuffer As Integer
        Public intAuxiliaryBuffer As Integer
        Public intMainPlane As Integer 'Main drawing layer
        Public intReserved As Integer
        Public intLayerMask1, intLayerMask2, intLayerMask3 As Integer
    End Structure

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes: Do not allow a resize of the window if full screen, this happens if you double click the title, but is prevented with this sub.

        'Check if full screen
        If intResolutionMode = 5 Then
            'Prevent moving the form by control box click
            If (m.Msg = intWINDOW_MESSAGE_SYSTEM_COMMAND And m.WParam.ToInt32() = intCONTROL_MOVE) Or
            (m.Msg = intWINDOW_MESSAGE_CLICK_BUTTON_DOWN And m.WParam.ToInt32() = intWINDOW_CAPTION) Then 'Prevent button down moving form
                'Exit
                Exit Sub
            End If
        End If

        'If a double click on the title bar is triggered
        If m.Msg = intWINDOW_MESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
            'Exit
            Exit Sub
        End If

        'Block window position change, allows Window to full-screen properly
        If m.Msg <> intWM_POSCHANGED Then
            'Still send message
            MyBase.WndProc(m) 'Must have mybase
        End If

    End Sub

    Sub New()

        'This call is required by the designer
        InitializeComponent() 'Do not remove

        'Load for testing the threading status
        thrRendering = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Rendering))

        'Check for multi-threading first
        If thrRendering.TrySetApartmentState(Threading.ApartmentState.MTA) Then
            'Set, multi-threading is possible
            blnThreadSupported = True
            'Stop the paint event, we will do the painting when we want to
            SetStyle(ControlStyles.Opaque, True) 'Reduce flicker for OpenGL
            SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            SetStyle(ControlStyles.UserPaint, False)
            'Form buffer
            DoubleBuffered = True
        End If

        'Set images into memory
        SetImagesIntoMemory()

        '0% GDI loaded
        btmLoadingBar = abtmLoadingBarPictureMemories(0)

        '0% OpenGL loaded
        intLoadingBarTextureIndex = 1 '0 was taken by picture background (1 to 4)

        'Set GDI sound 100%
        btmSoundPercent = abtmSoundMemories(100)

        'Set OpenGL sound 100%
        intSoundTextureIndex = 124 '100% just like above

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

        'Text
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmTextSize25TNRFiles, abtmTextSize25TNRMemories, "Images\Text\25 Size Times New Roman\",
                                                      ".png", 0, 1)
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmTextSize36TNRFiles, abtmTextSize36TNRMemories, "Images\Text\36 Size Times New Roman\",
                                                      ".png", 0, 1)
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmTextSize42TNRFiles, abtmTextSize42TNRMemories, "Images\Text\42 Size Times New Roman\",
                                                      ".png", 0, 1)
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmTextSize55TNRFiles, abtmTextSize55TNRMemories, "Images\Text\55 Size Times New Roman\",
                                                      ".png", 0, 1)
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmTextSize72TNRFiles, abtmTextSize72TNRMemories, "Images\Text\72 Size Times New Roman\",
                                                      ".png", 0, 1)

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
        LoadPictureFileAndCopyBitmapIntoMemory(btmGraphicsTextFile, btmGraphicsTextMemory, "Images\Options\Graphics.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmGDIPlusSoftwareTextFile, btmGDIPlusSoftwareTextMemory, "Images\Options\GDIPLUSSoftware.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNotGDIPlusSoftwareTextFile, btmNotGDIPlusSoftwareTextMemory,
                                               "Images\Options\notGDIPLUSSoftware.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmOpenGLHardwareTextFile, btmOpenGLHardwareTextMemory, "Images\Options\OpenGLHardware.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmNotOpenGLHardwareTextFile, btmNotOpenGLHardwareTextMemory,
                                               "Images\Options\notOpenGLHardware.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSoundTextFile, btmSoundTextMemory, "Images\Options\Sound.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmMCISSTextFile, btmMCISSTextMemory, "Images\Options\MCISS.png")
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
        LoadPictureFileAndCopyBitmapIntoMemory(btmVolumeTextFile, btmVolumeTextMemory, "Images\Options\Volume.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSoundBarFile, btmSoundBarMemory, "Images\Options\SoundBar.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmSliderFile, btmSliderMemory, "Images\Options\Slider.png")

        'Loading screen
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingBackgroundFile, btmLoadingBackgroundMemory, "Images\Loading Game\LoadingBackground.jpg")

        'Loading bar pictures
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingBarPictureFiles, abtmLoadingBarPictureMemories, "Images\Loading Game\LoadingBar",
                                                      ".png", 0, 10)

        'Loading screen continued
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingTextFile, btmLoadingTextMemory, "Images\Loading Game\LoadingText.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingStartTextFile, btmLoadingStartTextMemory, "Images\Loading Game\LoadingStartText.png")

        'Loading paragraphs
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingParagraphFiles, abtmLoadingParagraphMemories,
                                                      "Images\Loading Game\LoadingParagraph", ".png", 1, 1)

        'Highscores background
        LoadPictureFileAndCopyBitmapIntoMemory(btmHighscoresBackgroundFile, btmHighscoresBackgroundMemory,
                                               "Images\Highscores\HighscoresBackground.jpg")

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
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusHostTextHoverFile, btmVersusHostTextHoverMemory, "Images\Versus\VersusHostHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusOrTextFile, btmVersusOrTextMemory, "Images\Versus\VersusOr.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusJoinTextFile, btmVersusJoinTextMemory, "Images\Versus\VersusJoin.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusJoinTextHoverFile, btmVersusJoinTextHoverMemory, "Images\Versus\VersusJoinHover.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusIPAddressTextFile, btmVersusIPAddressTextMemory, "Images\Versus\VersusIPAddress.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusConnectTextFile, btmVersusConnectTextMemory, "Images\Versus\VersusConnect.png")
        LoadPictureFileAndCopyBitmapIntoMemory(btmVersusConnectTextHoverFile, btmVersusConnectTextHoverMemory,
                                               "Images\Versus\VersusConnectHover.png")

        'Loading versus screen
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmLoadingParagraphVersusFiles, abtmLoadingParagraphVersusMemories,
                                                      "Images\Loading Game\LoadingParagraphVersus", ".png", 1, 1)

        'Loading versus screen continued
        LoadPictureFileAndCopyBitmapIntoMemory(btmLoadingWaitingTextFile, btmLoadingWaitingTextMemory,
                                               "Images\Loading Game\LoadingWaitingText.png")

        'Story screen
        LoadPictureFileAndCopyBitmapIntoMemory(btmStoryBackgroundFile, btmStoryBackgroundMemory, "Images\Story\StoryBackground.jpg")

        'Story paragraphs
        LoadArrayPictureFilesAndCopyBitmapsIntoMemory(abtmStoryParagraphFiles, abtmStoryParagraphMemories, "Images\Story\Paragraph", ".png", 1, 1)

        'Game versus mismatch
        LoadPictureFileAndCopyBitmapIntoMemory(btmGameMismatchBackgroundFile, btmGameMismatchBackgroundMemory,
                                               "Images\Game Mismatch\GameMismatchBackground.jpg")

        'Load sound file percentages
        For intLoop As Integer = 0 To abtmSoundFiles.GetUpperBound(0)
            LoadPictureFileAndCopyBitmapIntoMemory(abtmSoundFiles(intLoop), abtmSoundMemories(intLoop), "Images\Options\Sound" & CStr(intLoop) &
                                                   ".png")
        Next

    End Sub

    Private Sub LoadPictureFileAndCopyBitmapIntoMemory(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap,
                                                       strImageDirectory As String)

        'Prepare to create new memory blank bitmap
        If IO.File.Exists(strDirectory & strImageDirectory) Then
            'Attempt to load file
            Try
                'Load picture file bitmap
                btmByRefBitmapFile = New Bitmap(Image.FromFile(strDirectory & strImageDirectory))
                'Create new blank image
                btmByRefBitmapMemory = New Bitmap(btmByRefBitmapFile.Width, btmByRefBitmapFile.Height, Imaging.PixelFormat.Format32bppPArgb)
                'Draw the file bitmap data to memory blank sheet, this also changes the pixelformat making rendering much faster with the memory
                GDIDrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)
                'Dispose of the file as memory will only be used now
                DisposeBitmap(btmByRefBitmapFile)
            Catch ex As Exception
                'Display and close
                gCloseApplicationWithErrorMessage("Memory is full or set very low. Please increase memory size. This application will close now.")
            End Try
        Else
            'Display and close
            gCloseApplicationWithErrorMessage("The " & strDirectory & strImageDirectory & " file is missing. This application will close now.")
        End If

    End Sub

    Private Sub DisposeBitmap(ByRef btmByRefBitmap As Bitmap)

        'Dispose
        btmByRefBitmap.Dispose()

        'Set
        btmByRefBitmap = Nothing

    End Sub

    Private Sub GDIDrawGraphics(ByRef graByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, pntPoint As Point)

        'Set options for fastest rendering
        SetGraphicOptions(graByRefGraphicsToDrawOn)

        'Draw
        graByRefGraphicsToDrawOn.DrawImageUnscaled(btmBitmapToDraw, pntPoint)

        'Dispose graphics
        DisposeGraphics(graByRefGraphicsToDrawOn)

    End Sub

    Private Sub SetGraphicOptions(graGraphics As Graphics)

        'Set options for fastest rendering
        With graGraphics
            .CompositingMode = Drawing2D.CompositingMode.SourceOver
            .CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
            .SmoothingMode = Drawing2D.SmoothingMode.None
            .InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
            .TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixel
            .PixelOffsetMode = Drawing2D.PixelOffsetMode.HighSpeed
        End With

    End Sub

    Private Sub DisposeGraphics(ByRef graByRefGraphics As Graphics)

        'Dispose graphics
        graByRefGraphics.Dispose()

        'Set
        graByRefGraphics = Nothing

    End Sub

    Private Sub GDIDrawGraphics(ByRef graByRefGraphicsToDrawOn As Graphics, btmBitmapToDraw As Bitmap, rectRectangle As Rectangle)

        'Set options for fastest rendering
        SetGraphicOptions(graByRefGraphicsToDrawOn)

        'Draw
        graByRefGraphicsToDrawOn.DrawImage(btmBitmapToDraw, rectRectangle) 'Scales it down by the rectangle

        'Dispose graphics
        DisposeGraphics(graByRefGraphicsToDrawOn)

    End Sub

    Private Sub LoadArrayPictureFilesAndCopyBitmapsIntoMemory(ByRef abtmByRefPictureFiles() As Bitmap, ByRef abtmByRefPictureMemories() As Bitmap,
                                                              strImageDirectoryWithoutFileType As String, strFileType As String,
                                                              intIndexStarting As Integer, intIndexIncrease As Integer)

        'Declare
        Dim intIndex As Integer = intIndexStarting

        'Loading bar pictures
        For intLoop As Integer = 0 To abtmByRefPictureFiles.GetUpperBound(0)
            'Load
            LoadPictureFileAndCopyBitmapIntoMemory(abtmByRefPictureFiles(intLoop), abtmByRefPictureMemories(intLoop),
                                                   strImageDirectoryWithoutFileType & intIndex.ToString & strFileType)
            'Increase
            intIndex += intIndexIncrease '0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100
        Next

    End Sub

    Private Sub LoadCreditsFilesAndCopyBitmapsIntoMemory(ByRef abtmByRefFiles() As Bitmap, ByRef abtmByRefMemories() As Bitmap,
                                                         strNameForCredits As String)

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

        'Play sound if not blocked by versus screen because back can be hit so many times there
        If Not blnDoNotPlayPressedSoundAfterPressingBack Then
            'Play sound
            udcButtonHoverSound.PlaySound(gintSoundVolume)
        Else
            'Reset
            blnDoNotPlayPressedSoundAfterPressingBack = False
        End If

    End Sub

    Private Sub ElapsedParagraph(sender As Object, e As EventArgs)

        'Check which type of wait
        Select Case strTypeOfParagraphWait

            Case "Single"
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
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
                Else
                    'Check which mode
                    Select Case intParagraphWaitMode
                        Case 0
                            'Set
                            intLoadingParagraphTextureIndex = 14
                        Case 1
                            'Set
                            intLoadingParagraphTextureIndex = 15
                        Case 2
                            'Set
                            intLoadingParagraphTextureIndex = 16
                        Case 3
                            'Set
                            intLoadingParagraphTextureIndex = 17
                            'Disable timer
                            tmrParagraph.Enabled = False
                    End Select
                End If

            Case "Versus"
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
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
                Else
                    'Check which mode
                    Select Case intParagraphWaitMode
                        Case 0
                            'Set
                            intVersusLoadingParagraphTextureIndex = 2
                        Case 1
                            'Set
                            intVersusLoadingParagraphTextureIndex = 3
                        Case 2
                            'Set
                            intVersusLoadingParagraphTextureIndex = 4
                        Case 3
                            'Set
                            intVersusLoadingParagraphTextureIndex = 5
                            'Disable timer
                            tmrParagraph.Enabled = False
                    End Select
                End If

        End Select

        'Increase or reset
        If intParagraphWaitMode = 3 Then
            intParagraphWaitMode = 0
        Else
            intParagraphWaitMode += 1
        End If

    End Sub

    Private Sub ElapsedFog(sender As Object, e As EventArgs)

        'Check to make sure screen is rendering
        If intCanvasMode = 0 Then

            'Check if OpenGL or GDI
            If Not blnOpenGL Then

                'Check if need to change interval
                If tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START Then
                    'Reset fog variables
                    ResetFogVariables(intFogBackRectangleMove, intFogBackX, blnProcessBackFog)
                    ResetFogVariables(intFogFrontRectangleMove, intFogFrontX, blnProcessFrontFog)
                    'Change interval
                    tmrFog.Interval = dblFOG_FRAME_DELAY
                End If

                'Increase fog back
                IncreaseFogVariables(intFogBackRectangleMove, intFOG_BACK_MEMORY_WIDTH, intFogBackX, blnProcessBackFog)

                'Increase fog front
                IncreaseFogVariables(intFogFrontRectangleMove, intFOG_FRONT_MEMORY_WIDTH, intFogFrontX, blnProcessFrontFog)

                'Check to reset
                If Not blnProcessBackFog And Not blnProcessFrontFog Then
                    tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START
                End If

            Else

                'Check if need to change interval
                If tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START Then
                    'Reset fog variables
                    intOpenGLBackFogX = 0
                    intOpenGLFrontFogX = 0
                    'Change interval
                    tmrFog.Interval = dblFOG_FRAME_DELAY
                End If

                'Increase fog variables
                intOpenGLBackFogX += intFOG_SPEED
                intOpenGLFrontFogX += intFOG_SPEED

                'Check to reset, check if both fogs have completely passed
                If intOpenGLBackFogX >= intFOG_BACK_MEMORY_WIDTH + intORIGINAL_SCREEN_WIDTH And
                intOpenGLFrontFogX >= intFOG_FRONT_MEMORY_WIDTH + intORIGINAL_SCREEN_WIDTH Then
                    tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START
                End If

            End If

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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDISetStoryPointerAndBitmap(41, 114, 0) 'First index
                Else
                    'Set
                    OpenGLSetStoryPointerAndBitmap(41, 114, 1) 'First index
                End If
            Case 1
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(1)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 2
                End If
            Case 2
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(2)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 3
                End If
            Case 3
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDIChangeStoryParagraphBitmapAndInterval(3, dblSTORY_SOUND_DELAY)
                Else
                    'Set
                    OpenGLChangeStoryParagraphBitmapAndInterval(4, dblSTORY_SOUND_DELAY)
                End If
            Case 4
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Play story paragraph 1 sound
                    PlayStoryParagraphSoundAndChangeInterval(0, dblSTORY_PARAGRAPH2_DELAY)
                Else
                    'Play story paragraph 1 sound
                    PlayStoryParagraphSoundAndChangeInterval(0, dblSTORY_PARAGRAPH2_DELAY)
                End If
            Case 5
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDIClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
                Else
                    'Set
                    OpenGLClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
                End If
            Case 6
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDISetStoryPointerAndBitmap(19, 84, 4)
                Else
                    'Set
                    OpenGLSetStoryPointerAndBitmap(19, 84, 5)
                End If
            Case 7
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(5)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 6
                End If
            Case 8
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(6)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 7
                End If
            Case 9
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDIChangeStoryParagraphBitmapAndInterval(7, dblSTORY_SOUND_DELAY)
                Else
                    'Set
                    OpenGLChangeStoryParagraphBitmapAndInterval(8, dblSTORY_SOUND_DELAY)
                End If
            Case 10
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Play story paragraph 2 sound
                    PlayStoryParagraphSoundAndChangeInterval(1, dblSTORY_PARAGRAPH3_DELAY)
                Else
                    'Play story paragraph 2 sound
                    PlayStoryParagraphSoundAndChangeInterval(1, dblSTORY_PARAGRAPH3_DELAY)
                End If
            Case 11
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDIClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
                Else
                    'Set
                    OpenGLClearStoryParagraphBitmapAndChangeInterval(dblLOADING_TRANSPARENCY_DELAY)
                End If
            Case 12
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDISetStoryPointerAndBitmap(15, 90, 8)
                Else
                    'Set
                    OpenGLSetStoryPointerAndBitmap(15, 90, 9)
                End If
            Case 13
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(9)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 10
                End If
            Case 14
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmStoryParagraph = abtmStoryParagraphMemories(10)
                Else
                    'Set
                    intStoryParagraphTextureIndex = 11
                End If
            Case 15
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    GDIChangeStoryParagraphBitmapAndInterval(11, dblSTORY_SOUND_DELAY)
                Else
                    'Set
                    OpenGLChangeStoryParagraphBitmapAndInterval(12, dblSTORY_SOUND_DELAY)
                End If
            Case 16
                'Play story paragraph 3 sound
                audcStoryParagraphSounds(2).PlaySound(gintSoundVolume)
                'Disable timer
                tmrStory.Enabled = False
        End Select

        'Increase
        intStoryWaitMode += 1

    End Sub

    Private Sub GDISetStoryPointerAndBitmap(intX As Integer, intY As Integer, intIndex As Integer)

        'Set
        pntStoryParagraph = New Point(intX, intY)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(intIndex)

    End Sub

    Private Sub OpenGLSetStoryPointerAndBitmap(intX As Integer, intY As Integer, intIndex As Integer)

        'Set
        pntStoryParagraph = New Point(intX, intY)

        'Set
        intStoryParagraphTextureIndex = intIndex

    End Sub

    Private Sub GDIChangeStoryParagraphBitmapAndInterval(intIndex As Integer, dblInterval As Double)

        'Set
        btmStoryParagraph = abtmStoryParagraphMemories(intIndex)

        'Set
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub OpenGLChangeStoryParagraphBitmapAndInterval(intIndex As Integer, dblInterval As Double)

        'Set
        intStoryParagraphTextureIndex = intIndex

        'Set
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub PlayStoryParagraphSoundAndChangeInterval(intIndex As Integer, dblInterval As Double)

        'Play story paragraph sound
        audcStoryParagraphSounds(intIndex).PlaySound(gintSoundVolume)

        'Change interval
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub GDIClearStoryParagraphBitmapAndChangeInterval(dblInterval As Double)

        'Set
        btmStoryParagraph = Nothing

        'Set
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub OpenGLClearStoryParagraphBitmapAndChangeInterval(dblInterval As Double)

        'Set
        intStoryParagraphTextureIndex = 0

        'Set
        tmrStory.Interval = dblInterval

    End Sub

    Private Sub ElapsedCredits(sender As Object, e As EventArgs)

        'Check which mode
        Select Case intCreditsWaitMode
            Case 0
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmJohnGonzales = abtmJohnGonzalesMemories(0)
                    btmZacharyStafford = abtmZacharyStaffordMemories(0)
                    btmCoryLewis = abtmCoryLewisMemories(0)
                Else
                    'Set
                    intJohnGonzlesTextureIndex = 1
                    intZacharyStaffordTextureIndex = 5
                    intCoryLewisTextureIndex = 9
                End If
            Case 1
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmJohnGonzales = abtmJohnGonzalesMemories(1)
                    btmZacharyStafford = abtmZacharyStaffordMemories(1)
                    btmCoryLewis = abtmCoryLewisMemories(1)
                Else
                    'Set
                    intJohnGonzlesTextureIndex = 2
                    intZacharyStaffordTextureIndex = 6
                    intCoryLewisTextureIndex = 10
                End If
            Case 2
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmJohnGonzales = abtmJohnGonzalesMemories(2)
                    btmZacharyStafford = abtmZacharyStaffordMemories(2)
                    btmCoryLewis = abtmCoryLewisMemories(2)
                Else
                    'Set
                    intJohnGonzlesTextureIndex = 3
                    intZacharyStaffordTextureIndex = 7
                    intCoryLewisTextureIndex = 11
                End If
            Case 3
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmJohnGonzales = abtmJohnGonzalesMemories(3)
                    btmZacharyStafford = abtmZacharyStaffordMemories(3)
                    btmCoryLewis = abtmCoryLewisMemories(3)
                Else
                    'Set
                    intJohnGonzlesTextureIndex = 4
                    intZacharyStaffordTextureIndex = 8
                    intCoryLewisTextureIndex = 12
                End If
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
            blnBackFromGame = True 'This makes the screen go back to menu and properly unload things
        End If

    End Sub

    Private Function blnMouseInRegion(pntMouse As Point, intImageWidth As Integer, intImageHeight As Integer, pntStartingPoint As Point) As Boolean

        'Return
        If pntMouse.X >= CInt(pntStartingPoint.X * gsngScreenWidthRatio) AndAlso
        pntMouse.X <= CInt((pntStartingPoint.X + intImageWidth) * gsngScreenWidthRatio) AndAlso
        pntMouse.Y >= CInt(pntStartingPoint.Y * gsngScreenHeightRatio) AndAlso
        pntMouse.Y <= CInt((pntStartingPoint.Y + intImageHeight) * gsngScreenHeightRatio) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Set
        blnRendering = False

        'Wait
        While thrRendering.IsAlive
        End While

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

        'Release OpenGL
        ReleaseOpenGL()

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

    Private Sub ReleaseOpenGL()

        'Release context
        Invoke(Sub() ReleaseDC(Handle, iptGDIContext)) 'Prevent cross-threading

        'Delete context
        DeleteDC(iptPermanentContext)

        'Make current
        wglMakeCurrent(IntPtr.Zero, IntPtr.Zero)

    End Sub

    Private Sub RemoveGameObjectsFromMemory(Optional blnStopScreamSound As Boolean = True)

        'Stop sound effects
        If blnStopScreamSound Then
            StopSoundObject(udcScreamSound)
        End If

        'Stop game background sounds
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            StopSoundObject(audcGameBackgroundSounds(intLoop))
        Next

        'Stop helicopter blade sound
        StopSoundObject(udcRotatingBladeSound)

        'Stop and dispose helicopter
        StopAndDisposeTimerObject(udcHelicopterGonzales)

        'Stop gagging sound
        For intLoop As Integer = 0 To audcSmallChainGagSounds.GetUpperBound(0)
            StopSoundObject(audcSmallChainGagSounds(intLoop))
        Next

        'Stop and dispose chained zombie
        StopAndDisposeTimerObject(gudcChainedZombie)

        'Stop water splash
        StopSoundObject(udcWaterSplashSound)

        'Stop face zombie eyes open sound
        StopSoundObject(udcFaceZombieEyesOpenSound)

        'Dispose face zombie
        StopAndDisposeTimerObject(gudcFaceZombie)

        'Stop reloading sound
        StopSoundObject(udcReloadingSound)

        'Stop and dispose character, remove handler
        StopAndDisposeTimerObject(udcCharacter)

        'Stop and dispose zombies
        If gaudcZombies IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombies.GetUpperBound(0)
                StopAndDisposeTimerObject(gaudcZombies(intLoop))
            Next
        End If

        'Stop and dispose versus character (host)
        StopAndDisposeTimerObject(udcCharacterOne)

        'Stop and dispose versus character (joiner)
        StopAndDisposeTimerObject(udcCharacterTwo)

        'Stop and dispose on screen word (Single player and hoster)
        StopAndDisposeTimerObject(udcOnScreenWordOne)

        'Stop and dispose on screen word (Joiner)
        StopAndDisposeTimerObject(udcOnScreenWordTwo)

        'Stop and dispose zombies (host)
        If gaudcZombiesOne IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesOne.GetUpperBound(0)
                StopAndDisposeTimerObject(gaudcZombiesOne(intLoop))
            Next
        End If

        'Stop and dispose zombies (joiner)
        If gaudcZombiesTwo IsNot Nothing Then
            For intLoop As Integer = 0 To gaudcZombiesTwo.GetUpperBound(0)
                StopAndDisposeTimerObject(gaudcZombiesTwo(intLoop))
            Next
        End If

    End Sub

    Private Sub StopSoundObject(udcSound As clsSound)

        'Check if the sound is not nothing
        If udcSound IsNot Nothing Then
            udcSound.StopSound()
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefHelicopterClass As clsHelicopter)

        'Check if the class is nothing
        If udcByRefHelicopterClass IsNot Nothing Then
            'Stop and dispose
            udcByRefHelicopterClass.StopAndDisposeTimer()
            udcByRefHelicopterClass = Nothing
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefChainedZombieClass As clsChainedZombie)

        'Check if the class is nothing
        If udcByRefChainedZombieClass IsNot Nothing Then
            'Stop and dispose
            udcByRefChainedZombieClass.StopAndDisposeTimer()
            udcByRefChainedZombieClass = Nothing
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefFaceZombieClass As clsFaceZombie)

        'Check if the class is nothing
        If udcByRefFaceZombieClass IsNot Nothing Then
            'Stop and dispose
            udcByRefFaceZombieClass.StopAndDisposeTimer()
            udcByRefFaceZombieClass = Nothing
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefCharacterClass As clsCharacter)

        'Check if the class is nothing
        If udcByRefCharacterClass IsNot Nothing Then
            'Stop and dispose
            udcByRefCharacterClass.StopAndDisposeTimer()
            udcByRefCharacterClass = Nothing
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefOnScreenWordClass As clsOnScreenWord)

        'Check if the class is nothing
        If udcByRefOnScreenWordClass IsNot Nothing Then
            'Stop and dispose
            udcByRefOnScreenWordClass.StopAndDisposeTimer()
            udcByRefOnScreenWordClass = Nothing
        End If

    End Sub

    Private Sub StopAndDisposeTimerObject(ByRef udcByRefZombieClass As clsZombie)

        'Check if the class is nothing
        If udcByRefZombieClass IsNot Nothing Then
            'Stop and dispose
            udcByRefZombieClass.StopAndDisposeTimer()
            udcByRefZombieClass = Nothing
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

        'Stop listening and empty variable
        StopListeningOrConnectingAndEmptyVariable(blnListening, thrListening)

        'Stop connecting and empty variable
        StopListeningOrConnectingAndEmptyVariable(blnConnecting, thrConnecting)

        'Empty write data
        CloseTCPIPData(gswClientData)

        'Empty read data
        CloseTCPIPData(srClientData)

        'Empty TCP server object
        EmptyTCPObject(tcplServer)

        'Empty TCP client object
        EmptyTCPObject(tcpcClient)

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

    Private Sub CloseTCPIPData(ByRef srByRefStreamWriter As IO.StreamWriter)

        'Check if it is nothing
        If srByRefStreamWriter IsNot Nothing Then
            srByRefStreamWriter.Close()
            srByRefStreamWriter = Nothing
        End If

    End Sub

    Private Sub CloseTCPIPData(ByRef srByRefStreamReader As IO.StreamReader)

        'Check if it is nothing
        If srByRefStreamReader IsNot Nothing Then
            srByRefStreamReader.Close()
            srByRefStreamReader = Nothing
        End If

    End Sub

    Private Sub StopListeningOrConnectingAndEmptyVariable(ByRef blnByRefListeningOrConnecting As Boolean,
                                                          ByRef thrByRefListeningOrConnecting As System.Threading.Thread)

        'Set
        blnByRefListeningOrConnecting = False

        'Listening or connecting TCP/IP
        If thrByRefListeningOrConnecting IsNot Nothing Then
            While thrByRefListeningOrConnecting.IsAlive
            End While
        End If

        'Set
        thrByRefListeningOrConnecting = Nothing

    End Sub

    Private Sub EmptyTCPObject(tcplServerClass As Net.Sockets.TcpListener)

        'Empty TCP object
        If tcplServerClass IsNot Nothing Then
            tcplServerClass.Stop()
            tcplServerClass = Nothing
        End If

    End Sub

    Private Sub EmptyTCPObject(tcpcClientClass As Net.Sockets.TcpClient)

        'Empty TCP object
        If tcpcClientClass IsNot Nothing Then
            tcpcClientClass.Close()
            tcpcClientClass = Nothing
        End If

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
            StopAndCloseSound(audcAmbianceSound(intLoop))
        Next

        'Button pressed
        StopAndCloseSound(udcButtonPressedSound)

        'Button hover
        StopAndCloseSound(udcButtonHoverSound)

        'Story paragraphs
        For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
            StopAndCloseSound(audcStoryParagraphSounds(intLoop))
        Next

        'Loaded game 100%
        StopAndCloseSound(udcFinishedLoading100PercentSound)

        'Loaded game press
        StopAndCloseSound(udcGameLoadedPressedSound)

        'Game backgrounds
        For intLoop As Integer = 0 To audcGameBackgroundSounds.GetUpperBound(0)
            StopAndCloseSound(audcGameBackgroundSounds(intLoop))
        Next

        'Scream
        StopAndCloseSound(udcScreamSound)

        'Gun shot
        StopAndCloseSound(udcGunShotSound)

        'Zombie deaths
        For intLoop As Integer = 0 To audcZombieDeathSounds.GetUpperBound(0)
            StopAndCloseSound(audcZombieDeathSounds(intLoop))
        Next

        'Reloading
        StopAndCloseSound(udcReloadingSound)

        'Step
        StopAndCloseSound(udcStepSound)

        'Water foot step left
        StopAndCloseSound(udcWaterFootStepLeftSound)

        'Water foot step right
        StopAndCloseSound(udcWaterFootStepRightSound)

        'Gravel foot step left
        StopAndCloseSound(udcGravelFootStepLeftSound)

        'Gravel foot step right
        StopAndCloseSound(udcGravelFootStepRightSound)

        'Opening metal door
        StopAndCloseSound(udcOpeningMetalDoorSound)

        'Light zap
        StopAndCloseSound(udcLightZapSound)

        'Zombie growl
        StopAndCloseSound(udcZombieGrowlSound)

        'Rotating blade
        StopAndCloseSound(udcRotatingBladeSound)

    End Sub

    Private Sub StopAndCloseSound(ByRef udcByRefSoundClass As clsSound)

        'Check if sound exists
        If udcByRefSoundClass IsNot Nothing Then
            udcByRefSoundClass.StopAndCloseSound()
            udcByRefSoundClass = Nothing
        End If

    End Sub

    Private Sub DataArrival(strData As String)

        'Notes: Data with TCP/IP is network streamed and not lined up perfectly, there must be delimiters, especially ending check, 
        '       Buffer Is an absolute must

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

                Case "7" 'Missed shot
                    'Shoot
                    udcCharacterTwo.Shoot()
                    'Show missed word
                    udcOnScreenWordTwo.ShowWord(90, 1, intON_SCREEN_WORD_MISSED_BLUE_X, intON_SCREEN_WORD_MISSED_BLUE_Y)

                Case "8" 'Not used by host

                Case "9" 'Not used by host

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

                Case "7" 'Missed shot
                    'Shoot
                    udcCharacterOne.Shoot()
                    'Show missed word
                    udcOnScreenWordOne.ShowWord(90, 0, intON_SCREEN_WORD_MISSED_RED_X, intON_SCREEN_WORD_MISSED_RED_Y)

                Case "8", "9" 'End game for joiner
                    'Set
                    blnPlayerWasPinned = True
                    'Set
                    gblnPreventKeyPressEvent = True
                    'Start black screen
                    BlackScreening(intBLACK_SCREEN_DEATH_DELAY)
                    'Stop reloading sound
                    udcReloadingSound.StopSound()
                    'Play
                    udcScreamSound.PlaySound(gintSoundVolume)
                    'Stop level music
                    audcGameBackgroundSounds(gintLevel - 1).StopSound()
                    'Game over, time stopped
                    StopTheStopWatches()
                    'Check the data
                    If strData = "8|" Then
                        intVersusWhoWonMode = 0 'Host lost
                    Else
                        intVersusWhoWonMode = 1 'Host won
                    End If

            End Select

        End If

    End Sub

    Private Function strGetBlockData(strData As String, Optional intArrayElement As Integer = 0) As String

        'Notes: Data looks like "X|~" or "X|String~" which equals "" or "String"

        'Remove ending delimiter
        strData = strData.Replace("~", "")

        'Return
        Return strSplitElement(strData, "|", intArrayElement)

    End Function

    Private Function strSplitElement(strToSplit As String, strDelimiter As String, intIndexToReturn As Integer) As String

        'Declare
        Dim astrTemp() As String = Split(strToSplit, strDelimiter)

        'Return
        Return astrTemp(intIndexToReturn)

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

    Private Sub GamesMismatched(strGameVersionFromConnecter As String)

        'Set
        strGameVersionFromConnection = strGameVersionFromConnecter

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(10, 0, False)

        'Start mismatch timer
        tmrGameMismatch.Enabled = True

    End Sub

    Private Sub ShowParagraphAndSetVariables(intCanvasModeToBe As Integer, intCanvasShowToBe As Integer, strTypeOfParagraphWaitType As String,
                                         blnGameIsVersusToBe As Boolean)

        'Change
        ShowNextScreenAndExitMenu(intCanvasModeToBe, intCanvasShowToBe)

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Empty paragraph first
            If blnGameIsVersusToBe Then
                'Set
                btmLoadingParagraphVersus = Nothing
            Else
                'Set
                btmLoadingParagraph = Nothing
            End If
        Else
            'Empty paragraph first
            If blnGameIsVersusToBe Then
                'Set
                intVersusLoadingParagraphTextureIndex = 0
            Else
                'Set
                intLoadingParagraphTextureIndex = 0
            End If
        End If

        'Set
        strTypeOfParagraphWait = strTypeOfParagraphWaitType

        'Set
        intParagraphWaitMode = 0

        'Use the paragraph timer
        tmrParagraph.Enabled = True

        'Set
        blnGameIsVersus = blnGameIsVersusToBe

        'Load the beginning game material thread
        SetAndStartThread(thrLoadBeginningGameMaterial, Sub() LoadBeginningGameMaterialThread())

    End Sub

    Private Sub LoadBeginningGameMaterialThread()

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            '0% GDI loaded
            btmLoadingBar = abtmLoadingBarPictureMemories(0)
        Else
            '0% OpenGL loaded
            intLoadingBarTextureIndex = 1
        End If

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

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            '10% GDI loaded
            btmLoadingBar = abtmLoadingBarPictureMemories(1)
        Else
            '10% OpenGL loaded
            intLoadingBarTextureIndex = 2
        End If

        'Check if previously loaded to the end
        If intMemoryLoadPosition = 187 Then
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
        gpntDestroyedBrickWall = New Point(1250, intORIGINAL_SCREEN_HEIGHT - 155)

        'Set
        strEntireLengthOfWords = ""

        'Set
        For intLoop As Integer = 0 To astrWordParts.GetUpperBound(0)
            astrWordParts(intLoop) = ""
        Next

        'Check if multiplayer game
        If blnGameIsVersus Then
            'Set
            intZombiesKilledIncreaseSpawnOne = 0
            intZombiesKilledIncreaseSpawnTwo = 0
            'Set
            intZombieKillsOne = 0
            intZombieKillsTwo = 0
            'Set
            intZombieKillsWaitingTwo = 0
            'Set
            strZombieKillBufferOne = ""
            strZombieKillBufferTwo = ""
            'Set
            intZombieIncreasedPinDistanceOne = 0
            intZombieIncreasedPinDistanceTwo = 0
        Else
            'Set
            intZombiesKilledIncreaseSpawn = 0
            'Set
            intZombieKills = 0
            'Set
            intZombieIncreasedPinDistance = 0
        End If

        'Set defaults for highscores and records used with the stats screen
        intZombieKillsCombined = 0
        strElapsedTime = ""
        intElapsedTime = 0
        intTypedWords = 0
        intWPM = 0
        blnSetStats = False
        blnComparedHighscore = False

    End Sub

    Private Sub LoadDefaultVariables()

        'Set
        If btmDeathScreen IsNot Nothing Then
            DisposeBitmap(btmDeathScreen)
        End If

        'Set
        blnRemovedGameObjectsFromMemory = False

        'Set
        blnStartedTimeElapsed = False

        'Set
        gpntGameBackground.X = 0

        'Set
        blnBlackScreenFinished = False

        'Set
        blnPlayerWasPinned = False

        'Set
        gblnPreventKeyPressEvent = False

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

        'Set
        blnWordWrong = False

        'Check
        If strEntireLengthOfWords = "" Then
            'Declare
            Dim intWordPartIndex As Integer = 1
            'Get first random word
            astrWordParts(0) = astrWords(gGetRandomNumber(0, astrWords.GetUpperBound(0)))
            'Get nine more words
            While intWordPartIndex <> 12
                'Get random words
                SetARandomWordInsideAWordPart(intWordPartIndex)
                'Increase
                intWordPartIndex += 1
            End While
        Else
            'Loop to Rearrange the word parts
            For intLoop As Integer = 0 To (astrWordParts.GetUpperBound(0) - 1)
                'Change
                astrWordParts(intLoop) = astrWordParts(intLoop + 1)
            Next
            'Reset last word
            astrWordParts(astrWordParts.GetUpperBound(0)) = ""
            'Get a last word
            SetARandomWordInsideAWordPart(astrWordParts.GetUpperBound(0))
            'Empty
            strEntireLengthOfWords = ""
        End If

        'Copy all words to the main string
        For intLoop As Integer = 0 To astrWordParts.GetUpperBound(0)
            'Include
            strEntireLengthOfWords &= astrWordParts(intLoop)
            'Add space
            If intLoop <> astrWordParts.GetUpperBound(0) Then
                strEntireLengthOfWords &= " "
            End If
        Next

    End Sub

    Private Sub SetARandomWordInsideAWordPart(intIndex As Integer)

        'Get a last word
        While astrWordParts(intIndex) = "" Or astrWordParts(intIndex) = astrWordParts(intIndex - 1)
            'Random word
            astrWordParts(intIndex) = astrWords(gGetRandomNumber(0, astrWords.GetUpperBound(0)))
        End While

    End Sub

    Private Sub LoadGameAudio()

        'Load game background sounds
        LoopLoadSoundClassIfNotExistent(audcGameBackgroundSounds, "Sounds\GameBackground", 1)

        'Load scream sound
        LoadSoundClassIfNotExistent(udcScreamSound, "Sounds\Scream.mp3", 1)

        'Load gun shot sound
        LoadSoundClassIfNotExistent(udcGunShotSound, "Sounds\GunShot.mp3", 10)

        'Load zombie death sounds
        LoopLoadSoundClassIfNotExistent(audcZombieDeathSounds, "Sounds\ZombieDeath", 10)

        'Load reloading sound
        LoadSoundClassIfNotExistent(udcReloadingSound, "Sounds\Reloading.mp3", 2) 'Incase multiplayer

        'Load step sound
        LoadSoundClassIfNotExistent(udcStepSound, "Sounds\Step.mp3", 6)

        'Load water foot left sound
        LoadSoundClassIfNotExistent(udcWaterFootStepLeftSound, "Sounds\WaterFootStepLeft.mp3", 3)

        'Load water foot right sound
        LoadSoundClassIfNotExistent(udcWaterFootStepRightSound, "Sounds\WaterFootStepRight.mp3", 3)

        'Load gravel foot left sound
        LoadSoundClassIfNotExistent(udcGravelFootStepLeftSound, "Sounds\GravelFootStepLeft.mp3", 3)

        'Load gravel foot right sound
        LoadSoundClassIfNotExistent(udcGravelFootStepRightSound, "Sounds\GravelFootStepRight.mp3", 3)

        'Load opening metal door sound
        LoadSoundClassIfNotExistent(udcOpeningMetalDoorSound, "Sounds\OpeningMetalDoor.mp3", 1) 'Happens only once during gameplay

        'Load light zap sound
        LoadSoundClassIfNotExistent(udcLightZapSound, "Sounds\LightZap.mp3", 5)

        'Load zombie growl sound
        LoadSoundClassIfNotExistent(udcZombieGrowlSound, "Sounds\ZombieGrowl.mp3", 1)

        'Load rotating blade sound of the helicopter
        LoadSoundClassIfNotExistent(udcRotatingBladeSound, "Sounds\RotatingBlade.mp3", 1)

        'Load chained zombie gag sound
        LoopLoadSoundClassIfNotExistent(audcSmallChainGagSounds, "Sounds\SmallChainGag", 2)

        'Load water splash sound
        LoadSoundClassIfNotExistent(udcWaterSplashSound, "Sounds\WaterSplash.mp3", 5)

        'Load face zombie eyes open sound
        LoadSoundClassIfNotExistent(udcFaceZombieEyesOpenSound, "Sounds\FaceZombieEyesOpen.mp3", 1)

    End Sub

    Private Sub LoopLoadSoundClassIfNotExistent(ByRef audcByRefSoundClass() As clsSound, strDirectoryPartial As String, intNumberOfSounds As Integer)

        'Load zombie death sounds
        For intLoop As Integer = 0 To audcByRefSoundClass.GetUpperBound(0)
            LoadSoundClassIfNotExistent(audcByRefSoundClass(intLoop), strDirectoryPartial & CStr(intLoop + 1) & ".mp3", intNumberOfSounds)
        Next

    End Sub

    Private Sub LoadSoundClassIfNotExistent(ByRef udcByRefSoundClass As clsSound, strDirectoryEnd As String, intNumberOfSounds As Integer)

        'Load sound class
        If udcByRefSoundClass Is Nothing Then
            udcByRefSoundClass = New clsSound(Me, strDirectory & strDirectoryEnd, intNumberOfSounds)
        End If

    End Sub

    Private Sub RestartLoadingGame(blnIncreaseMemoryPosition As Boolean)

        'Increase loading position
        If blnIncreaseMemoryPosition Then
            intMemoryLoadPosition += 1
        End If

        'Load the loading game thread
        SetAndStartThread(thrLoadingGame, Sub() LoadingGameThread())

    End Sub

    Private Sub LoadingGameThread()

        'Notes: This will load but wait until memory has completed before doing another load, each file is loaded individually, 
        '       waiting for a response from memory as well

        'Load continually
        Select Case intMemoryLoadPosition
            Case 0 To 4
                'File load array
                LoadGameFileWithIndexByDirectory(0, abtmGameBackgroundFiles, ablnGameBackgroundMemoriesCopied, "Images\Game Play\GameBackground",
                                                 "GameBackground", ".jpg")
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
                '20% loaded
                SetLoadingBarPercentage(48, 2)
            Case 49, 50
                'File load array
                LoadGameFileWithIndex(49, abtmCharacterStandRedFiles, ablnCharacterStandRedMemoriesCopied, "Images\Character\Red\Standing\", ".png")
            Case 51, 52
                'File load array
                LoadGameFileWithIndex(51, abtmCharacterShootRedFiles, ablnCharacterShootRedMemoriesCopied, "Images\Character\Red\Shoot Once\", ".png")
            Case 53 To 74
                'File load array
                LoadGameFileWithIndex(53, abtmCharacterReloadRedFiles, ablnCharacterReloadRedMemoriesCopied, "Images\Character\Red\Reload\", ".png")
                '30% loaded
                SetLoadingBarPercentage(74, 3)
            Case 75, 76
                'File load array
                LoadGameFileWithIndex(75, abtmCharacterStandBlueFiles, ablnCharacterStandBlueMemoriesCopied, "Images\Character\Blue\Standing\", ".png")
            Case 77, 78
                'File load array
                LoadGameFileWithIndex(77, abtmCharacterShootBlueFiles, ablnCharacterShootBlueMemoriesCopied, "Images\Character\Blue\Shoot Once\",
                                      ".png")
            Case 79 To 100
                'File load array
                LoadGameFileWithIndex(79, abtmCharacterReloadBlueFiles, ablnCharacterReloadBlueMemoriesCopied, "Images\Character\Blue\Reload\", ".png")
                '40% loaded
                SetLoadingBarPercentage(100, 4)
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
                '50% loaded
                SetLoadingBarPercentage(118, 5)
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
                '60% loaded
                SetLoadingBarPercentage(136, 6)
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
                '70% loaded
                SetLoadingBarPercentage(154, 7)
            Case 155 To 159
                'File load array
                LoadGameFileWithIndex(155, abtmHelicopterFiles, ablnHelicopterMemoriesCopied, "Images\Helicopters\HelicopterGonzales\", ".jpg")
            Case 160
                'File load
                LoadGameFile(btmAK47MagazineFile, blnAK47MagazineMemoryCopied, "Images\Game Play\AK47Magazine.png")
                '80% loaded
                SetLoadingBarPercentage(160, 8)
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
                LoadGameFileWithIndex(178, abtmFaceZombieFiles, ablnFaceZombieMemoriesCopied, "Images\Game Play\GameBackground2\Face Zombie\", ".png")
            Case 180
                'File load
                LoadGameFile(btmPipeValveFile, blnPipeValveMemoryCopied, "Images\Game Play\GameBackground2\PipeValve.png")
            Case 181
                'File load
                LoadGameFile(btmDestroyedBrickWallFile, blnDestroyedBrickWallMemoryCopied, "Images\Game Play\GameBackground1\DestroyedBrickWall.png")
            Case 182, 183
                'File load array
                LoadGameFileWithIndex(182, abtmOnScreenWordMissedRedFiles, ablnOnScreenWordMissedRedMemoriesCopied,
                                      "Images\Game Play\On Screen Words\Missed Red\", ".png")
            Case 184, 185
                'File load array
                LoadGameFileWithIndex(184, abtmOnScreenWordMissedBlueFiles, ablnOnScreenWordMissedBlueMemoriesCopied,
                                      "Images\Game Play\On Screen Words\Missed Blue\", ".png")
            Case 186
                '90% loaded
                SetLoadingBarPercentage(186, 9)
            Case 187 'Be careful changing this number, find = check if previously loaded to the end
                'Check if single player
                If Not blnGameIsVersus Then
                    'Character
                    udcCharacter = New clsCharacter(Me, intCHARACTER_X, intCHARACTER_Y, "udcCharacter", udcReloadingSound, udcGunShotSound,
                                                    udcStepSound, udcWaterFootStepLeftSound, udcWaterFootStepRightSound,
                                                    udcGravelFootStepLeftSound, udcGravelFootStepRightSound)
                    'On screen word
                    udcOnScreenWordOne = New clsOnScreenWord
                    'Zombies
                    LoadZombies("Level 1 Single Player")
                Else
                    'Load in a special way
                    If blnHost Then
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, intCHARACTER_HOSTER_X, intCHARACTER_HOSTER_Y, "udcCharacterOne", udcReloadingSound,
                                                           udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound, udcWaterFootStepRightSound,
                                                           udcGravelFootStepLeftSound, udcGravelFootStepRightSound) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, intCHARACTER_JOINER_X, intCHARACTER_JOINER_Y, "udcCharacterTwo", udcReloadingSound,
                                                           udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound, udcWaterFootStepRightSound,
                                                           udcGravelFootStepLeftSound, udcGravelFootStepRightSound, True) 'Join
                    Else
                        'Character one
                        udcCharacterOne = New clsCharacter(Me, intCHARACTER_HOSTER_X, intCHARACTER_HOSTER_Y, "udcCharacterOne", udcReloadingSound,
                                                           udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound, udcWaterFootStepRightSound,
                                                           udcGravelFootStepLeftSound, udcGravelFootStepRightSound, True) 'Host
                        'Character two
                        udcCharacterTwo = New clsCharacter(Me, intCHARACTER_JOINER_X, intCHARACTER_JOINER_Y, "udcCharacterTwo", udcReloadingSound,
                                                           udcGunShotSound, udcStepSound, udcWaterFootStepLeftSound, udcWaterFootStepRightSound,
                                                           udcGravelFootStepLeftSound, udcGravelFootStepRightSound) 'Join
                    End If
                    'On screen words
                    udcOnScreenWordOne = New clsOnScreenWord
                    udcOnScreenWordTwo = New clsOnScreenWord
                    'Zombies
                    LoadZombies("Level 1 Multiplayer")
                    'Set if hosting
                    If blnHost And Not blnReadyEarly Then
                        blnWaiting = True
                    End If
                End If
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    '100% GDI
                    btmLoadingBar = abtmLoadingBarPictureMemories(10)
                Else
                    '100% OpenGL
                    intLoadingBarTextureIndex = 11
                End If
                'Play sound
                udcFinishedLoading100PercentSound.PlaySound(gintSoundVolume)
                'Set
                blnFinishedLoading = True 'Means completely loaded
        End Select

    End Sub

    Private Sub LoadGameFileWithIndexByDirectory(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ablnMemoryCopied() As Boolean,
                                                 strImageDirectoryWithoutFolderNumber As String, strFileNameWithoutNumber As String,
                                                 strFileType As String)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract

        'File load
        LoadGameFile(abtmByRefFile(intIndexChanger), ablnMemoryCopied(intIndexChanger), strImageDirectoryWithoutFolderNumber &
                     CStr(intIndexChanger + 1) & "\" & strFileNameWithoutNumber & CStr(intIndexChanger + 1) & strFileType)

    End Sub

    Private Sub LoadGameFileWithIndex(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap, ablnMemoryCopied() As Boolean,
                                      strImageSubDirectory As String, strFileType As String)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract

        'File load
        LoadGameFile(abtmByRefFile(intIndexChanger), ablnMemoryCopied(intIndexChanger), strImageSubDirectory & CStr(intIndexChanger + 1) &
                     strFileType)

    End Sub

    Private Sub LoadGameFile(ByRef btmByRefFile As Bitmap, blnMemoryCopied As Boolean, strImageDirectory As String)

        'Load file
        If btmByRefFile Is Nothing And Not blnMemoryCopied Then
            'Attempt to load file
            If IO.File.Exists(strDirectory & strImageDirectory) Then
                'Attempt to load a game file
                Try
                    'Load a game file
                    btmByRefFile = New Bitmap(Image.FromFile(strDirectory & strImageDirectory))
                Catch ex As Exception
                    'Make sure the error was not caused by cancelling a thread
                    If ex.Message <> "Thread was being aborted." Then
                        'Display and close
                        gCloseApplicationWithErrorMessage("Memory is full or set very low. Please increase memory size. " &
                                                          "This application will close now.")
                    End If
                End Try
            Else
                'Display and close
                gCloseApplicationWithErrorMessage("The " & strDirectory & strImageDirectory & " file is missing. This application will close now.")
            End If
        End If

    End Sub

    Private Sub LoadZombies(strGameType As String)

        'Check if multiplayer or not
        Select Case strGameType

            Case "Level 1 Single Player"
                'Load single player zombies
                LoadZombiesByType(gaudcZombies, intZOMBIE_X, intZOMBIE_Y, "udcCharacter", True)

            Case "Level 1 Multiplayer"
                'Check if hosting, if not hosting then ghost like property the zombies
                LoadZombiesByType(gaudcZombiesOne, intZOMBIE_HOSTER_X, intZOMBIE_HOSTER_Y, "udcCharacterOne", False, Not blnHost)
                LoadZombiesByType(gaudcZombiesTwo, intZOMBIE_JOINER_X, intZOMBIE_JOINER_Y, "udcCharacterTwo", False, Not blnHost)

        End Select

    End Sub

    Private Sub LoadZombiesByType(ByRef audcByRefZombiesType() As clsZombie, intZombieX As Integer, intZombieY As Integer,
                                  strThisObjectNameCorrespondingToCharacter As String, blnAnimationStartRandom As Boolean,
                                  Optional blnImitation As Boolean = False)

        'Declare
        Dim intZombieXToBe As Integer = intZombieX
        Dim intZombieYToBe As Integer = intZombieY
        Dim intZombieSpeedIncrease As Integer = 0
        Dim intIncreaseAfterHorde As Integer = 0
        Dim intZombieBetweenZombieWidth As Integer = intZOMBIE_BETWEEN_ZOMBIE_WIDTH
        Dim blnTurnOffWidth As Boolean = False

        'Zombies
        For intLoop As Integer = 0 To (intNUMBER_OF_ZOMBIES_CREATED - 1)
            'Re-dim first
            ReDim Preserve audcByRefZombiesType(intLoop)
            'Setup zombie
            audcByRefZombiesType(intLoop) = New clsZombie(Me, intZombieXToBe + intZombieBetweenZombieWidth, intZombieYToBe, 5 +
                                            intZombieSpeedIncrease, strThisObjectNameCorrespondingToCharacter, audcZombieDeathSounds,
                                            udcWaterSplashSound, blnAnimationStartRandom, blnImitation)
            'Increase
            intIncreaseAfterHorde += 1
            'Increase zombie speed
            If intIncreaseAfterHorde = intINCREASE_ZOMBIE_SPEED_AFTER_DEATHS Then
                'Increase
                intZombieSpeedIncrease += 1
                'Reset
                intIncreaseAfterHorde = 0
            End If
            'Increase
            If intLoop >= intNUMBER_OF_ZOMBIES_AT_ONE_TIME Then
                'Check if set before
                If Not blnTurnOffWidth Then
                    'Set
                    blnTurnOffWidth = True
                    'Set
                    intZombieBetweenZombieWidth = 0
                    'Set
                    intZombieXToBe = intORIGINAL_SCREEN_WIDTH
                End If
            End If
            'Check
            If Not blnTurnOffWidth Then
                intZombieBetweenZombieWidth += intZOMBIE_BETWEEN_ZOMBIE_WIDTH
            End If
        Next

    End Sub

    Private Sub LoadingBitmapsIntoMemory()

        'Note: This waits for files to be completed and then loads it into memory

        'Load into memory
        Select Case intMemoryLoadPosition
            Case 0 To 4
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(0, abtmGameBackgroundFiles, abtmGameBackgroundMemories, 0,
                                                                aiptGameBackgroundTextures, intINTERNAL_FORMAT_RGB,
                                                                ablnGameBackgroundMemoriesCopied) 'OpenGL 0 to 4
            Case 5
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmWordBarFile, btmWordBarMemory, iptWordBarTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnWordBarMemoryCopied)
            Case 6, 7
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(6, abtmCharacterStandFiles, gabtmCharacterStandMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterStandMemoriesCopied) 'OpenGL 0 to 1
            Case 8, 9
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(8, abtmCharacterShootFiles, gabtmCharacterShootMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterShootMemoriesCopied) 'OpenGL 2 to 3
            Case 10 To 31
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(10, abtmCharacterReloadFiles, gabtmCharacterReloadMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterReloadMemoriesCopied) 'OpenGL 4 to 25
            Case 32 To 48
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(32, abtmCharacterRunningFiles, gabtmCharacterRunningMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterRunningMemoriesCopied) 'OpenGL 26 to 42
            Case 49, 50
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(49, abtmCharacterStandRedFiles, gabtmCharacterStandRedMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterStandRedMemoriesCopied) 'OpenGL 43 to 44
            Case 51, 52
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(51, abtmCharacterShootRedFiles, gabtmCharacterShootRedMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterShootRedMemoriesCopied) 'OpenGL 45 to 46
            Case 53 To 74
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(53, abtmCharacterReloadRedFiles, gabtmCharacterReloadRedMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterReloadRedMemoriesCopied) 'OpenGL 47 to 68
            Case 75, 76
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(75, abtmCharacterStandBlueFiles, gabtmCharacterStandBlueMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterStandBlueMemoriesCopied) 'OpenGL 69 to 70
            Case 77, 78
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(77, abtmCharacterShootBlueFiles, gabtmCharacterShootBlueMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterShootBlueMemoriesCopied) 'OpenGL 71 to 72
            Case 79 To 100
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(79, abtmCharacterReloadBlueFiles, gabtmCharacterReloadBlueMemories, 6,
                                                                aiptCharacterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnCharacterReloadBlueMemoriesCopied) 'OpenGL 73 to 94
            Case 101 To 104
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(101, abtmZombieWalkFiles, gabtmZombieWalkMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieWalkMemoriesCopied) 'OpenGL 0 to 3
            Case 105 To 110
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(105, abtmZombieDeath1Files, gabtmZombieDeath1Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeath1MemoriesCopied) 'OpenGL 4 to 9
            Case 111 To 116
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(111, abtmZombieDeath2Files, gabtmZombieDeath2Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeath2MemoriesCopied) 'OpenGL 10 to 15
            Case 117, 118
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(117, abtmZombiePinFiles, gabtmZombiePinMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombiePinMemoriesCopied) 'OpenGL 16 to 17
            Case 119 To 122
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(119, abtmZombieWalkRedFiles, gabtmZombieWalkRedMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieWalkRedMemoriesCopied) 'OpenGL 18 to 21
            Case 123 To 128
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(123, abtmZombieDeathRed1Files, gabtmZombieDeathRed1Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeathRed1MemoriesCopied) 'OpenGL 22 to 27
            Case 129 To 134
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(129, abtmZombieDeathRed2Files, gabtmZombieDeathRed2Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeathRed2MemoriesCopied) 'OpenGL 28 to 33
            Case 135, 136
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(135, abtmZombiePinRedFiles, gabtmZombiePinRedMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombiePinRedMemoriesCopied) 'OpenGL 34 to 35
            Case 137 To 140
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(137, abtmZombieWalkBlueFiles, gabtmZombieWalkBlueMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieWalkBlueMemoriesCopied) 'OpenGL 36 to 39
            Case 141 To 146
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(141, abtmZombieDeathBlue1Files, gabtmZombieDeathBlue1Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeathBlue1MemoriesCopied) 'OpenGL 40 to 45
            Case 147 To 152
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(147, abtmZombieDeathBlue2Files, gabtmZombieDeathBlue2Memories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombieDeathBlue2MemoriesCopied) 'OpenGL 46 to 51
            Case 153, 154
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(153, abtmZombiePinBlueFiles, gabtmZombiePinBlueMemories, 101,
                                                                aiptZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnZombiePinBlueMemoriesCopied) 'OpenGL 52 to 53
            Case 155 To 159
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(155, abtmHelicopterFiles, gabtmHelicopterMemories, 155,
                                                                aiptHelicopterTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnHelicopterMemoriesCopied) 'OpenGL 0 to 4
            Case 160
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmAK47MagazineFile, btmAK47MagazineMemory, iptAK47MagazineTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnAK47MagazineMemoryCopied)
            Case 161
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmDeathOverlayFile, btmDeathOverlayMemory, iptDeathOverlayTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnDeathOverlayMemoryCopied)
            Case 162
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmWinOverlayFile, btmWinOverlayMemory, iptWinOverlayTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnWinOverlayMemoryCopied)
            Case 163
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmYouWonFile, btmYouWonMemory, iptYouWonTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnYouWonMemoryCopied)
            Case 164
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmYouLostFile, btmYouLostMemory, iptYouLostTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnYouLostMemoryCopied)
            Case 165 To 167
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(165, abtmBlackScreenFiles, abtmBlackScreenMemories, 165,
                                                                aiptBlackScreenTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnBlackScreenMemoriesCopied) 'OpenGL 0 to 2
            Case 168 To 170
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(168, abtmPath1Files, abtmPath1Memories, 168, aiptPathTextures,
                                                                intINTERNAL_FORMAT_RGB, ablnPath1MemoriesCopied) 'OpenGL 0 to 2
            Case 171 To 173
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(171, abtmPath2Files, abtmPath2Memories, 168, aiptPathTextures,
                                                                intINTERNAL_FORMAT_RGB, ablnPath2MemoriesCopied) 'OpenGL 3 to 5
            Case 174 To 176
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(174, abtmChainedZombieFiles, gabtmChainedZombieMemories, 174,
                                                                aiptChainedZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnChainedZombieMemoriesCopied) 'OpenGL 0 to 2
            Case 177
                'Memory load, this water is a clone copy process in the engine
                CopyBitmapIntoMemoryAfterDrawingScreenNoOpenGL(btmGameBackground2WaterFile, btmGameBackground2WaterMemory,
                                                               blnGameBackground2WaterMemoryCopied)
            Case 178, 179
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(178, abtmFaceZombieFiles, gabtmFaceZombieMemories, 178,
                                                                aiptFaceZombieTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnFaceZombieMemoriesCopied) 'OpenGL 0 to 1
            Case 180
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmPipeValveFile, btmPipeValveMemory, iptPipeValveTexture,
                                                       intINTERNAL_FORMAT_RGBA, blnPipeValveMemoryCopied)
            Case 181
                'Memory load
                CopyBitmapIntoMemoryAfterDrawingScreen(btmDestroyedBrickWallFile, btmDestroyedBrickWallMemory,
                                                       iptDestroyedBrickWallTexture, intINTERNAL_FORMAT_RGBA,
                                                       blnDestroyedBrickWallMemoryCopied)
            Case 182, 183
                'Memory load array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(182, abtmOnScreenWordMissedRedFiles, gabtmOnScreenWordMissedRedMemories, 182,
                                                                aiptOnScreenWordMissedTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnOnScreenWordMissedRedMemoriesCopied) 'OpenGL 0 to 1
            Case 184, 185
                'Memory laod array
                CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(184, abtmOnScreenWordMissedBlueFiles, gabtmOnScreenWordMissedBlueMemories, 182,
                                                                aiptOnScreenWordMissedTextures, intINTERNAL_FORMAT_RGBA,
                                                                ablnOnScreenWordMissedBlueMemoriesCopied) 'OpenGL 2 to 3
            Case 186 'Be careful changing this number, find = check if previously loaded to the end
                'Set
                If btmGameBackground2WaterMemory IsNot Nothing Then
                    intWaterHeight = btmGameBackground2WaterMemory.Height
                End If
                'Memory copy level
                CopyLevelBitmap(btmGameBackgroundMemory, abtmGameBackgroundMemories(0))
                'Restart the thread, should do the last 100%
                RestartLoadingGame(True)
        End Select

    End Sub

    Private Sub CopyBitmapIntoMemoryAfterDrawingScreenWithIndex(intIndexToSubtract As Integer, ByRef abtmByRefFile() As Bitmap,
                                                                ByRef abtmByRefMemory() As Bitmap, intIndexToSubtractOpenGL As Integer,
                                                                ByRef iptByRefTexture() As IntPtr, intInternalFormat As Integer,
                                                                ByRef ablnByRefMemoryCopied() As Boolean,
                                                                Optional blnRotateFlip As Boolean = False,
                                                                Optional rftRotateFlipType As System.Drawing.RotateFlipType =
                                                                RotateFlipType.Rotate180FlipNone)

        'Declare
        Dim intIndexChanger As Integer = intMemoryLoadPosition - intIndexToSubtract
        Dim intIndexChangerOpenGL As Integer = intMemoryLoadPosition - intIndexToSubtractOpenGL

        'File load
        CopyBitmapIntoMemoryAfterDrawingScreen(abtmByRefFile(intIndexChanger), abtmByRefMemory(intIndexChanger),
                                               iptByRefTexture(intIndexChangerOpenGL), intInternalFormat,
                                               ablnByRefMemoryCopied(intIndexChanger),
                                               blnRotateFlip, rftRotateFlipType)

    End Sub

    Private Sub CopyBitmapIntoMemoryAfterDrawingScreen(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap,
                                                       ByRef iptByRefTexture As IntPtr, intInternalFormat As Integer,
                                                       ByRef blnByRefBitmapMemoryCopied As Boolean, Optional blnRotateFlip As Boolean = False,
                                                       Optional rftRotateFlipType As System.Drawing.RotateFlipType = RotateFlipType.Rotate180FlipNone)

        'Load bitmap into memory, see if file was loaded and if haven't already loaded into memory
        If btmByRefBitmapFile IsNot Nothing And Not blnByRefBitmapMemoryCopied Then

            'Prepare to load image into memory
            Try
                'New image into memory
                btmByRefBitmapMemory = New Bitmap(btmByRefBitmapFile.Width, btmByRefBitmapFile.Height, Imaging.PixelFormat.Format32bppPArgb)
                'Memory copy
                GDIDrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)
                'Check if need to rotate flip
                If blnRotateFlip Then
                    btmByRefBitmapMemory.RotateFlip(rftRotateFlipType)
                End If
                'Dispose file lock
                DisposeBitmap(btmByRefBitmapFile)
            Catch ex As Exception
                'Make sure the error was not caused by cancelling a thread
                If ex.Message <> "Thread was being aborted." Then
                    'Display and close
                    gCloseApplicationWithErrorMessage("Memory is full or set very low. Please increase memory size. This application will close now.")
                End If
            End Try

            'Try to load OpenGL texture
            CanLoadOpenGLTexture(iptByRefTexture, btmByRefBitmapMemory, intInternalFormat)

            'Set
            blnByRefBitmapMemoryCopied = True

        End If

        'Check to see if it was already copied
        If blnByRefBitmapMemoryCopied Then

            'Restart loading game thread
            RestartLoadingGame(True)

        End If

    End Sub

    Private Sub CanLoadOpenGLTexture(ByRef iptByRefTexture As IntPtr, btmBitmapMemory As Bitmap, intInternalFormat As Integer)

        'Check if can load and then load texture
        If blnOpenGLLoaded Then
            'Load texture
            LoadOpenGLTexture(iptByRefTexture, btmBitmapMemory, intInternalFormat)
        End If

    End Sub

    Private Sub CopyBitmapIntoMemoryAfterDrawingScreenNoOpenGL(ByRef btmByRefBitmapFile As Bitmap, ByRef btmByRefBitmapMemory As Bitmap,
                                                               ByRef blnByRefBitmapMemoryCopied As Boolean,
                                                               Optional blnRotateFlip As Boolean = False,
                                                               Optional rftRotateFlipType As System.Drawing.RotateFlipType =
                                                               RotateFlipType.Rotate180FlipNone)

        'Load bitmap into memory, see if file was loaded and if haven't already loaded into memory
        If btmByRefBitmapFile IsNot Nothing And Not blnByRefBitmapMemoryCopied Then

            'Prepare to load image into memory
            Try
                'New image into memory
                btmByRefBitmapMemory = New Bitmap(btmByRefBitmapFile.Width, btmByRefBitmapFile.Height, Imaging.PixelFormat.Format32bppPArgb)
                'Memory copy
                GDIDrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmByRefBitmapFile, pntTopLeft)
                'Check if need to rotate flip
                If blnRotateFlip Then
                    btmByRefBitmapMemory.RotateFlip(rftRotateFlipType)
                End If
                'Dispose file lock
                DisposeBitmap(btmByRefBitmapFile)
            Catch ex As Exception
                'Make sure the error was not caused by cancelling a thread
                If ex.Message <> "Thread was being aborted." Then
                    'Display and close
                    gCloseApplicationWithErrorMessage("Memory is full or set very low. Please increase memory size. This application will close now.")
                End If
            End Try

            'Set
            blnByRefBitmapMemoryCopied = True

        End If

        'Check to see if it was already copied
        If blnByRefBitmapMemoryCopied Then

            'Restart loading game thread
            RestartLoadingGame(True)

        End If

    End Sub

    Private Sub SetLoadingBarPercentage(intMemoryLoadPositionToCheck As Integer, intPictureMemoryIndex As Integer)

        'Check if need to set loading bar percentage
        If intMemoryLoadPosition = intMemoryLoadPositionToCheck Then
            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'GDI set if not set already
                btmLoadingBar = abtmLoadingBarPictureMemories(intPictureMemoryIndex)
            Else
                'OpenGL set if not set already
                intLoadingBarTextureIndex = intPictureMemoryIndex + 1
            End If
        End If

    End Sub

    Private Sub CopyLevelBitmap(ByRef btmByRefBitmapMemory As Bitmap, btmBitmapMemoryToCopy As Bitmap)

        'Dispose old
        If btmByRefBitmapMemory IsNot Nothing Then
            DisposeBitmap(btmByRefBitmapMemory)
        End If

        'Prepare to copy level
        Try
            'Make new image into memory
            btmByRefBitmapMemory = New Bitmap(btmBitmapMemoryToCopy.Width, btmBitmapMemoryToCopy.Height, Imaging.PixelFormat.Format32bppPArgb)
            'Memory copy
            GDIDrawGraphics(Graphics.FromImage(btmByRefBitmapMemory), btmBitmapMemoryToCopy, pntTopLeft)
        Catch ex As Exception
            'Make sure the error was not caused by cancelling a thread
            If ex.Message <> "Thread was being aborted." Then
                'Display and close
                gCloseApplicationWithErrorMessage("Could not copy the level bitmap, memory might be low. This application will close now.")
            End If
        End Try

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

        'Create stop watches
        CreateStopWatches()

        'Play background sound music
        audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(Math.Round(gintSoundVolume / 4)), True)

    End Sub

    Private Sub CreateStopWatches()

        'Create the time type stop watch
        gswhTimeTyped = New Stopwatch

        'Create the elapsed time in game
        swhTimeElapsed = New Stopwatch

    End Sub

    Private Sub ShowNextScreenAndExitMenu(intCanvasModeToSet As Integer, intCanvasShowToSet As Integer)

        'Set
        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(intCanvasModeToSet, intCanvasShowToSet)

        'Stop sound
        audcAmbianceSound(0).StopSound()

        'Disable fog timer
        tmrFog.Enabled = False

        'Stop process
        If Not blnOpenGL Then
            'Reset for GDI
            blnProcessBackFog = False
            blnProcessFrontFog = False
        Else
            'Reset for OpenGL
            intOpenGLFrontFogX = 0
            intOpenGLBackFogX = 0
        End If

    End Sub

    Private Sub SetAndStartThread(ByRef thrByRefThread As System.Threading.Thread, subSub As SubDelegate)

        'Set the thread
        thrByRefThread = New System.Threading.Thread(New System.Threading.ThreadStart(Sub() subSub()))

        'Start thread
        thrByRefThread.Start()

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

    Private Sub StopTheStopWatches()

        'Stop time typed
        gswhTimeTyped.Stop()

        'Stop time elapsed in game
        swhTimeElapsed.Stop()

    End Sub

    Private Sub ConnectionLost() 'Need this because it is an address of a multi-thread

        'Go back to menu
        GeneralBackButtonClick(New Point(-1, -1), False, True) 'Point doesn't matter here, forcing back button activity

    End Sub

    Private Sub BlackScreening(intScreenDelay As Integer)

        'Set
        intBlackScreenWaitMode = 0

        'Set timer delay
        tmrBlackScreen.Interval = intScreenDelay

        'Enable
        tmrBlackScreen.Enabled = True

    End Sub

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Hide the ugly gray screen
        Hide()

        'Check if multi-threaded or not
        If Not blnThreadSupported Then
            'Display
            MessageBox.Show("This computer doesn't support multi-threading. This application will close now.", "Last Stand", MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            'Exit
            End
        Else
            'Force focus
            Focus()
            'Enable FPS timer, this is the only place to enable it first, if enabled sooner, it won't work
            tmrFPS.Enabled = True
            'Load necessary sounds for the beginning engine
            LoadNecessarySoundsForBeginningEngine() 'This needs a handle from the form window
            'Get highscores early because grabbing information from the database access files is slow
            LoadHighscoresIntoAString()
            'Set percentage multiplers for screen modes
            gsngScreenWidthRatio = CSng(ClientSize.Width) / CSng(intORIGINAL_SCREEN_WIDTH)
            gsngScreenHeightRatio = CSng(ClientSize.Height) / CSng(intORIGINAL_SCREEN_HEIGHT)
            'Menu sound
            audcAmbianceSound(0).PlaySound(gintSoundVolume, True)
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, ClientSize.Width, ClientSize.Height) 'Full screen
            'Set
            blnRendering = True
            'Start rendering
            thrRendering.Start()
            'Set border
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            'Enable fog
            tmrFog.Enabled = True 'Already interval dblFOG_FRAME_WAIT_TO_START and variables are empty
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

    Private Sub LoadHighscoresIntoAString()

        'Declare
        Dim strSQL As String = "SELECT * FROM HighscoresTable ORDER BY intRank ASC"
        Dim strConnection As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirectory &
                                      "Highscores.mdb;Jet OLEDB:Database Password=32x1aZ@!"
        Dim dtbDataTable As New DataTable

        'Fill data table
        FillDataTable(strSQL, strConnection, dtbDataTable)

        'Load string
        For intLoop As Integer = 0 To dtbDataTable.Rows.Count - 1
            'Set highscore array
            astrHighscores(intLoop) = dtbDataTable.Rows(intLoop).Item(0).ToString & ". " &
                                      dtbDataTable.Rows(intLoop).Item(1).ToString & strDRAW_STATS_SECONDS_WORD & " " &
                                      dtbDataTable.Rows(intLoop).Item(2).ToString & " WPM " &
                                      dtbDataTable.Rows(intLoop).Item(3).ToString & " zombie kills - " &
                                      dtbDataTable.Rows(intLoop).Item(4).ToString
        Next

    End Sub

    Private Sub FillDataTable(strSQL As String, strConnection As String, dtbDataTable As DataTable)

        'Load into a data table first
        Using odaDataAdapter As New OleDb.OleDbDataAdapter(strSQL, strConnection)
            'Fill data table
            odaDataAdapter.Fill(dtbDataTable)
        End Using

    End Sub

    Private Sub Rendering()

        'Declare
        Static sblnLoadOpenGLOnce As Boolean = True

        'Loop
        While blnRendering
            'Check if need to setup OpenGL into the window
            If blnOpenGL Then
                'Reset
                glLoadIdentity()
                'Empty variables if back from game
                EmptyVariablesInRender()
                'Check mode before showing objects and graphics
                CheckCanvasModeAndShowObjectsAndGraphics()
                'Check to draw hot keys
                CheckToDrawHotKeys()
                'Setup view port
                glViewport(0, 0, ClientSize.Width, ClientSize.Height)
                'Swap buffers
                SwapBuffers(iptGDIContext)
            Else
                'Empty variables if back from game
                EmptyVariablesInRender()
                'Check mode before showing objects and graphics
                CheckCanvasModeAndShowObjectsAndGraphics()
                'Check to draw hot keys
                CheckToDrawHotKeys()
                'Paint on screen with rectangle
                GDIDrawGraphics(Me.CreateGraphics(), btmCanvas, rectFullScreen)
            End If
            'If changing screen, we must change resolution in this thread or else strange things happen
            ScreenResolutionChanged()
            'Load parts of the game here after a screen has been shown, slowly load, this only happens during the loading screens
            LoadBitmapsIntoMemoryDuringRender()
            'Check if need to load new level
            CheckIfNeedToLoadNewLevel()
            'Check if GDI and need to load for first time
            If Not blnOpenGL Then
                'Load OpenGL once, it is here so the screen won't show until after load
                If sblnLoadOpenGLOnce Then
                    'Set
                    sblnLoadOpenGLOnce = False
                    'Try and load the OpenGL to this window
                    Try
                        'Initialize OpenGL
                        OpenGLInitialize()
                        'Set
                        blnOpenGLLoaded = True
                    Catch ex As Exception
                        'No debug
                    End Try
                    'Un-hide
                    Invoke(Sub() Show()) 'Prevent cross-threading
                End If
            End If
            'Increase frames per second
            intFPSCalculated += 1
        End While

        'Dispose of canvas bitmap
        DisposeBitmap(btmCanvas)

    End Sub

    Private Sub EmptyVariablesInRender()

        'Check if back from game
        If blnBackFromGame Then
            'Disable timer
            tmrParagraph.Enabled = False
            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'Set
                btmLoadingParagraph = Nothing
                'Set
                btmLoadingParagraphVersus = Nothing
                '0% loaded
                btmLoadingBar = abtmLoadingBarPictureMemories(0)
                'Reset credits
                btmJohnGonzales = Nothing
                btmZacharyStafford = Nothing
                btmCoryLewis = Nothing
                'Reset
                btmStoryParagraph = Nothing
            Else
                'Set
                intLoadingParagraphTextureIndex = 0
                'Set
                intVersusLoadingParagraphTextureIndex = 0
                '0% loaded
                intLoadingBarTextureIndex = 1
                'Reset credits
                intJohnGonzlesTextureIndex = 0
                intZacharyStaffordTextureIndex = 0
                intCoryLewisTextureIndex = 0
                'Reset
                intStoryParagraphTextureIndex = 0
            End If
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
            ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(0, 0, blnPlayPressedSoundNow)
            'Set
            blnPlayPressedSoundNow = False
            'Disable timer
            tmrBlackScreen.Enabled = False
            'Set
            blnFinishedLoading = False
            'Disable timer
            tmrCredits.Enabled = False
            'Disable story timer
            tmrStory.Enabled = False
            'Set interval
            tmrStory.Interval = dblLOADING_TRANSPARENCY_DELAY
            'Set
            intStoryWaitMode = 0
            'Disable mismatch timer
            tmrGameMismatch.Enabled = False
            'Stop sounds
            For intLoop As Integer = 0 To audcStoryParagraphSounds.GetUpperBound(0)
                audcStoryParagraphSounds(intLoop).StopSound()
            Next
            'Change interval
            tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START
            'Start fog
            tmrFog.Enabled = True
            'Set
            blnBackFromGame = False
        End If

    End Sub

    Private Sub CheckCanvasModeAndShowObjectsAndGraphics()

        'Check mode before showing objects and graphics
        Select Case intCanvasMode
            Case 0 'Menu
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIRenderMenuScreen()
                Else
                    'OpenGL
                    OpenGLRenderMenuScreen()
                End If
            Case 1 'Options screen
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIRenderOptionsScreen()
                Else
                    'OpenGL
                    OpenGLRenderOptionsScreen()
                End If
            Case 2 'Play game, loading screen first
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDILoadingGameScreen()
                Else
                    'OpenGL
                    OpenGLLoadingGameScreen()
                End If
            Case 3 'Playing the game
                StartedGameScreen()
            Case 4 'Highscores screen
                HighscoresScreen() 'Has both GDI and OpenGL
            Case 5 'Credits screen
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDICreditsScreen()
                Else
                    'OpenGL
                    OpenGLCreditsScreen()
                End If
            Case 6 'Versus screen
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIVersusScreen()
                Else
                    'OpenGL
                    OpenGLVersusScreen()
                End If
            Case 7 'Loading versus connected game
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDILoadingVersusConnectedScreen()
                Else
                    'OpenGL
                    OpenGLLoadingVersusConnectedScreen()
                End If
            Case 8 'Playing versus game
                StartedVersusGameScreen()
            Case 9 'Story
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIStoryScreen()
                Else
                    'OpenGL
                    OpenGLStoryScreen()
                End If
            Case 10 'Mismatch game versions
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIGameVersionMismatch()
                Else
                    'OpenGL
                    OpenGLGameVersionMismatch()
                End If
            Case 11, 12 'Path system
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'GDI
                    GDIPathChoices()
                Else
                    'OpenGL
                    OpenGLPathChoices()
                End If
        End Select

    End Sub

    Private Sub CheckToDrawHotKeys()

        'Check for god mode
        If blnGodMode Then
            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'Draw FPS counter
                GDIDrawText(btmCanvas, 25, "RED", "GOD MODE SINGLE PLAYER", New Point(10, 10), 2)
            Else
                'Draw FPS counter
                OpenGLDrawText(25, "RED", "GOD MODE SINGLE PLAYER", New Point(10, 10), 2)
            End If
        Else
            'Check to draw FPS counter
            If blnDrawFPS Then
                'Make sure it is not the option screen
                If intCanvasMode <> 1 Then
                    'Check if GDI or OpenGL
                    If Not blnOpenGL Then
                        'Draw FPS counter
                        GDIDrawText(btmCanvas, 25, "RED", "FPS: " & intFPSDisplay.ToString, New Point(10, 10), 2)
                    Else
                        'Draw FPS counter
                        OpenGLDrawText(25, "RED", "FPS: " & intFPSDisplay.ToString, New Point(10, 10), 2)
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub GDIDrawText(ByRef btmByRefBitmapToDrawOn As Bitmap, intFontSize As Integer, strColor As String, strText As String,
                            pntPoint As Point, Optional intPixelWidthSubtractionSpacing As Integer = 0)

        'Declare
        Dim intReferenceWidth As Integer = 0
        Dim strParsedOneCharacter As String = ""
        Dim intColorIndexSkip As Integer = 0
        Dim blnTwoCharacters As Boolean = False

        'Check if different than white
        ChangeTextColorIndex(strColor, intColorIndexSkip)

        'Loop until all letters are drawn
        While strText.Length > 0
            'Parse one character
            ParseOneCharacter(strText, strParsedOneCharacter, blnTwoCharacters)
            'Check font size
            Select Case intFontSize
                Case 25
                    'Draw character
                    GDIDrawCharacter(strParsedOneCharacter, btmByRefBitmapToDrawOn, pntPoint, intReferenceWidth, abtmTextSize25TNRMemories,
                                     intColorIndexSkip, intFontSize)
                Case 36
                    'Draw character
                    GDIDrawCharacter(strParsedOneCharacter, btmByRefBitmapToDrawOn, pntPoint, intReferenceWidth, abtmTextSize36TNRMemories,
                                     intColorIndexSkip, intFontSize)
                Case 42
                    'Draw character
                    GDIDrawCharacter(strParsedOneCharacter, btmByRefBitmapToDrawOn, pntPoint, intReferenceWidth, abtmTextSize42TNRMemories,
                                     intColorIndexSkip, intFontSize)
                Case 55
                    'Draw character
                    GDIDrawCharacter(strParsedOneCharacter, btmByRefBitmapToDrawOn, pntPoint, intReferenceWidth, abtmTextSize55TNRMemories,
                                     intColorIndexSkip, intFontSize)
                Case 72
                    'Draw character
                    GDIDrawCharacter(strParsedOneCharacter, btmByRefBitmapToDrawOn, pntPoint, intReferenceWidth, abtmTextSize72TNRMemories,
                                     intColorIndexSkip, intFontSize)
            End Select
            'Check for two characters
            CheckTwoCharactersRemoveTextAndChangePointXReference(blnTwoCharacters, strText, pntPoint, intReferenceWidth, intPixelWidthSubtractionSpacing)
        End While

    End Sub

    Private Sub ChangeTextColorIndex(strColor As String, ByRef intByRefColorIndexSkip As Integer)

        'Check if different than white
        Select Case UCase(strColor) 'Leave white at index starting = 0
            Case "RED"
                'Set color index skip
                intByRefColorIndexSkip = 96
            Case "BLUE"
                'Set color index skip
                intByRefColorIndexSkip = 192
        End Select

    End Sub

    Private Sub ParseOneCharacter(strText As String, ByRef strByRefParsedOneCharacter As String, ByRef blnByRefTwoCharacters As Boolean)

        'Check length of string
        If strText.Length > 1 Then
            'Set one character
            strByRefParsedOneCharacter = strText.Substring(0, 1)
            'Check if there is a certain format
            If strByRefParsedOneCharacter = """" OrElse strByRefParsedOneCharacter = "'" Then
                'Check length
                Select Case strText.Length 'Ignore case 1
                    Case 2
                        'Check if the last bit is acceptable
                        TwoCharactersString(strText, strByRefParsedOneCharacter, blnByRefTwoCharacters)
                    Case > 2
                        'Check if the last bit is acceptable
                        TwoCharactersString(strText.Substring(0, 2), strByRefParsedOneCharacter, blnByRefTwoCharacters)
                End Select
            End If
        Else
            'Set one character
            strByRefParsedOneCharacter = strText
        End If

    End Sub

    Private Sub TwoCharactersString(strTextToBe As String, ByRef strByRefParseOneCharacter As String,
                                    ByRef blnByRefTwoCharacters As Boolean)

        'Check if the last bit is acceptable
        If strTextToBe = """0" OrElse strTextToBe = """1" OrElse strTextToBe = "'0" OrElse strTextToBe = "'1" Then
            'Include the next character
            strByRefParseOneCharacter = strTextToBe
            'Set
            blnByRefTwoCharacters = True
        End If

    End Sub

    Private Sub GDIDrawCharacter(strParsedOneCharacter As String, ByRef btmByRefBitmapToDrawOn As Bitmap, pntPoint As Point,
                              ByRef intByRefReferenceWidth As Integer, abtmTextSizeMemories() As Bitmap, intColorIndexSkip As Integer,
                              intFontSize As Integer)

        'Declare
        Dim intYChange As Integer = 0

        'Set the y change if necessary
        SetYChangeByParsedOneCharacter(strParsedOneCharacter, intFontSize, intYChange)

        'Check for type of letter
        Select Case strParsedOneCharacter
            Case " "
                'Set width
                intByRefReferenceWidth = 11
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      CInt(strParsedOneCharacter) + intColorIndexSkip)
            Case "A"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      10 + intColorIndexSkip)
            Case "B"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      11 + intColorIndexSkip)
            Case "C"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      12 + intColorIndexSkip)
            Case "D"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      13 + intColorIndexSkip)
            Case "E"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      14 + intColorIndexSkip)
            Case "F"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      15 + intColorIndexSkip)
            Case "G"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      16 + intColorIndexSkip)
            Case "H"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      17 + intColorIndexSkip)
            Case "I"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      18 + intColorIndexSkip)
            Case "J"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      19 + intColorIndexSkip)
            Case "K"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      20 + intColorIndexSkip)
            Case "L"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      21 + intColorIndexSkip)
            Case "M"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      22 + intColorIndexSkip)
            Case "N"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      23 + intColorIndexSkip)
            Case "O"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      24 + intColorIndexSkip)
            Case "P"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      25 + intColorIndexSkip)
            Case "Q"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      26 + intColorIndexSkip)
            Case "R"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      27 + intColorIndexSkip)
            Case "S"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      28 + intColorIndexSkip)
            Case "T"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      29 + intColorIndexSkip)
            Case "U"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      30 + intColorIndexSkip)
            Case "V"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      31 + intColorIndexSkip)
            Case "W"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      32 + intColorIndexSkip)
            Case "X"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      33 + intColorIndexSkip)
            Case "Y"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      34 + intColorIndexSkip)
            Case "Z"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth, abtmTextSizeMemories,
                                                      35 + intColorIndexSkip)
            Case "a"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 36 + intColorIndexSkip)
            Case "b"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 37 + intColorIndexSkip)
            Case "c"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 38 + intColorIndexSkip)
            Case "d"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 39 + intColorIndexSkip)
            Case "e"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 40 + intColorIndexSkip)
            Case "f"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 41 + intColorIndexSkip)
            Case "g"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 42 + intColorIndexSkip)
            Case "h"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 43 + intColorIndexSkip)
            Case "i"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 44 + intColorIndexSkip)
            Case "j"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 45 + intColorIndexSkip)
            Case "k"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 46 + intColorIndexSkip)
            Case "l"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 47 + intColorIndexSkip)
            Case "m"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 48 + intColorIndexSkip)
            Case "n"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 49 + intColorIndexSkip)
            Case "o"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 50 + intColorIndexSkip)
            Case "p"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 51 + intColorIndexSkip)
            Case "q"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 52 + intColorIndexSkip)
            Case "r"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 53 + intColorIndexSkip)
            Case "s"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 54 + intColorIndexSkip)
            Case "t"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 55 + intColorIndexSkip)
            Case "u"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 56 + intColorIndexSkip)
            Case "v"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 57 + intColorIndexSkip)
            Case "w"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 58 + intColorIndexSkip)
            Case "x"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 59 + intColorIndexSkip)
            Case "y"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 60 + intColorIndexSkip)
            Case "z"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 61 + intColorIndexSkip)
            Case "."
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 62 + intColorIndexSkip)
            Case ","
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 63 + intColorIndexSkip)
            Case "?"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 64 + intColorIndexSkip)
            Case "<"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 65 + intColorIndexSkip)
            Case ">"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 66 + intColorIndexSkip)
            Case "|"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 67 + intColorIndexSkip)
            Case ":"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 68 + intColorIndexSkip)
            Case ";"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 69 + intColorIndexSkip)
            Case """0", """"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 70 + intColorIndexSkip)
            Case """1"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 71 + intColorIndexSkip)
            Case "'0", "'"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 72 + intColorIndexSkip)
            Case "'1"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 73 + intColorIndexSkip)
            Case "{"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 74 + intColorIndexSkip)
            Case "}"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 75 + intColorIndexSkip)
            Case "["
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 76 + intColorIndexSkip)
            Case "]"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 77 + intColorIndexSkip)
            Case "/"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 78 + intColorIndexSkip)
            Case "\"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 79 + intColorIndexSkip)
            Case "~"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 80 + intColorIndexSkip)
            Case "`"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 81 + intColorIndexSkip)
            Case "!"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 82 + intColorIndexSkip)
            Case "@"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 83 + intColorIndexSkip)
            Case "#"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 84 + intColorIndexSkip)
            Case "$"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 85 + intColorIndexSkip)
            Case "%"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 86 + intColorIndexSkip)
            Case "^"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 87 + intColorIndexSkip)
            Case "&"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 88 + intColorIndexSkip)
            Case "*"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 89 + intColorIndexSkip)
            Case "("
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 90 + intColorIndexSkip)
            Case ")"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, pntPoint, intByRefReferenceWidth,
                                                      abtmTextSizeMemories, 91 + intColorIndexSkip)
            Case "-"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 92 + intColorIndexSkip)
            Case "_"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 93 + intColorIndexSkip)
            Case "="
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 94 + intColorIndexSkip)
            Case "+"
                'Draw letter at a point
                GDIDrawCharacterIncreaseWidthVariable(btmByRefBitmapToDrawOn, New Point(pntPoint.X, pntPoint.Y + intYChange),
                                                      intByRefReferenceWidth, abtmTextSizeMemories, 95 + intColorIndexSkip)
        End Select

    End Sub

    Private Sub SetYChangeByParsedOneCharacter(strParsedOneCharacter As String, intFontSize As Integer, ByRef intByRefYChange As Integer)

        'Set the y change if necessary
        Select Case strParsedOneCharacter
            Case "a", "c", "e", "g", "m", "n", "o", "p", "q", "r", "s", "u", "v", "w", "x", "y", "z", ";", ":", "-", "="
                'Y height change
                SetTextYChange(intFontSize, intByRefYChange, 4, 4, 5, 8, 9, 12, 16)
            Case "t", "+"
                'Y height change
                SetTextYChange(intFontSize, intByRefYChange, 1, 2, 2, 3, 3, 4, 6)
            Case ".", ",", "_"
                'Y height change
                SetTextYChange(intFontSize, intByRefYChange, 10, 10, 14, 19, 22, 29, 40)
            Case "<", ">"
                'Y height change
                SetTextYChange(intFontSize, intByRefYChange, 2, 2, 3, 4, 4, 6, 8)
        End Select

    End Sub

    Private Sub SetTextYChange(intFontSize As Integer, ByRef intByRefYChange As Integer, intFontSize18 As Integer,
                               intFontSize20 As Integer, intFontSize25 As Integer, intFontSize36 As Integer,
                               intFontsize42 As Integer, intFontsize55 As Integer, intFontSize72 As Integer)

        Select Case intFontSize
            Case 18
                'Set
                intByRefYChange = intFontSize18
            Case 20
                'Set
                intByRefYChange = intFontSize20
            Case 25
                'Set
                intByRefYChange = intFontSize25
            Case 36
                'Set
                intByRefYChange = intFontSize36
            Case 42
                'Set
                intByRefYChange = intFontsize42
            Case 55
                'Set
                intByRefYChange = intFontsize55
            Case 72
                'Set
                intByRefYChange = intFontSize72
        End Select

    End Sub

    Private Sub GDIDrawCharacterIncreaseWidthVariable(ByRef btmByRefBitmapToDrawOn As Bitmap, pntPoint As Point,
                                                      ByRef intByRefReferenceWidth As Integer, abtmTextSizeMemories() As Bitmap,
                                                      intIndex As Integer)

        'Draw letter at a point
        GDIDrawGraphics(Graphics.FromImage(btmByRefBitmapToDrawOn), abtmTextSizeMemories(intIndex), pntPoint)

        'Set width
        intByRefReferenceWidth = abtmTextSizeMemories(intIndex).Width

    End Sub

    Private Sub RemoveTextString(ByRef strByRefText As String, intLength As Integer)

        'Check length before removal
        If strByRefText.Length > intLength Then
            'Remove letter
            strByRefText = strByRefText.Substring(intLength)
        Else
            'Empty
            strByRefText = ""
        End If

    End Sub

    Private Sub CheckTwoCharactersRemoveTextAndChangePointXReference(ByRef blnByRefTwoCharacters As Boolean, ByRef strByRefText As String,
                                                                     ByRef pntByRefPoint As Point, intReferenceWidth As Integer,
                                                                     intPixelWidthSubtractionSpacing As Integer)

        'Check
        If blnByRefTwoCharacters Then
            'Reset
            blnByRefTwoCharacters = False
            'Check length before removal
            RemoveTextString(strByRefText, 2)
        Else
            'Check length before removal
            RemoveTextString(strByRefText, 1)
        End If
        'Set reference point
        pntByRefPoint.X += intReferenceWidth - intPixelWidthSubtractionSpacing

    End Sub

    Private Sub OpenGLDrawText(intFontSize As Integer, strColor As String, strText As String, pntPoint As Point,
                               Optional intPixelWidthSubtractionSpacing As Integer = 0)

        'Declare
        Dim intReferenceWidth As Integer = 0
        Dim strParsedOneCharacter As String = ""
        Dim intColorIndexSkip As Integer = 0
        Dim blnTwoCharacters As Boolean = False

        'Check if different than white
        ChangeTextColorIndex(strColor, intColorIndexSkip)

        'Loop until all letters are drawn
        While strText.Length > 0
            'Parse one character
            ParseOneCharacter(strText, strParsedOneCharacter, blnTwoCharacters)
            'Check which type of size
            Select Case intFontSize
                Case 25
                    'Draw character
                    OpenGLDrawCharacter(strParsedOneCharacter, pntPoint, intReferenceWidth, aiptTextSize25Textures,
                                        intColorIndexSkip, abtmTextSize25TNRMemories, intFontSize)
                Case 36
                    'Draw character
                    OpenGLDrawCharacter(strParsedOneCharacter, pntPoint, intReferenceWidth, aiptTextSize36Textures,
                                        intColorIndexSkip, abtmTextSize36TNRMemories, intFontSize)
                Case 42
                    'Draw character
                    OpenGLDrawCharacter(strParsedOneCharacter, pntPoint, intReferenceWidth, aiptTextSize42Textures,
                                        intColorIndexSkip, abtmTextSize42TNRMemories, intFontSize)
                Case 55
                    'Draw character
                    OpenGLDrawCharacter(strParsedOneCharacter, pntPoint, intReferenceWidth, aiptTextSize55Textures,
                                        intColorIndexSkip, abtmTextSize55TNRMemories, intFontSize)
                Case 72
                    'Draw character
                    OpenGLDrawCharacter(strParsedOneCharacter, pntPoint, intReferenceWidth, aiptTextSize72Textures,
                                        intColorIndexSkip, abtmTextSize72TNRMemories, intFontSize)
            End Select
            'Check for two characters
            CheckTwoCharactersRemoveTextAndChangePointXReference(blnTwoCharacters, strText, pntPoint, intReferenceWidth,
                                                                 intPixelWidthSubtractionSpacing)
        End While

    End Sub

    Private Sub OpenGLDrawCharacter(strParsedOneCharacter As String, pntPoint As Point, ByRef intByRefReferenceWidth As Integer,
                                    aiptTextures() As IntPtr, intColorIndexSkip As Integer, abtmTextMemories() As Bitmap,
                                    intFontSize As Integer)

        'Declare
        Dim intYChange As Integer = 0

        'Set the y change if necessary
        SetYChangeByParsedOneCharacter(strParsedOneCharacter, intFontSize, intYChange)

        'Check for type of letter
        Select Case strParsedOneCharacter
            Case " "
                'Set width
                intByRefReferenceWidth = 11
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, CInt(strParsedOneCharacter) + intColorIndexSkip,
                                                         pntPoint, abtmTextMemories, intByRefReferenceWidth,
                                                         strParsedOneCharacter)
            Case "A"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 10 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "B"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 11 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "C"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 12 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "D"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 13 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "E"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 14 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "F"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 15 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "G"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 16 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "H"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 17 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "I"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 18 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "J"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 19 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "K"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 20 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "L"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 21 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "M"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 22 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "N"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 23 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "O"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 24 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "P"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 25 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "Q"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 26 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "R"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 27 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "S"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 28 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "T"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 29 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "U"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 30 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "V"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 31 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "W"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 32 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "X"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 33 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "Y"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 34 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "Z"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 35 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "a"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 36 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "b"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 37 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "c"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 38 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "d"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 39 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "e"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 40 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "f"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 41 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "g"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 42 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "h"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 43 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "i"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 44 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "j"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 45 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "k"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 46 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "l"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 47 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "m"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 48 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "n"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 49 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "o"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 50 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "p"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 51 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "q"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 52 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "r"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 53 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "s"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 54 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "t"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 55 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "u"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 56 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "v"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 57 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "w"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 58 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "x"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 59 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "y"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 60 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "z"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 61 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "."
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 62 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case ","
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 63 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "?"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 64 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "<"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 65 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case ">"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 66 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "|"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 67 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case ":"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 68 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case ";"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 69 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case """0", """"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 70 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case """1"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 71 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "'0", "'"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 72 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "'1"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 73 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "{"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 74 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "}"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 75 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "["
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 76 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "]"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 77 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "/"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 78 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "\"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 79 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "~"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 80 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "`"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 81 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "!"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 82 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "@"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 83 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "#"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 84 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "$"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 85 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "%"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 86 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "^"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 87 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "&"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 88 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "*"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 89 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "("
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 90 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case ")"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 91 + intColorIndexSkip, pntPoint, abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "-"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 92 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "_"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 93 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "="
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 94 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
            Case "+"
                'Draw letter at a point
                OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures, 95 + intColorIndexSkip,
                                                         New Point(pntPoint.X, pntPoint.Y + intYChange), abtmTextMemories,
                                                         intByRefReferenceWidth, strParsedOneCharacter)
        End Select

    End Sub

    Private Sub OpenGLDrawCharacterIncreaseWidthVariable(aiptTextures() As IntPtr, intIndex As Integer, pntPoint As Point,
                                                         abtmTextMemory() As Bitmap, ByRef intByRefReferenceWidth As Integer,
                                                         strParsedOneCharacter As String)

        'Draw letter at a point
        OpenGLDrawImage(aiptTextures(intIndex), pntPoint, abtmTextMemory(intIndex), True)

        'Set width
        intByRefReferenceWidth = abtmTextMemory(intIndex).Width

    End Sub

    Private Sub ScreenResolutionChanged()

        'Change resolution if it does need to be changed
        If blnScreenChanged Then
            'Change window state
            If intResolutionMode <> 5 Then '5 = fullscreen
                'Normal
                Invoke(Sub() Me.WindowState = FormWindowState.Normal) 'Prevent cross-threading
                'Change
                Invoke(Sub() Me.Width = intScreenWidth) 'Prevent cross-threading
                Invoke(Sub() Me.Height = intScreenHeight) 'Prevent cross-threading
            Else
                'Force full screen, let windows do stuff for us
                Invoke(Sub() Me.WindowState = FormWindowState.Maximized) 'Prevent cross-threading
            End If
            'Set
            gsngScreenWidthRatio = CSng(ClientSize.Width) / CSng(intORIGINAL_SCREEN_WIDTH)
            gsngScreenHeightRatio = CSng(ClientSize.Height) / CSng(intORIGINAL_SCREEN_HEIGHT)
            'Set screen rectangle
            rectFullScreen.Width = ClientSize.Width
            rectFullScreen.Height = ClientSize.Height
            'Center the form
            If intResolutionMode <> 5 Then
                Invoke(Sub() Me.Top = CInt((My.Computer.Screen.WorkingArea.Height / 2) - (Me.Height / 2))) 'Prevent cross-threading
                Invoke(Sub() Me.Left = CInt((My.Computer.Screen.WorkingArea.Width / 2) - (Me.Width / 2))) 'Prevent cross-threading
            End If
            'Reset
            blnScreenChanged = False
        End If

    End Sub

    Private Sub LoadBitmapsIntoMemoryDuringRender()

        'Load parts of the game here after a screen has been shown, slowly load, this only happens during the loading screens
        Select Case intCanvasMode
            Case 2, 7
                LoadingBitmapsIntoMemory()
        End Select

    End Sub

    Private Sub CheckIfNeedToLoadNewLevel()

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

    End Sub

    Private Sub OpenGLInitialize()

        'Set GDI context to paint into the window
        Invoke(Sub() iptGDIContext = GetDC(Handle)) 'Prevent cross-threading

        'Setup the structure of pixel format
        With strucPixelFormatDescriptor
            'Set size of in the pixel format structure variable
            .intSize = System.Runtime.InteropServices.Marshal.SizeOf(strucPixelFormatDescriptor)
            'Set version
            .intVersion = intPFD_VERSION
            'Format support
            .intFormatSupport = intPFD_DRAW_TO_WINDOW Or intPFD_SUPPORT_OPENGL Or intPFD_TYPE_RGBA
            'RGBA type
            .intTypeRGBA = intPFD_TYPE_RGBA
            'Bits
            .intBits = intPFD_BITS
            '16-bit z-buffer
            .intSixteenBitZBuffer = intPFD_SIXTEEN_BIT_Z_BUFFER
            'Main drawing layer
            .intMainPlane = intPFD_MAIN_PLANE
        End With

        'Choose pixel format
        uintPixelFormatTemp = ChoosePixelFormat(iptGDIContext, strucPixelFormatDescriptor)

        'Set pixel format
        SetPixelFormat(iptGDIContext, uintPixelFormatTemp, strucPixelFormatDescriptor)

        'Set permanent context
        iptPermanentContext = wglCreateContext(iptGDIContext)

        'Try to activate the rendering context
        blnActivatedRenderingContext = wglMakeCurrent(iptGDIContext, iptPermanentContext) 'By this point, no error and window is attached

        'Viewport
        glViewport(0, 0, ClientSize.Width, ClientSize.Height)

        'Projection
        glMatrixMode(intGL_PROJECTION)

        'Remove
        glLoadIdentity()

        'Set version string
        strOpenGLVersion = strSplitElement(Runtime.InteropServices.Marshal.PtrToStringAnsi(glGetString(intGL_VERSION)),
                           " - ", 0)

        'Model view
        glMatrixMode(intGL_MODELVIEW)

        'Reset the projection matrix
        glLoadIdentity()

        'Enable smooth shading
        glShadeModel(intGL_SMOOTH)

        'Load textures for OpenGL
        LoadOpenGLOptionsTextures()

    End Sub

    Private Sub LoadOpenGLOptionsTextures()

        'Load options background memory
        LoadOpenGLTexture(aiptOptionsTextures(0), btmOptionsBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load options back hover
        LoadOpenGLTexture(aiptOptionsTextures(1), btmBackHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load options back non-hover
        LoadOpenGLTexture(aiptOptionsTextures(2), btmBackTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load graphics text memory
        LoadOpenGLTexture(aiptOptionsTextures(3), btmGraphicsTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not GDI+ software text memory
        LoadOpenGLTexture(aiptOptionsTextures(4), btmNotGDIPlusSoftwareTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load OpenGL hardware text memory
        LoadOpenGLTexture(aiptOptionsTextures(5), btmOpenGLHardwareTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load sound text memory
        LoadOpenGLTexture(aiptOptionsTextures(6), btmSoundTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load MCI-SS text memory
        LoadOpenGLTexture(aiptOptionsTextures(7), btmMCISSTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load draw resolution text memory
        LoadOpenGLTexture(aiptOptionsTextures(8), btmResolutionTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load 800x600
        LoadOpenGLTexture(aiptOptionsTextures(9), btm800x600TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not 800x600
        LoadOpenGLTexture(aiptOptionsTextures(10), btmNot800x600TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load 1024x768
        LoadOpenGLTexture(aiptOptionsTextures(11), btm1024x768TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not 1024x768
        LoadOpenGLTexture(aiptOptionsTextures(12), btmNot1024x768TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load 1280x800
        LoadOpenGLTexture(aiptOptionsTextures(13), btm1280x800TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not 1280x800
        LoadOpenGLTexture(aiptOptionsTextures(14), btmNot1280x800TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load 1280x1024
        LoadOpenGLTexture(aiptOptionsTextures(15), btm1280x1024TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not 1280x1024
        LoadOpenGLTexture(aiptOptionsTextures(16), btmNot1280x1024TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load 1440x900
        LoadOpenGLTexture(aiptOptionsTextures(17), btm1440x900TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not 1440x900
        LoadOpenGLTexture(aiptOptionsTextures(18), btmNot1440x900TextMemory, intINTERNAL_FORMAT_RGBA)

        'Load fullscreen
        LoadOpenGLTexture(aiptOptionsTextures(19), btmFullScreenTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load not fullscreen
        LoadOpenGLTexture(aiptOptionsTextures(20), btmNotFullScreenTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load volume
        LoadOpenGLTexture(aiptOptionsTextures(21), btmVolumeTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load sound bar
        LoadOpenGLTexture(aiptOptionsTextures(22), btmSoundBarMemory, intINTERNAL_FORMAT_RGB)

        'Load sound percentage
        LoadOpenGLTexture(aiptOptionsTextures(23), btmSoundPercent, intINTERNAL_FORMAT_RGBA)

        'Load sound percentage numbers
        LoadArrayOpenGLTextures(abtmSoundMemories, aiptOptionsTextures, intINTERNAL_FORMAT_RGBA, 24)

        'Load slider
        LoadOpenGLTexture(aiptOptionsTextures(125), btmSliderMemory, intINTERNAL_FORMAT_RGBA)

        'Load OpenGL text textures
        LoadOpenGLTextTextures()

        'Load menu
        LoadOpenGLTexture(aiptMenuTextures(0), btmMenuBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load back fog
        LoadOpenGLTexture(aiptMenuTextures(1), btmFogBackMemory, intINTERNAL_FORMAT_RGBA)

        'Load archer
        LoadOpenGLTexture(aiptMenuTextures(2), btmArcherMemory, intINTERNAL_FORMAT_RGBA)

        'Load front fog
        LoadOpenGLTexture(aiptMenuTextures(3), btmFogFrontMemory, intINTERNAL_FORMAT_RGBA)

        'Load start menu text
        LoadOpenGLTexture(aiptMenuTextures(4), btmStartTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load start menu hover text
        LoadOpenGLTexture(aiptMenuTextures(5), btmStartHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load highscores menu text
        LoadOpenGLTexture(aiptMenuTextures(6), btmHighscoresTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load highscores menu hover text
        LoadOpenGLTexture(aiptMenuTextures(7), btmHighscoresHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load story menu text
        LoadOpenGLTexture(aiptMenuTextures(8), btmStoryTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load story menu hover text
        LoadOpenGLTexture(aiptMenuTextures(9), btmStoryHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load options menu text
        LoadOpenGLTexture(aiptMenuTextures(10), btmOptionsTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load options menu hover text
        LoadOpenGLTexture(aiptMenuTextures(11), btmOptionsHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load credits menu text
        LoadOpenGLTexture(aiptMenuTextures(12), btmCreditsTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load credits menu hover text
        LoadOpenGLTexture(aiptMenuTextures(13), btmCreditsHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus menu text
        LoadOpenGLTexture(aiptMenuTextures(14), btmVersusTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus menu hover text
        LoadOpenGLTexture(aiptMenuTextures(15), btmVersusHoverTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load last stand text
        LoadOpenGLTexture(aiptMenuTextures(16), btmLastStandTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load loading background texture
        LoadOpenGLTexture(aiptLoadingTextures(0), btmLoadingBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load loading bar textures
        LoadArrayOpenGLTextures(abtmLoadingBarPictureMemories, aiptLoadingTextures, intINTERNAL_FORMAT_RGBA, 1)

        'Load loading text (not finished)
        LoadOpenGLTexture(aiptLoadingTextures(12), btmLoadingTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load loading text (finished)
        LoadOpenGLTexture(aiptLoadingTextures(13), btmLoadingStartTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load loading paragraph
        LoadArrayOpenGLTextures(abtmLoadingParagraphMemories, aiptLoadingTextures, intINTERNAL_FORMAT_RGBA, 14)

        'Load highscores background
        LoadOpenGLTexture(aiptHighscoreTextures(0), btmHighscoresBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load story background
        LoadOpenGLTexture(aiptStoryTextures(0), btmStoryBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load story paragraphs
        LoadArrayOpenGLTextures(abtmStoryParagraphMemories, aiptStoryTextures, intINTERNAL_FORMAT_RGBA, 1)

        'Load credits background
        LoadOpenGLTexture(aiptCreditTextures(0), btmCreditsBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load John Gonzales
        LoadArrayOpenGLTextures(abtmJohnGonzalesMemories, aiptCreditTextures, intINTERNAL_FORMAT_RGBA, 1)

        'Load Zachary Stafford
        LoadArrayOpenGLTextures(abtmZacharyStaffordMemories, aiptCreditTextures, intINTERNAL_FORMAT_RGBA, 5)

        'Load Cory Lewis
        LoadArrayOpenGLTextures(abtmCoryLewisMemories, aiptCreditTextures, intINTERNAL_FORMAT_RGBA, 9)

        'Load versus screen background
        LoadOpenGLTexture(aiptVersusScreenTextures(0), btmVersusBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load versus host text
        LoadOpenGLTexture(aiptVersusScreenTextures(1), btmVersusHostTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus host hover text
        LoadOpenGLTexture(aiptVersusScreenTextures(2), btmVersusHostTextHoverMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus screen or text
        LoadOpenGLTexture(aiptVersusScreenTextures(3), btmVersusOrTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus join text
        LoadOpenGLTexture(aiptVersusScreenTextures(4), btmVersusJoinTextHoverMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus join hover
        LoadOpenGLTexture(aiptVersusScreenTextures(5), btmVersusJoinTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus nickname
        LoadOpenGLTexture(aiptVersusScreenTextures(6), btmVersusNickNameTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load ip address text
        LoadOpenGLTexture(aiptVersusScreenTextures(7), btmVersusIPAddressTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus connect text
        LoadOpenGLTexture(aiptVersusScreenTextures(8), btmVersusConnectTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus connect hover text
        LoadOpenGLTexture(aiptVersusScreenTextures(9), btmVersusConnectTextHoverMemory, intINTERNAL_FORMAT_RGBA)

        'Load game mismatch background
        LoadOpenGLTexture(aiptGameMismatchTextures(0), btmGameMismatchBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load versus loading background
        LoadOpenGLTexture(aiptVersusLoadingTextures(0), btmLoadingBackgroundMemory, intINTERNAL_FORMAT_RGB)

        'Load loading waiting text
        LoadOpenGLTexture(aiptVersusLoadingTextures(1), btmLoadingWaitingTextMemory, intINTERNAL_FORMAT_RGBA)

        'Load versus loading paragraph textures
        LoadArrayOpenGLTextures(abtmLoadingParagraphVersusMemories, aiptVersusLoadingTextures, intINTERNAL_FORMAT_RGBA, 2)

    End Sub

    Private Sub LoadOpenGLTextTextures()

        'Load 25 size times new roman text
        LoadArrayOpenGLTextures(abtmTextSize25TNRMemories, aiptTextSize25Textures, intINTERNAL_FORMAT_RGBA)

        'Load 36 size times new roman text
        LoadArrayOpenGLTextures(abtmTextSize36TNRMemories, aiptTextSize36Textures, intINTERNAL_FORMAT_RGBA)

        'Load 42 size times new roman text
        LoadArrayOpenGLTextures(abtmTextSize42TNRMemories, aiptTextSize42Textures, intINTERNAL_FORMAT_RGBA)

        'Load 55 size times new roman text
        LoadArrayOpenGLTextures(abtmTextSize55TNRMemories, aiptTextSize55Textures, intINTERNAL_FORMAT_RGBA)

        'Load 72 size times new roman text
        LoadArrayOpenGLTextures(abtmTextSize72TNRMemories, aiptTextSize72Textures, intINTERNAL_FORMAT_RGBA)

    End Sub

    Private Sub LoadOpenGLTexture(ByRef iptByRefTexture As IntPtr, btmBitmapMemory As Bitmap, intInternalFormat As Integer)

        'Notes: This sub runs in a multi-threaded environment.

        'Declare
        Dim btmdBitmapData As Imaging.BitmapData

        'Generate the texture
        glGenTextures(1, iptByRefTexture)

        'Bind the texture
        glBindTexture(intGL_TEXTURE_2D, iptByRefTexture)

        'Declare and lock
        btmdBitmapData = btmBitmapMemory.LockBits(New Rectangle(0, 0, btmBitmapMemory.Width,
                         btmBitmapMemory.Height), Imaging.ImageLockMode.ReadOnly, btmBitmapMemory.PixelFormat)

        'Set texture image
        glTexImage2D(intGL_TEXTURE_2D, 0, intInternalFormat, btmdBitmapData.Width, btmdBitmapData.Height, 0, intGL_BGRA_EXT,
                     intGL_UNSIGNED_BYTE, btmdBitmapData.Scan0)

        'Texture information parameters
        glTexParameteri(intGL_TEXTURE_2D, intGL_TEXTURE_WRAP_S, intGL_CLAMP_TO_EDGE)
        glTexParameteri(intGL_TEXTURE_2D, intGL_TEXTURE_WRAP_T, intGL_CLAMP_TO_EDGE)
        glTexParameteri(intGL_TEXTURE_2D, intGL_TEXTURE_MIN_FILTER, intGL_LINEAR)
        glTexParameteri(intGL_TEXTURE_2D, intGL_TEXTURE_MAG_FILTER, intGL_LINEAR)

        'Unlock
        btmBitmapMemory.UnlockBits(btmdBitmapData)

    End Sub

    Private Sub LoadArrayOpenGLTextures(abtmMemories() As Bitmap, ByRef aiptByRefTextures() As IntPtr, intInternalFormat As Integer,
                                        Optional intIndexIncrease As Integer = 0)

        'Load textures
        For intLoop As Integer = 0 To abtmMemories.GetUpperBound(0)
            LoadOpenGLTexture(aiptByRefTextures(intLoop + intIndexIncrease), abtmMemories(intLoop), intInternalFormat)
        Next

    End Sub

    Private Sub OpenGLDrawFullScreenImage(iptTexture As IntPtr)

        'Enable texture
        glEnable(intGL_TEXTURE_2D)

        'Bind texture
        glBindTexture(intGL_TEXTURE_2D, iptTexture)

        'Begin triangles
        glBegin(intGL_TRIANGLES)

        'Vertices
        glTexCoord2f(0F, 0F)
        glVertex3f(-1.0F, 1.0F, 0F)
        glTexCoord2f(0F, 1.0F)
        glVertex3f(-1.0F, -1.0F, 0F)
        glTexCoord2f(1.0F, 1.0F)
        glVertex3f(1.0F, -1.0F, 0F)
        glTexCoord2f(0F, 0F)
        glVertex3f(-1.0F, 1.0F, 0F)
        glTexCoord2f(1.0F, 0F)
        glVertex3f(1.0F, 1.0F, 0F)
        glTexCoord2f(1.0F, 1.0F)
        glVertex3f(1.0F, -1.0F, 0F)

        'End
        glEnd()

    End Sub

    Private Sub OpenGLDrawImage(iptTexture As IntPtr, pntTexture As Point, btmBitmapMemory As Bitmap,
                                Optional blnHasTransparency As Boolean = False)

        'Enable texture
        glEnable(intGL_TEXTURE_2D)

        'Check if transparent
        If blnHasTransparency Then
            'Enable blend
            glEnable(intGL_BLEND)
            'Enable blend
            glBlendFunc(intGL_SRC_ALPHA, intGL_ONE_MINUS_SRC_ALPHA)
        End If

        'Bind texture
        glBindTexture(intGL_TEXTURE_2D, iptTexture)

        'Begin triangles
        glBegin(intGL_TRIANGLES)

        'Declare
        Dim sngX As Single = CSng((btmBitmapMemory.Width / ClientSize.Width) * gsngScreenWidthRatio)
        Dim sngXOffset As Single = CSng(-1.0F + ((2.0F * pntTexture.X * gsngScreenWidthRatio) / ClientSize.Width) + sngX)
        Dim sngY As Single = CSng((btmBitmapMemory.Height / ClientSize.Height) * gsngScreenHeightRatio)
        Dim sngYOffset As Single = CSng(-1.0F + ((2.0F * pntTexture.Y * gsngScreenHeightRatio) / ClientSize.Height) + sngY)

        'Vertices
        glTexCoord2f(0F, 0F)
        glVertex3f(-sngX + sngXOffset, sngY - sngYOffset, 0F)
        glTexCoord2f(0F, 1.0F)
        glVertex3f(-sngX + sngXOffset, -sngY - sngYOffset, 0F)
        glTexCoord2f(1.0F, 1.0F)
        glVertex3f(sngX + sngXOffset, -sngY - sngYOffset, 0F)
        glTexCoord2f(0F, 0F)
        glVertex3f(-sngX + sngXOffset, sngY - sngYOffset, 0F)
        glTexCoord2f(1.0F, 0F)
        glVertex3f(sngX + sngXOffset, sngY - sngYOffset, 0F)
        glTexCoord2f(1.0F, 1.0F)
        glVertex3f(sngX + sngXOffset, -sngY - sngYOffset, 0F)

        'End
        glEnd()

    End Sub

    Private Sub NextLevel(btmGameBackgroundLevel As Bitmap)

        'Load default variables
        LoadDefaultVariables()

        'Memory copy level
        CopyLevelBitmap(btmGameBackgroundMemory, btmGameBackgroundLevel)

        'Set
        intZombieKills = 0 'Must be reset to set the beginning number of zombies

        'Set
        intZombiesKilledIncreaseSpawn = 0 'Must be reset

        'Character
        udcCharacter = New clsCharacter(Me, intCHARACTER_X, intCHARACTER_Y, "udcCharacter", udcReloadingSound, udcGunShotSound, udcStepSound,
                                        udcWaterFootStepLeftSound, udcWaterFootStepRightSound, udcGravelFootStepLeftSound, udcGravelFootStepRightSound)

        'On screen word
        udcOnScreenWordOne = New clsOnScreenWord

        'Zombies
        LoadZombies("Level 1 Single Player")

        'Load into levels
        Select Case gintLevel
            Case 2
                'Set
                blnOpenedEyes = False
                'Set
                gpntPipeValve.X = 1250
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

        'Play background sound music
        audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(Math.Round(gintSoundVolume / 4)), True)

    End Sub

    Private Sub GDIRenderMenuScreen()

        'Declare
        Dim rectSourceBack As Rectangle
        Dim rectSourceFront As Rectangle

        'Prepare clone of fog back
        GDIPrepareFog(blnProcessBackFog, intFogBackRectangleMove, intFOG_BACK_MEMORY_WIDTH, rectSourceBack, intFogBackX,
                      intFOG_BACK_ADJUSTED_HEIGHT, btmFogBackCloneScreenShown, btmFogBackMemory)

        'Prepare clone of fog front
        GDIPrepareFog(blnProcessFrontFog, intFogFrontRectangleMove, intFOG_FRONT_MEMORY_WIDTH, rectSourceFront, intFogFrontX,
                      intFOG_FRONT_ADJUSTED_HEIGHT, btmFogFrontCloneScreenShown, btmFogFrontMemory)

        'Paint onto canvas the menu background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmMenuBackgroundMemory, pntTopLeft)

        'Draw fog in back
        DrawFog(btmFogBackCloneScreenShown, intFOG_BACK_MEMORY_WIDTH, intFogBackRectangleMove, intFogBackX, intFOG_BACK_Y)

        'Draw Archer
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmArcherMemory, pntArcher)

        'Draw fog in front
        DrawFog(btmFogFrontCloneScreenShown, intFOG_FRONT_MEMORY_WIDTH, intFogFrontRectangleMove, intFogFrontX, intFOG_FRONT_Y)

        'Draw start text
        GDIDrawNonHoverOrHoverText(1, btmStartTextMemory, pntStartText, btmStartHoverTextMemory, pntStartHoverText)

        'Draw highscores text
        GDIDrawNonHoverOrHoverText(2, btmHighscoresTextMemory, pntHighscoresText, btmHighscoresHoverTextMemory, pntHighscoresHoverText)

        'Draw story text
        GDIDrawNonHoverOrHoverText(3, btmStoryTextMemory, pntStoryText, btmStoryHoverTextMemory, pntStoryHoverText)

        'Draw options text
        GDIDrawNonHoverOrHoverText(4, btmOptionsTextMemory, pntOptionsText, btmOptionsHoverTextMemory, pntOptionsHoverText)

        'Draw credits text
        GDIDrawNonHoverOrHoverText(5, btmCreditsTextMemory, pntCreditsText, btmCreditsHoverTextMemory, pntCreditsHoverText)

        'Draw versus text
        GDIDrawNonHoverOrHoverText(6, btmVersusTextMemory, pntVersusText, btmVersusHoverTextMemory, pntVersusHoverText)

        'Draw last stand text
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLastStandTextMemory, pntLastStandText)

    End Sub

    Private Sub GDIPrepareFog(blnProcessFog As Boolean, intFogRectangleMove As Integer, intFogMemoryWidth As Integer,
                              ByRef rectByRefSource As Rectangle, intFogWidth As Integer, intFogMemoryHeight As Integer,
                              ByRef btmByRefFogCloneScreenShown As Bitmap, btmFogMemory As Bitmap)

        'Prepare clone of fog
        If blnProcessFog Then
            'Check
            If intFogRectangleMove <> 0 Then
                'Make rectangle
                If intFogRectangleMove > intORIGINAL_SCREEN_WIDTH Then
                    'Check for last end of the picture on the left side
                    If intFogMemoryWidth - intFogRectangleMove < 0 Then
                        'Set
                        rectByRefSource = New Rectangle(0, 0, intFogWidth, intFogMemoryHeight)
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

    Private Sub DrawFog(ByRef btmByRefFogCloneScreenShown As Bitmap, intFogMemoryWidth As Integer, intFogRectangleMove As Integer,
                        intFogX As Integer, intFogDistanceY As Integer)

        'Draw fog
        If btmByRefFogCloneScreenShown IsNot Nothing Then
            'Draw depending
            If intFogMemoryWidth - intFogRectangleMove < 0 Then
                'Draw
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmByRefFogCloneScreenShown, New Point(intORIGINAL_SCREEN_WIDTH - intFogX,
                                intFogDistanceY))
            Else
                'Draw
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmByRefFogCloneScreenShown, New Point(0, intFogDistanceY))
            End If
            'Dispose of fog clone
            DisposeBitmap(btmByRefFogCloneScreenShown)
        End If

    End Sub

    Private Sub GDIDrawNonHoverOrHoverText(intWhatCanvasShowShouldBe As Integer, btmNonHoverMemory As Bitmap,
                                           pntNonHover As Point, btmHoverMemory As Bitmap, pntHover As Point)

        'Draw versus text
        If intCanvasShow = intWhatCanvasShowShouldBe Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmHoverMemory, pntHover)
        Else
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmNonHoverMemory, pntNonHover)
        End If

    End Sub

    Private Sub OpenGLRenderMenuScreen()

        'Draw the menu background
        OpenGLDrawFullScreenImage(aiptMenuTextures(0))

        'Draw back fog
        OpenGLDrawImage(aiptMenuTextures(1), New Point(-intFOG_BACK_MEMORY_WIDTH + intOpenGLBackFogX, intFOG_BACK_Y), btmFogBackMemory, True)

        'Draw archer
        OpenGLDrawImage(aiptMenuTextures(2), pntArcher, btmArcherMemory, True)

        'Draw front fog
        OpenGLDrawImage(aiptMenuTextures(3), New Point(-intFOG_FRONT_MEMORY_WIDTH + intOpenGLFrontFogX, intFOG_FRONT_Y), btmFogFrontMemory, True)

        'Draw start text
        OpenGLDrawNonHoverOrHoverText(1, aiptMenuTextures(4), btmStartTextMemory, pntStartText, aiptMenuTextures(5),
                                      btmStartHoverTextMemory, pntStartHoverText)

        'Draw highscores text
        OpenGLDrawNonHoverOrHoverText(2, aiptMenuTextures(6), btmHighscoresTextMemory, pntHighscoresText, aiptMenuTextures(7),
                                      btmHighscoresHoverTextMemory, pntHighscoresHoverText)

        'Draw story text
        OpenGLDrawNonHoverOrHoverText(3, aiptMenuTextures(8), btmStoryTextMemory, pntStoryText, aiptMenuTextures(9),
                                      btmStoryHoverTextMemory, pntStoryHoverText)

        'Draw options text
        OpenGLDrawNonHoverOrHoverText(4, aiptMenuTextures(10), btmOptionsTextMemory, pntOptionsText, aiptMenuTextures(11),
                                      btmOptionsHoverTextMemory, pntOptionsHoverText)

        'Draw credits text
        OpenGLDrawNonHoverOrHoverText(5, aiptMenuTextures(12), btmCreditsTextMemory, pntCreditsText, aiptMenuTextures(13),
                                      btmCreditsHoverTextMemory, pntCreditsHoverText)

        'Draw versus text
        OpenGLDrawNonHoverOrHoverText(6, aiptMenuTextures(14), btmVersusTextMemory, pntVersusText, aiptMenuTextures(15),
                                      btmVersusHoverTextMemory, pntVersusHoverText)

        'Draw last stand text
        OpenGLDrawImage(aiptMenuTextures(16), pntLastStandText, btmLastStandTextMemory, True)

    End Sub

    Private Sub OpenGLDrawNonHoverOrHoverText(intWhatCanvasShowShouldBe As Integer, iptNonHoverTexture As IntPtr,
                                              btmNonHoverMemory As Bitmap, pntNonHover As Point, iptHoverTexture As IntPtr,
                                              btmHoverMemory As Bitmap, pntHover As Point)

        'Draw start text
        If intCanvasShow = intWhatCanvasShowShouldBe Then
            OpenGLDrawImage(iptHoverTexture, pntHover, btmHoverMemory, True)
        Else
            OpenGLDrawImage(iptNonHoverTexture, pntNonHover, btmNonHoverMemory, True)
        End If

    End Sub

    Private Sub GDIRenderOptionsScreen()

        'Draw options background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmOptionsBackgroundMemory, pntTopLeft)

        'Draw graphics text
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmGraphicsTextMemory, pntGraphicsText)

        'Draw GDI+
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmGDIPlusSoftwareTextMemory, pntGDIPlusSoftwareText)

        'Draw OpenGL if loaded
        If blnOpenGLLoaded Then
            'Draw OpenGL
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmNotOpenGLHardwareTextMemory, pntOpenGLHardwareText)
            'Draw the OpenGL version
            GDIDrawText(btmCanvas, 25, "WHITE", strOpenGLVersion, New Point(386, 73), 2)
        End If

        'Draw sound text
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmSoundTextMemory, pntSoundText)

        'Draw MCI-SS
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmMCISSTextMemory, pntMCISSText)

        'Draw resolution text
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmResolutionTextMemory, pntResolutionText)

        'Check which resolution
        GDICheckResolutionMode(0, btm800x600TextMemory, btmNot800x600TextMemory, pnt800x600Text)
        GDICheckResolutionMode(1, btm1024x768TextMemory, btmNot1024x768TextMemory, pnt1024x768Text)
        GDICheckResolutionMode(2, btm1280x800TextMemory, btmNot1280x800TextMemory, pnt1280x800Text)
        GDICheckResolutionMode(3, btm1280x1024TextMemory, btmNot1280x1024TextMemory, pnt1280x1024Text)
        GDICheckResolutionMode(4, btm1440x900TextMemory, btmNot1440x900TextMemory, pnt1440x900Text)
        GDICheckResolutionMode(5, btmFullScreenTextMemory, btmNotFullScreenTextMemory, pntFullscreenText)

        'Draw volume text
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmVolumeTextMemory, pntVolumeText)

        'Draw sound bar
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmSoundBarMemory, pntSoundBar)

        'Draw sound percentage
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmSoundPercent, pntSoundPercent)

        'Draw slider
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmSliderMemory, pntSlider)

        'Draw version
        GDIDrawText(btmCanvas, 36, "RED", "Version", New Point(18, 480), 2)

        'Draw version
        GDIDrawText(btmCanvas, 36, "WHITE", strGAME_VERSION, New Point(153, 482), 2)

        'Draw frames per second
        GDIDrawText(btmCanvas, 36, "RED", "Frames Per Second", New Point(438, 480), 2)

        'Draw the number of frames per second
        GDIDrawText(btmCanvas, 36, "WHITE", intFPSDisplay.ToString, New Point(753, 482), 2)

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub GDICheckResolutionMode(intModeSelected As Integer, btmResolutionText As Bitmap, btmNotResolutionText As Bitmap,
                                       pntResolutionText As Point)

        'Check resolution before drawing
        If intResolutionMode = intModeSelected Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmResolutionText, pntResolutionText)
        Else
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmNotResolutionText, pntResolutionText)
        End If

    End Sub

    Private Sub GDIShowBackButtonOrHoverBackButton()

        'Show back button or hover back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmBackHoverTextMemory, pntBackHoverText)
        Else
            'Draw back text
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmBackTextMemory, pntBackText)
        End If

    End Sub

    Private Sub OpenGLRenderOptionsScreen()

        'Draw background
        OpenGLDrawFullScreenImage(aiptOptionsTextures(0))

        'Draw the OpenGL version
        OpenGLDrawText(25, "WHITE", strOpenGLVersion, New Point(386, 73), 2)

        'Draw graphics word
        OpenGLDrawImage(aiptOptionsTextures(3), pntGraphicsText, btmGraphicsTextMemory, True)

        'Draw not GDI+
        OpenGLDrawImage(aiptOptionsTextures(4), pntGDIPlusSoftwareText, btmGDIPlusSoftwareTextMemory, True)

        'Draw OpenGL
        OpenGLDrawImage(aiptOptionsTextures(5), pntOpenGLHardwareText, btmOpenGLHardwareTextMemory, True)

        'Draw sound text memory
        OpenGLDrawImage(aiptOptionsTextures(6), pntSoundText, btmSoundTextMemory, True)

        'Draw MCI-SS text memory
        OpenGLDrawImage(aiptOptionsTextures(7), pntMCISSText, btmMCISSTextMemory, True)

        'Draw resolution text memory
        OpenGLDrawImage(aiptOptionsTextures(8), pntResolutionText, btmResolutionTextMemory, True)

        'Check which resolution
        OpenGLCheckResolutionMode(0, aiptOptionsTextures(9), aiptOptionsTextures(10), pnt800x600Text,
                                  btm800x600TextMemory, btmNot800x600TextMemory)
        OpenGLCheckResolutionMode(1, aiptOptionsTextures(11), aiptOptionsTextures(12), pnt1024x768Text,
                                  btm1024x768TextMemory, btmNot1024x768TextMemory)
        OpenGLCheckResolutionMode(2, aiptOptionsTextures(13), aiptOptionsTextures(14), pnt1280x800Text,
                                  btm1280x800TextMemory, btmNot1280x800TextMemory)
        OpenGLCheckResolutionMode(3, aiptOptionsTextures(15), aiptOptionsTextures(16), pnt1280x1024Text,
                                  btm1280x1024TextMemory, btmNot1280x1024TextMemory)
        OpenGLCheckResolutionMode(4, aiptOptionsTextures(17), aiptOptionsTextures(18), pnt1440x900Text,
                                  btm1440x900TextMemory, btmNot1440x900TextMemory)
        OpenGLCheckResolutionMode(5, aiptOptionsTextures(19), aiptOptionsTextures(20), pntFullscreenText,
                                  btmFullScreenTextMemory, btmNotFullScreenTextMemory)

        'Draw volume
        OpenGLDrawImage(aiptOptionsTextures(21), pntVolumeText, btmVolumeTextMemory, True)

        'Draw sound bar
        OpenGLDrawImage(aiptOptionsTextures(22), pntSoundBar, btmSoundBarMemory, True)

        'Draw sound percentage, these are all the same bitmap width and height so use element 0
        OpenGLDrawImage(aiptOptionsTextures(intSoundTextureIndex), pntSoundPercent, abtmSoundMemories(0), True)

        'Draw slider
        OpenGLDrawImage(aiptOptionsTextures(125), pntSlider, btmSliderMemory, True)

        'Draw version
        OpenGLDrawText(36, "RED", "Version", New Point(18, 480), 2)

        'Draw version number
        OpenGLDrawText(36, "WHITE", strGAME_VERSION, New Point(153, 482), 2)

        'Draw frames per second
        OpenGLDrawText(36, "RED", "Frames Per Second", New Point(438, 480), 2)

        'Draw frames per second number
        OpenGLDrawText(36, "WHITE", intFPSDisplay.ToString, New Point(753, 482), 2)

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLCheckResolutionMode(intModeSelected As Integer, iptTexture As IntPtr, iptNotTexture As IntPtr, pntResolutionText As Point,
                                      btmBitmapMemory As Bitmap, btmNotBitmapMemory As Bitmap)

        'Check resolution before drawing
        If intResolutionMode = intModeSelected Then
            'Draw resolution text
            OpenGLDrawImage(iptTexture, pntResolutionText, btmBitmapMemory, True)
        Else
            'Draw resolution text
            OpenGLDrawImage(iptNotTexture, pntResolutionText, btmNotBitmapMemory, True)
        End If

    End Sub

    Private Sub OpenGLShowBackButtonOrHoverBackButton()

        'Show back button or hover back button
        If intCanvasShow = 1 Then
            'Draw back button hover
            OpenGLDrawImage(aiptOptionsTextures(1), pntBackHoverText, btmBackHoverTextMemory, True)
        Else
            'Draw back button non-hover
            OpenGLDrawImage(aiptOptionsTextures(2), pntBackText, btmBackTextMemory, True)
        End If

    End Sub

    Private Sub GDILoadingGameScreen()

        'Draw loading background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw Loading text
        If blnFinishedLoading Then
            'Start
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
        Else
            'Loading
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
        End If

        'Draw paragraph
        If btmLoadingParagraph IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingParagraph, pntLoadingParagraph)
        End If

        'Text for controls
        GDIDrawText(btmCanvas, 25, "WHITE", strCONTROLS_DETAIL, New Point(200, 390), 2)

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLLoadingGameScreen()

        'Draw loading background
        OpenGLDrawFullScreenImage(aiptLoadingTextures(0))

        'Draw loading bar
        OpenGLDrawImage(aiptLoadingTextures(intLoadingBarTextureIndex), pntLoadingBar, abtmLoadingBarPictureMemories(0),
                        True) 'Use any array index because all same width, height

        'Draw Loading text
        If blnFinishedLoading Then
            'Start
            OpenGLDrawImage(aiptLoadingTextures(13), pntLoadingStartText, btmLoadingStartTextMemory, True)
        Else
            'Loading
            OpenGLDrawImage(aiptLoadingTextures(12), pntLoadingText, btmLoadingTextMemory, True)
        End If

        'Draw paragraph
        If intLoadingParagraphTextureIndex <> 0 Then
            OpenGLDrawImage(aiptLoadingTextures(intLoadingParagraphTextureIndex), pntLoadingParagraph,
                            abtmLoadingParagraphMemories(0), True) 'Use any array index
        End If

        'Text for controls
        OpenGLDrawText(25, "WHITE", strCONTROLS_DETAIL, New Point(200, 390), 2)

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StartedGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Check if beat level
            If Not blnPlayerWasPinned Then

                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Paint black background
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), abtmBlackScreenMemories(2), pntTopLeft)
                Else
                    'Paint black background
                    OpenGLDrawImage(aiptBlackScreenTextures(2), pntTopLeft, abtmBlackScreenMemories(2), True)
                End If

                'Check which level was completed
                Select Case gintLevel
                    Case 1
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Check if GDI or OpenGL
                        If Not blnOpenGL Then
                            'Set
                            btmPath = abtmPath1Memories(0)
                        Else
                            'Set
                            intPathTextureIndex = 0
                        End If
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(11, 0, False) 'This exits the game screen to path choices
                    Case 2, 3
                        'Set
                        blnLightZap1 = False
                        blnLightZap2 = False
                        'Check if GDI or OpenGL
                        If Not blnOpenGL Then
                            'Set
                            btmPath = abtmPath2Memories(0)
                        Else
                            'Set
                            intPathTextureIndex = 3
                        End If
                        'Set
                        ChangeCanvasModeAndChangeCanvasShowAndPlayClickSound(12, 0, False) 'This exits the game screen to path choices
                    Case 4
                        'Check if GDI or OpenGL
                        If Not blnOpenGL Then
                            'Write to screen
                            GDIDrawText(btmCanvas, 36, "WHITE", "This level is not created yet and therefore", New Point(63, 115), 2)
                            GDIDrawText(btmCanvas, 36, "WHITE", "you have become lost in the swamp and", New Point(63, 145), 2)
                            GDIDrawText(btmCanvas, 36, "WHITE", "perished.", New Point(63, 175), 2)
                        Else
                            'Write to screen
                            OpenGLDrawText(36, "WHITE", "This level is not created yet and therefore", New Point(63, 115), 2)
                            OpenGLDrawText(36, "WHITE", "you have become lost in the swamp and", New Point(63, 145), 2)
                            OpenGLDrawText(36, "WHITE", "perished.", New Point(63, 175), 2)
                        End If
                    Case 5
                        'Check if GDI or OpenGL
                        If Not blnOpenGL Then
                            'Show win overlay
                            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmWinOverlayMemory, pntTopLeft)
                        Else
                            'Show win overlay
                            OpenGLDrawImage(iptWinOverlayTexture, pntTopLeft, btmWinOverlayMemory, True)
                        End If
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

                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Show death screen fading
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)
                    'Show death overlay
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmDeathOverlayMemory, pntTopLeft)
                Else
                    'Show death screen fading
                    OpenGLDrawFullScreenImage(iptDeathScreenTexture)
                    'Show death overlay
                    OpenGLDrawImage(iptDeathOverlayTexture, pntTopLeft, btmDeathOverlayMemory, True)
                End If

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

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Show back button or hover back button
            GDIShowBackButtonOrHoverBackButton()
        Else
            'Show back button or hover back button
            OpenGLShowBackButtonOrHoverBackButton()
        End If

    End Sub

    Private Sub DrawStats(intZombieKillsToBe As Integer)

        'Check if stats has been created
        If Not blnSetStats Then
            'Set
            blnSetStats = True
            'Set timespan
            tsTimeSpan = swhTimeElapsed.Elapsed
            'Set seconds string
            strElapsedTime = CStr(CInt(tsTimeSpan.TotalSeconds)) & strDRAW_STATS_SECONDS_WORD
            'Set timespan
            tsTimeSpan = gswhTimeTyped.Elapsed
            'Get typed time
            intElapsedTime = CInt(tsTimeSpan.TotalSeconds)
            'Set WPM
            If intTypedWords = 0 Then
                'Set
                intWPM = 0
            Else
                'Convert first
                Dim dblFirstDivision As Double = CDbl(intElapsedTime) / 60
                Dim dblSecondDivision As Double = CDbl(intTypedWords) / dblFirstDivision
                'Set
                intWPM = CInt(Math.Round(dblSecondDivision))
            End If
        End If

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw zombie kills
            GDIDrawText(btmCanvas, 55, "WHITE", CStr(intZombieKillsToBe), New Point(540, 220), 2)
            'Draw survival time
            GDIDrawText(btmCanvas, 55, "WHITE", strElapsedTime, New Point(540, 319), 2)
            'Draw WPM
            GDIDrawText(btmCanvas, 55, "WHITE", intWPM.ToString, New Point(540, 415), 2)
        Else
            'Draw zombie kills
            OpenGLDrawText(55, "WHITE", CStr(intZombieKillsToBe), New Point(540, 220), 2)
            'Draw survival time
            OpenGLDrawText(55, "WHITE", strElapsedTime, New Point(540, 319), 2)
            'Draw WPM
            OpenGLDrawText(55, "WHITE", intWPM.ToString, New Point(540, 415), 2)
        End If

    End Sub

    Private Sub CompareHighscoresAccess()

        'Declare
        Dim strSQL As String = "SELECT intTime FROM HighscoresTable ORDER BY intRank ASC"
        Dim strConnection As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirectory &
                                      "Highscores.mdb;Jet OLEDB:Database Password=32x1aZ@!"
        Dim dtbDataTable As New DataTable
        Dim intRank As Integer = 0
        Dim blnRankBeat As Boolean = False

        'Fill data table with ranks
        FillDataTable(strSQL, strConnection, dtbDataTable)

        'Compare time
        For intLoop As Integer = 0 To dtbDataTable.Rows.Count - 1
            If CInt(tsTimeSpan.TotalSeconds) <= CInt(dtbDataTable.Rows(intLoop).Item(0).ToString) Then
                'Set
                intRank = intLoop + 1
                'Set
                blnRankBeat = True
                'Exit
                Exit For
            End If
        Next

        'Check if rank beat
        If blnRankBeat Then
            'Write to database
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

        'Enter the information into the database
        EnterTheInformationIntoTheDatabase(strInputBox, intRank)

        'Reload the highscore database
        LoadHighscoresIntoAString()

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

    Private Sub EnterTheInformationIntoTheDatabase(strName As String, intRank As Integer)

        'Notes: Command builder is needed for creating update SQL statements to be executed.

        'Declares
        Dim strSQL As String = "SELECT * FROM HighscoresTable ORDER BY intRank ASC"
        Dim strConnection As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDirectory &
                                      "Highscores.mdb;Jet OLEDB:Database Password=32x1aZ@!"
        Dim dtbDataTable As New DataTable

        'Use objects that will dispose themselves and the connection
        Using odbConnection As New OleDb.OleDbConnection(strConnection)
            Using odcCommand As New OleDb.OleDbCommand(strSQL, odbConnection)
                Using odaDataAdapter As New OleDb.OleDbDataAdapter(odcCommand)
                    Using ocbCommandBuilder As New OleDb.OleDbCommandBuilder(odaDataAdapter)
                        'Set fixes
                        ocbCommandBuilder.QuotePrefix = "["
                        ocbCommandBuilder.QuoteSuffix = "]"
                        'Fill data table
                        odaDataAdapter.Fill(dtbDataTable)
                        'Loop to copy all data downward
                        If intRank <> 10 Then
                            For intLoop As Integer = (dtbDataTable.Rows.Count - 1) To intRank Step -1
                                dtbDataTable.Rows(intLoop).Item("intTime") = dtbDataTable.Rows(intLoop - 1).Item("intTime")
                                dtbDataTable.Rows(intLoop).Item("intWPM") = dtbDataTable.Rows(intLoop - 1).Item("intWPM")
                                dtbDataTable.Rows(intLoop).Item("intKills") = dtbDataTable.Rows(intLoop - 1).Item("intKills")
                                dtbDataTable.Rows(intLoop).Item("strPlayerName") = dtbDataTable.Rows(intLoop - 1).Item("strPlayerName")
                            Next
                        End If
                        'Replace rank spot
                        dtbDataTable.Rows(intRank - 1).Item("intTime") = CStr(CInt(tsTimeSpan.TotalSeconds))
                        dtbDataTable.Rows(intRank - 1).Item("intWPM") = CStr(intWPM)
                        dtbDataTable.Rows(intRank - 1).Item("intKills") = CStr(intZombieKillsCombined)
                        dtbDataTable.Rows(intRank - 1).Item("strPlayerName") = strName
                        'Update the database
                        odaDataAdapter.Update(dtbDataTable)
                    End Using
                End Using
            End Using
        End Using

    End Sub

    Private Sub PlaySinglePlayerGame()

        'Enable time elapsed
        StartTimeElapsed()

        'Check for special effects in background
        Select Case gintLevel
            Case 5
                'Check for helicopter
                If gpntGameBackground.X <= -600 Then
                    'Draw for both GDI and OpenGL background bitmap
                    GDIDrawGraphics(Graphics.FromImage(btmGameBackgroundMemory), udcHelicopterGonzales.Image, udcHelicopterGonzales.Point)
                End If
        End Select

        'Draw dead zombies permanently
        DrawDeadZombiesPermanently(gaudcZombies, intZombiesKilledIncreaseSpawn)

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
                gblnPreventKeyPressEvent = True
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
                'Game over, time stopped
                StopTheStopWatches()
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

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw character
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), udcCharacter.Image, udcCharacter.Point)
        Else
            'Draw character
            OpenGLDrawImage(aiptCharacterTextures(udcCharacter.TextureIndex), udcCharacter.Point, udcCharacter.Image, True)
        End If

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
                        If intTempDistance <= intZOMBIE_PINNING_X Then
                            'Check for god mode
                            If Not blnGodMode Then
                                'Check if level not beat
                                If Not blnBeatLevel Then
                                    'Set distance for future pin
                                    intZombieIncreasedPinDistance = intTempDistance - 12
                                    'Make zombie pin
                                    gaudcZombies(intLoop).Pin()
                                    'Set
                                    blnPlayerWasPinned = True
                                    'Set
                                    gblnPreventKeyPressEvent = True
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
                                    'Game over, time stopped
                                    StopTheStopWatches()
                                    'Keep the zombie kills updated
                                    intZombieKillsCombined += intZombieKills
                                End If
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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Draw zombies dying, pinning or walking
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), gaudcZombies(intLoop).Image, gaudcZombies(intLoop).Point)
                Else
                    'Draw zombies dying, pinning or walking
                    OpenGLDrawImage(aiptZombieTextures(gaudcZombies(intLoop).TextureIndex), gaudcZombies(intLoop).Point,
                                    gaudcZombies(intLoop).Image, True)
                End If
            End If
        Next

        'Check level for special effect
        Select Case gintLevel
            Case 1
                'Destroyed brick wall
                If gpntGameBackground.X < -450 Then
                    'Check if GDI or OpenGL
                    If Not blnOpenGL Then
                        'Draw brick wall
                        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmDestroyedBrickWallMemory, gpntDestroyedBrickWall)
                    Else
                        'Draw brick wall
                        OpenGLDrawImage(iptDestroyedBrickWallTexture, gpntDestroyedBrickWall, btmDestroyedBrickWallMemory, True)
                    End If
                End If
            Case 2
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Chained zombie
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), gudcChainedZombie.Image, gudcChainedZombie.Point)
                Else
                    'Chained zombie
                    OpenGLDrawImage(aiptChainedZombieTextures(gudcChainedZombie.TextureIndex), gudcChainedZombie.Point,
                                    gudcChainedZombie.Image, True)
                End If
                'Draw level 2 water on screen
                DrawLevel2WaterOnScreen()
                'Pipe valve
                If gpntGameBackground.X <= -550 And gpntGameBackground.X >= -2200 Then
                    'Check if GDI or OpenGL
                    If Not blnOpenGL Then
                        'Draw pipe value
                        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmPipeValveMemory, gpntPipeValve)
                    Else
                        'Draw pipe valve
                        OpenGLDrawImage(iptPipeValveTexture, gpntPipeValve, btmPipeValveMemory, True)
                    End If
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
                            'Make the zombie move back and forth
                            gudcFaceZombie.Move()
                        End If
                    End If
                    'Check if GDI or OpenGL
                    If Not blnOpenGL Then
                        'Draw zombie face
                        GDIDrawGraphics(Graphics.FromImage(btmCanvas), gudcFaceZombie.Image, gudcFaceZombie.Point)
                    Else
                        'Draw zombie face
                        OpenGLDrawImage(aiptFaceZombieTextures(gudcFaceZombie.TextureIndex), gudcFaceZombie.Point,
                                        gudcFaceZombie.Image, True)
                    End If
                End If
        End Select

        'Draw word bar with wording
        DrawWordBarWithWording("RED")

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Show magazine with bullet count
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmAK47MagazineMemory, pntAK47Magazine)
            'Draw bullet count on magazine
            GDIDrawText(btmCanvas, 25, "RED", (30 - udcCharacter.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
        Else
            'Show magazine with bullet count
            OpenGLDrawImage(iptAK47MagazineTexture, pntAK47Magazine, btmAK47MagazineMemory, True)
            'Draw bullet count on magazine
            OpenGLDrawText(25, "RED", (30 - udcCharacter.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
        End If

        'Draw missed word on screen
        DrawOnScreenWord(udcOnScreenWordOne)

        'Check if need to black screen and make copy if died
        CheckEndGame()

    End Sub

    Private Sub StartTimeElapsed()

        'Enable time elapsed
        If Not swhTimeElapsed.IsRunning AndAlso Not blnStartedTimeElapsed Then
            'Start
            swhTimeElapsed.Start()
            'Set
            blnStartedTimeElapsed = True
        End If

    End Sub

    Private Sub DrawDeadZombiesPermanently(audcZombiesType() As clsZombie, ByRef intByRefZombiesKilledIncreaseSpawn As Integer)

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
                GDIDrawGraphics(Graphics.FromImage(btmGameBackgroundMemory), audcZombiesType(intLoop).Image, pntTemp)
                'Increase count
                intByRefZombiesKilledIncreaseSpawn += 1
                'Start a new zombie
                audcZombiesType(intByRefZombiesKilledIncreaseSpawn + intNUMBER_OF_ZOMBIES_AT_ONE_TIME - 1).Start()
            End If
        Next

    End Sub

    Private Sub PaintOnCanvasAndCloneScreen()

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT)

        'Clone the only necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackgroundMemory.Clone(rectSource,
                                            Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw the background to the screen with the cloned version
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmGameBackgroundCloneScreenShown, pntTopLeft)
        Else
            'Load OpenGL texture once and update
            LoadOpenGLTextureOnceAndUpdate(iptGameBackgroundTexture, btmGameBackgroundCloneScreenShown, intINTERNAL_FORMAT_RGB)
            'Draw the background to the screen with the cloned version
            OpenGLDrawFullScreenImage(iptGameBackgroundTexture)
        End If

        'Dispose because clone just makes the picture expand with more data
        DisposeBitmap(btmGameBackgroundCloneScreenShown)

    End Sub

    Private Sub LoadOpenGLTextureOnceAndUpdate(ByRef iptByRefTexture As IntPtr, btmBitmapMemory As Bitmap,
                                               intInternalFormat As Integer)

        'Check if texture has never been loaded
        If iptByRefTexture = IntPtr.Zero Then
            'Load texture
            LoadOpenGLTexture(iptByRefTexture, btmBitmapMemory, intInternalFormat)
        Else
            'Update texture
            UpdateOpenGLTexture(iptByRefTexture, btmBitmapMemory, intInternalFormat)
        End If

    End Sub

    Private Sub UpdateOpenGLTexture(ByRef iptByRefTexture As IntPtr, btmBitmapMemory As Bitmap, intInternalFormat As Integer)

        'Notes: This sub runs in a multi-threaded environment.

        'Declare
        Dim btmdBitmapData As Drawing.Imaging.BitmapData

        'Bind the texture
        glBindTexture(intGL_TEXTURE_2D, iptByRefTexture)

        'Declare and lock
        btmdBitmapData = btmBitmapMemory.LockBits(New Rectangle(0, 0, btmBitmapMemory.Width,
                         btmBitmapMemory.Height), Drawing.Imaging.ImageLockMode.ReadOnly, btmBitmapMemory.PixelFormat)

        'Set texture image
        glTexImage2D(intGL_TEXTURE_2D, 0, intInternalFormat, btmdBitmapData.Width, btmdBitmapData.Height, 0, intGL_BGRA_EXT,
                     intGL_UNSIGNED_BYTE, btmdBitmapData.Scan0)

        'Unlock
        btmBitmapMemory.UnlockBits(btmdBitmapData)

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
                    If udcCharacter.GetPictureFrame = 0 Then
                        'Reload
                        udcCharacter.Reload()
                        'Set
                        udcCharacter.StatusModeAboutToDo = clsCharacter.eintStatusMode.Stand
                    End If

                Case clsCharacter.eintStatusMode.Run
                    'Clear the key press buffer
                    strKeyPressBuffer = ""
                    'Make the character run
                    If udcCharacter.GetPictureFrame = 0 Then
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
                            udcCharacter.Reload()

                        Case clsCharacter.eintStatusMode.Stand, clsCharacter.eintStatusMode.Shoot
                            'Make the character shoot with buffer
                            CheckTheKeyPressBuffer(udcCharacter, gaudcZombies, intZombieKills)

                    End Select

            End Select

        End If

    End Sub

    Private Sub CheckTheKeyPressBuffer(udcCharacterType As clsCharacter, audcZombiesType() As clsZombie, ByRef intByRefZombieKills As Integer)

        'Check the key press buffer
        If strKeyPressBuffer <> "" Then

            'Declare to be split
            Dim astrTemp() As String = Split(strKeyPressBuffer, ".") 'Each element is used because delimiter came first instead of last

            'Prepare to send data if host
            Dim astrZombiesToKill(0) As String 'Can be used to detect ""
            Dim strZombieDeathData As String = ""

            'Check typing time
            If Not gswhTimeTyped.IsRunning Then
                gswhTimeTyped.Start()
            End If

            'Loop through
            For intLoop As Integer = 0 To (astrTemp.GetUpperBound(0) - 1) 'The last element is null but still counts as an element

                'Check if typing matches
                If strEntireLengthOfWords.Substring(0, 1) = LCase(astrTemp(intLoop)) Then 'Correct
                    'Figure out how to remove properly
                    If strEntireLengthOfWords.Substring(1, 1) = " " Then
                        'Increase typed words
                        intTypedWords += 1
                        'Shoot
                        udcCharacterType.Shoot()
                        'Check if spelled correctly
                        If Not blnWordWrong Then
                            'Check to make sure there is a spawned zombie to kill and isn't dying
                            If blnZombieToKillExists(audcZombiesType) Then
                                'Declare
                                Dim intIndexOfClosestZombie As Integer = intGetIndexOfClosestZombie(audcZombiesType)
                                'Increase
                                intByRefZombieKills += 1
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
                                        gSendData(5) 'This makes zombie waiting kills count up
                                    End If
                                Else
                                    'Kill zombie
                                    audcZombiesType(intIndexOfClosestZombie).Dying()
                                End If
                            End If
                        Else
                            'Send data missed shot
                            If blnGameIsVersus Then
                                'Check
                                If blnHost Then
                                    'Show word for host
                                    udcOnScreenWordOne.ShowWord(90, 0, intON_SCREEN_WORD_MISSED_RED_X, intON_SCREEN_WORD_MISSED_RED_Y)
                                Else
                                    'Show word for joiner
                                    udcOnScreenWordTwo.ShowWord(90, 1, intON_SCREEN_WORD_MISSED_BLUE_X, intON_SCREEN_WORD_MISSED_BLUE_Y)
                                End If
                                'Send data
                                gSendData(7)
                            Else
                                'Show missed word on screen
                                udcOnScreenWordOne.ShowWord(90, 0, intON_SCREEN_WORD_MISSED_RED_X, intON_SCREEN_WORD_MISSED_RED_Y)
                            End If
                        End If
                        'Remove two
                        strEntireLengthOfWords = strEntireLengthOfWords.Substring(2)
                        'Get a new word
                        LoadARandomWord()
                        'Check if needs to reload
                        If udcCharacterType.BulletsUsed >= 30 Then
                            'Set
                            udcCharacterType.StatusModeAboutToDo = clsCharacter.eintStatusMode.Reload
                            'Reset
                            strKeyPressBuffer = ""
                            'Exit
                            Exit For
                        End If
                    Else
                        'Remove one
                        strEntireLengthOfWords = strEntireLengthOfWords.Substring(1)
                    End If
                Else 'Incorrect
                    'Set
                    blnWordWrong = True
                End If

                'Remove the piece
                strKeyPressBuffer = strKeyPressBuffer.Substring(2)

            Next

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

    Private Function intGetIndexOfClosestZombie(audcZombiesType() As clsZombie) As Integer

        'Declare
        Dim intClosestX As Integer = Integer.MaxValue
        Dim intIndex As Integer = 0

        'Loop to get closest zombie
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            'If spawned, if not dying, and if closest point is less than another point
            If audcZombiesType(intLoop).Spawned And Not audcZombiesType(intLoop).IsDying And intClosestX > audcZombiesType(intLoop).Point.X Then
                'Set closest
                intClosestX = audcZombiesType(intLoop).Point.X
                'Set index
                intIndex = intLoop
            End If
        Next

        'Return
        Return intIndex

    End Function

    Private Sub DrawLevel2WaterOnScreen()

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH, intWaterHeight)

        'Clone only the necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackground2WaterMemory.Clone(rectSource,
                                            Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw the background water to screen
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmGameBackgroundCloneScreenShown, New Point(0, intORIGINAL_SCREEN_HEIGHT -
                            intWaterHeight)) 'Disposed of later
        Else
            'Load OpenGL texture once and update
            LoadOpenGLTextureOnceAndUpdate(iptGameBackground2WaterTexture, btmGameBackgroundCloneScreenShown, intINTERNAL_FORMAT_RGBA)
            'Draw the background water to screen
            OpenGLDrawImage(iptGameBackground2WaterTexture, New Point(0, intORIGINAL_SCREEN_HEIGHT - intWaterHeight),
                            btmGameBackgroundCloneScreenShown, True)
        End If

        'Dipose because clone just makes the picture expand with more data
        DisposeBitmap(btmGameBackgroundCloneScreenShown)

    End Sub

    Private Sub DrawWordBarWithWording(strColor As String)

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw word bar
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmWordBarMemory, pntWordBar)
        Else
            'Draw word bar
            OpenGLDrawImage(iptWordBarTexture, pntWordBar, btmWordBarMemory, True)
        End If

        'Declare
        Dim strEntireLengthOfWordsMinusLetters As String = strEntireLengthOfWords

        'Check the word
        If strEntireLengthOfWordsMinusLetters.Length() >= 18 Then
            strEntireLengthOfWordsMinusLetters = strEntireLengthOfWordsMinusLetters.Substring(0, 18) & "..."
        End If

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw text in the word bar
            GDIDrawText(btmCanvas, 36, "WHITE", strEntireLengthOfWordsMinusLetters, New Point(260, 47), 3)
            'Word overlay
            GDIDrawText(btmCanvas, 36, strColor, strEntireLengthOfWordsMinusLetters.Substring(0, 1), New Point(260, 47), 3)
        Else
            'Draw text in the word bar
            OpenGLDrawText(36, "WHITE", strEntireLengthOfWordsMinusLetters, New Point(260, 47), 3)
            'Word overlay
            OpenGLDrawText(36, strColor, strEntireLengthOfWordsMinusLetters.Substring(0, 1), New Point(260, 47), 3)
        End If

    End Sub

    Private Sub DrawOnScreenWord(udcOnScreenWord As clsOnScreenWord)

        'Draw
        If udcOnScreenWord.OnScreen Then
            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'Draw
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), udcOnScreenWord.Image, udcOnScreenWord.Point)
            Else
                'Draw
                OpenGLDrawImage(aiptOnScreenWordMissedTextures(udcOnScreenWord.TextureIndex), udcOnScreenWord.Point,
                                udcOnScreenWord.Image, True)
            End If
        End If

    End Sub

    Private Sub CheckEndGame()

        'Check if need to black screen
        If blnBeatLevel Or blnPlayerWasPinned Then
            'Paint black overlay
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), abtmBlackScreenMemories(intBlackScreenWaitMode), pntTopLeft)
            'Check if OpenGL
            If blnOpenGL Then
                'Paint black overlay
                OpenGLDrawImage(aiptBlackScreenTextures(intBlackScreenWaitMode), pntTopLeft,
                                abtmBlackScreenMemories(intBlackScreenWaitMode), True)
            End If
            'Check
            If intBlackScreenWaitMode = 2 And blnBeatLevel Then
                'Set
                blnBlackScreenFinished = True
            End If
        End If

        'Check if need to copy screen
        If blnPlayerWasPinned Then
            'Check
            If intBlackScreenWaitMode = 2 Then
                'Before fading the screen, copy it to show for the death overlay
                Dim rectSource As New Rectangle(0, 0, intORIGINAL_SCREEN_WIDTH, intORIGINAL_SCREEN_HEIGHT)
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'After paint make a copy
                    btmDeathScreen = btmCanvas.Clone(rectSource, Imaging.PixelFormat.Format32bppPArgb)
                Else
                    'Make an OpenGL screenshot
                    MakeOpenGLScreenShot()
                    'Load OpenGL texture once And update
                    LoadOpenGLTextureOnceAndUpdate(iptDeathScreenTexture, btmDeathScreen, intINTERNAL_FORMAT_RGB)
                End If
                'Set
                blnBlackScreenFinished = True
            End If
        End If

    End Sub

    Private Sub MakeOpenGLScreenShot()

        'Notes: Use GDI to draw on the background clone and that will be the bitmap loaded as a texture into OpenGL. Objects
        '       drawn on the background behind other objects are in the btmGameBackground already.

        'Clone the background
        btmDeathScreen = btmGameBackgroundMemory.Clone(New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH,
                             intORIGINAL_SCREEN_HEIGHT), Imaging.PixelFormat.Format32bppPArgb)

        'Check if single player or multiplayer
        If Not blnGameIsVersus Then
            'Draw character
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), udcCharacter.Image, udcCharacter.Point)
            'Draw all alive zombies
            DrawAllAliveZombiesScreenShot(gaudcZombies)
            'Check level for special effect
            Select Case gintLevel
                Case 1
                    'Draw brick wall
                    GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmDestroyedBrickWallMemory, gpntDestroyedBrickWall)
                Case 2
                    'Chained zombie
                    GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), gudcChainedZombie.Image, gudcChainedZombie.Point)
                    'OpenGL level 2 water copy on the screen shot
                    Level2WaterCopyScreenShot()
                    'Draw pipe value
                    GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmPipeValveMemory, gpntPipeValve)
                    'Draw zombie face
                    GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), gudcFaceZombie.Image, gudcFaceZombie.Point)
            End Select
            'OpenGL draw word bar on the screen shot
            DrawWordBarScreenShot("RED")
            'Show magazine with bullet count
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmAK47MagazineMemory, pntAK47Magazine)
            'Draw bullet count on magazine
            GDIDrawText(btmDeathScreen, 25, "RED", (30 - udcCharacter.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            'Draw on screen word
            DrawOnScreenWordScreenShot(udcOnScreenWordOne)
        Else
            'Draw character hoster
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), udcCharacterOne.Image, udcCharacterOne.Point)
            'Draw all alive zombies
            DrawAllAliveZombiesScreenShot(gaudcZombiesOne)
            'Draw character joiner
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), udcCharacterTwo.Image, udcCharacterTwo.Point)
            'Draw all alive zombies
            DrawAllAliveZombiesScreenShot(gaudcZombiesTwo)
            'Draw word bar with wording
            If blnHost Then
                DrawWordBarScreenShot("RED")
            Else
                DrawWordBarScreenShot("BLUE")
            End If
            'Draw nickname
            If blnHost Then
                GDIDrawText(btmDeathScreen, 25, "RED", strNickName, New Point(63, 117), 2) 'Host sees own name
                GDIDrawText(btmDeathScreen, 25, "BLUE", strNickNameConnected, New Point(118, 142), 2) 'Host sees joiner name
            Else
                GDIDrawText(btmDeathScreen, 25, "RED", strNickNameConnected, New Point(63, 117), 2) 'Joiner sees host name
                GDIDrawText(btmDeathScreen, 25, "BLUE", strNickName, New Point(118, 142), 2) 'Joiner sees own name
            End If
            'Show magazine with bullet count
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmAK47MagazineMemory, pntAK47Magazine)
            'Draw bullet count on magazine
            If blnHost Then
                GDIDrawText(btmDeathScreen, 25, "RED", (30 - udcCharacterOne.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            Else
                GDIDrawText(btmDeathScreen, 25, "BLUE", (30 - udcCharacterTwo.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            End If
            'Draw on screen word
            DrawOnScreenWordScreenShot(udcOnScreenWordOne)
            'Draw on screen word
            DrawOnScreenWordScreenShot(udcOnScreenWordTwo)
        End If

        'Draw black fade
        GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), abtmBlackScreenMemories(2), pntTopLeft)

    End Sub

    Private Sub DrawAllAliveZombiesScreenShot(audcZombiesType() As clsZombie)

        'Draw all alive zombies
        For intLoop As Integer = 0 To audcZombiesType.GetUpperBound(0)
            If Not audcZombiesType(intLoop).IsDead Then
                'Draw zombies dying, pinning or walking
                GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), audcZombiesType(intLoop).Image, audcZombiesType(intLoop).Point)
            End If
        Next

    End Sub

    Private Sub Level2WaterCopyScreenShot()

        'Declare
        Dim rectSource As New Rectangle(Math.Abs(gpntGameBackground.X), 0, intORIGINAL_SCREEN_WIDTH, intWaterHeight)

        'Clone only the necessary spot
        btmGameBackgroundCloneScreenShown = btmGameBackground2WaterMemory.Clone(rectSource,
                                            Imaging.PixelFormat.Format32bppPArgb) 'If out of memory here, could be x + width is too short

        'Draw the background water to screen
        GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmGameBackgroundCloneScreenShown, New Point(0, intORIGINAL_SCREEN_HEIGHT -
                        intWaterHeight)) 'Disposed of later

        'Dipose because clone just makes the picture expand with more data
        DisposeBitmap(btmGameBackgroundCloneScreenShown)

    End Sub

    Private Sub DrawWordBarScreenShot(strColor As String)

        'Draw word bar
        GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), btmWordBarMemory, pntWordBar)

        'Declare
        Dim strEntireLengthOfWordsMinusLetters As String = strEntireLengthOfWords

        'Check the word
        If strEntireLengthOfWordsMinusLetters.Length() >= 18 Then
            strEntireLengthOfWordsMinusLetters = strEntireLengthOfWordsMinusLetters.Substring(0, 18) & "..."
        End If

        'Draw text in the word bar
        GDIDrawText(btmDeathScreen, 36, "WHITE", strEntireLengthOfWordsMinusLetters, New Point(260, 47), 3)

        'Word overlay
        GDIDrawText(btmDeathScreen, 36, strColor, strEntireLengthOfWordsMinusLetters.Substring(0, 1), New Point(260, 47), 3)

    End Sub

    Private Sub DrawOnScreenWordScreenShot(udcOnScreenWordType As clsOnScreenWord)

        'Draw on screen word
        If udcOnScreenWordType.OnScreen Then
            GDIDrawGraphics(Graphics.FromImage(btmDeathScreen), udcOnScreenWordType.Image, udcOnScreenWordType.Point)
        End If

    End Sub

    Private Sub HighscoresScreen()

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw highscores background
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmHighscoresBackgroundMemory, pntTopLeft)
        Else
            'Draw highscores background
            OpenGLDrawFullScreenImage(aiptHighscoreTextures(0))
        End If

        'Declare
        Static strRankNumbers As String = ""

        'Set rank numbers
        SetRankNumbers(strRankNumbers)

        'Declare
        Dim intHighscoreY As Integer = 0

        'Loop to draw scores
        For intLoop As Integer = 0 To astrHighscores.GetUpperBound(0)
            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'Draw highscores
                GDIDrawText(btmCanvas, 25, "RED", astrHighscores(intLoop), New Point(18, 105 + intHighscoreY), 2)
                GDIDrawText(btmCanvas, 25, "WHITE", (intLoop + 1).ToString & ".", New Point(18, 105 + intHighscoreY), 2)
            Else
                'Draw highscores
                OpenGLDrawText(25, "RED", astrHighscores(intLoop), New Point(18, 105 + intHighscoreY), 2)
                OpenGLDrawText(25, "WHITE", (intLoop + 1).ToString & ".", New Point(18, 105 + intHighscoreY), 2)
            End If
            'Increase
            intHighscoreY += 27
        Next

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Access written on screen as type of database
            GDIDrawText(btmCanvas, 25, "RED", "Database type is", New Point(18, 15), 2)
            GDIDrawText(btmCanvas, 25, "WHITE", "Access .MDB File (Password Protected)", New Point(202, 15), 2)
            'Show back button or hover back button
            GDIShowBackButtonOrHoverBackButton()
        Else
            'Access written on screen as type of database
            OpenGLDrawText(25, "RED", "Database type is", New Point(18, 15), 2)
            OpenGLDrawText(25, "WHITE", "Access .MDB File (Password Protected)", New Point(202, 15), 2)
            'Show back button or hover back button
            OpenGLShowBackButtonOrHoverBackButton()
        End If

    End Sub

    Private Sub SetRankNumbers(ByRef strByRefRankNumbers As String)

        'Set rank numbers
        If strByRefRankNumbers = "" Then
            'Loop
            For intLoop As Integer = 1 To 10
                'Add a delimiter of "."
                strByRefRankNumbers &= CStr(intLoop) & "."
                'Check if need to add a new line
                If intLoop <> 10 Then
                    'Add a new line
                    strByRefRankNumbers &= vbNewLine
                End If
            Next
        End If

    End Sub

    Private Sub GDICreditsScreen()

        'Draw credits background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmCreditsBackgroundMemory, pntTopLeft)

        'Draw John Gonzales
        If btmJohnGonzales IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmJohnGonzales, pntJohnGonzales)
        End If

        'Draw Zachary Stafford
        If btmZacharyStafford IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmZacharyStafford, pntZacharyStafford)
        End If

        'Draw Cory Lewis
        If btmCoryLewis IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmCoryLewis, pntCoryLewis)
        End If

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLCreditsScreen()

        'Draw credits background
        OpenGLDrawFullScreenImage(aiptCreditTextures(0))

        'Draw John Gonzales
        If intJohnGonzlesTextureIndex <> 0 Then
            OpenGLDrawImage(aiptCreditTextures(intJohnGonzlesTextureIndex), pntJohnGonzales, abtmJohnGonzalesMemories(0),
                            True) 'Index has same width and spec
        End If

        'Draw Zachary Stafford
        If intZacharyStaffordTextureIndex <> 0 Then
            OpenGLDrawImage(aiptCreditTextures(intZacharyStaffordTextureIndex), pntZacharyStafford, abtmZacharyStaffordMemories(0),
                            True) 'Index has same width and spec
        End If

        'Draw Cory Lewis
        If intCoryLewisTextureIndex <> 0 Then
            OpenGLDrawImage(aiptCreditTextures(intCoryLewisTextureIndex), pntCoryLewis, abtmCoryLewisMemories(0),
                            True) 'Index has same width and spec
        End If

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub GDIVersusScreen()

        'Draw background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmVersusBackgroundMemory, pntTopLeft)

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

        'Other
        Select Case intCanvasVersusShow
            Case 0 'Nickname
                'Draw host button
                GDIDrawNonHoverOrHoverText(2, btmVersusHostTextMemory, pntVersusHost, btmVersusHostTextHoverMemory, pntVersusHostHover)
                'Draw or
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmVersusOrTextMemory, pntVersusOr)
                'Draw join button
                GDIDrawNonHoverOrHoverText(3, btmVersusJoinTextMemory, pntVersusJoin, btmVersusJoinTextHoverMemory, pntVersusJoinHover)
                'Draw nickname
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmVersusNickNameTextMemory, pntVersusBlackOutline)
                'Draw player name text
                GDIDrawText(btmCanvas, 72, "WHITE", strNickName, New Point(63, 185), 2)
            Case 1 'Host
                'Draw hosting text
                GDIDrawText(btmCanvas, 42, "RED", "Hosting...", New Point(329, 240), 2)
            Case 2 'Join
                'Draw IP address
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmVersusIPAddressTextMemory, pntVersusBlackOutline)
                'Draw IP address to type
                GDIDrawText(btmCanvas, 72, "WHITE", strIPAddressConnect, New Point(63, 185), 2)
                'Draw connect button
                GDIDrawNonHoverOrHoverText(4, btmVersusConnectTextMemory, pntVersusConnect, btmVersusConnectTextHoverMemory, pntVersusConnectHover)
            Case 3 'Connecting
                'Draw connecting text
                GDIDrawText(btmCanvas, 42, "RED", "Connecting...", New Point(293, 240), 2)
        End Select

        'Draw ip address text
        GDIDrawText(btmCanvas, 42, "RED", strIPAddress, New Point(20, 27), 2)

        'Draw port forwarding
        GDIDrawText(btmCanvas, 36, "WHITE", "Router port forwarding: 10101", New Point(180, 452), 2)

    End Sub

    Private Sub OpenGLVersusScreen()

        'Draw background
        OpenGLDrawFullScreenImage(aiptVersusScreenTextures(0))

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

        'Other
        Select Case intCanvasVersusShow
            Case 0 'Nickname
                'Draw host button
                OpenGLDrawNonHoverOrHoverText(2, aiptVersusScreenTextures(1), btmVersusHostTextMemory, pntVersusHost,
                                              aiptVersusScreenTextures(2), btmVersusHostTextHoverMemory, pntVersusHostHover)
                'Draw or
                OpenGLDrawImage(aiptVersusScreenTextures(3), pntVersusOr, btmVersusOrTextMemory, True)
                'Draw join button
                OpenGLDrawNonHoverOrHoverText(3, aiptVersusScreenTextures(4), btmVersusJoinTextMemory, pntVersusJoin,
                                              aiptVersusScreenTextures(5), btmVersusJoinTextHoverMemory, pntVersusJoinHover)
                'Draw nickname
                OpenGLDrawImage(aiptVersusScreenTextures(6), pntVersusBlackOutline, btmVersusNickNameTextMemory, True)
                'Draw player name text
                OpenGLDrawText(72, "WHITE", strNickName, New Point(63, 185), 2)
            Case 1 'Host
                'Draw hosting text
                OpenGLDrawText(42, "RED", "Hosting...", New Point(329, 240), 2)
            Case 2 'Join
                'Draw IP address
                OpenGLDrawImage(aiptVersusScreenTextures(7), pntVersusBlackOutline, btmVersusIPAddressTextMemory, True)
                'Draw IP address to type
                OpenGLDrawText(72, "WHITE", strIPAddressConnect, New Point(63, 185), 2)
                'Draw connect button
                OpenGLDrawNonHoverOrHoverText(4, aiptVersusScreenTextures(8), btmVersusConnectTextMemory, pntVersusConnect,
                                              aiptVersusScreenTextures(9), btmVersusConnectTextHoverMemory, pntVersusConnectHover)
            Case 3 'Connecting
                'Draw connecting text
                OpenGLDrawText(42, "RED", "Connecting...", New Point(293, 240), 2)
        End Select

        'Draw ip address text
        OpenGLDrawText(42, "RED", strIPAddress, New Point(20, 27), 2)

        'Draw port forwarding
        OpenGLDrawText(36, "WHITE", "Router port forwarding: 10101", New Point(180, 452), 2)

    End Sub

    Private Sub GDILoadingVersusConnectedScreen()

        'Draw loading background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBackgroundMemory, pntTopLeft)

        'Draw loading bar
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingBar, pntLoadingBar)

        'Draw loading, waiting, and start text
        If blnWaiting Then
            'Draw waiting
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingWaitingTextMemory, pntLoadingWaitingText)
        Else
            'Check if finished loading
            If blnFinishedLoading Then
                'Start
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingStartTextMemory, pntLoadingStartText)
            Else
                'Loading
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingTextMemory, pntLoadingText)
            End If
        End If

        'Draw paragraph
        If btmLoadingParagraphVersus IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmLoadingParagraphVersus, pntLoadingParagraph)
        End If

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLLoadingVersusConnectedScreen()

        'Draw loading background
        OpenGLDrawFullScreenImage(aiptVersusLoadingTextures(0))

        'Draw loading bar
        OpenGLDrawImage(aiptLoadingTextures(intLoadingBarTextureIndex), pntLoadingBar,
                        abtmLoadingBarPictureMemories(0), True) 'Any index, same width and height specs

        'Draw loading, waiting, and start text
        If blnWaiting Then
            'Draw waiting
            OpenGLDrawImage(aiptVersusLoadingTextures(1), pntLoadingWaitingText, btmLoadingWaitingTextMemory, True)
        Else
            'Check if finished loading
            If blnFinishedLoading Then
                'Start
                OpenGLDrawImage(aiptLoadingTextures(13), pntLoadingStartText, btmLoadingStartTextMemory, True)
            Else
                'Loading
                OpenGLDrawImage(aiptLoadingTextures(12), pntLoadingText, btmLoadingTextMemory, True)
            End If
        End If

        'Draw paragraph
        If intVersusLoadingParagraphTextureIndex <> 0 Then
            OpenGLDrawImage(aiptVersusLoadingTextures(intVersusLoadingParagraphTextureIndex), pntLoadingParagraph,
                            abtmLoadingParagraphVersusMemories(0), True) 'Use any index, same width and height specs
        End If

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub StartedVersusGameScreen()

        'Check if black screen displayed
        If blnBlackScreenFinished Then

            'Check if GDI or OpenGL
            If Not blnOpenGL Then
                'Show death overlay
                GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmDeathScreen, pntTopLeft)
                'Check which mode to draw
                If intVersusWhoWonMode = 0 Then
                    'Draw won
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmYouWonMemory, pntTopLeft)
                Else
                    'Draw won
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmYouLostMemory, pntTopLeft)
                End If
            Else
                'Show death overlay
                OpenGLDrawFullScreenImage(iptDeathScreenTexture)
                'Check the index of the pointer
                If intVersusWhoWonMode = 0 Then
                    'Draw won
                    OpenGLDrawImage(iptYouWonTexture, pntTopLeft, btmYouWonMemory, True)
                Else
                    'Draw lost
                    OpenGLDrawImage(iptYouLostTexture, pntTopLeft, btmYouLostMemory, True)
                End If
            End If

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

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Show back button or hover back button
            GDIShowBackButtonOrHoverBackButton()
        Else
            'Show back button or hover back button
            OpenGLShowBackButtonOrHoverBackButton()
        End If

    End Sub

    Private Sub PlayMultiplayerGame()

        'Enable time elapsed
        StartTimeElapsed()

        'Draw dead zombies permanently for hoster
        DrawDeadZombiesPermanently(gaudcZombiesOne, intZombiesKilledIncreaseSpawnOne)

        'Draw dead zombies permanently for joiner
        DrawDeadZombiesPermanently(gaudcZombiesTwo, intZombiesKilledIncreaseSpawnTwo)

        'Paint on canvas and clone the only necessary spot of the background
        PaintOnCanvasAndCloneScreen()

        'Check character status if not game over
        If Not blnPlayerWasPinned Then
            'Check character status
            If blnHost Then
                CheckCharacterMultiplayerStatus(udcCharacterOne, gaudcZombiesOne, intZombieKillsOne)
            Else
                CheckCharacterMultiplayerStatus(udcCharacterTwo, gaudcZombiesTwo, intZombieKillsTwo)
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
                'Increase
                intZombieKillsTwo += 1
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
            CheckMultiplayerZombieKillBuffer(strZombieKillBufferOne, udcCharacterOne, gaudcZombiesOne, intZombieKillsOne, True)
            'Zombies two
            CheckMultiplayerZombieKillBuffer(strZombieKillBufferTwo, udcCharacterTwo, gaudcZombiesTwo,
                                             intZombieKillsTwo) 'Joiner on the join screen doesn't need to shoot
        End If

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw character hoster
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), udcCharacterOne.Image, udcCharacterOne.Point)
        Else
            'Draw character hoster
            OpenGLDrawImage(aiptCharacterTextures(udcCharacterOne.TextureIndex), udcCharacterOne.Point, udcCharacterOne.Image, True)
        End If

        'Draw first zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesOne, intZOMBIE_PINNING_X, intZombieIncreasedPinDistanceOne, 8)

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw character joiner
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), udcCharacterTwo.Image, udcCharacterTwo.Point)
        Else
            'Draw character joiner
            OpenGLDrawImage(aiptCharacterTextures(udcCharacterTwo.TextureIndex), udcCharacterTwo.Point, udcCharacterTwo.Image, True)
        End If

        'Draw second zombies
        DrawMultiplayerZombiesAndSendData(gaudcZombiesTwo, intZOMBIE_PINNING_X + intJOINER_ADDED_X, intZombieIncreasedPinDistanceTwo, 9)

        'Send X Positions
        gSendData(3, strGetXPositionsOfZombies())

        'Draw word bar with wording
        If blnHost Then
            DrawWordBarWithWording("RED")
        Else
            DrawWordBarWithWording("BLUE")
        End If

        'Check if GDI or OpenGL
        If Not blnOpenGL Then
            'Draw nickname
            If blnHost Then
                GDIDrawText(btmCanvas, 25, "RED", strNickName, New Point(63, 117), 2) 'Host sees own name
                GDIDrawText(btmCanvas, 25, "BLUE", strNickNameConnected, New Point(118, 142), 2) 'Host sees joiner name
            Else
                GDIDrawText(btmCanvas, 25, "RED", strNickNameConnected, New Point(63, 117), 2) 'Joiner sees host name
                GDIDrawText(btmCanvas, 25, "BLUE", strNickName, New Point(118, 142), 2) 'Joiner sees own name
            End If
            'Show magazine with bullet count
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmAK47MagazineMemory, pntAK47Magazine)
            'Draw bullet count on magazine
            If blnHost Then
                GDIDrawText(btmCanvas, 25, "RED", (30 - udcCharacterOne.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            Else
                GDIDrawText(btmCanvas, 25, "BLUE", (30 - udcCharacterTwo.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            End If
        Else
            'Draw nickname
            If blnHost Then
                OpenGLDrawText(25, "RED", strNickName, New Point(63, 117), 2) 'Host sees own name)
                OpenGLDrawText(25, "BLUE", strNickNameConnected, New Point(118, 142), 2) 'Host sees joiner name)
            Else
                OpenGLDrawText(25, "RED", strNickNameConnected, New Point(63, 117), 2) 'Joiner sees host name
                OpenGLDrawText(25, "BLUE", strNickName, New Point(118, 142), 2) 'Joiner sees own name
            End If
            'Show magazine with bullet count
            OpenGLDrawImage(iptAK47MagazineTexture, pntAK47Magazine, btmAK47MagazineMemory, True)
            'Draw bullet count on magazine
            If blnHost Then
                OpenGLDrawText(25, "RED", (30 - udcCharacterOne.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            Else
                OpenGLDrawText(25, "BLUE", (30 - udcCharacterTwo.BulletsUsed).ToString, New Point(pntAK47Magazine.X + 1, pntAK47Magazine.Y + 30), 2)
            End If
        End If

        'Draw missed word on screen
        DrawOnScreenWord(udcOnScreenWordOne)

        'Draw missed word on screen
        DrawOnScreenWord(udcOnScreenWordTwo)

        'Check if need to black screen and make copy if died
        CheckEndGame()

    End Sub

    Private Sub CheckCharacterMultiplayerStatus(udcCharacterType As clsCharacter, audcZombiesType() As clsZombie,
                                                ByRef intByRefZombieKills As Integer)

        'Check to see if character must prepare to do something first
        Select Case udcCharacterType.StatusModeAboutToDo

            Case clsCharacter.eintStatusMode.Reload
                'Clear the key press buffer
                strKeyPressBuffer = ""
                'Make the character reload
                If udcCharacterType.GetPictureFrame = 0 Then
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
                        udcCharacterType.Reload() 'Only reload after finishing a gun shot frame

                    Case clsCharacter.eintStatusMode.Stand, clsCharacter.eintStatusMode.Shoot
                        'Make the character shoot with buffer
                        CheckTheKeyPressBuffer(udcCharacterType, audcZombiesType, intByRefZombieKills)

                End Select

        End Select

    End Sub

    Private Sub CheckMultiplayerZombieKillBuffer(ByRef strByRefZombieKillBufferType As String, udcCharacterType As clsCharacter,
                                                 audcZombiesType() As clsZombie, ByRef intByRefZombieKills As Integer,
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
                    'Shoot
                    udcCharacterType.Shoot()
                    'Increase
                    intByRefZombieKills += 1
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
                                gblnPreventKeyPressEvent = True
                                'Start black screen
                                BlackScreening(intBLACK_SCREEN_DEATH_DELAY)
                                'Stop reloading sound
                                udcReloadingSound.StopSound()
                                'Play
                                udcScreamSound.PlaySound(gintSoundVolume)
                                'Stop level music
                                audcGameBackgroundSounds(gintLevel - 1).StopSound()
                                'Game over, time stopped
                                StopTheStopWatches()
                                'Send data
                                gSendData(intDataCase) 'Joiner won if 8 and 9 joiner lost
                                'Show who won to host
                                Select Case intDataCase
                                    Case 8
                                        intVersusWhoWonMode = 1
                                    Case 9
                                        intVersusWhoWonMode = 0
                                End Select
                            End If
                        End If
                    Else
                        'Check distance
                        If intTempDistance <= intByRefZombieIncreasedPinDistance Then
                            'Increase distance
                            intByRefZombieIncreasedPinDistance -= 25
                            'Make zombie pin
                            audcZombiesType(intLoop).Pin()
                        End If
                    End If
                End If
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Draw zombies dying, pinning or walking
                    GDIDrawGraphics(Graphics.FromImage(btmCanvas), audcZombiesType(intLoop).Image, audcZombiesType(intLoop).Point)
                Else
                    'Draw zombies dying, pinning or walking
                    OpenGLDrawImage(aiptZombieTextures(audcZombiesType(intLoop).TextureIndex), audcZombiesType(intLoop).Point,
                                    audcZombiesType(intLoop).Image, True)
                End If
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

    Private Sub GDIStoryScreen()

        'Draw story background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmStoryBackgroundMemory, pntTopLeft)

        'Draw story text
        If btmStoryParagraph IsNot Nothing Then
            GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmStoryParagraph, pntStoryParagraph)
        End If

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLStoryScreen()

        'Draw story background
        OpenGLDrawFullScreenImage(aiptStoryTextures(0))

        'Draw story text
        If intStoryParagraphTextureIndex <> 0 Then
            'Check which type of index because of different width and height specs
            Select Case intStoryParagraphTextureIndex
                Case 1 To 4
                    OpenGLDrawImage(aiptStoryTextures(intStoryParagraphTextureIndex),
                                    pntStoryParagraph, abtmStoryParagraphMemories(0), True) '0 to 3 (same width and height)
                Case 5 To 8
                    OpenGLDrawImage(aiptStoryTextures(intStoryParagraphTextureIndex),
                                    pntStoryParagraph, abtmStoryParagraphMemories(4), True) '4 to 7 (same width and height)
                Case 9 To 12
                    OpenGLDrawImage(aiptStoryTextures(intStoryParagraphTextureIndex),
                                    pntStoryParagraph, abtmStoryParagraphMemories(8), True) '8 to 11 (same width and height)
            End Select
        End If

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub GDIGameVersionMismatch()

        'Draw mismatch background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmGameMismatchBackgroundMemory, pntTopLeft)

        'Display
        GDIDrawText(btmCanvas, 36, "RED", "Your version: " & strGAME_VERSION, New Point(452, 113), 2)

        'Draw text
        If blnHost Then
            GDIDrawText(btmCanvas, 36, "RED", "Joiner version: " & strGameVersionFromConnection, New Point(452, 213), 2)
        Else
            GDIDrawText(btmCanvas, 36, "RED", "Host version: " & strGameVersionFromConnection, New Point(452, 213), 2)
        End If

    End Sub

    Private Sub OpenGLGameVersionMismatch()

        'Draw mismatch background
        OpenGLDrawFullScreenImage(aiptGameMismatchTextures(0))

        'Display
        OpenGLDrawText(36, "RED", "Your version: " & strGAME_VERSION, New Point(452, 113), 2)

        'Draw text
        If blnHost Then
            OpenGLDrawText(36, "RED", "Joiner version: " & strGameVersionFromConnection, New Point(452, 213), 2)
        Else
            OpenGLDrawText(36, "RED", "Host version: " & strGameVersionFromConnection, New Point(452, 213), 2)
        End If

    End Sub

    Private Sub GDIPathChoices()

        'Draw background
        GDIDrawGraphics(Graphics.FromImage(btmCanvas), btmPath, pntTopLeft)

        'Text
        GDIDrawText(btmCanvas, 36, "RED", "Pick your path...", New Point(18, 480), 2)

        'Show back button or hover back button
        GDIShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub OpenGLPathChoices()

        'Draw background
        OpenGLDrawFullScreenImage(aiptPathTextures(intPathTextureIndex))

        'Text
        OpenGLDrawText(36, "RED", "Pick your path...", New Point(18, 480), 2)

        'Show back button or hover back button
        OpenGLShowBackButtonOrHoverBackButton()

    End Sub

    Private Sub frmGame_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp

        'Check for FKey presses
        Select Case e.KeyCode
            Case Keys.F12
                'Set
                blnDrawFPS = Not blnDrawFPS
            Case Keys.F11
                'Set
                blnGodMode = Not blnGodMode
        End Select

    End Sub

    Private Sub frmGame_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

        'Notes: This sub has to be synced with the rendering thread. The problem is everything here must be captured and checked by the render thread.

        'Keys pressed
        'Debug.Print(CStr(e.KeyChar))

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

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Wait
        While blnBackFromGame
            Application.DoEvents() 'Need this or blnBackFromGame will never change in this event, it changes in a thread, but event cannot read it properly
        End While

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
                strIPAddress = strGetLocalIPAddress()
                'Set
                blnFirstTimeNicknameTyping = True
                blnFirstTimeIPAddressTyping = True
                'Change
                ShowNextScreenAndExitMenu(6, 0)

        End Select

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
                'Change interval
                tmrFog.Interval = dblFOG_FRAME_WAIT_TO_START
                'Enable fog
                tmrFog.Enabled = True

            Case blnMouseInRegion(pntMouse, 147, 24, pntGDIPlusSoftwareText)
                'Set
                blnOpenGL = False

            Case blnMouseInRegion(pntMouse, 178, 24, pntOpenGLHardwareText)
                'Only click if OpenGL did load
                If blnOpenGLLoaded Then
                    'Set
                    blnOpenGL = True
                End If

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
        If blnFinishedLoading And blnMouseInRegion(pntMouse, 806, 67, pntLoadingBar) Then
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
            'Create stop watches
            CreateStopWatches()
            'Play background sound music
            audcGameBackgroundSounds(gintLevel - 1).PlaySound(CInt(Math.Round(gintSoundVolume / 4)), True)
        Else
            'Back was clicked
            GeneralBackButtonClick(pntMouse, True)
        End If

    End Sub

    Private Sub VersusMouseClickScreen(pntMouse As Point)

        'Check which button is clicked
        Select Case intCanvasVersusShow

            Case 0 'Default screen

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                        'Check nickname
                        DefaultNickName()
                        'Go back to menu
                        GeneralBackButtonClick(New Point(-1, -1), True, True) 'Point doesn't matter here, forcing back button activity

                    Case blnMouseInRegion(pntMouse, 220, 88, pntVersusHost) 'Host was clicked
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Set
                        blnHost = True
                        'Check nickname
                        DefaultNickName()
                        'Set
                        intCanvasVersusShow = 1
                        'Avoid error of hosting twice on the computer
                        Try
                            'Set
                            tcplServer = New System.Net.Sockets.TcpListener(System.Net.IPAddress.Any, 10101)
                            'Start
                            tcplServer.Start()
                            'Set
                            blnListening = True
                            'Start thread listening
                            thrListening = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Listening))
                            thrListening.Start()
                        Catch ex As Exception 'Instead go back
                            'Check nickname
                            DefaultNickName()
                            'Go back to menu
                            GeneralBackButtonClick(New Point(-1, -1), True, True) 'Point doesn't matter here, forcing back button activity
                        End Try

                    Case blnMouseInRegion(pntMouse, 200, 85, pntVersusJoin) 'Join was clicked
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Set
                        blnHost = False
                        'Set
                        strIPAddressConnect = strGetLocalIPAddress()
                        'Check nickname
                        DefaultNickName()
                        'Set
                        intCanvasVersusShow = 2

                End Select

            Case 1 'Hosting

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                        'Set
                        blnDoNotPlayPressedSoundAfterPressingBack = True
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Stop listening and empty variable
                        StopListeningOrConnectingAndEmptyVariable(blnListening, thrListening)
                        'Empty TCP server object
                        EmptyTCPObject(tcplServer)
                        'Set
                        blnHost = False
                        'Change mode
                        intCanvasVersusShow = 0

                End Select

            Case 2 'Nickname screen

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                        'Set
                        blnDoNotPlayPressedSoundAfterPressingBack = True
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Set
                        intCanvasVersusShow = 0

                    Case blnMouseInRegion(pntMouse, 302, 62, pntVersusConnect) 'Connect button was clicked
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Set
                        intCanvasVersusShow = 3
                        'Set
                        blnConnecting = True
                        'Start thread connecting
                        SetAndStartThread(thrConnecting, Sub() Connecting())

                End Select

            Case 3 'Connecting

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back button was clicked
                        'Set
                        blnDoNotPlayPressedSoundAfterPressingBack = True
                        'Play click sound
                        udcButtonPressedSound.PlaySound(gintSoundVolume)
                        'Abort client connecting, it is faster
                        AbortThread(thrConnecting)
                        'Stop connecting and empty variable
                        StopListeningOrConnectingAndEmptyVariable(blnConnecting, thrConnecting)
                        'Empty TCP client object
                        EmptyTCPObject(tcpcClient)
                        'Change mode
                        intCanvasVersusShow = 2

                End Select

        End Select

    End Sub

    Private Sub Listening()

        'Loop
        While blnListening
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
        While blnConnecting
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
                gSendData(0) 'Waiting for completed connection
                'Wait or else complications for exiting and reconnecting
                System.Threading.Thread.Sleep(1)
            End While
        End If

    End Sub

    Private Sub DefaultNickName()

        'Check nick name
        If Trim(strNickName) = "" Then
            strNickName = "Player"
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

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x 
        '       And y points to ensure it always works.

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
                VersusMouseOverScreen(pntMouse)

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

    Private Sub OptionsMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case True

            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText) 'Back has been moused over
                'Hover sound
                HoverText(1, "OptionsBack")

            Case blnSliderWithMouseDown 'Slider has been moused over
                'Change
                ChangeSoundVolume(pntMouse)
                'Reset mouse over variables
                ResetMouseOverVariables()

            Case Else
                'Reset mouse over variables
                ResetMouseOverVariables()

        End Select

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

    Private Sub frmGame_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp

        'If options
        If intCanvasMode = 1 Then
            'Sound changing done
            blnSliderWithMouseDown = False
        End If

    End Sub

    Private Sub ChangeSoundVolume(pntMouse As Point)

        'Notes: Found out if point is between 42 * width ratio and 341 * width ratio, these points were found on the
        '       original 525x840 picture.

        'Declare
        Dim intSoundIndex As Integer = 0

        'Check percent
        If pntMouse.X <= CInt(intSOUND_START_PIXEL * gsngScreenWidthRatio) Then
            'Change sound percent and volume
            ChangeSoundPercentPictureAndVolume(0, 0, intSoundIndex)
            'Slider
            pntSlider.X = intSOUND_LEFT_END_OF_SLIDER
        ElseIf pntMouse.X >= CInt(intSOUND_END_PIXEL * gsngScreenWidthRatio) Then
            'Change sound percent and volume
            ChangeSoundPercentPictureAndVolume(100, 1000, intSoundIndex)
            'Slider
            pntSlider.X = intSOUND_RIGHT_END_OF_SLIDER
        Else
            'Change sound percent and volume
            ChangeSoundPercentPictureAndVolume(CInt(((pntMouse.X / gsngScreenWidthRatio) - intSOUND_START_PIXEL) /
                                               intSOUND_DIVIDER_PERCENT), CInt((((pntMouse.X / gsngScreenWidthRatio) -
                                               intSOUND_START_PIXEL) / intSOUND_DIVIDER_PERCENT) *
                                               intSOUND_MULTIPLER_PERCENT), intSoundIndex)
            'Slider
            pntSlider.X = CInt(pntMouse.X / gsngScreenWidthRatio) - intSOUND_SLIDER_HALF_PIXEL
        End If

        'Set OpenGL picture texture index
        intSoundTextureIndex = intSoundIndex

        'After setting volume, change option sound
        audcAmbianceSound(1).ChangeVolumeWhileSoundIsPlaying()

    End Sub

    Private Sub ChangeSoundPercentPictureAndVolume(intSoundPercentIndex As Integer, intSoundVolumeIndex As Integer,
                                                   ByRef intByRefSoundIndex As Integer)

        'Change sound picture
        btmSoundPercent = abtmSoundMemories(intSoundPercentIndex)

        'Set sound volume
        gintSoundVolume = intSoundVolumeIndex

        'Set
        intByRefSoundIndex = intSoundPercentIndex + intSOUND_OPENGL_ARRAY_OFFSET

    End Sub

    Private Sub ResetMouseOverVariables()

        'Set timer off
        tmrOptionsMouseOver.Enabled = False

        'Reset
        strOptionsButtonSpot = ""

        'Repaint menu background
        intCanvasShow = 0

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

    Private Sub VersusMouseOverScreen(pntMouse As Point)

        'Check which options has been moused over
        Select Case intCanvasVersusShow

            Case 0 'Default screen

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
                        'Hover sound
                        HoverText(1, "VersusBack")

                    Case blnMouseInRegion(pntMouse, 220, 88, pntVersusHost)
                        'Hover sound
                        HoverText(2, "VersusHost")

                    Case blnMouseInRegion(pntMouse, 200, 85, pntVersusJoin)
                        'Hover sound
                        HoverText(3, "VersusJoin")

                    Case Else
                        'Reset mouse over variables
                        ResetMouseOverVariables()

                End Select

            Case 1 'Hosting screen

                'Do based on true, back button
                VersusMouseOverScreenBackButton(pntMouse, "VersusHostingBack")

            Case 2 'Connect screen

                'Do based on true
                Select Case True

                    Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
                        'Hover sound
                        HoverText(1, "VersusConnectBack")

                    Case blnMouseInRegion(pntMouse, 302, 62, pntVersusConnect)
                        'Hover sound
                        HoverText(4, "VersusConnect")

                    Case Else
                        'Reset mouse over variables
                        ResetMouseOverVariables()

                End Select

            Case 3 'Connecting screen

                'Do based on true, back button
                VersusMouseOverScreenBackButton(pntMouse, "VersusConnectingBack")

        End Select

    End Sub

    Private Sub VersusMouseOverScreenBackButton(pntMouse As Point, strOptionsButtonSpotToBe As String)

        'Do based on true
        Select Case True
            Case blnMouseInRegion(pntMouse, 95, 37, pntBackText)
                'Hover sound
                HoverText(1, strOptionsButtonSpotToBe)
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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath1Memories(1)
                Else
                    'Set
                    intPathTextureIndex = 1
                End If
                'Play light switch
                If Not blnLightZap1 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap1 = True
                End If

            Case blnMouseInRegion(pntMouse, 184, 153, New Point(547, 240)) 'Path right
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath1Memories(2)
                Else
                    'Set
                    intPathTextureIndex = 2
                End If
                'Play light switch
                If Not blnLightZap2 Then
                    'Play zap
                    udcLightZapSound.PlaySound(gintSoundVolume)
                    'Set
                    blnLightZap2 = True
                End If

            Case Else
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath1Memories(0)
                Else
                    'Set
                    intPathTextureIndex = 0
                End If
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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath2Memories(1)
                Else
                    'Set
                    intPathTextureIndex = 4
                End If
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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath2Memories(2)
                Else
                    'Set
                    intPathTextureIndex = 5
                End If
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
                'Check if GDI or OpenGL
                If Not blnOpenGL Then
                    'Set
                    btmPath = abtmPath2Memories(0)
                Else
                    'Set
                    intPathTextureIndex = 3
                End If
                'Reset
                blnLightZap1 = False
                blnLightZap2 = False
                'Reset mouse over variables
                ResetMouseOverVariables()

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

    Private Function strOpenGLError() As String

        'Get the error
        Select Case glGetError()
            Case 0
                Return "GL_NO_ERROR"
            Case intGL_INVALID_ENUM
                Return "GL_INVALID_ENUM"
            Case intGL_INVALID_VALUE
                Return "GL_INVALID_VALUE"
            Case intGL_INVALID_OPERATION
                Return "GL_INVALID_OPERATION"
            Case intGL_STACK_OVERFLOW
                Return "GL_STACK_OVERFLOW"
            Case intGL_STACK_UNDERFLOW
                Return "GL_STACK_UNDERFLOW"
            Case intGL_OUT_OF_MEMORY
                Return "GL_OUT_OF_MEMORY"
            Case intGL_INVALID_FRAMEBUFFER_OPERATION
                Return "GL_INVALID_FRAMEBUFFER_OPERATION"
            Case intGL_CONTEXT_LOST
                Return "GL_CONTEXT_LOST"
            Case intGL_TABLE_TOO_LARGE
                Return "GL_TABLE_TOO_LARGE"
            Case Else
                Return "GL_UNKNOWN_ERROR"
        End Select

    End Function

End Class