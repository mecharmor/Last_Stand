﻿'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Constants
    Private Const GAMEVERSION As String = "1.08"
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
    Private Const LOADINGTRANSPARENCYDELAY As Integer = 400
    Private Const ENDGAMEDELAYBLACKSCREEN As Integer = 500

    'Declare beginning necessary engine needs
    Private intScreenWidth As Integer = 800
    Private intScreenHeight As Integer = 600
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False
    Private rectFullScreen As Rectangle
    Private gGraphics As Graphics
    Private btmCanvas As New Bitmap(ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT, System.Drawing.Imaging.PixelFormat.Format32bppPArgb) 'Original resolution of our game
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
    Private pntLoadingParagraph As New Point(123, 261)
    Private thrParagraph As System.Threading.Thread
    Private thrLoading As System.Threading.Thread

    'Game screen
    Private btmGameBackground As Bitmap
    Private pntGameBackground As New Point(0, 0)
    Private btmEndGame25 As Bitmap
    Private btmEndGame50 As Bitmap
    Private btmEndGame75 As Bitmap
    Private btmEndGame100 As Bitmap
    Private btmEndGame As Bitmap
    Private thrEndGame As System.Threading.Thread
    Private btmWordBar As Bitmap
    Private pntWordBar As New Point(482, 27)
    Private udcCharacter As clsCharacter
    Private audcZombies(4) As clsZombie
    Private intZombieSpeed As Integer = 15 ' Default
    Private blnBackFromGame As Boolean = False
    Private blnGameOverFirstTime As Boolean = False
    Private blnEndingGameCantType As Boolean = False
    Private btmAK47Magazine As Bitmap
    Private pntAK47Magazine As New Point(59, 877)
    Private btmDeathOverlay As Bitmap
    Private udcScreamSound As clsSound
    Private intZombieKills As Integer = 0
    Private intCanvasMovementZombieDeathAdjustment As Integer = 0

    'Stop watch
    Private swhStopWatch As Stopwatch
    Private tsTimeSpan As TimeSpan
    Private strElapsedTime As String

    'Words
    Private astrWords(0) As String 'Used to fill with words
    Private intWordIndex As Integer = 0
    Private strTheWord As String = ""
    Private strWord As String = ""

    'Highscores screen
    Private btmHighscoresBackground As New Bitmap(Image.FromFile(strDirectory & "Images\Highscores\HighscoresBackground.jpg"))
    Private strHighscores As String = ""

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
    Private intZombieSpeedOne As Integer = 0
    Private intZombieSpeedTwo As Integer = 0
    Private audcZombiesOne(4) As clsZombie
    Private audcZombiesTwo(4) As clsZombie
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

        'Set for loading
        btmLoadingBar = btmLoadingBar0

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Fog
        AbortThread(thrFog)

        'Rendering
        AbortThread(thrRendering)

        'Sound delay
        AbortThread(thrHoverSoundDelay)

        'Loading
        AbortThread(thrLoading)

        'Paragraphing
        AbortThread(thrParagraph)

        'Ending game
        AbortThread(thrEndGame)

        'Credits
        AbortThread(thrCredits)

        'Story
        AbortThread(thrStory)

        'Remove game objects
        RemoveGameObjectsFromMemory()

        'Empty versus multiplayer variables
        EmptyMultiplayerVariables()

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

        'Paragraph versus
        AbortThread(thrParagraphVersus)

        'Listening TCP/IP
        AbortThread(thrListening)

        'Connecting TCP/IP
        AbortThread(thrConnecting)

        'Mismatch
        AbortThread(thrGameMismatch)

        'Empty versus multiplayer variables
        If gswClientData IsNot Nothing Then
            gswClientData.Close()
        End If
        If srClientData IsNot Nothing Then
            srClientData.Close()
        End If
        If tcplServer IsNot Nothing Then
            tcplServer.Stop()
        End If
        If tcpcClient IsNot Nothing Then
            tcpcClient.Close()
        End If

        'Make this last because of try and end try can create glitch problems
        If udcVersusConnectedThread IsNot Nothing Then
            'Remove handlers
            RemoveHandler udcVersusConnectedThread.GotData, AddressOf DataArrival
            RemoveHandler udcVersusConnectedThread.ConnectionGone, AddressOf ConnectionLost
            'Abort threads
            udcVersusConnectedThread.AbortCheckConnectionThread()
            udcVersusConnectedThread.AbortReadLineThread()
        End If

    End Sub

    Private Sub RemoveGameObjectsFromMemory()

        'Stop and dispose character, remove handler
        If udcCharacter IsNot Nothing Then
            'Stop reloading sound
            udcCharacter.StopReloadingSound()
            'Stop and dispose
            udcCharacter.StopAndDispose()
            udcCharacter = Nothing
        End If

        'Stop and dispose zombies
        For intLoop As Integer = 0 To audcZombies.GetUpperBound(0)
            If audcZombies(intLoop) IsNot Nothing Then
                audcZombies(intLoop).StopAndDispose()
                audcZombies(intLoop) = Nothing
            End If
        Next

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
        For intLoop As Integer = 0 To audcZombiesOne.GetUpperBound(0)
            If audcZombiesOne(intLoop) IsNot Nothing Then
                audcZombiesOne(intLoop).StopAndDispose()
                audcZombiesOne(intLoop) = Nothing
            End If
        Next
        For intLoop As Integer = 0 To audcZombiesTwo.GetUpperBound(0)
            If audcZombiesTwo(intLoop) IsNot Nothing Then
                audcZombiesTwo(intLoop).StopAndDispose()
                audcZombiesTwo(intLoop) = Nothing
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
                'Set
                blnBackFromGame = False
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
                'Stop and dispose game objects
                RemoveGameObjectsFromMemory()
                'Empty versus multiplayer variables
                EmptyMultiplayerVariables()
                'Wait or else a complication will start when reconnecting after an exit of versus game
                System.Threading.Thread.Sleep(1)
                'Set
                btmLoadingParagraph = Nothing
                'Set
                btmLoadingBar = btmLoadingBar0
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(0, 0, blnPlayPressedSoundNow)
                'Set
                blnPlayPressedSoundNow = False
                'Restart fog
                RestartFog()
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
        DrawText(gGraphics, "Version " & GAMEVERSION, 50, Color.Black, 33, 953, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Version " & GAMEVERSION, 50, Color.Black, 35, 955, 1000, 125) 'Black shadow
        DrawText(gGraphics, "Version " & GAMEVERSION, 50, Color.Red, 30, 950, 1000, 125) 'Overlay

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

    Private Sub StartedGameScreen()

        'Check for game over death overlay screen
        If btmEndGame Is btmEndGame100 Then

            'Draw end game screen
            DrawEndGameScreen(intZombieKills, udcCharacter)

        Else

            'Move graphics to copy and print dead first
            gGraphics = Graphics.FromImage(btmGameBackground)

            'Set graphic options
            SetDefaultGraphicOptions()

            'Draw dead zombies permanently
            For intLoop As Integer = 0 To audcZombies.GetUpperBound(0)
                If audcZombies(intLoop).IsDead Then
                    Dim pntTemp As New Point(audcZombies(intLoop).ZombiePoint.X + CInt(Math.Abs(pntGameBackground.X)), audcZombies(intLoop).ZombiePoint.Y)
                    'Draw dead
                    gGraphics.DrawImageUnscaled(audcZombies(intLoop).ZombieImage, pntTemp)
                    'Increase count
                    intZombieKills += 1
                    'Add zombie
                    UseZombieWave(intZombieKills, audcZombies, intLoop, ORIGINALSCREENWIDTH, 325, intZombieSpeed, "udcCharacter")
                End If
            Next

            'Paint background and word
            PaintBackgroundAndWord()

            'Draw character
            gGraphics.DrawImageUnscaled(udcCharacter.CharacterImage, udcCharacter.CharacterPoint)

            'Draw zombies
            For intLoop As Integer = 0 To audcZombies.GetUpperBound(0)
                'Check distance
                If audcZombies(intLoop).ZombiePoint.X <= 200 And Not audcZombies(intLoop).IsPinning Then
                    'Set
                    audcZombies(intLoop).Pin()
                    'Check for first time pin
                    CheckingForFirstTimePin(False)
                End If
                'Draw zombies dying, pinning or walking
                If Not audcZombies(intLoop).IsDead Then
                    If udcCharacter.IsRunning Then
                        'Change
                        For intLoopRunning As Integer = 0 To audcZombies.GetUpperBound(0)
                            audcZombies(intLoopRunning).ZombiePoint = New Point(audcZombies(intLoopRunning).ZombiePoint.X - 5, audcZombies(intLoopRunning).ZombiePoint.Y)
                        Next
                    End If
                    gGraphics.DrawImageUnscaled(audcZombies(intLoop).ZombieImage, audcZombies(intLoop).ZombiePoint)
                End If
            Next

            'Show magazine with bullet count
            gGraphics.DrawImageUnscaled(btmAK47Magazine, pntAK47Magazine)

            'Draw bullet count on magazine
            DrawText(gGraphics, CStr(30 - udcCharacter.BulletsUsed), 40, Color.Red, pntAK47Magazine.X - 15, pntAK47Magazine.Y + 50, 100, 75)

            'If game over
            If btmEndGame IsNot Nothing Then
                'Fade to black
                gGraphics.DrawImageUnscaled(btmEndGame, pntTopLeft)
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

    Private Sub CheckingForFirstTimePin(blnEndGame As Boolean, Optional strData As String = "")

        'Don't die
        Exit Sub

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
        PaintOnCanvas()

        'Check for running
        If Not blnGameIsVersus Then
            If udcCharacter.IsRunning Then
                'Change
                pntGameBackground.X -= 20
            End If
        End If

        'Draw the background, the armory area
        gGraphics.DrawImageUnscaled(btmGameBackground, pntGameBackground)

        'Draw word bar
        gGraphics.DrawImageUnscaled(btmWordBar, pntWordBar)

        'Draw text in the word bar
        DrawText(gGraphics, strTheWord, 50, Color.Black, 505, 95, 1000, 100) 'Shadow
        DrawText(gGraphics, strTheWord, 50, Color.Black, 503, 93, 1000, 100) 'Shadow
        DrawText(gGraphics, strTheWord, 50, Color.White, 500, 90, 1000, 100)
        If blnGameIsVersus Then
            If blnHost Then
                DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Red, 500, 90, 1000, 100) 'Overlay
            Else
                DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Blue, 500, 90, 1000, 100) 'Overlay
            End If
        Else
            DrawText(gGraphics, strTheWord.Substring(0, intWordIndex), 50, Color.Red, 500, 90, 1000, 100) 'Overlay
        End If

    End Sub

    Private Sub DrawEndGameScreen(intZombieKillsType As Integer, udcCharacterType As clsCharacter)

        'Paint on canvas
        PaintOnCanvas()

        'Screen is black, show death overlay
        gGraphics.DrawImageUnscaled(btmDeathOverlay, pntTopLeft)

        'Draw zombie kills
        DrawText(gGraphics, CStr(intZombieKillsType), 85, Color.Black, 915, 437, 1000, 125)
        DrawText(gGraphics, CStr(intZombieKillsType), 85, Color.White, 910, 432, 1000, 125)

        'Draw survival time
        DrawText(gGraphics, strElapsedTime, 85, Color.Black, 925, 605, 1000, 125)
        DrawText(gGraphics, strElapsedTime, 85, Color.White, 920, 600, 1000, 125)

        'Declare
        Dim intReloadTimeUsed As Integer = udcCharacterType.ReloadTimes * 3
        Dim intElapsedTime As Integer = CInt(Replace(strElapsedTime, " Sec", "")) - intReloadTimeUsed
        Dim intWPM As Integer = CInt((intZombieKillsType / (intElapsedTime / 60)))

        'Draw WPM
        DrawText(gGraphics, CStr(intWPM), 85, Color.Black, 925, 763, 1000, 125)
        DrawText(gGraphics, CStr(intWPM), 85, Color.White, 920, 758, 1000, 125)

        'Draw who won if versus game
        If blnGameIsVersus Then
            If btmVersusWhoWon IsNot Nothing Then
                If btmVersusWhoWon Is btmYouWon Then
                    gGraphics.DrawImageUnscaled(btmVersusWhoWon, pntYouWon)
                Else
                    gGraphics.DrawImageUnscaled(btmVersusWhoWon, pntYouLost)
                End If
            End If
        End If

    End Sub

    Private Sub UseZombieWave(intZombieKillsType As Integer, audcZombiesType() As clsZombie, intIndex As Integer, intX As Integer, intY As Integer,
                              intWalkingSpeed As Integer, strThisObjectNameCorrespondingToCharacter As String)

        'Check number of zombies killed
        Select Case strThisObjectNameCorrespondingToCharacter
            Case "udcCharacter"
                Select Case intZombieKillsType
                    Case 1 To 4, 6 To 9, 11 To 14, 16 To 19, 21 To 24, 26 To 29
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed, strThisObjectNameCorrespondingToCharacter)
                    Case 5, 10, 15, 20, 25, 30
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                    Case 31 To 33, 36 To 38, 41 To 43, 46 To 48, 51 To 53, 56 To 58
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 5, strThisObjectNameCorrespondingToCharacter)
                    Case 34, 35, 39, 40, 44, 45, 49, 50
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                    Case Else
                        If intZombieKills <> 50 Then
                            'Re-create a new zombie
                            ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                        End If
                End Select
            Case Else
                Select Case intZombieKillsType
                    Case 1 To 4, 6 To 9, 11 To 14, 16 To 19, 21 To 24, 26 To 29
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed, strThisObjectNameCorrespondingToCharacter)
                    Case 5, 10, 15, 20, 25, 30
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                    Case 31 To 33, 36 To 38, 41 To 43, 46 To 48, 51 To 53, 56 To 58
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 5, strThisObjectNameCorrespondingToCharacter)
                    Case 34, 35, 39, 40, 44, 45, 49, 50, 54, 55, 59, 60
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                    Case 61, 62, 66, 67, 71, 72, 76, 77, 81, 82, 86, 87
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 10, strThisObjectNameCorrespondingToCharacter)
                    Case 63 To 65, 68 To 70, 73 To 75, 78 To 80, 83 To 85, 88 To 90
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                    Case Else
                        'Re-create a new zombie
                        ReCreateZombie(audcZombiesType, intIndex, intX, intY, intWalkingSpeed + 45, strThisObjectNameCorrespondingToCharacter)
                End Select
        End Select

    End Sub

    Private Sub ReCreateZombie(audcZombiesType() As clsZombie, intIndex As Integer, intX As Integer, intY As Integer, intWalkingSpeed As Integer, strThisObjectNameCorrespondingToCharacter As String)

        'Create a new zombie
        audcZombiesType(intIndex) = New clsZombie(Me, intX, intY, intWalkingSpeed, strThisObjectNameCorrespondingToCharacter)
        audcZombiesType(intIndex).Start()

    End Sub

    Private Sub EndingGame(Optional blnWon As Boolean = False)

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

        'Stop the watch
        StopTheWatch()

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

    Private Sub StopTheWatch()

        'Stop the watch
        swhStopWatch.Stop()

        'Set
        tsTimeSpan = swhStopWatch.Elapsed
        strElapsedTime = CStr(CInt(tsTimeSpan.TotalSeconds)) & " Sec"

    End Sub

    Private Sub HighscoresScreen()

        'Paint on canvas
        PaintOnCanvas()

        'Draw highscores background
        gGraphics.DrawImageUnscaled(btmHighscoresBackground, pntTopLeft)

        'Draw highscores
        DrawText(gGraphics, strHighscores, 34, Color.Black, 32, 222, ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 33, 223, ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 34, 224, ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Black, 35, 225, ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT) 'Shadow
        DrawText(gGraphics, strHighscores, 34, Color.Red, 30, 220, ORIGINALSCREENWIDTH, ORIGINALSCREENHEIGHT)

        'Back button
        If intCanvasShow = 1 Then
            'Draw back text as hovered
            gGraphics.DrawImageUnscaled(btmBackHoverText, pntBackHoverText)
        Else
            'Draw back text
            gGraphics.DrawImageUnscaled(btmBackText, pntBackText)
        End If

    End Sub

    Private Sub LoadHighscoresString()

        'Reset
        strHighscores = ""

        'Declare
        Dim ioSR As IO.StreamReader

        'Set
        ioSR = IO.File.OpenText(strDirectory & "Images\Highscores\Highscores.txt")

        'Use loop to get scores
        While ioSR.Peek <> -1
            strHighscores &= ioSR.ReadLine & vbNewLine
        End While

        'Close
        ioSR.Close()

    End Sub

    Private Sub CreditsScreen()

        'Paint on canvas
        PaintOnCanvas()

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
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales25
        btmZacharyStafford = btmZacharyStafford25
        btmCoryLewis = btmCoryLewis25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales50
        btmZacharyStafford = btmZacharyStafford50
        btmCoryLewis = btmCoryLewis50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales75
        btmZacharyStafford = btmZacharyStafford75
        btmCoryLewis = btmCoryLewis75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmJohnGonzales = btmJohnGonzales100
        btmZacharyStafford = btmZacharyStafford100
        btmCoryLewis = btmCoryLewis100

    End Sub

    Private Sub VersusScreen()

        'Paint on canvas
        PaintOnCanvas()

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
        PaintOnCanvas()

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
        If btmEndGame Is btmEndGame100 Then

            'Draw end game screen
            If blnHost Then
                DrawEndGameScreen(intZombieKillsOne, udcCharacterOne)
            Else
                DrawEndGameScreen(intZombieKillsTwo, udcCharacterTwo)
            End If

        Else

            'Move graphics to copy and print dead first
            gGraphics = Graphics.FromImage(btmGameBackground)

            'Set graphic options
            SetDefaultGraphicOptions()

            'Draw dead zombies permanently
            For intLoop As Integer = 0 To audcZombiesOne.GetUpperBound(0)
                If audcZombiesOne(intLoop).IsDead Then
                    'Check if hosting
                    If blnHost Then
                        'Send data
                        gSendData("6|" & CStr(intLoop)) 'This marks the zombie to be dead for the joiner
                    End If
                    'Draw dead
                    gGraphics.DrawImageUnscaled(audcZombiesOne(intLoop).ZombieImage, audcZombiesOne(intLoop).ZombiePoint)
                    'Increase
                    intZombieKillsOne += 1
                    'Add zombie
                    UseZombieWave(intZombieKillsOne, audcZombiesOne, intLoop, ORIGINALSCREENWIDTH, 300, intZombieSpeedOne, "udcCharacterOne")
                End If
            Next
            For intLoop As Integer = 0 To audcZombiesTwo.GetUpperBound(0)
                If audcZombiesTwo(intLoop).IsDead Then
                    'Check if hosting
                    If blnHost Then
                        'Send data
                        gSendData("7|" & CStr(intLoop)) 'This marks the zombie to be dead for the joiner
                    End If
                    'Draw dead
                    gGraphics.DrawImageUnscaled(audcZombiesTwo(intLoop).ZombieImage, audcZombiesTwo(intLoop).ZombiePoint)
                    'Increase
                    intZombieKillsTwo += 1
                    'Add zombie
                    UseZombieWave(intZombieKillsTwo, audcZombiesTwo, intLoop, ORIGINALSCREENWIDTH, 350, intZombieSpeedTwo, "udcCharacterTwo")
                End If
            Next

            'Paint background and word
            PaintBackgroundAndWord()

            'Draw character hoster
            gGraphics.DrawImageUnscaled(udcCharacterOne.CharacterImage, udcCharacterOne.CharacterPoint)

            'Draw zombies
            If blnHost Then
                'First horde of zombies for hoster
                For intLoop As Integer = 0 To audcZombiesOne.GetUpperBound(0)
                    'Check distance
                    If audcZombiesOne(intLoop).ZombiePoint.X <= 200 And Not audcZombiesOne(intLoop).IsPinning Then
                        'Set
                        audcZombiesOne(intLoop).Pin()
                        'Check for first time pin
                        CheckingForFirstTimePin(False, "8|") 'Joiner won
                    End If
                    'Draw zombies dying, pinning or walking
                    If Not audcZombiesOne(intLoop).IsDead Then
                        gGraphics.DrawImageUnscaled(audcZombiesOne(intLoop).ZombieImage, audcZombiesOne(intLoop).ZombiePoint)
                    End If
                Next
                'Draw character joiner
                gGraphics.DrawImageUnscaled(udcCharacterTwo.CharacterImage, udcCharacterTwo.CharacterPoint)
                'Second horde of zombies for joiner
                For intLoop As Integer = 0 To audcZombiesTwo.GetUpperBound(0)
                    'Check distance
                    If audcZombiesTwo(intLoop).ZombiePoint.X <= 250 And Not audcZombiesTwo(intLoop).IsPinning Then
                        'Set
                        audcZombiesTwo(intLoop).Pin()
                        'Check for first time pin
                        CheckingForFirstTimePin(True, "9|") 'Joiner lost
                    End If
                    'Draw zombies dying, pinning or walking
                    If Not audcZombiesTwo(intLoop).IsDead Then
                        gGraphics.DrawImageUnscaled(audcZombiesTwo(intLoop).ZombieImage, audcZombiesTwo(intLoop).ZombiePoint)
                    End If
                Next
            Else 'What the joiner sees
                'First horde of zombies for hoster
                For intLoop As Integer = 0 To audcZombiesOne.GetUpperBound(0)
                    'Check distance
                    If audcZombiesOne(intLoop).ZombiePoint.X <= 200 And Not audcZombiesOne(intLoop).IsPinning Then
                        'Set
                        audcZombiesOne(intLoop).Pin()
                    End If
                    'Draw zombies dying, pinning or walking
                    If Not audcZombiesOne(intLoop).IsDead Then
                        gGraphics.DrawImageUnscaled(audcZombiesOne(intLoop).ZombieImage, audcZombiesOne(intLoop).ZombiePoint)
                    End If
                Next
                'Draw character joiner
                gGraphics.DrawImageUnscaled(udcCharacterTwo.CharacterImage, udcCharacterTwo.CharacterPoint)
                'Second horde of zombies for joiner
                For intLoop As Integer = 0 To audcZombiesTwo.GetUpperBound(0)
                    'Check distance
                    If audcZombiesTwo(intLoop).ZombiePoint.X <= 250 And Not audcZombiesTwo(intLoop).IsPinning Then
                        'Set
                        audcZombiesTwo(intLoop).Pin()
                    End If
                    'Draw zombies dying, pinning or walking
                    If Not audcZombiesTwo(intLoop).IsDead Then
                        gGraphics.DrawImageUnscaled(audcZombiesTwo(intLoop).ZombieImage, audcZombiesTwo(intLoop).ZombiePoint)
                    End If
                Next
            End If

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
            If btmEndGame IsNot Nothing Then
                'Fade to black
                gGraphics.DrawImageUnscaled(btmEndGame, pntTopLeft)
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
        PaintOnCanvas()

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
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        pntStoryParagraph = New Point(83, 229)

        'Set
        btmStoryParagraph = btmStoryParagraph1_25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph1_50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph1_75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

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
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        pntStoryParagraph = New Point(38, 168)

        'Set
        btmStoryParagraph = btmStoryParagraph2_25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph2_50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph2_75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

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
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        pntStoryParagraph = New Point(31, 180)

        'Set
        btmStoryParagraph = btmStoryParagraph3_25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmStoryParagraph = btmStoryParagraph3_100

        'Sleep
        System.Threading.Thread.Sleep(1000)

        'Play story paragraph 3 sound
        udcStoryParagraph2Sound = New clsSound(Me, strDirectory & "Sounds\StoryParagraph3.mp3", 86000, gintSoundVolume)

    End Sub

    Private Sub GameVersionMismatch()

        'Paint on canvas
        PaintOnCanvas()

        'Draw story background
        gGraphics.DrawImageUnscaled(btmGameMismatchBackground, pntTopLeft)

        'Display
        MismatchingText("Your version: " & GAMEVERSION, 878, 197)

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
                Dim udcButtonHoverSound As New clsSound(Me, strDirectory & "Sounds\ButtonHover.mp3", 3000, gintSoundVolume)
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

    End Sub

    Private Sub HoverSoundWaiting()

        'Sleep for 3000 ms
        System.Threading.Thread.Sleep(3000)

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
            Dim udcButtonPressedSound As New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 3000, gintSoundVolume, False)
        End If

    End Sub

    Private Sub Paragraphing()

        'Set
        btmLoadingParagraph = Nothing

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraph = btmLoadingParagraph100

    End Sub

    Private Sub LoadingGame()

        'Notes: This procedure is for loading the game to play

        'Set
        btmLoadingBar = btmLoadingBar0

        'Set
        pntGameBackground.X = 0

        'Set
        intCanvasMovementZombieDeathAdjustment = 0

        'Load game objects
        LoadGameObjects(10)

        'Set
        btmLoadingBar = btmLoadingBar10

        'Set
        intZombieKills = 0

        'Set to be fresh every game
        btmGameBackground = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground.jpg"))

        'Set
        btmEndGame = Nothing

        'Set
        blnEndingGameCantType = False

        'Set
        blnGameOverFirstTime = True

        'Load game objects
        LoadGameObjects(20)

        'Set
        btmLoadingBar = btmLoadingBar20

        'Set zombie speed
        intZombieSpeed = 10

        'Load game objects
        LoadGameObjects(30)

        'Set
        btmLoadingBar = btmLoadingBar30

        'Load game objects
        LoadGameObjects(40)

        'Set
        btmLoadingBar = btmLoadingBar40

        'Load game objects
        LoadGameObjects(50)

        'Set
        btmLoadingBar = btmLoadingBar50

        'Load game objects
        LoadGameObjects(60)

        'Set
        btmLoadingBar = btmLoadingBar60

        'Load game objects
        LoadGameObjects(70)

        'Set
        btmLoadingBar = btmLoadingBar70

        'Load game objects
        LoadGameObjects(80)

        'Set
        btmLoadingBar = btmLoadingBar80

        'Load game objects
        LoadGameObjects(90)

        'Set
        btmLoadingBar = btmLoadingBar90

        'Load game objects
        LoadGameObjects(99)

        'Set
        btmLoadingBar = btmLoadingBar99

        'Grab a random word
        LoadARandomWord()

        'Load game objects
        LoadGameObjects(100)

        'Character
        udcCharacter = New clsCharacter(Me, 100, 325, "udcCharacter")

        'Zombies
        audcZombies(0) = New clsZombie(Me, ORIGINALSCREENWIDTH, 325, intZombieSpeed, "udcCharacter")
        audcZombies(1) = New clsZombie(Me, ORIGINALSCREENWIDTH + 500, 325, intZombieSpeed, "udcCharacter")
        audcZombies(2) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1000, 325, intZombieSpeed, "udcCharacter")
        audcZombies(3) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1500, 325, intZombieSpeed, "udcCharacter")
        audcZombies(4) = New clsZombie(Me, ORIGINALSCREENWIDTH + 2000, 325, intZombieSpeed, "udcCharacter")

        'Set
        btmLoadingBar = btmLoadingBar100

        'Set
        intCanvasShow = 1 'Means completely loaded

    End Sub

    Private Sub LoadPieceOfBitmaps(ByRef abtmByRefBitmap() As Bitmap, strMainFolder As String, strSubFolder As String)

        'Load
        For intLoop As Integer = 0 To abtmByRefBitmap.GetUpperBound(0)
            abtmByRefBitmap(intLoop) = New Bitmap(Image.FromFile(strDirectory & "Images\" & strMainFolder & "\" & strSubFolder & "\" & CStr(intLoop + 1) & ".png"))
        Next

    End Sub

    Private Sub LoadGameObjects(intLoadPercentage As Integer)

        'Declare
        Static sblnFirstTimeLoadingGameObjects As Boolean = True

        'Check
        If sblnFirstTimeLoadingGameObjects Then
            'Load by certain percentage
            Select Case intLoadPercentage
                Case 10
                    'Words
                    AddWordsToArray()
                    'Word bar
                    btmWordBar = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\WordBar.png"))
                Case 20
                    'Character
                    LoadPieceOfBitmaps(gbtmCharacterStand, "Character\Generic", "Standing")
                    LoadPieceOfBitmaps(gbtmCharacterShoot, "Character\Generic", "Shoot Once")
                    LoadPieceOfBitmaps(gbtmCharacterReload, "Character\Generic", "Reload")
                    LoadPieceOfBitmaps(gbtmCharacterRunning, "Character\Generic", "Running")
                Case 30
                    'Character red
                    LoadPieceOfBitmaps(gbtmCharacterStandRed, "Character\Red", "Standing")
                    LoadPieceOfBitmaps(gbtmCharacterShootRed, "Character\Red", "Shoot Once")
                    LoadPieceOfBitmaps(gbtmCharacterReloadRed, "Character\Red", "Reload")
                Case 40
                    'Character blue
                    LoadPieceOfBitmaps(gbtmCharacterStandBlue, "Character\Blue", "Standing")
                    LoadPieceOfBitmaps(gbtmCharacterShootBlue, "Character\Blue", "Shoot Once")
                    LoadPieceOfBitmaps(gbtmCharacterReloadBlue, "Character\Blue", "Reload")
                Case 50
                    'Zombies
                    LoadPieceOfBitmaps(gbtmZombieWalk, "Zombies\Generic", "Movement")
                    LoadPieceOfBitmaps(gbtmZombieDeath1, "Zombies\Generic", "Death 1")
                    LoadPieceOfBitmaps(gbtmZombieDeath2, "Zombies\Generic", "Death 2")
                    LoadPieceOfBitmaps(gbtmZombiePin, "Zombies\Generic", "Pinning")
                Case 60
                    'Zombies red
                    LoadPieceOfBitmaps(gbtmZombieWalkRed, "Zombies\Red", "Movement")
                    LoadPieceOfBitmaps(gbtmZombieDeath1Red, "Zombies\Red", "Death 1")
                    LoadPieceOfBitmaps(gbtmZombieDeath2Red, "Zombies\Red", "Death 2")
                    LoadPieceOfBitmaps(gbtmZombiePinRed, "Zombies\Red", "Pinning")
                Case 70
                    'Zombies blue
                    LoadPieceOfBitmaps(gbtmZombieWalkBlue, "Zombies\Blue", "Movement")
                    LoadPieceOfBitmaps(gbtmZombieDeath1Blue, "Zombies\Blue", "Death 1")
                    LoadPieceOfBitmaps(gbtmZombieDeath2Blue, "Zombies\Blue", "Death 2")
                    LoadPieceOfBitmaps(gbtmZombiePinBlue, "Zombies\Blue", "Pinning")
                Case 80
                    'Magazine
                    btmAK47Magazine = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\AK47Magazine.png"))
                Case 90
                    'Death overlay
                    btmDeathOverlay = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\DeathOverlay.png"))
                    'Versus win or lose
                    btmYouWon = New Bitmap(Image.FromFile(strDirectory & "Images\Versus\YouWon.png"))
                    btmYouLost = New Bitmap(Image.FromFile(strDirectory & "Images\Versus\YouLost.png"))
                Case 99
                    'End game
                    btmEndGame25 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame25.png"))
                    btmEndGame50 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame50.png"))
                    btmEndGame75 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame75.png"))
                    btmEndGame100 = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\EndGame100.png"))
                Case 100
                    'Set
                    sblnFirstTimeLoadingGameObjects = False
            End Select
        End If

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
                'Stop fog thread
                AbortThread(thrFog)
                'Set
                ChangeCanvasModeAndChangeCanvasShowAndPlayZombieSound(2, 0)
                'Stop sound
                If udcAmbianceSound IsNot Nothing Then
                    udcAmbianceSound.StopAndCloseSound()
                End If
                'Start loading game
                thrLoading = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGame))
                thrLoading.Start()
                'Start paragraphing
                thrParagraph = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Paragraphing))
                thrParagraph.Start()
                'Reset fog
                ResetFog()
                'Exit
                Exit Sub
            End If

            'Highscores was clicked
            If blnMouseInRegion(pntMouse, 413, 99, pntHighscoresText) Then
                'Change
                ShowNextScreenAndExitMenu(4, 0)
                'Reset fog
                ResetFog()
                'Load highscores into string
                LoadHighscoresString()
                'Exit
                Exit Sub
            End If

            'Story was clicked
            If blnMouseInRegion(pntMouse, 218, 87, pntStoryText) Then
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
                'Change
                ShowNextScreenAndExitMenu(1, 0)
                'Reset fog
                ResetFog()
                'Exit
                Exit Sub
            End If

            'Credits was clicked
            If blnMouseInRegion(pntMouse, 285, 78, pntCreditsText) Then
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
                'Wait until
                While blnBackFromGame
                    System.Threading.Thread.Sleep(1)
                End While
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
                    'Set
                    intCanvasMode = 3
                    'Set
                    intCanvasShow = 0
                    'Start character
                    udcCharacter.Start()
                    'Start zombies
                    For intLoop As Integer = 0 To audcZombies.GetUpperBound(0)
                        audcZombies(intLoop).Start()
                    Next
                    'Start stop watch
                    swhStopWatch = New Stopwatch
                    swhStopWatch.Start()
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
                'Stop screaming sound
                If udcScreamSound IsNot Nothing Then
                    udcScreamSound.StopAndCloseSound()
                End If
                'Set
                blnPlayPressedSoundNow = True
                'Set
                blnBackFromGame = True
                'Menu sound, play last to make sure other stuff sets, was having a problem if not in some cases
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, gintSoundVolume, True) '38 seconds + extra
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

    End Sub

    Private Sub StartVersusGameObjects()

        'Start characters
        udcCharacterOne.Start()
        udcCharacterTwo.Start()

        'Start zombies
        For intLoop As Integer = 0 To audcZombiesOne.GetUpperBound(0)
            audcZombiesOne(intLoop).Start()
        Next
        For intLoop As Integer = 0 To audcZombiesTwo.GetUpperBound(0)
            audcZombiesTwo(intLoop).Start()
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
        intFogPass1X = -(ORIGINALSCREENWIDTH * 2)
        intFogPass2X = -((ORIGINALSCREENWIDTH * 2) * 2)
        pntFogBackPass1.X = -(ORIGINALSCREENWIDTH * 2)
        pntFogFrontPass1.X = -(ORIGINALSCREENWIDTH * 2)
        pntFogBackPass2.X = -((ORIGINALSCREENWIDTH * 2) * 2)
        pntFogFrontPass2.X = -((ORIGINALSCREENWIDTH * 2) * 2)

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

        'Keys pressed
        'Debug.Print(CStr(Asc(e.KeyChar)))

        'Check if playing game
        If intCanvasMode = 3 Then

            'Exit if reloading, or game over
            If udcCharacter.IsReloading Or blnEndingGameCantType Or udcCharacter.BulletsUsed = 30 Then
                Exit Sub
            End If

            'Check for spacebar and count bullets
            If CheckForSpacebar(e, udcCharacter) Then
                Exit Sub
            End If

            'Check the word being typed
            CheckTheWordBeingTyped(e, udcCharacter, audcZombies)

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

            'Check if hoster or joiner
            If blnHost Then
                'Exit if reloading, or game over
                If udcCharacterOne.IsReloading Or blnEndingGameCantType Or udcCharacterOne.BulletsUsed = 30 Then
                    Exit Sub
                End If
                'Check for spacebar and count bullets
                If CheckForSpacebar(e, udcCharacterOne, True) Then
                    Exit Sub
                End If
                'Check the word being typed
                CheckTheWordBeingTyped(e, udcCharacterOne, audcZombiesOne, True)
            Else
                'Exit if reloading, or game over
                If udcCharacterTwo.IsReloading Or blnEndingGameCantType Or udcCharacterTwo.BulletsUsed = 30 Then
                    Exit Sub
                End If
                'Check for spacebar and count bullets
                If CheckForSpacebar(e, udcCharacterTwo, True) Then
                    Exit Sub
                End If
                'Check the word being typed
                CheckTheWordBeingTyped(e, udcCharacterTwo, audcZombiesTwo, True, True)
            End If

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

    Private Sub CheckTheWordBeingTyped(e As KeyPressEventArgs, udcCharacterType As clsCharacter, audcZombiesType() As clsZombie,
                                       Optional blnSendData As Boolean = False, Optional blnIgnoreKillingZombie As Boolean = False)

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
                'Kill closest zombie
                If Not blnIgnoreKillingZombie Then
                    'Declare
                    Dim intIndex As Integer = intGetIndexOfClosestZombie(audcZombiesType)
                    'Wait until a zombie can be killed
                    While True
                        If intIndex <> audcZombiesType.GetUpperBound(0) + 1 Then
                            'Die
                            audcZombiesType(intIndex).Dying()
                            'Send data
                            If blnSendData Then
                                gSendData("3|" & intIndex) 'send zombie kill
                            End If
                            'Exit
                            Exit While
                        Else
                            'Sleep
                            System.Threading.Thread.Sleep(5)
                            'Re-check index
                            intIndex = intGetIndexOfClosestZombie(audcZombiesType)
                        End If
                    End While
                    'Character shoots
                    udcCharacterType.Shot()
                Else
                    'Check if still needs to send data
                    If blnSendData Then
                        gSendData("3|") 'Joiner needs to shoot and kill zombie
                    End If
                End If
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
        Dim intIndex As Integer = audcZombie.GetUpperBound(0) + 1

        'Loop to get closest zombie
        For intLoop As Integer = 0 To audcZombie.GetUpperBound(0)
            If Not audcZombie(intLoop).IsDying And Not audcZombie(intLoop).IsDead Then
                If intClosestX > audcZombie(intLoop).ZombiePoint.X Then
                    intClosestX = audcZombie(intLoop).ZombiePoint.X
                    intIndex = intLoop
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
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus25

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus50

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus75

        'Sleep
        System.Threading.Thread.Sleep(LOADINGTRANSPARENCYDELAY)

        'Set
        btmLoadingParagraphVersus = btmLoadingParagraphVersus100

    End Sub

    Private Sub LoadingVersusGame()

        'Notes: This procedure is for loading the game to play versus

        'Set
        btmLoadingBar = btmLoadingBar0

        'Set
        pntGameBackground.X = 0

        'Set
        intCanvasMovementZombieDeathAdjustment = 0

        'Load game objects
        LoadGameObjects(10)

        'Set
        btmLoadingBar = btmLoadingBar10

        'Set
        blnGameIsVersus = True

        'Set
        intZombieKillsOne = 0
        intZombieKillsTwo = 0

        'Set to be fresh every game
        btmGameBackground = New Bitmap(Image.FromFile(strDirectory & "Images\Game Play\GameBackground.jpg"))

        'Set
        btmEndGame = Nothing

        'Set
        blnEndingGameCantType = False

        'Set
        blnGameOverFirstTime = True

        'Load game objects
        LoadGameObjects(20)

        'Set
        btmLoadingBar = btmLoadingBar20

        'Set zombie speed
        intZombieSpeedOne = 10
        intZombieSpeedTwo = 10

        'Load game objects
        LoadGameObjects(30)

        'Set
        btmLoadingBar = btmLoadingBar30

        'Load game objects
        LoadGameObjects(40)

        'Set
        btmLoadingBar = btmLoadingBar40

        'Load game objects
        LoadGameObjects(50)

        'Set
        btmLoadingBar = btmLoadingBar50

        'Load game objects
        LoadGameObjects(60)

        'Set
        btmLoadingBar = btmLoadingBar60

        'Load game objects
        LoadGameObjects(70)

        'Set
        btmLoadingBar = btmLoadingBar70

        'Load game objects
        LoadGameObjects(80)

        'Set
        btmLoadingBar = btmLoadingBar80

        'Load game objects
        LoadGameObjects(90)

        'Set
        btmLoadingBar = btmLoadingBar90

        'Load game objects
        LoadGameObjects(99)

        'Set
        btmLoadingBar = btmLoadingBar99

        'Grab a random word
        LoadARandomWord()

        'Load game objects
        LoadGameObjects(100)

        'Load in a special way
        If blnHost Then
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", False) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", True) 'Join
        Else
            'Character one
            udcCharacterOne = New clsCharacter(Me, 100, 300, "udcCharacterOne", True) 'Host
            'Character two
            udcCharacterTwo = New clsCharacter(Me, 200, 350, "udcCharacterTwo", False) 'Join
        End If

        'Zombies for host
        audcZombiesOne(0) = New clsZombie(Me, ORIGINALSCREENWIDTH, 300, intZombieSpeedOne, "udcCharacterOne")
        audcZombiesOne(1) = New clsZombie(Me, ORIGINALSCREENWIDTH + 500, 300, intZombieSpeedOne, "udcCharacterOne")
        audcZombiesOne(2) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1000, 300, intZombieSpeedOne, "udcCharacterOne")
        audcZombiesOne(3) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1500, 300, intZombieSpeedOne, "udcCharacterOne")
        audcZombiesOne(4) = New clsZombie(Me, ORIGINALSCREENWIDTH + 2000, 300, intZombieSpeedOne, "udcCharacterOne")

        'Zombies for joiner
        audcZombiesTwo(0) = New clsZombie(Me, ORIGINALSCREENWIDTH + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
        audcZombiesTwo(1) = New clsZombie(Me, ORIGINALSCREENWIDTH + 500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
        audcZombiesTwo(2) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
        audcZombiesTwo(3) = New clsZombie(Me, ORIGINALSCREENWIDTH + 1500 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")
        audcZombiesTwo(4) = New clsZombie(Me, ORIGINALSCREENWIDTH + 2000 + 100, 350, intZombieSpeedTwo, "udcCharacterTwo")

        'Set if hosting
        If blnHost And Not blnReadyEarly Then
            blnWaiting = True
        End If

        'Set
        btmLoadingBar = btmLoadingBar100

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
        If GAMEVERSION <> strGetBlockData(strData, 1) Then
            GamesMismatched(strGetBlockData(strData, 1))
        Else
            PrepareVersustToPlayAfterMismatchPass()
        End If

    End Sub

    Private Sub PrepareVersustToPlayAfterMismatchPass()

        'Set
        intCanvasMode = 7

        'Start loading game
        thrLoadingVersus = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingVersusGame))
        thrLoadingVersus.Start()

        'Start paragraphing
        thrParagraphVersus = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf ParagraphingVersus))
        thrParagraphVersus.Start()

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
                        gSendData("1|" & GAMEVERSION) 'Will exit if not matching
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
                    Dim intIndex As Integer = intGetIndexOfClosestZombie(audcZombiesTwo)
                    'Wait until a zombie can be killed
                    While True
                        If intIndex <> audcZombiesTwo.GetUpperBound(0) + 1 Then
                            'Send data
                            gSendData("4|" & intIndex)
                            'Die
                            audcZombiesTwo(intIndex).Dying()
                            'Exit
                            Exit While
                        Else
                            'Sleep
                            System.Threading.Thread.Sleep(5)
                            'Re-check index
                            intIndex = intGetIndexOfClosestZombie(audcZombiesTwo)
                        End If
                    End While
                    'Character shot
                    udcCharacterTwo.Shot()
                Case "5" 'Show join reloading
                    udcCharacterTwo.Reload()
            End Select

        Else 'Joiner

            'Check data
            Select Case (strGetBlockData(strData))
                Case "0" 'Send data back
                    gSendData("0|") 'Waiting for host to connect
                Case "1"
                    'Send data of game version
                    gSendData("1|" & GAMEVERSION) 'Will exit if not matching
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
                    'Kill zombie
                    audcZombiesOne(CInt(strGetBlockData(strData, 1))).Dying(False)
                    'Character shot
                    udcCharacterOne.Shot()
                Case "4" 'Kill zombie
                    'Kill
                    audcZombiesTwo(CInt(strGetBlockData(strData, 1))).Dying(False)
                    'Chatacter shot
                    udcCharacterTwo.Shot()
                Case "5" 'Show host reloading
                    'Reload
                    udcCharacterOne.Reload()
                Case "6" 'Host zombie needs to be prepared to die
                    audcZombiesOne(CInt(strGetBlockData(strData, 1))).IsDead = True
                Case "7" 'Join zombie needs to be prepared to die
                    audcZombiesTwo(CInt(strGetBlockData(strData, 1))).IsDead = True
                Case "8", "9" 'End game, hoster died
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
                    If strGetBlockData(strData) = "8" Then
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

    Private Sub frmGame_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        'Running Forward
        Select Case e.KeyCode
            Case Keys.Right
                If Not blnGameIsVersus Then
                    If Not udcCharacter.IsRunning Then
                        udcCharacter.Running()
                    End If
                End If
        End Select

    End Sub

End Class