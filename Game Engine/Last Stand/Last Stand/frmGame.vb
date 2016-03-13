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
    Private strDirectory As String = AppDomain.CurrentDomain.BaseDirectory
    Private blnScreenChanged As Boolean = False

    'Menu necessary needs
    Private btmMenuBackground As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\MenuBackground.jpg"))
    Private btmFogFrontPass1 As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\FogFront.png"))
    Private btmFogBackPass1 As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\FogBack.png"))
    Private intFogPass1X As Integer = -(ORIGINALSCREENWIDTH * 2)
    Private intFogFrontPass1Y As Integer = 650
    Private intFogBackPass1Y As Integer = 250
    Private pntFogFrontPass1 As New Point(intFogPass1X, intFogFrontPass1Y)
    Private pntFogBackPass1 As New Point(intFogPass1X, intFogBackPass1Y)
    Private btmFogFrontPass2 As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\FogFront.png"))
    Private btmFogBackPass2 As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\FogBack.png"))
    Private intFogPass2X As Integer = -((ORIGINALSCREENWIDTH * 2) * 2)
    Private intFogFrontPass2Y As Integer = 650
    Private intFogBackPass2Y As Integer = 250
    Private pntFogFrontPass2 As New Point(intFogPass2X, intFogFrontPass2Y)
    Private pntFogBackPass2 As New Point(intFogPass2X, intFogBackPass2Y)
    Private thrFog As System.Threading.Thread
    Private btmArcher As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Archer.png"))
    Private pntArcher As New Point(117, 0)
    Private btmLastStandText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\LastStand.png"))
    Private pntLastStandText As New Point(147, 833)

    'Menu buttons
    Private btmStartText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Start.png"))
    Private pntStartText As New Point(1081, 31)
    Private btmStartHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\StartHover.png"))
    Private pntStartHoverText As New Point(1059, 25)
    Private btmHighscoresText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Highscores.png"))
    Private pntHighscoresText As New Point(1198, 141)
    Private btmHighscoresHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\HighscoresHover.png"))
    Private pntHighscoresHoverText As New Point(1157, 131)
    Private btmStoryText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Story.png"))
    Private pntStoryText As New Point(1246, 283)
    Private btmStoryHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\StoryHover.png"))
    Private pntStoryHoverText As New Point(1222, 275)
    Private btmOptionsText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Options.png"))
    Private pntOptionsText As New Point(1207, 410)
    Private btmOptionsHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\OptionsHover.png"))
    Private pntOptionsHoverText As New Point(1175, 400)
    Private btmCreditsText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Credits.png"))
    Private pntCreditsText As New Point(1352, 536)
    Private btmCreditsHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\CreditsHover.png"))
    Private pntCreditsHoverText As New Point(1323, 527)
    Private btmVersusText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\Versus.png"))
    Private pntVersusText As New Point(284, 71)
    Private btmVersusHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\Menu\VersusHover.png"))
    Private pntVersusHoverText As New Point(256, 63)

    'Hover
    Private thrHoverSoundDelay As System.Threading.Thread

    'Sound
    Private udcAmbianceSound As clsSound

    'General, common uses
    Private btmBackText As New Bitmap(Image.FromFile(strDirectory & "\Images\General\Back.png"))
    Private pntBackText As New Point(1439, 46)
    Private btmBackHoverText As New Bitmap(Image.FromFile(strDirectory & "\Images\General\BackHover.png"))
    Private pntBackHoverText As New Point(1418, 35)

    'Options screen
    Private btmOptionsBackground As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\OptionsBackground.jpg"))
    Private btmResolutionText As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\Resolution.png"))
    Private pntResolutionText As New Point(40, 41)
    Private btm800x600Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\800x600.png"))
    Private btmNot800x600Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\not800x600.png"))
    Private pnt800x600Text As New Point(85, 142)
    Private btm1024x768Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\1024x768.png"))
    Private btmNot1024x768Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\not1024x768.png"))
    Private pnt1024x768Text As New Point(85, 192)
    Private btm1280x800Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\1280x800.png"))
    Private btmNot1280x800Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\not1280x800.png"))
    Private pnt1280x800Text As New Point(85, 242)
    Private btm1280x1024Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\1280x1024.png"))
    Private btmNot1280x1024Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\not1280x1024.png"))
    Private pnt1280x1024Text As New Point(85, 293)
    Private btm1440x900Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\1440x900.png"))
    Private btmNot1440x900Text As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\not1440x900.png"))
    Private pnt1440x900Text As New Point(85, 342)
    Private btmFullscreenText As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\Fullscreen.png"))
    Private btmNotFullscreenText As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\notFullscreen.png"))
    Private pntFullscreenText As New Point(85, 391)
    Private btmSoundText As New Bitmap(Image.FromFile(strDirectory & "\Images\Options\Sound.png"))
    Private pntSoundText As New Point(40, 447)
    Private intResolutionMode As Integer = 0 'Default 800x600
    Private intFullscreenWidth As Integer = 0 'Means not fullscreen
    Private intFullscreenHeight As Integer = 0 'Means not fullscreen

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
            udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, True) '38 seconds + extra
            'Set full screen rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth - WIDTHSUBTRACTION, intScreenHeight - HEIGHTSUBTRACTION) 'Full screen
            'Set for fog
            thrFog = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf KeepFogMoving))
            thrFog.Start()
            'Start rendering
            thrRendering.Start()
        End If

    End Sub

    Private Sub RenderMenuScreen(btmMousedOver As Bitmap)

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
        If btmMousedOver Is btmStartText Then
            gGraphics.DrawImageUnscaled(btmStartHoverText, pntStartHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmStartText, pntStartText)
        End If

        'Draw highscores text
        If btmMousedOver Is btmHighscoresText Then
            gGraphics.DrawImageUnscaled(btmHighscoresHoverText, pntHighscoresHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmHighscoresText, pntHighscoresText)
        End If

        'Draw story text
        If btmMousedOver Is btmStoryText Then
            gGraphics.DrawImageUnscaled(btmStoryHoverText, pntStoryHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmStoryText, pntStoryText)
        End If

        'Draw credits text
        If btmMousedOver Is btmOptionsText Then
            gGraphics.DrawImageUnscaled(btmOptionsHoverText, pntOptionsHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmOptionsText, pntOptionsText)
        End If

        'Draw options text
        If btmMousedOver Is btmCreditsText Then
            gGraphics.DrawImageUnscaled(btmCreditsHoverText, pntCreditsHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmCreditsText, pntCreditsText)
        End If

        'Draw versus text
        If btmMousedOver Is btmVersusText Then
            gGraphics.DrawImageUnscaled(btmVersusHoverText, pntVersusHoverText)
        Else
            gGraphics.DrawImageUnscaled(btmVersusText, pntVersusText)
        End If

        'Draw last stand text
        gGraphics.DrawImageUnscaled(btmLastStandText, pntLastStandText)

    End Sub

    Private Sub Rendering()

        'Loop
        While True
            'Paint on canvas first
            gGraphics = Graphics.FromImage(btmCanvas)
            'Set graphic options
            SetDefaultGraphicOptions()
            'Check mode before painting on canvas
            Select Case intCanvasMode
                Case 0 'Menu
                    'Check which to show
                    Select Case intCanvasShow
                        Case 0 'Blank no animation
                            'Render Menu
                            RenderMenuScreen(Nothing)
                        Case 1 'Start
                            'Render Menu
                            RenderMenuScreen(btmStartText)
                        Case 2 'Highscores
                            'Render Menu
                            RenderMenuScreen(btmHighscoresText)
                        Case 3 'Story
                            'Render Menu
                            RenderMenuScreen(btmStoryText)
                        Case 4 'Options
                            'Render Menu
                            RenderMenuScreen(btmOptionsText)
                        Case 5 'Credits
                            'Render Menu
                            RenderMenuScreen(btmCreditsText)
                        Case 6 'Versus
                            'Render Menu
                            RenderMenuScreen(btmVersusText)
                    End Select
                Case 1 'Options screen
                    'Check which to show
                    Select Case intCanvasShow
                        Case 0 'Blank no animation
                            RenderOptionsScreen(False)
                        Case 1 'Back button
                            RenderOptionsScreen(True)
                    End Select
            End Select
            'Select to paint on screen
            gGraphics = Me.CreateGraphics()
            'Set graphic options
            SetDefaultGraphicOptions()
            'Paint on screen
            gGraphics.DrawImage(btmCanvas, rectFullScreen)
            'If changing screen, we must change resolution in this thread or else strange things happen
            If blnScreenChanged Then
                'Change window state
                If intResolutionMode <> 5 Then '5 = fullscreen
                    'Normal
                    Me.Invoke(Sub() Me.WindowState = FormWindowState.Normal)
                    'Change
                    Me.Invoke(Sub() Me.Width = intScreenWidth)
                    Me.Invoke(Sub() Me.Height = intScreenHeight)
                Else
                    'Force full screen, let windows do stuff for us
                    Me.Invoke(Sub() Me.WindowState = FormWindowState.Maximized)
                    'Save width and height
                    intFullscreenWidth = Me.Width
                    intFullscreenHeight = Me.Height
                End If
                'Set
                gdblScreenWidthRatio = CDbl((Me.Width - WIDTHSUBTRACTION) / ORIGINALSCREENWIDTH)
                gdblScreenHeightRatio = CDbl((Me.Height - HEIGHTSUBTRACTION) / ORIGINALSCREENHEIGHT)
                'Set screen rectangle
                rectFullScreen.Width = Me.Width - WIDTHSUBTRACTION
                rectFullScreen.Height = Me.Height - HEIGHTSUBTRACTION
                'Center the form
                If intResolutionMode <> 5 Then
                    Me.Invoke(Sub() Me.Top = CInt((My.Computer.Screen.WorkingArea.Height / 2) - (Me.Height / 2)))
                    Me.Invoke(Sub() Me.Left = CInt((My.Computer.Screen.WorkingArea.Width / 2) - (Me.Width / 2)))
                End If
                'Reset
                blnScreenChanged = False
            End If
            'Do events
            Application.DoEvents()
        End While

    End Sub

    Private Sub RenderOptionsScreen(blnBackButtonSelected As Boolean)

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

        'Check
        If blnBackButtonSelected Then
            'Draw back text
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

    Private Function blnMousedOver(pntMouse As Point, intImageWidth As Integer, intImageHeight As Integer, pntStartingPoint As Point,
                                   intCanvasShowToSet As Integer, blnOnTopOf As Boolean) As Boolean

        'Check for mouse over
        If blnMouseInRegion(pntMouse, intImageWidth, intImageHeight, pntStartingPoint) Then
            'Set canvas show
            intCanvasShow = intCanvasShowToSet
            'Set zombie sound only once
            OnlyPlayZombieMenuSoundOnceThenWait(blnOnTopOf)
            'Return
            Return True
        Else
            'Return
            Return False
        End If

    End Function

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x and y points to ensure it always works.

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)
        Static sblnOnTopOf As Boolean = False

        'Menu stuff
        If intCanvasMode = 0 Then

            'Start has been moused over
            If blnMousedOver(pntMouse, 212, 69, pntStartText, 1, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Highscores has been moused over
            If blnMousedOver(pntMouse, 413, 99, pntHighscoresText, 2, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Story has been moused over
            If blnMousedOver(pntMouse, 218, 87, pntStoryText, 3, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Options has been moused over
            If blnMousedOver(pntMouse, 289, 89, pntOptionsText, 4, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Credits has been moused over
            If blnMousedOver(pntMouse, 285, 78, pntCreditsText, 5, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Versus has been moused over
            If blnMousedOver(pntMouse, 256, 74, pntVersusText, 6, sblnOnTopOf) Then
                'They moused over, executed code, now exit
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
            If blnMousedOver(pntMouse, 190, 74, pntBackText, 1, sblnOnTopOf) Then
                'They moused over, executed code, now exit
                Exit Sub
            End If

            'Reset
            sblnOnTopOf = False

            'Repaint options background
            intCanvasShow = 0

        End If

    End Sub

    Private Sub OnlyPlayZombieMenuSoundOnceThenWait(ByRef blnByRefOnTopOf As Boolean)

        'Check if mouse was on top of button once before
        If Not blnByRefOnTopOf Then
            'Set
            blnByRefOnTopOf = True
            'Hover sound
            If thrHoverSoundDelay Is Nothing Then
                'Play sound
                Dim udcButtonHoverSound As New clsSound(Me, AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonHover.mp3", 3000)
                'Start a waiting thread of 2500 ms
                thrHoverSoundDelay = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf HoverSoundWaiting))
                thrHoverSoundDelay.Start()
            End If
        End If

    End Sub

    Private Sub HoverSoundWaiting()

        'Sleep for 3000 ms
        System.Threading.Thread.Sleep(3000)
        thrHoverSoundDelay = Nothing

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Notes: Sometimes the scale of a screen can totally make pixels not match with formulas, in this case we use hard coded x and y points to ensure it always works.

        'Declare
        Dim pntMouse As Point = Me.PointToClient(MousePosition)

        'If menu
        If intCanvasMode = 0 Then

            'Options was clicked
            If blnMouseInRegion(pntMouse, 289, 89, pntOptionsText) Then
                'Stop fog thread
                thrFog.Abort()
                'Set
                intCanvasMode = 1
                intCanvasShow = 0
                'Stop sound
                udcAmbianceSound.StopAndCloseSound()
                'Play sound
                Dim udcButtonPressedSound As New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 3000)
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
                'Back was pressed
                BackButtonPressed()
                'Menu sound
                udcAmbianceSound = New clsSound(Me, strDirectory & "Sounds\Ambiance.mp3", 39000, True) '38 seconds + extra
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

    End Sub

    Private Sub BackButtonPressed()

        'Set
        intCanvasMode = 0
        intCanvasShow = 0

        'Play sound
        Dim udcButtonPressedSound As New clsSound(Me, strDirectory & "Sounds\ButtonPressed.mp3", 3000)

        'Restart fog
        RestartFog()

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

End Class