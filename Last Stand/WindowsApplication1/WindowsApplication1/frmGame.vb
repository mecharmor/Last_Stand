'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Screen
    Private intScreenWidth As Integer = 0
    Private intScreenHeight As Integer = 0
    Private dblScreenWidthRatio As Double = 0
    Private dblScreenHeightRatio As Double = 0

    'Canvas
    Private intCanvasMode As Integer = 0
    Private btmCanvas As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Canvas.jpg"), 1920, 1200)
    Private pntTopLeft As New Point(0, 0)
    Private intSystemHeightCaption As Integer = 0 'Usually 22 pixels
    Private rectFullScreen As Rectangle 'Setup for full screen later

    'Menu
    Private btmMenu As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Menu.jpg"), 1920, 1200)

    'Last stand
    Private btmLastStand As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LastStand.png"), 1350, 200)
    Private pntLastStand As New Point(282, 954)

    'Start
    Private btmStart As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Start.png"), 209, 66)
    Private pntStart As New Point(1250, 50)

    'Hover start
    Private btmStartHover As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/StartHover.png"), 295, 95)
    Private pntStartHover As New Point(1207, 36)

    'Story
    Private btmLoadingBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBackground.jpg"), 1920, 1200)

    'Loading paragraph
    Private thrLoadingParagraph As System.Threading.Thread
    Private btmLoadingParagraph25 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph25.png"), 1424, 472)
    Private btmLoadingParagraph50 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph50.png"), 1424, 472)
    Private btmLoadingParagraph75 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph75.png"), 1424, 472)
    Private btmLoadingParagraph100 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph100.png"), 1424, 472)
    Private pntLoadingParagraph As New Point(264, 350)
    Private intLoadingParagraph As Integer = 0

    'Loading bar
    Private btmLoading0 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/0.png"), 1613, 134)
    Private btmLoading10 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/10.png"), 1613, 134)
    Private btmLoading20 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/20.png"), 1613, 134)
    Private btmLoading30 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/30.png"), 1613, 134)
    Private btmLoading40 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/40.png"), 1613, 134)
    Private btmLoading50 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/50.png"), 1613, 134)
    Private btmLoading60 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/60.png"), 1613, 134)
    Private btmLoading70 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/70.png"), 1613, 134)
    Private btmLoading80 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/80.png"), 1613, 134)
    Private btmLoading90 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/90.png"), 1613, 134)
    Private btmLoading99 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/99.png"), 1613, 134) 'Troll bar
    Private btmLoading100 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/100.png"), 1613, 134)
    Private pntLoading As New Point(144, 945)
    Private intLoadingGameObjects As Integer = 0
    Private thrLoadingGameObjects As System.Threading.Thread

    'Loading text
    Private btmLoadingText As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/LoadingText.png"), 430, 101)
    Private pntLoadingText As New Point(754, 960)

    'Start text
    Private btmLoadingStartText As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/LoadingStartText.png"), 273, 83)
    Private pntLoadingStartText As New Point(820, 970)

    'Background
    Private btmBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Background.jpg"), 1920, 1200)

    'Zombies
    Private udcZombies(9) As Zombie

    'Character
    Private udcCharacter As Character

    'Graphics
    Private gGraphics As Graphics
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False

    'Sound
    Private udcAmbianceSound As Sound
    Private udcButtonHoverStartSound As Sound
    Private thrStartSoundWaiting As System.Threading.Thread
    Private udcGunShotSound As Sound
    Private udcShellEjectedSound As Sound

    'Constants
    Private Const WINDOWMESSAGE_SYSTEM_COMMAND As Integer = 274
    Private Const CONTROL_MOVE As Integer = 61456
    Private Const WINDOWMESSAGE_CLICK_BUTTONDOWN As Integer = 161
    Private Const WINDOWCAPTION As Integer = 2
    Private Const WINDOWMESSAGE_TITLE_BAR_DOUBLE_CLICKED As Integer = &HA3

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

        'Notes:'Do not allow a resize of the window, this happens if you double click the title, but is prevented with this sub, prevents moving the window around too

        'Prevent moving the form by control box click
        If (m.Msg = WINDOWMESSAGE_SYSTEM_COMMAND) And (m.WParam.ToInt32() = CONTROL_MOVE) Then
            Return
        End If

        'Prevent button down moving form
        If (m.Msg = WINDOWMESSAGE_CLICK_BUTTONDOWN) And (m.WParam.ToInt32() = WINDOWCAPTION) Then
            Return
        End If

        'If a double click on the title bar is triggered
        If m.Msg = WINDOWMESSAGE_TITLE_BAR_DOUBLE_CLICKED Then
            Return
        End If

        'Still send message but without resizing
        MyBase.WndProc(m)

    End Sub

    Private Sub frmGame_Load(sender As Object, e As EventArgs) Handles Me.Load

        'Check if multi-threading was possible
        If Not blnThreadSupported Then
            'Display
            MessageBox.Show("This computer doesn't support multi-threading. This application will close now.", "Last Stand", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Exit
            Me.Close()
        Else
            'Menu sound
            udcAmbianceSound = New Sound("Ambiance", AppDomain.CurrentDomain.BaseDirectory & "Sounds\Ambiance.mp3")
            udcAmbianceSound.PlaySound(False)
            'Screen
            intScreenWidth = Me.Width
            intScreenHeight = Me.Height
            'Set height caption, title bar measurement
            intSystemHeightCaption = SystemInformation.CaptionHeight
            'Current screen ratio
            dblScreenWidthRatio = CDbl(intScreenWidth) / 1920
            dblScreenHeightRatio = CDbl(intScreenHeight) / 1200
            'Set full rectangle
            rectFullScreen = New Rectangle(0, -intSystemHeightCaption, intScreenWidth, intScreenHeight) 'Full screen
            'Start rendering
            thrRendering.Start()
        End If

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Render thread needs to be aborted
        If thrRendering IsNot Nothing Then
            If thrRendering.IsAlive Then
                thrRendering.Abort()
            End If
        End If

        'Start sound waiting thread
        If thrStartSoundWaiting IsNot Nothing Then
            If thrStartSoundWaiting.IsAlive Then
                thrStartSoundWaiting.Abort()
            End If
        End If

        'Loading paragraph thread
        If thrLoadingParagraph IsNot Nothing Then
            If thrLoadingParagraph.IsAlive Then
                thrLoadingParagraph.Abort()
            End If
        End If

        'Loading game objects
        If thrLoadingGameObjects IsNot Nothing Then
            If thrLoadingGameObjects.IsAlive Then
                thrLoadingGameObjects.Abort()
            End If
        End If

        'Stop zombie if he exists
        If udcZombies(0) IsNot Nothing Then
            For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                If udcZombies(intLoop) IsNot Nothing Then
                    udcZombies(intLoop).StopZombie()
                End If
            Next
        End If

        'Stop character
        If udcCharacter IsNot Nothing Then
            udcCharacter.StopCharacter()
        End If

    End Sub

    Sub New()

        ' This call is required by the designer
        InitializeComponent() 'Do not remove

        'Prepare rendering thread
        thrRendering = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Rendering))

        'Check for multi-threading first
        If thrRendering.TrySetApartmentState(Threading.ApartmentState.MTA) Then
            'Set, multi-threading is possible
            blnThreadSupported = True
            'Stop the paint event, we will do the painting when we want to
            MyBase.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
            MyBase.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            MyBase.SetStyle(ControlStyles.UserPaint, False)
            'Lock down the window size
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
        End If

    End Sub

    Private Sub Rendering()

        'Loop
        While True
            'Paint on canvas first
            gGraphics = Graphics.FromImage(btmCanvas)
            'Set options
            SetDefaultGraphicOptions(gGraphics)
            'Check which case of the canvas
            Select Case intCanvasMode
                Case 0
                    'Paint onto invisible canvas
                    gGraphics.DrawImageUnscaled(btmMenu, pntTopLeft)
                    gGraphics.DrawImageUnscaled(btmLastStand, pntLastStand)
                    gGraphics.DrawImageUnscaled(btmStart, pntStart)
                Case 1
                    'Paint onto invisible canvas
                    gGraphics.DrawImageUnscaled(btmMenu, pntTopLeft)
                    gGraphics.DrawImageUnscaled(btmLastStand, pntLastStand)
                    gGraphics.DrawImageUnscaled(btmStartHover, pntStartHover)
                Case 2
                    'Paint onto invisible canvas
                    gGraphics.DrawImageUnscaled(btmLoadingBackground, pntTopLeft)
                    'Check loading
                    Select Case intLoadingGameObjects
                        Case 0
                            gGraphics.DrawImageUnscaled(btmLoading0, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 10
                            gGraphics.DrawImageUnscaled(btmLoading10, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 20
                            gGraphics.DrawImageUnscaled(btmLoading20, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 30
                            gGraphics.DrawImageUnscaled(btmLoading30, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 40
                            gGraphics.DrawImageUnscaled(btmLoading40, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 50
                            gGraphics.DrawImageUnscaled(btmLoading50, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 60
                            gGraphics.DrawImageUnscaled(btmLoading60, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 70
                            gGraphics.DrawImageUnscaled(btmLoading70, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 80
                            gGraphics.DrawImageUnscaled(btmLoading80, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 90
                            gGraphics.DrawImageUnscaled(btmLoading90, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 99
                            gGraphics.DrawImageUnscaled(btmLoading99, pntLoading)
                            'Loading... text
                            gGraphics.DrawImageUnscaled(btmLoadingText, pntLoadingText)
                        Case 100
                            gGraphics.DrawImageUnscaled(btmLoading100, pntLoading)
                            'Start text
                            gGraphics.DrawImageUnscaled(btmLoadingStartText, pntLoadingStartText)
                    End Select
                    'Loading paragraph
                    Select Case intLoadingParagraph
                        Case 0
                            'Waiting
                        Case 1
                            gGraphics.DrawImageUnscaled(btmLoadingParagraph25, pntLoadingParagraph)
                        Case 2
                            gGraphics.DrawImageUnscaled(btmLoadingParagraph50, pntLoadingParagraph)
                        Case 3
                            gGraphics.DrawImageUnscaled(btmLoadingParagraph75, pntLoadingParagraph)
                        Case 4
                            gGraphics.DrawImageUnscaled(btmLoadingParagraph100, pntLoadingParagraph)
                    End Select
                Case 3
                    'Paint onto invisible canvas
                    gGraphics.DrawImageUnscaled(btmBackground, pntTopLeft)
                    'Show character
                    gGraphics.DrawImageUnscaled(udcCharacter.btmCharacter, udcCharacter.rectCharacter)
                    'Show zombies
                    For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                        If udcZombies(intLoop) IsNot Nothing Then
                            gGraphics.DrawImageUnscaled(udcZombies(intLoop).btmZombie, udcZombies(intLoop).rectZombie)
                        End If
                    Next
                    'Check distance
                    For IntLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                        If udcZombies(IntLoop) IsNot Nothing Then
                            'Close game if too close to character
                            If udcZombies(IntLoop).rectZombie.X <= 200 Then
                                Me.Invoke(Sub() Me.Close()) 'Exit
                            End If
                        End If
                    Next
            End Select
            'Set graphics
            gGraphics = Me.CreateGraphics()
            'Set options
            SetDefaultGraphicOptions(gGraphics)
            'Paint canvas to screen
            gGraphics.DrawImage(btmCanvas, rectFullScreen)
            'Do events
            Application.DoEvents()
        End While

    End Sub

    Private Sub SetDefaultGraphicOptions(gGraphicsDeviceContext As Graphics)

        'Set options for fastest rendering
        gGraphicsDeviceContext.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor
        gGraphicsDeviceContext.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
        gGraphicsDeviceContext.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None
        gGraphicsDeviceContext.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed
        gGraphicsDeviceContext.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixel

    End Sub

    Private Function MouseInRegion(intImageWidth As Integer, intImageHeight As Integer, pntStartingPoint As Point) As Boolean

        'Return
        If MousePosition.X >= CInt(pntStartingPoint.X * dblScreenWidthRatio) And
        MousePosition.X <= CInt((pntStartingPoint.X + intImageWidth) * dblScreenWidthRatio) And
        MousePosition.Y >= CInt(pntStartingPoint.Y * dblScreenHeightRatio) And
        MousePosition.Y <= CInt((pntStartingPoint.Y + intImageHeight) * dblScreenHeightRatio) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'If menu
        If intCanvasMode < 2 Then

            'Start has been clicked
            If MouseInRegion(209, 66, pntStart) Then
                'Set
                intCanvasMode = 2
                'Stop sound
                udcAmbianceSound.StopSound()
                'Play sound
                Dim udcButtonPressedStart As New Sound("ButtonPressedStart", AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonPressedStart.mp3")
                udcButtonPressedStart.PlaySound(False)
                'Start loading
                thrLoadingParagraph = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingParagraph))
                thrLoadingParagraph.Start()
                'Start loading of game objects
                thrLoadingGameObjects = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf LoadingGameObjects))
                thrLoadingGameObjects.Start()
            End If

            'Exit
            Exit Sub

        End If

        'If ready to start game
        If intCanvasMode = 2 And intLoadingGameObjects = 100 Then

            'Start has been clicked
            If MouseInRegion(1613, 134, pntLoading) Then
                'Set
                intCanvasMode = 3
                'Start character
                udcCharacter.Start()
                'Start zombies
                For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                    If udcZombies(intLoop) IsNot Nothing Then
                        udcZombies(intLoop).Start()
                    End If
                Next
            End If

            'Exit
            Exit Sub

        End If

        'If game is playing
        If intCanvasMode = 3 Then

            'Shoot immediately
            udcCharacter.CharacterShot()
            udcGunShotSound = New Sound("GunShot", AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3")
            udcGunShotSound.PlaySound(False)
            System.Threading.Thread.Sleep(500)
            udcShellEjectedSound = New Sound("ShellEjected", AppDomain.CurrentDomain.BaseDirectory & "Sounds\ShellEjected.mp3")
            udcShellEjectedSound.PlaySound(False)

            'Kill zombie
            For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                If udcZombies(intLoop) IsNot Nothing Then
                    If udcZombies(intLoop).IsAlive Then
                        udcZombies(intLoop).Dead()
                        Exit Sub
                    End If
                End If
            Next

            'Exit
            Exit Sub

        End If

    End Sub

    Private Sub LoadingGameObjects()

        'Set
        intLoadingGameObjects = 10
        intLoadingGameObjects = 20
        intLoadingGameObjects = 30
        intLoadingGameObjects = 40
        intLoadingGameObjects = 50

        'Character
        udcCharacter = New Character(CInt(100 * dblScreenWidthRatio))

        'Set
        intLoadingGameObjects = 60

        'Zombie
        udcZombies(0) = New Zombie(CInt(intScreenWidth * dblScreenWidthRatio))

        'Set
        intLoadingGameObjects = 70

        'Zombie
        udcZombies(1) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 200), False, 50)

        'Set
        intLoadingGameObjects = 80

        'Zombie
        udcZombies(2) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 400))

        'Set
        intLoadingGameObjects = 90

        'Zombie
        udcZombies(3) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 500))

        'Set
        intLoadingGameObjects = 99 'Trolling

        'Zombie
        udcZombies(4) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 700))
        'Zombie
        udcZombies(5) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 800))
        'Zombie
        udcZombies(6) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 900), False, 35)
        'Zombie
        udcZombies(7) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 1000))
        'Zombie
        udcZombies(8) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 1100))
        'Zombie
        udcZombies(9) = New Zombie(CInt((intScreenWidth * dblScreenWidthRatio) + 1200), False, 65)

        'Wait
        Application.DoEvents()

        'Set
        intLoadingGameObjects = 100

    End Sub

    Private Sub LoadingParagraph()

        'Wait
        System.Threading.Thread.Sleep(300)
        'Change
        intLoadingParagraph = 1 'Opacity 25%
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 2 'Opacity 50%
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 3 'Opacity 75%
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 4 'Opacity 100%

    End Sub

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Declare
        Static sblnOnTopOf As Boolean = False

        'Exit if
        If intCanvasMode < 2 Then

            'Start has been moused over
            If MouseInRegion(209, 66, pntStart) Then
                'Paint hovered over start
                intCanvasMode = 1
                'Only play once, don't keep looping
                If Not sblnOnTopOf Then
                    'Set
                    sblnOnTopOf = True
                    'Hover sound
                    If thrStartSoundWaiting Is Nothing Then
                        'Play sound
                        udcButtonHoverStartSound = New Sound("ButtonHoverStart", AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonHoverStart.mp3")
                        udcButtonHoverStartSound.PlaySound(False)
                        'Start a waiting thread of 2500 ms
                        thrStartSoundWaiting = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf StartSoundWaiting))
                        thrStartSoundWaiting.Start()
                    End If
                End If
            Else
                'Reset
                sblnOnTopOf = False
                'Repaint original
                intCanvasMode = 0
            End If

        End If

    End Sub

    Private Sub StartSoundWaiting()

        'Sleep for 2500 ms
        System.Threading.Thread.Sleep(2500)
        thrStartSoundWaiting = Nothing

    End Sub

End Class