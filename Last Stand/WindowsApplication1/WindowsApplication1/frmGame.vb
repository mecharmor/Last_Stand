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
    Private rectCanvas As New Rectangle(0, 0, 1920, 1200)
    Private rectFullScreen As Rectangle

    'Menu
    Private btmMenu As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Menu.jpg"), 1920, 1200)

    'Last stand
    Private btmLastStand As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LastStand.png"), 1350, 200)
    Private rectLastStand As New Rectangle(282, 954, 1350, 200)

    'Start
    Private btmStart As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Start.png"), 209, 66)
    Private rectStart As New Rectangle(1250, 50, 209, 66)

    'Hover start
    Private btmStartHover As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/StartHover.png"), 295, 95)
    Private rectStartHover As New Rectangle(1207, 36, 295, 95)

    'Story
    Private btmLoadingBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBackground.jpg"), 1920, 1200)

    'Loading paragraph
    Private thrLoadingParagraph As System.Threading.Thread
    Private btmLoadingParagraph25 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph25.png"), 1424, 472)
    Private btmLoadingParagraph50 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph50.png"), 1424, 472)
    Private btmLoadingParagraph75 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph75.png"), 1424, 472)
    Private btmLoadingParagraph100 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingParagraph/LoadingParagraph100.png"), 1424, 472)
    Private rectLoadingParagraph As New Rectangle(264, 350, 1424, 472)
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
    Private rectLoading As New Rectangle(144, 945, 1613, 134)
    Private intLoadingGameObjects As Integer = 0
    Private thrLoadingGameObjects As System.Threading.Thread

    'Loading text
    Private btmLoadingText As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/LoadingText.png"), 430, 101)
    Private rectLoadingText As New Rectangle(754, 960, 430, 101)

    'Start text
    Private btmLoadingStartText As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/LoadingBar/LoadingStartText.png"), 273, 83)
    Private rectLoadingStartText As New Rectangle(820, 970, 273, 83)

    'Background
    Private btmBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Background.jpg"), 1920, 1200)

    'Zombie
    Private udcZombies(9) As Zombie

    'Character
    Private udcCharacter As Character

    'Game variables
    Private blnStartedGame As Boolean = False

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
            intScreenWidth = Me.Width '1696
            intScreenHeight = Me.Height '1066

            'Current screen ratio
            dblScreenWidthRatio = CDbl(intScreenWidth) / 1920
            dblScreenHeightRatio = CDbl(intScreenHeight) / 1200

            'Set full rectangle
            rectFullScreen = New Rectangle(0, 0, intScreenWidth, intScreenHeight) 'Full screen

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
            gGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor
            gGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
            gGraphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None
            gGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed
            gGraphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixel
            'Check which case of the canvas
            Select Case intCanvasMode
                Case 0
                    'Paint onto invisible canvas
                    gGraphics.DrawImage(btmMenu, rectCanvas)
                    gGraphics.DrawImage(btmLastStand, rectLastStand)
                    gGraphics.DrawImage(btmStart, rectStart)
                Case 1
                    'Paint onto invisible canvas
                    gGraphics.DrawImage(btmMenu, rectCanvas)
                    gGraphics.DrawImage(btmLastStand, rectLastStand)
                    gGraphics.DrawImage(btmStartHover, rectStartHover)
                Case 2
                    'Paint onto invisible canvas
                    gGraphics.DrawImage(btmLoadingBackground, rectCanvas)
                    'Check loading
                    Select Case intLoadingGameObjects
                        Case 0
                            gGraphics.DrawImage(btmLoading0, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 10
                            gGraphics.DrawImage(btmLoading10, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 20
                            gGraphics.DrawImage(btmLoading20, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 30
                            gGraphics.DrawImage(btmLoading30, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 40
                            gGraphics.DrawImage(btmLoading40, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 50
                            gGraphics.DrawImage(btmLoading50, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 60
                            gGraphics.DrawImage(btmLoading60, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 70
                            gGraphics.DrawImage(btmLoading70, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 80
                            gGraphics.DrawImage(btmLoading80, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 90
                            gGraphics.DrawImage(btmLoading90, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 99
                            gGraphics.DrawImage(btmLoading99, rectLoading)
                            'Loading... text
                            gGraphics.DrawImage(btmLoadingText, rectLoadingText)
                        Case 100
                            gGraphics.DrawImage(btmLoading100, rectLoading)
                            'Start text
                            gGraphics.DrawImage(btmLoadingStartText, rectLoadingStartText)
                    End Select
                    'Loading paragraph
                    Select Case intLoadingParagraph
                        Case 0
                            'Waiting
                        Case 1
                            gGraphics.DrawImage(btmLoadingParagraph25, rectLoadingParagraph)
                        Case 2
                            gGraphics.DrawImage(btmLoadingParagraph50, rectLoadingParagraph)
                        Case 3
                            gGraphics.DrawImage(btmLoadingParagraph75, rectLoadingParagraph)
                        Case 4
                            gGraphics.DrawImage(btmLoadingParagraph100, rectLoadingParagraph)
                    End Select
                Case 3
                    'Paint onto invisible canvas
                    gGraphics.DrawImage(btmBackground, rectCanvas)
                    'Show character
                    gGraphics.DrawImage(udcCharacter.btmCharacter, udcCharacter.rectCharacter)
                    'Show zombies
                    For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                        If udcZombies(intLoop) IsNot Nothing Then
                            gGraphics.DrawImage(udcZombies(intLoop).btmZombie, udcZombies(intLoop).rectZombie)
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
            'Do events
            Application.DoEvents()
            'Set graphics
            gGraphics = Me.CreateGraphics()
            'Set options
            gGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor
            gGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None
            gGraphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None
            gGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighSpeed
            gGraphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixel
            'Paint canvas to screen
            gGraphics.DrawImage(btmCanvas, rectFullScreen)
        End While

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'If menu
        If intCanvasMode < 2 Then

            'Start has been clicked
            If MousePosition.X >= CInt(1250 * dblScreenWidthRatio) And MousePosition.X <= CInt((1250 + 209) * dblScreenWidthRatio) And
            MousePosition.Y >= CInt((50 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
            MousePosition.Y <= CInt((50 + 66 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then
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
            If MousePosition.X >= CInt(144 * dblScreenWidthRatio) And MousePosition.X <= CInt((144 + 1613) * dblScreenWidthRatio) And
            MousePosition.Y >= CInt((945 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
            MousePosition.Y <= CInt((945 + 134 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then
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
        udcCharacter = New Character(100)

        'Set
        intLoadingGameObjects = 60

        'Zombie
        udcZombies(0) = New Zombie(intScreenWidth)

        'Set
        intLoadingGameObjects = 70

        'Zombie
        udcZombies(1) = New Zombie(intScreenWidth + 200, False, 50)

        'Set
        intLoadingGameObjects = 80

        'Zombie
        udcZombies(2) = New Zombie(intScreenWidth + 400)

        'Set
        intLoadingGameObjects = 90

        'Zombie
        udcZombies(3) = New Zombie(intScreenWidth + 500)

        'Set
        intLoadingGameObjects = 99 'Trolling

        'Zombie
        udcZombies(4) = New Zombie(intScreenWidth + 700)
        'Zombie
        udcZombies(5) = New Zombie(intScreenWidth + 800)
        'Zombie
        udcZombies(6) = New Zombie(intScreenWidth + 900, False, 35)
        'Zombie
        udcZombies(7) = New Zombie(intScreenWidth + 1000)
        'Zombie
        udcZombies(8) = New Zombie(intScreenWidth + 1100)
        'Zombie
        udcZombies(9) = New Zombie(intScreenWidth + 1200, False, 65)

        'Set
        intLoadingGameObjects = 100

    End Sub

    Private Sub LoadingParagraph()

        'Wait
        System.Threading.Thread.Sleep(300)
        'Change
        intLoadingParagraph = 1
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 2
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 3
        'Wait
        System.Threading.Thread.Sleep(200)
        'Change
        intLoadingParagraph = 4

    End Sub

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Declare
        Static sblnOnTopOf As Boolean = False

        'Exit if
        If intCanvasMode < 2 Then

            'Start has been moused over
            If MousePosition.X >= CInt(1250 * dblScreenWidthRatio) And MousePosition.X <= CInt((1250 + 209) * dblScreenWidthRatio) And
            MousePosition.Y >= CInt((50 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
            MousePosition.Y <= CInt((50 + 66 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then
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