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

    'Start
    Private btmStart As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Start.png"), 209, 66)
    Private rectStart As New Rectangle(1250, 50, 209, 66)

    'Hover start
    Private btmStartHover As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/StartHover.png"), 295, 95)
    Private rectStartHover As New Rectangle(1207, 36, 295, 95)

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
    Private udcGunShotSound As Sound

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
                    gGraphics.DrawImage(btmStart, rectStart)
                Case 1
                    'Paint onto invisible canvas
                    gGraphics.DrawImage(btmMenu, rectCanvas)
                    gGraphics.DrawImage(btmStartHover, rectStartHover)
                Case 2
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
            'Reduce CPU usage
            System.Threading.Thread.Sleep(10) '60 ms passes before a frame
        End While

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Exit if
        If intCanvasMode < 2 Then

            'Start has been clicked
            If MousePosition.X >= CInt(1250 * dblScreenWidthRatio) And MousePosition.X <= CInt((1250 + 209) * dblScreenWidthRatio) And
            MousePosition.Y >= CInt((50 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
            MousePosition.Y <= CInt((50 + 66 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then

                'Character
                udcCharacter = New Character(0 + 100)

                'Zombies
                udcZombies(0) = New Zombie(intScreenWidth)
                udcZombies(1) = New Zombie(intScreenWidth + 200)

                'Do events
                Application.DoEvents()

                'Set
                intCanvasMode = 2

                'Stop sound
                udcAmbianceSound.StopSound()

                'Play sound
                Dim udcButtonPressedStart As New Sound("ButtonPressedStart", AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonPressedStart.mp3")
                udcButtonPressedStart.PlaySound(False)

            End If

        Else

            udcCharacter.blnShot = True
            udcGunShotSound = New Sound("GunShot", AppDomain.CurrentDomain.BaseDirectory & "Sounds\GunShot.mp3")
            udcGunShotSound.PlaySound(False)

        End If

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
                'Only play once, don't keep looping
                If Not sblnOnTopOf Then
                    'Set
                    sblnOnTopOf = True
                    'Hover sound
                    udcButtonHoverStartSound = New Sound("ButtonHoverStart", AppDomain.CurrentDomain.BaseDirectory & "Sounds\ButtonHoverStart.mp3")
                    udcButtonHoverStartSound.PlaySound(False)
                End If
                'Paint hovered over start
                intCanvasMode = 1
            Else
                'Reset
                sblnOnTopOf = False
                'Repaint original
                intCanvasMode = 0
            End If

        End If

    End Sub

End Class