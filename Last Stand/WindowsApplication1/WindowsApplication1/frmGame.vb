﻿'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class frmGame

    'Screen
    Private intScreenWidth As Integer = 0
    Private intScreenHeight As Integer = 0
    Private dblScreenWidthRatio As Double = 0
    Private dblScreenHeightRatio As Double = 0

    'Menu
    Private btmMenu As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Menu.jpg"), 1920, 1200)
    Private rectMenu As Rectangle

    'Start
    Private btmStart As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Start.png"), 209, 66)
    Private rectStart As Rectangle

    'Background
    Private btmBackground As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Background.jpg"), 1280, 720)
    Private btmBackgroundCopy As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Background.jpg"), 1280, 720)
    Private rectBackground As Rectangle

    'Zombie
    Private udcZombies(9) As Zombie

    'Game variables
    Private blnStartedGame As Boolean = False

    'Graphics
    Private gGraphics As Graphics
    Private thrRendering As System.Threading.Thread
    Private blnThreadSupported As Boolean = False

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

            'Exit
            Me.Close()

        Else

            'Screen
            intScreenWidth = Me.Width '1696
            intScreenHeight = Me.Height '1066

            'Current screen ratio
            dblScreenWidthRatio = CDbl(intScreenWidth) / 1920
            dblScreenHeightRatio = CDbl(intScreenHeight) / 1200

            'Set rectangles
            rectMenu = New Rectangle(0, 0, intScreenWidth, intScreenHeight)
            rectStart = New Rectangle(1250, 50, 209, 66)
            rectBackground = New Rectangle(0, 0, intScreenWidth, intScreenHeight)

            'Paint onto invisible canvas
            gGraphics = Graphics.FromImage(btmMenu)
            gGraphics.DrawImage(btmStart, rectStart)

            'Set graphics
            gGraphics = Me.CreateGraphics() 'Prepare to draw to the form

            'Start rendering
            thrRendering.Start()

        End If

    End Sub

    Private Sub frmGame_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Render thread needs to be aborted
        If thrRendering.IsAlive Then
            thrRendering.Abort()
        End If

        'Stop zombie if he exists
        If blnStartedGame Then
            For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                udcZombies(intLoop).StopZombie()
            Next
        End If

    End Sub

    Sub New()

        ' This call is required by the designer
        InitializeComponent() 'Do not remove

        'Prepare rendering thread
        thrRendering = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Rendering))

        'Check for multi-threading first
        If Not thrRendering.TrySetApartmentState(Threading.ApartmentState.MTA) Then

            'Display
            MessageBox.Show("This computer doesn't support multi-threading. This application will close now.", "Last Stand", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else

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
            'If game not started
            If Not blnStartedGame Then
                'Draw menu
                gGraphics.DrawImage(btmMenu, rectMenu)
            Else
                gGraphics = Graphics.FromImage(btmBackground)
                Try
                    For intLoop As Integer = 0 To udcZombies.GetUpperBound(0)
                        gGraphics.DrawImage(udcZombies(intLoop).btmZombie, udcZombies(intLoop).rectZombie)
                    Next
                Catch ex As Exception
                End Try
                gGraphics = Me.CreateGraphics()
                gGraphics.DrawImage(btmBackground, rectBackground)
                gGraphics = Graphics.FromImage(btmBackground)
                gGraphics.DrawImage(btmBackgroundCopy, New Rectangle(0, 0, 1280, 720))
            End If
            'Reduce CPU usage
            System.Threading.Thread.Sleep(60) '60 ms passes before a frame
        End While

    End Sub

    Private Sub frmGame_Click(sender As Object, e As EventArgs) Handles Me.Click

        'Start has been clicked
        If MousePosition.X >= CInt(1250 * dblScreenWidthRatio) And MousePosition.X <= CInt((1250 + 209) * dblScreenWidthRatio) And
        MousePosition.Y >= CInt((50 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
        MousePosition.Y <= CInt((50 + 66 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then
            blnStartedGame = True
            udcZombies(0) = New Zombie(intScreenWidth)
            udcZombies(1) = New Zombie(intScreenWidth + 50)
            udcZombies(2) = New Zombie(intScreenWidth - 14)
            udcZombies(3) = New Zombie(intScreenWidth + 15)
            udcZombies(4) = New Zombie(intScreenWidth + 20)
            udcZombies(5) = New Zombie(intScreenWidth - 10)
            udcZombies(6) = New Zombie(intScreenWidth + 100)
            udcZombies(7) = New Zombie(intScreenWidth + 67)
            udcZombies(8) = New Zombie(intScreenWidth - 59)
            udcZombies(9) = New Zombie(intScreenWidth + 200)
        End If

    End Sub

    Private Sub frmGame_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        'Start has been moused over
        If MousePosition.X >= CInt(1250 * dblScreenWidthRatio) And MousePosition.X <= CInt((1250 + 209) * dblScreenWidthRatio) And
        MousePosition.Y >= CInt((50 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) And
        MousePosition.Y <= CInt((50 + 66 + SystemInformation.CaptionHeight) * dblScreenHeightRatio) Then
            Debug.Print("Start is moused over")
        Else
            'Repaint original
        End If

    End Sub

End Class