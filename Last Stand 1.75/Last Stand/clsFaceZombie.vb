'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsFaceZombie

    'Constants
    Private Const dblZOMBIE_MOVING_DELAY As Double = 62.5

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 0

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Texture index for OpenGL
    Private intTextureIndex As Integer

    'Sound
    Private _udcFaceZombieEyesOpenSound As clsSound

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, udcFaceZombieEyesOpenSound As clsSound, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set frame
        intFrame = 0

        'Set animation
        btmImage = gabtmFaceZombieMemories(0)
        intTextureIndex = 0

        'Set
        pntPoint = New Point(intSpawnX, intSpawnY)

        'Set sound
        _udcFaceZombieEyesOpenSound = udcFaceZombieEyesOpenSound

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handler for the timer animation
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start(dblZOMBIE_MOVING_DELAY)
        End If

    End Sub

    Public Sub Start(Optional dblAnimatingDelay As Double = dblZOMBIE_MOVING_DELAY)

        'Set timer delay
        tmrAnimation.Interval = dblAnimatingDelay

        'Start the animating thread
        tmrAnimation.Enabled = True

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Check frame
        Select Case intFrame

            Case 0
                'Move right
                Point = New Point(Point.X + 4, Point.Y)

            Case 1
                'Move left
                Point = New Point(Point.X - 4, Point.Y)

            Case 2
                'Do nothing

            Case 3
                'Move left
                Point = New Point(Point.X - 4, Point.Y)

            Case 4
                'Move right
                Point = New Point(Point.X + 4, Point.Y)

        End Select

        'Increase frame
        If intFrame = 4 Then
            'Reset
            intFrame = 0
        Else
            'Increase
            intFrame += 1
        End If

    End Sub

    Public Sub StopAndDisposeTimer()

        'Stop and dispose timer variable
        gStopAndDisposeTimerVariable(tmrAnimation)

        'Remove handler
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

    Public ReadOnly Property Image() As Bitmap

        'Return
        Get
            Return btmImage
        End Get

    End Property

    Public Property Point() As Point

        'Return
        Get
            Return pntPoint
        End Get

        'Set
        Set(value As Point)
            pntPoint = value
        End Set

    End Property

    Public ReadOnly Property TextureIndex() As Integer

        'Return
        Get
            Return intTextureIndex
        End Get

    End Property

    Public Sub OpenEyes()

        'Change frame
        btmImage = gabtmFaceZombieMemories(1)
        intTextureIndex = 1

        'Play sound
        _udcFaceZombieEyesOpenSound.PlaySound(gintSoundVolume)

    End Sub

    Public Sub Move()

        'Disable
        tmrAnimation.Enabled = False

        'Set
        tmrAnimation.Interval = dblZOMBIE_MOVING_DELAY

        'Set
        intFrame = 0

        'Set
        tmrAnimation.Enabled = True

        'Enable
        tmrAnimation.Enabled = True

    End Sub

End Class