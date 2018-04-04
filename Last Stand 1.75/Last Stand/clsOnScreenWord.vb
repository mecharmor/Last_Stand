'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsOnScreenWord

    'Declare
    Private intFrame As Integer = 1 'Defaulted
    Private blnOnScreen As Boolean = False

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Texture index for OpenGL
    Private intTextureIndex As Integer

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(Optional blnStartAnimation As Boolean = False)

        'Set animation
        btmImage = gabtmOnScreenWordMissedRedMemories(0)
        intTextureIndex = 0

        'Set timer
        tmrAnimation.AutoReset = True

        'Add handler for the timer animation
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional dblAnimatingDelay As Double = 90)

        'Set timer delay
        tmrAnimation.Interval = dblAnimatingDelay

        'Start the animating thread
        tmrAnimation.Enabled = True

        'Set
        blnOnScreen = True

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Check frame
        Select Case intFrame

            Case 0
                'Default do nothing
                tmrAnimation.Enabled = False
                'Set
                blnOnScreen = False

            Case 1
                'Set animation
                btmImage = gabtmOnScreenWordMissedRedMemories(1)
                intTextureIndex = 1

            Case 2
                'Set animation
                btmImage = gabtmOnScreenWordMissedBlueMemories(1)
                intTextureIndex = 3

        End Select

        'Reset frame
        intFrame = 0

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

    Public Sub StopAndDisposeTimer()

        'Stop and dispose timer variable
        gStopAndDisposeTimerVariable(tmrAnimation)

        'Remove handler
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

    Public Sub ShowWord(dblWordShowDelay As Double, intCurrentFrame As Integer, intX As Integer, intY As Integer)

        'Disable
        tmrAnimation.Enabled = False

        'Set
        tmrAnimation.Interval = dblWordShowDelay

        'Set frame
        intFrame = intCurrentFrame + 1

        'Set picture
        Select Case intCurrentFrame

            Case 0
                'Set animation
                btmImage = gabtmOnScreenWordMissedRedMemories(0)
                intTextureIndex = 0

            Case 1
                'Set animation
                btmImage = gabtmOnScreenWordMissedBlueMemories(0)
                intTextureIndex = 2

        End Select

        'Set point
        pntPoint = New Point(intX, intY)

        'Turn on timer
        tmrAnimation.Start()

        'Set
        blnOnScreen = True

        'Enable
        tmrAnimation.Enabled = True

    End Sub

    Public ReadOnly Property OnScreen() As Boolean

        'Return
        Get
            Return blnOnScreen
        End Get

    End Property

End Class