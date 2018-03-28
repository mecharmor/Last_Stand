'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsChainedZombie

    'Constant
    Private Const dblZOMBIE_ANIMATION_DELAY As Double = 4000
    Private Const dblZOMBIE_ANIMATION_TWITCH As Double = 100
    Private Const dblZOMBIE_ANIMATION_WAITING As Double = 1000

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 0

    'Gagging sound
    Private _audcSmallChainGagSounds() As clsSound

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, audcSmallChainGagSounds() As clsSound, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set animation
        btmImage = gabtmChainedZombieMemories(0)

        'Set
        pntPoint = New Point(intSpawnX, intSpawnY)

        'Set timer
        tmrAnimation.AutoReset = True

        'Set
        _audcSmallChainGagSounds = audcSmallChainGagSounds

        'Add handlers
        AddHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional dblAnimatingDelay As Double = dblZOMBIE_ANIMATION_DELAY)

        'Set timer delay
        tmrAnimation.Interval = dblAnimatingDelay

        'Start the animating thread by enabling the timer
        tmrAnimation.Enabled = True

    End Sub

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Static
        Static sdblTwitchNumber As Double = 0
        Static sblnTwitch As Boolean = True

        'Check the frame
        Select Case intFrame
            Case 0
                btmImage = gabtmChainedZombieMemories(1)
            Case 1
                btmImage = gabtmChainedZombieMemories(2)
            Case 2
                btmImage = gabtmChainedZombieMemories(0)
        End Select

        'Check for twitching
        If sblnTwitch Then
            'Change interval
            tmrAnimation.Interval = dblZOMBIE_ANIMATION_TWITCH
            'Play gag sound
            _audcSmallChainGagSounds(gGetRandomNumber(0, 1)).PlaySound(gintSoundVolume)
            'Set
            sblnTwitch = False
        Else 'Looks like twitching
            'Change
            sdblTwitchNumber += dblZOMBIE_ANIMATION_TWITCH
            'Check
            If sdblTwitchNumber = dblZOMBIE_ANIMATION_DELAY Then
                'Change interval
                tmrAnimation.Interval = dblZOMBIE_ANIMATION_WAITING
                'Reset
                sdblTwitchNumber = 0
                'Set
                sblnTwitch = True
            End If
        End If

        'Increase frame
        intFrame += 1

        'Check frame
        If intFrame = 3 Then
            intFrame = 0
        End If

    End Sub

    Public Sub StopAndDispose()

        'Disable timers
        tmrAnimation.Enabled = False

        'Stop and dispose timers
        tmrAnimation.Stop()
        tmrAnimation.Dispose()

        'Remove handlers
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

    End Sub

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

    Public ReadOnly Property Image() As Bitmap

        'Return
        Get
            Return btmImage
        End Get

    End Property

End Class