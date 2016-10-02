'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsChainedZombie

    'Constant
    Private Const intZOMBIE_ANIMATION_DELAY As Integer = 4000
    Private Const intZOMBIE_ANIMATION_TWITCH As Integer = 100
    Private Const intZOMBIE_ANIMATION_WAITING As Integer = 1000

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1
    Private blnTwitch As Boolean = True

    'Gagging sound
    Private _audcSmallChainGagSounds() As clsSound

    'Ending thread
    Private blnThreadDisposing As Boolean = False

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Thread
    Private thrAnimating As System.Threading.Thread

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

    Private Sub ElapsedAnimation(sender As Object, e As EventArgs)

        'Disable timer
        tmrAnimation.Enabled = False

        'Start thread
        If Not blnThreadDisposing Then
            thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
            thrAnimating.Start()
        End If

    End Sub


    Public Sub Start(Optional intAnimatingDelay As Integer = intZOMBIE_ANIMATION_DELAY)

        'Set timer delay
        tmrAnimation.Interval = CDbl(intAnimatingDelay)

        'Start the animating thread by enabling the timer
        tmrAnimation.Enabled = True

    End Sub

    Public Sub StopAndDispose()

        'Set
        blnThreadDisposing = True

        'Disable timers
        tmrAnimation.Enabled = False

        'Stop and dispose timers
        tmrAnimation.Stop()
        tmrAnimation.Dispose()

        'Remove handlers
        RemoveHandler tmrAnimation.Elapsed, AddressOf ElapsedAnimation

        'Abort animating
        If thrAnimating IsNot Nothing Then
            'Wait
            While thrAnimating.IsAlive
            End While
            'Set
            thrAnimating = Nothing
        End If

    End Sub

    Private Sub Animating()

        'Static
        Static sintTwitchNumber As Integer = 0

        'Check the frame
        Select Case intFrame
            Case 1
                btmImage = gabtmChainedZombieMemories(1)
            Case 2
                btmImage = gabtmChainedZombieMemories(2)
        End Select

        'Check for twitching
        If blnTwitch Then
            'Change interval
            tmrAnimation.Interval = intZOMBIE_ANIMATION_TWITCH
            'Declare
            Dim rndNumber As New Random
            'Play gag sound
            _audcSmallChainGagSounds(rndNumber.Next(0, 2)).PlaySound(gintSoundVolume)
            'Set
            blnTwitch = False
        Else 'Looks like twitching
            'Change
            sintTwitchNumber += intZOMBIE_ANIMATION_TWITCH
            'Check
            If sintTwitchNumber = intZOMBIE_ANIMATION_DELAY Then
                'Change interval
                tmrAnimation.Interval = intZOMBIE_ANIMATION_WAITING
                'Reset
                sintTwitchNumber = 0
                'Set
                blnTwitch = True
            End If
        End If

        'Increase frame
        intFrame += 1

        'Check frame
        If intFrame = 3 Then
            intFrame = 1
        End If

        'Enable timer, unless need to dispose
        If Not blnThreadDisposing Then
            tmrAnimation.Enabled = True
        End If

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