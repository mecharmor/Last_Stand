'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsHelicopter

    'Constant
    Private Const HELICOPTER_ROTATINGBLADE_DELAY As Integer = 120

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1

    'Sound
    Private _udcRotatingBladeSound As clsSound

    'Ending thread
    Private blnThreadDisposing As Boolean = False

    'Bitmaps
    Private btmHelicopter As Bitmap
    Private pntHelicopter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread

    'Timer
    Private tmrAnimation As New System.Timers.Timer

    Public Sub New(frmToPass As Form, udcRotatingBladeSound As clsSound, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set animation
        btmHelicopter = gabtmHelicopterMemories(0)

        'Set
        pntHelicopter = New Point(2879, 0)

        'Set sound
        _udcRotatingBladeSound = udcRotatingBladeSound

        'Start the sound in a loop
        _udcRotatingBladeSound.PlaySound(gintSoundVolume, True)

        'Set timer
        tmrAnimation.AutoReset = True

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


    Public Sub Start(Optional intAnimatingDelay As Integer = HELICOPTER_ROTATINGBLADE_DELAY)

        'Set timer delay
        tmrAnimation.Interval = CDbl(intAnimatingDelay)

        'Start the animating thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

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

        'Check the frame
        Select Case intFrame
            Case 1
                btmHelicopter = gabtmHelicopterMemories(1)
            Case 2
                btmHelicopter = gabtmHelicopterMemories(2)
            Case 3
                btmHelicopter = gabtmHelicopterMemories(3)
            Case 4
                btmHelicopter = gabtmHelicopterMemories(4)
            Case 5
                btmHelicopter = gabtmHelicopterMemories(0)
        End Select

        'Increase frame
        intFrame += 1
        If intFrame = 6 Then
            intFrame = 1
        End If

        'Enable timer, unless need to dispose
        If Not blnThreadDisposing Then
            tmrAnimation.Enabled = True
        End If

    End Sub

    Public Property HelicopterPoint() As Point

        'Return
        Get
            Return pntHelicopter
        End Get

        'Set
        Set(value As Point)
            pntHelicopter = value
        End Set

    End Property

    Public ReadOnly Property HelicopterImage() As Bitmap

        'Return
        Get
            Return btmHelicopter
        End Get

    End Property

End Class