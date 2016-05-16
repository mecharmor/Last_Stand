'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsHelicopter

    'Declare
    Private _frmToPass As Form
    Private intFrame As Integer = 1

    'For error preventing
    Private blnKeepUsingAnimatingThread As Boolean = True

    'Bitmaps
    Private btmHelicopter As Bitmap
    Private pntHelicopter As Point

    'Thread
    Private thrAnimating As System.Threading.Thread
    Private _intAnimatingDelay As Integer = 0

    'Blade sound
    Private udcRotatingBladeSound As clsSound

    Public Sub New(frmToPass As Form, Optional blnStartAnimation As Boolean = False)

        'Set
        _frmToPass = frmToPass

        'Set animation
        btmHelicopter = gbtmHelicopter(0)

        'Set
        pntHelicopter = New Point(2879, 0)

        'Start
        If blnStartAnimation Then
            Start()
        End If

    End Sub

    Public Sub Start(Optional intAnimatingDelay As Integer = 100)

        'Set
        _intAnimatingDelay = intAnimatingDelay

        'Start thread
        thrAnimating = New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf Animating))
        thrAnimating.Start()

        'Start rotating blade sound
        udcRotatingBladeSound = New clsSound(_frmToPass, AppDomain.CurrentDomain.BaseDirectory & "Sounds\RotatingBlade.mp3", 98000, gintSoundVolume)

    End Sub

    Public Sub StopAndDispose()

        'Abort animating
        If thrAnimating IsNot Nothing Then
            If thrAnimating.IsAlive Then
                thrAnimating.Abort()
                thrAnimating = Nothing
            End If
        End If

    End Sub

    Public WriteOnly Property KeepUsingAnimatingThread() As Boolean

        'Set
        Set(value As Boolean)
            blnKeepUsingAnimatingThread = value
        End Set

    End Property

    Public WriteOnly Property RotationDelaySpeed() As Integer

        'Set
        Set(value As Integer)
            _intAnimatingDelay = value
        End Set

    End Property

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

    Public Sub StopRotatingBladeSound()

        'Stop sound
        If udcRotatingBladeSound IsNot Nothing Then
            udcRotatingBladeSound.StopAndCloseSound()
        End If

    End Sub

    Private Sub Animating()

        'Continue
        While blnKeepUsingAnimatingThread

            'Check the frame
            Select Case intFrame
                Case 1
                    btmHelicopter = gbtmHelicopter(1)
                Case 2
                    btmHelicopter = gbtmHelicopter(2)
                Case 3
                    btmHelicopter = gbtmHelicopter(3)
                Case 4
                    btmHelicopter = gbtmHelicopter(4)
                Case 5
                    btmHelicopter = gbtmHelicopter(0)
            End Select

            'Increase frame
            intFrame += 1
            If intFrame = 6 Then
                intFrame = 1
            End If

            'Sleep
            System.Threading.Thread.Sleep(_intAnimatingDelay)

        End While

    End Sub

End Class