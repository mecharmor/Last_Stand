'Options
Option Explicit On
Option Strict On
Option Infer Off

Public Class clsFaceZombie

    'Declare
    Private _frmToPass As Form

    'Bitmaps
    Private btmImage As Bitmap
    Private pntPoint As Point

    'Sound
    Private _udcFaceZombieEyesOpenSound As clsSound

    Public Sub New(frmToPass As Form, intSpawnX As Integer, intSpawnY As Integer, udcFaceZombieEyesOpenSound As clsSound)

        'Set
        _frmToPass = frmToPass

        'Set animation
        btmImage = gabtmFaceZombieMemories(0)

        'Set
        pntPoint = New Point(intSpawnX, intSpawnY)

        'Set sound
        _udcFaceZombieEyesOpenSound = udcFaceZombieEyesOpenSound

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

    Public Sub OpenEyes()

        'Change frame
        btmImage = gabtmFaceZombieMemories(1)

        'Play sound
        _udcFaceZombieEyesOpenSound.PlaySound(gintSoundVolume)

    End Sub

End Class
