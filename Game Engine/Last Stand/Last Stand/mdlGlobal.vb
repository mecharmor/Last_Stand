'Options
Option Explicit On
Option Strict On
Option Infer Off

Module mdlGlobal

    'Screen width ratio
    Public gdblScreenWidthRatio As Double = 0
    Public gdblScreenHeightRatio As Double = 0

    'Used for sound
    Public gintAlias As Integer = 0
    Public gintSoundVolume As Integer = 1000

    'Character
    Public gbtmCharacterStand(1) As Bitmap
    Public gbtmCharacterShoot(1) As Bitmap
    Public gbtmCharacterReload(21) As Bitmap

    'Zombies
    Public gbtmZombieWalk(3) As Bitmap
    Public gbtmZombieDeath(5) As Bitmap
    Public gbtmZombiePin(1) As Bitmap

End Module