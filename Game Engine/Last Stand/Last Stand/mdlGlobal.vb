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
    Public gbtmCharacterStand1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/Standing/1.png"))
    Public gbtmCharacterStand2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/Standing/2.png"))
    Public gbtmCharacterShoot1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/Shoot Once/1.png"))
    Public gbtmCharacterShoot2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/Shoot Once/2.png"))

    'Zombies
    Public gbtmZombieWalk1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/1.png"))
    Public gbtmZombieWalk2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/2.png"))
    Public gbtmZombieWalk3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/3.png"))
    Public gbtmZombieWalk4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/4.png"))
    Public gbtmZombieDeath1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/1.png"))
    Public gbtmZombieDeath2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/2.png"))
    Public gbtmZombieDeath3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/3.png"))
    Public gbtmZombieDeath4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/4.png"))
    Public gbtmZombieDeath5 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/5.png"))
    Public gbtmZombieDeath6 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/6.png"))

End Module