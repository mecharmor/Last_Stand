'Options
Option Explicit On
Option Strict On
Option Infer Off

Module GlobalModule

    'Screen width ratio
    Public gdblScreenWidthRatio As Double = 0
    Public gdblScreenHeightRatio As Double = 0

    'Character
    Public gbtmCharacterStand1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/1.png"), 643, 710)
    Public gbtmCharacterStand2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/2.png"), 643, 710)
    Public gbtmCharacterShoot1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/3.png"), 643, 710)
    Public gbtmCharacterShoot2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Character/4.png"), 643, 710)

    'Zombies
    Public gbtmZombieWalk1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/1.png"), 765, 752)
    Public gbtmZombieWalk2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/2.png"), 765, 752)
    Public gbtmZombieWalk3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/3.png"), 765, 752)
    Public gbtmZombieWalk4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Movement/4.png"), 765, 752)
    Public gbtmZombieDeath1 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/1.png"), 765, 752)
    Public gbtmZombieDeath2 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/2.png"), 765, 752)
    Public gbtmZombieDeath3 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/3.png"), 765, 752)
    Public gbtmZombieDeath4 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/4.png"), 765, 752)
    Public gbtmZombieDeath5 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/5.png"), 765, 752)
    Public gbtmZombieDeath6 As New Bitmap(Image.FromFile(AppDomain.CurrentDomain.BaseDirectory & "Images/Zombies/Generic/Death/6.png"), 765, 752)

End Module
