Attribute VB_Name = "Vars"
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public V As Integer

Public Up As Integer
Public Down As Integer
Public Right As Integer
Public Lef As Integer
Public STP As Integer
Public Space As Integer
Public Lefty(4) As Boolean
Public Righty(4) As Boolean
Public Shift As Integer

Public PlaneW As Long
Public PlaneB As Long
Public RPlaneW As Long
Public RPlaneB As Long
Public LPlaneW As Long
Public LPlaneB As Long
Public BG As Long
Public BarrelB As Long
Public BarrelW As Long
Public BoomB As Long
Public BoomW As Long
Public WBoomB As Long
Public WBoomW As Long
Public GuyB As Long
Public GuyW As Long
Public GraveBW As Long
Public groundW As Long
Public groundB As Long
Public RopeW As Long
Public RopeB As Long
Public WaterB As Long
Public SandB As Long
Public DirtW As Long
Public DirtB As Long
Public TankW As Long
Public TankB As Long

Public GuyClimb(50) As Boolean
Public RopeP As Integer
Public Fall(50) As Boolean
Public Death(50) As Boolean

Public Direction As Integer
Public RSwitch As Boolean
Public LSwitch As Boolean

Public GroundTop As Integer
Public Collide As Boolean

Public PX As Integer
Public slowdown As Integer
Public BGPosition As Integer

Public X As Integer
Public Y As Integer
Public ind As Integer

Public SlowDownBarrel(4) As Integer
Public BarrelX(4) As Single
Public BarrelY(4) As Single
Public BarrelP(4) As Integer
Public BarrelV(4) As Single
Public BarrelMove(4) As Single

Public Bomb(4) As Boolean

Public SlowDownBoom(4) As Integer
Public Boom(4) As Boolean
Public BoomP(4) As Integer
Public BoomX(4) As Single
Public BoomY(4) As Single

Public Guy(50) As Boolean
Public GuyX(50) As Integer
Public GuyY(50) As Single
Public GuyP(50) As Integer
Public GuySlowdown(50) As Integer
Public GuyStart(50) As Integer
Public GuySpeed(50) As Integer

Public GraveX(50) As Integer
Public Grave(50) As Boolean
Public GraveEnd(50) As Integer

Public Shoot As Boolean
Public angle(20) As Single
Public BulletX(20) As Single
Public indb As Integer

Public TalleyX As Integer
Public TalleyY As Integer
Public TalleyCount As Integer
Public TalleyIndex As Integer

Public DirtX(7) As Integer
Public Dirt(7) As Boolean
Public DInd As Integer
Public DirtP(7) As Integer

Public TankX(15) As Integer
Public Tank(15) As Boolean
Public TankSlowDown(15) As Integer
Public TAngleX(50) As Single
Public TAngleY(50) As Single
Public Tind As Integer
Public TAind As Integer
Public TankSpeed(15) As Integer
Public TSChange As Integer

Public temp As Integer
Public Switch As Boolean
Public SSlowdown As Integer
Public DSlowdown As Integer
Public LRCounter As Integer
Public EndDirection As Integer
Public Dead As Boolean
Public TankHit(15) As Integer
