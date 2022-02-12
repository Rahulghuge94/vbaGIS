VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaGIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Public Function Distance(point1() As Double, point2() As Double) As Double
'Function calculates distance between two points..
Dim xdif, ydif As Double
xdif = point2(0) - point1(0)
ydif = point2(1) - point1(1)
Distance = Sqr((xdif * xdif) + (ydif * ydif))
End Function
Public Function Atn2(y As Double, x As Double) As Double
Private Const Pi As Double = 3.14159265358979
  If x > 0 Then
    Atn2 = Atn(y / x)
  ElseIf x < 0 Then
    Atn2 = Sgn(y) * (Pi - Atn(Abs(y / x)))
  ElseIf y = 0 Then
    Atn2 = 0
  Else
    Atn2 = Sgn(y) * Pi / 2
  End If
End Function
Public Function getCoords2d(coordinat As Variant)
'Function returns coordinates of polylines,lightweight polylines in array of 2d coordinates..
    
Dim coords() As Double
arrcount = UBound(coordinat) + 1
coordsLength = (arrcount / 2) - 1
'Debug.Print coordsLength, arrcount
ReDim coords(coordsLength, 1)
counter = 0
For i = 1 To UBound(coordinat) - 1
    If i Mod 2 <> 0 Then
        coords(counter, 0) = coordinat(i - 1)
        coords(counter, 1) = coordinat(i)
        counter = counter + 1
    End If
Next
'Debug.Print coords(0, 0), coords(0, 1), coords(1, 0), coords(1, 1)
getCoords2d = coords
End Function

'function to check if point is right or left of infinite line. where p0 and p1 are start and end point of line and p2 is point.
Public Function isLeft(P0() As Double, P1() As Double, P2() As Double):
isLeft = ((P1(0) - P0(0)) * (P2(1) - P0(1)) - (P2(0) - P0(0)) * (P1(1) - P0(1)))
End Function

Public Function ArrayResize2d(Arr() As Double, x As Integer, y As Integer)
Dim Ar() As Double
ReDim Ar(x, y) As Double
For i = 0 To UBound(Arr)
For j = 0 To UBound(Arr, 2)
Ar(i, j) = Arr(i, j)
Next
Next
ArrayResize2d = Ar
End Function

'Function to create variant array of polyline coordinnates
Public Function varArray(cord As Variant)
Dim cords() As Variant, temp(0 To 1) As Variant
j = 0
ReDim cords(0 To ((UBound(cord) + 1) / 2) - 1)
For i = 0 To UBound(cord)
If i Mod 2 <> 0 Then
cords(j) = Array(cord(i - 1), cord(i))
j = j + 1
End If
Next
varArray = cords
End Function

'function to calculate if point is inside or outside of polygon using winding nummber theorem. p is point and v is array of vertices of polygon.
Public Function pointInsidePolygon(p() As Double, V() As Variant):
Dim P1(0 To 1) As Double, P2(0 To 1) As Double, wn As Integer
n = UBound(V)
wn = 0
i = 0
Do While i < n
    
    If V(i)(1) <= p(1) Then
       If V(i + 1)(1) > p(1) Then
          P1(0) = V(i)(0): P1(1) = V(i)(1): P2(0) = V(i + 1)(0): P2(1) = V(i + 1)(1)
          tmp = isLeft(P1, P2, p)
          If tmp > 0 Then
          wn = wn + 1
          End If
       End If
    Else
        If V(i + 1)(1) <= p(1) Then
           P1(0) = V(i)(0): P1(1) = V(i)(1): P2(0) = V(i + 1)(0): P2(1) = V(i + 1)(1)
           tmp = isLeft(P1, P2, p)
           If tmp < 0 Then
           wn = wn - 1
           End If
        End If
    End If
i = i + 1
Loop
If wn <> 0 Then
pointInsidePolygon = True
Else
pointInsidePolygon = False
End If
End Function

Public Function scaleOb(coords() As Variant, xscale As Double, yscale As Double)
'Scale object wrt xscale and yscale
For i = LBound(coords) To UBound(coords)
coords(i, 0) = xscale * coords(i, 0)
coords(i, 1) = yscale * coords(i, 1)
Next
scaleOb = coords
End Function

