Option Explicit

'''''''''' NOTES:
'''''''''' For the purposes of structural steel, all numeric values are in inches for now. Previous declarations of height and width as integers remain as legacy. New declarations as double probably unnecessary - integers should suffice
Public Height As Integer
Public Width As Integer
Public FOType As String         'Possible FO types for structural steel: "PDoor","OHDoor","Window","MiscFO"
Public rEdgePosition As Double
Public bEdgeHeight As Double
Public Wall As String
Public Description As String
Public FOMaterials As Collection
Public StructuralSteelOption As String



Private Sub Class_Initialize()
'default bottom edge to floor level
bEdgeHeight = 0
'new FO Materials Collection
Set FOMaterials = New Collection
End Sub

Public Function tEdgeHeight() As Double
    tEdgeHeight = bEdgeHeight + Height
End Function

Public Function lEdgePosition() As Double
    lEdgePosition = rEdgePosition + Width
End Function

'Public Sub SetWall(WallInput As String)
'Select Case WallInput
'Case "Endwall 1"
'    Wall = "e1"
'Case "Sidewall 2"
'    Wall = "s2"
'Case "Endwall 3"
'    Wall = "e3"
'Case "Sidewall 4"
'    Wall = "s4"
'End Select
'End Sub


