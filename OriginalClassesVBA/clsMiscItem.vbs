Option Explicit


Public Quantity As Integer
Public Shape As String
Public Name As String
Public Measurement As String
Public Color As String
Public TotalCost As Variant
Public UnitCost As Variant
Public FootageCost As Variant
Public DeleteFlag As Boolean
Public clsType As String
' Used for OH Doors:
Public Width As Integer  'ft
Public Height As Integer  'ft
' Used for Windows:
Public Area As Integer  'SF

Private Sub Class_Initialize()
    clsType = "MiscItem"
    'default color and shape to "N/A"
    Color = "N/A"
    Shape = "N/A"
    FootageCost = "N/A"
    'measurement to N/A
    Measurement = "N/A"
End Sub
