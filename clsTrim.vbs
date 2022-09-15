Option Explicit

Public tMeasurement As String    ''' Trim Measurement String
Public tLength As Integer
Public tType As String              ''' Trim type (rake, short eave, high eave, etc.) field
Public Quantity As Integer
Public DeleteFlag As Boolean        ''' boolean for removing Trim from collection
Public Color As String
Public clsType As String
Public tShape As String
Public TotalCost As Variant
Public UnitCost As Variant
Public FootageCost As Variant

Private Sub Class_Initialize()
    clsType = "Trim"
    FootageCost = "N/A"
    tShape = "N/A"
End Sub
