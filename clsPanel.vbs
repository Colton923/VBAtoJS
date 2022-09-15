Option Explicit
 
Public PanelLength As Double        ''' Fractional Inches
Public PanelMeasurement As String   ''' Formatted Imperial Measurement
Public Quantity As Integer
Public DeleteFlag As Boolean        ''' boolean for removing panel type from collection
Public PanelShape As String
Public PanelType As String
Public clsType As String
Public PanelColor As String
Public TotalCost As Variant
Public UnitCost As Variant
Public FootageCost As Variant
Public SkipFlag As Boolean
Public rEdgePosition As Integer
Public bEdgeHeight As Double


Private Sub Class_Initialize()
    clsType = "Panel"
    FootageCost = "N/A"
End Sub

Public Function lEdgePosition() As Double
    lEdgePosition = rEdgePosition + (3 * 12)
End Function

Public Function tEdgeHeight() As Double
    tEdgeHeight = bEdgeHeight + PanelLength
End Function
