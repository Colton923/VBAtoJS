Option Explicit

Public bLength As Double    'Ft
Public bHeight As Double    'Ft
Public rPitch As Integer
Public RafterLength As Double   'in
Public s2RafterSheetLength As Double
Public s4RafterSheetLength As Double
Public bWidth As Double 'Ft
Public rShape As String
Public s2Overhang As Double
Public s4Overhang As Double
Public e1Overhang As Double
Public e3Overhang As Double
Public s2Extension As Double
Public s4Extension As Double
Public e1Extension As Double
Public e3Extension As Double
Public e1ExtensionPanelQty As Integer
Public e3ExtensionPanelQty As Integer
Public Gutters As Boolean
Public BaseTrim As Boolean
'endwall wall panel overlaps
Public e1WallPanelOverlaps As Integer
Public e3WallPanelOverlaps As Integer
'extension pitches
Public s2ExtensionPitch As Integer
Public s4ExtensionPitch As Integer
'extension Heights
Public s2ExtensionHeight As Double
Public s4ExtensionHeight As Double
'extension widths
Public s2ExtensionWidth As Double
Public s4ExtensionWidth As Double
'Panel Shapes
Public wPanelShape As String    'sidewall panel shapes
Public rPanelShape As String    'roof panel shapes
'Panel Types, Colors
Public rPanelType As String
Public rPanelColor As String
Public wPanelType As String
Public wPanelColor As String
'Trim Colors
Public RakeTrimColor As String
Public OutsideCornerTrimColor As String
'soffit booleans
Public e1GableOverhangSoffit As Boolean
Public e3GableOverhangSoffit As Boolean
Public s2EaveOverhangSoffit As Boolean
Public s4EaveOverhangSoffit As Boolean
Public e1GableExtensionSoffit As Boolean
Public e3GableExtensionSoffit As Boolean
Public s2EaveExtensionSoffit As Boolean
Public s4EaveExtensionSoffit As Boolean
'Var for totaling eave extension length
Private EaveExtLength As Double
'roof panel overage
Public bLengthRoofPanelOverage As Integer
'Interior Columns Collection
Public InteriorColumns As Collection
Public s2ColumnWidth As Double
Public s4ColumnWidth As Double
'Weld Clips
Public WeldClips As Double
'Structural Steel total cost
Public SSTotalCost As Double

''''''''''''''''''''''''''''''''''''''''''''''''' FO Collections
Public e1FOs As New Collection
Public s2FOs As New Collection
Public e3FOs As New Collection
Public s4FOs As New Collection
Public fieldlocateFOs As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''' Column Collections
Public e1Columns As New Collection
Public s2Columns As New Collection
Public e3Columns As New Collection
Public s4Columns As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''' Girt Collections
Public e1Girts As New Collection
Public s2Girts As New Collection
Public e3Girts As New Collection
Public s4Girts As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''' Rafter Collections
Public e1Rafters As New Collection
Public intRafters As New Collection
Public e3Rafters As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''' Roof Purlin Collection
Public RoofPurlins As New Collection

''''''''''''''''''''''''''''''''''''''''''''''''' Overhang Members
Public e1OverhangMembers As New Collection
Public s2OverhangMembers As New Collection
Public e3OverhangMembers As New Collection
Public s4OverhangMembers As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''' Extension Members
Public e1ExtensionMembers As New Collection
Public s2ExtensionMembers As New Collection
Public e3ExtensionMembers As New Collection
Public s4ExtensionMembers As New Collection

''''''''''''''''''''''''''''''''''''''''''''''''' Base Angle Trim
Public BaseAngleTrim As New Collection

''''''''''''''''''''''''''''''''''''''''''''''''' Weld Plates
Public WeldPlates As New Collection



Public Function RoofLength() As Double
    RoofLength = (bLength * 12) + e1Overhang + e1Extension + e3Overhang + e3Extension
End Function

Public Function RoofFtLength() As Double
    RoofFtLength = ((bLength * 12) + e1Overhang + e1Extension + e3Overhang + e3Extension) / 12

End Function
'high side eave height
Public Function HighSideEaveHeight() As Double  'Inches
    HighSideEaveHeight = (bHeight * 12) + (bWidth * rPitch)
End Function
Public Function s2ExtensionRafterLength() As Double
    If s2Extension = 0 Then
        s2ExtensionRafterLength = 0
    Else
        s2ExtensionRafterLength = (s2Extension / 12) * Sqr((12 ^ 2) + (s2ExtensionPitch ^ 2))
    End If
End Function

Public Function s4ExtensionRafterLength() As Double
    If s4Extension = 0 Then
        s4ExtensionRafterLength = 0
    Else
        s4ExtensionRafterLength = (s4Extension / 12) * Sqr((12 ^ 2) + (s4ExtensionPitch ^ 2))
    End If
End Function

'''''''''''''''''''''''''''''''' Extension Intersections '''''''''''''''''''''''
'Note: Intersecting extension panels are accounted for as eave extension panels
Public Function s2e1ExtensionIntersection() As Boolean
    Select Case EstSht.Range("s2e1_Intersection").Value
    Case "N/A", "Exclude"
        s2e1ExtensionIntersection = False
    Case "Include"
        s2e1ExtensionIntersection = True
    End Select
End Function

Public Function s2e3ExtensionIntersection() As Boolean
    Select Case EstSht.Range("s2e3_Intersection").Value
    Case "N/A", "Exclude"
        s2e3ExtensionIntersection = False
    Case "Include"
        s2e3ExtensionIntersection = True
    End Select
End Function
Public Function s4e1ExtensionIntersection() As Boolean
    Select Case EstSht.Range("s4e1_Intersection").Value
    Case "N/A", "Exclude"
        s4e1ExtensionIntersection = False
    Case "Include"
        s4e1ExtensionIntersection = True
    End Select
End Function

Public Function s4e3ExtensionIntersection() As Boolean
    Select Case EstSht.Range("s4e3_Intersection").Value
    Case "N/A", "Exclude"
        s4e3ExtensionIntersection = False
    Case "Include"
        s4e3ExtensionIntersection = True
    End Select
End Function


''''''''''''''''''''''''''''''' Eave Extension Lengths (from endwall to endwall)
Public Function s2EaveExtensionBuildingLength() As Integer
    EaveExtLength = (bLength * 12) + e1Overhang + e3Overhang
    If s2e1ExtensionIntersection = True Then EaveExtLength = EaveExtLength + e1Extension
    If s2e3ExtensionIntersection = True Then EaveExtLength = EaveExtLength + e3Extension
    s2EaveExtensionBuildingLength = EaveExtLength
End Function

Public Function s4EaveExtensionBuildingLength() As Integer
    EaveExtLength = (bLength * 12) + e1Overhang + e3Overhang
    If s4e1ExtensionIntersection = True Then EaveExtLength = EaveExtLength + e1Extension
    If s4e3ExtensionIntersection = True Then EaveExtLength = EaveExtLength + e3Extension
    s4EaveExtensionBuildingLength = EaveExtLength
End Function

Public Function NetSingleRoofPanelQty() As Integer
    NetSingleRoofPanelQty = Application.WorksheetFunction.RoundUp((((bLength * 12) + e1Overhang + e3Overhang + e1Extension + e3Extension) / 12) / 3, 0)
End Function


'Wall Exclusions
Public Function WallStatus(Wall As String) As String
WallStatus = EstSht.Range(Wall & "_WallStatus").Value
End Function
'Partial Walls' Length Above Finished Floor
Public Function LengthAboveFinishedFloor(Wall As String) As Integer             ' Ft

If EstSht.Range(Wall & "_WallStatus").Value = "Include" Then
    LengthAboveFinishedFloor = 0
ElseIf EstSht.Range(Wall & "_WallStatus").Value = "Partial" Then
    LengthAboveFinishedFloor = EstSht.Range(Wall & "_WallStatus").offset(0, 2).Value
ElseIf EstSht.Range(Wall & "_WallStatus").Value = "Gable Only" Then
    LengthAboveFinishedFloor = bHeight
End If

End Function

'Liner Panel Options
Public Function LinerPanels(Location As String) As String
LinerPanels = EstSht.Range(Location & "_LinerPanels").Value
If LinerPanels = "" Then LinerPanels = "None"
End Function

'Wainscot
Public Function Wainscot(Wall As String) As String
Wainscot = EstSht.Range(Wall & "_Wainscot").Value
If Wainscot = "" Then Wainscot = "None"
End Function

'expandable endwall
Public Function ExpandableEndwall(eWall As String) As Boolean
If EstSht.Range(eWall & "_Expandable").Value <> "Yes" Then
    ExpandableEndwall = False
Else
    ExpandableEndwall = True
End If
End Function
'function for height to the very top of the building (that is, the top surface, not the bottom of the rafter) at a given horizontal distance
'SHOULD ONLY BE CALLED AFTER INT COLUMNS ARE GENERATED
Public Function DistanceToRoof(Wall As String, DistanceFromRightCorner As Double, Optional StartingHeight As Double)
Dim DistanceFromCenter As Double


'ActualPitch = (((bWidth * (rPitch / 12))) / (bWidth - ((s2ColumnWidth + s4ColumnWidth) / 12))) * 12
If rShape = "Gable" Then
    Select Case Wall
    Case "s2", "s4"
        DistanceToRoof = (bHeight * 12) - StartingHeight
    Case "e1"
        'less than halfway to peak
        If (DistanceFromRightCorner / 12) <= (bWidth / 2) Then
            'DistanceToRoof = (((DistanceFromRightCorner - s4ColumnWidth) / 12) * ActualPitch) + bHeight * 12 - StartingHeight
            DistanceToRoof = (((DistanceFromRightCorner / 12)) * rPitch) + (bHeight * 12) - StartingHeight
        'past peak
        ElseIf (DistanceFromRightCorner / 12) > (bWidth / 2) Then
            DistanceFromCenter = DistanceFromRightCorner - (bWidth / 2) * 12
            'DistanceToRoof = ((bHeight * 12 + (((bWidth - s2ColumnWidth / 12) / 2) * ActualPitch)) - ((DistanceFromCenter / 12) * rPitch)) - StartingHeight
            DistanceToRoof = (((bWidth) - ((DistanceFromRightCorner) / 12)) * rPitch) + (bHeight * 12) - StartingHeight
        End If
    Case "e3"
        'less than halfway to peak
        If (DistanceFromRightCorner / 12) <= (bWidth / 2) Then
            'DistanceToRoof = (((DistanceFromRightCorner - s2ColumnWidth) / 12) * ActualPitch) + bHeight * 12 - StartingHeight
            DistanceToRoof = (((DistanceFromRightCorner / 12)) * rPitch) + (bHeight * 12) - StartingHeight
        'past peak
        ElseIf (DistanceFromRightCorner / 12) > (bWidth / 2) Then
            DistanceFromCenter = DistanceFromRightCorner - (bWidth / 2) * 12
            'DistanceToRoof = ((bHeight * 12 + (((bWidth - s4ColumnWidth / 12) / 2) * ActualPitch)) - ((DistanceFromCenter / 12) * ActualPitch)) - StartingHeight
            DistanceToRoof = (((bWidth) - ((DistanceFromRightCorner) / 12)) * rPitch) + (bHeight * 12) - StartingHeight
        End If
    End Select
ElseIf rShape = "Single Slope" Then
    Select Case Wall
    Case "e1"
        'Inside Distance - Distance from s4 Column = Actual Distance of slope
        'Distance of Slope * rPitch = Height above eave height
        'Distance above eavh height + eave height = distance to roof
        DistanceToRoof = (((bWidth) - ((DistanceFromRightCorner) / 12)) * rPitch) + (bHeight * 12) - StartingHeight
    Case "s2"
        DistanceToRoof = (bHeight * 12) - StartingHeight
    Case "e3"
        'CL - Inside of s2 Column = Actual Distance of Slope
        'Distance of Slope * rPitch = Height above eave height
        'Distance above eavh height + eave height = distance to roof
        DistanceToRoof = (((DistanceFromRightCorner / 12)) * rPitch) + (bHeight * 12) - StartingHeight
    Case "s4"
        DistanceToRoof = bHeight * 12 + (bWidth * rPitch) - StartingHeight
    End Select
End If

End Function

'function for distance from right corner of an endwall at a given height
Public Function DistanceFromCorner(Wall As String, HeightAlongRoof As Double)
Dim DistanceFromCenter As Double
If rShape = "Gable" Then
    If HeightAlongRoof < bWidth * 12 / 2 Then
        If Wall = "e1" Then
            DistanceFromCorner = (((HeightAlongRoof - bHeight * 12) / rPitch) * 12)
        Else
            DistanceFromCorner = ((HeightAlongRoof - bHeight * 12) / rPitch) * 12
        End If
    Else 'right now these are the same
        If Wall = "e3" Then
            DistanceFromCorner = ((HeightAlongRoof - bHeight * 12) / rPitch) * 12
        Else
            DistanceFromCorner = ((HeightAlongRoof - bHeight * 12) / rPitch) * 12
        End If
    End If
ElseIf rShape = "Single Slope" Then
    If Wall = "e1" Then '0 is the tallest point
        DistanceFromCorner = bWidth * 12 - ((HeightAlongRoof - bHeight * 12) / rPitch) * 12
    Else 'for e3 0 is the lowest point
        DistanceFromCorner = ((HeightAlongRoof - bHeight * 12) / rPitch) * 12
    End If
End If

End Function

Private Sub Class_Initialize()
Dim FOCell As Range
Dim FO As clsFO
Dim BayCell As Range
Dim TotalBayLength As Double
Dim Column As clsMember
Dim Bay As Integer

'set basic building parameters
With EstSht
bHeight = .Range("Building_Height").Value
bWidth = .Range("Building_Width").Value
bLength = .Range("Building_Length").Value
rPitch = .Range("Roof_Pitch").Value
rShape = .Range("Roof_Shape").Value

'create Int Columns collection
Set InteriorColumns = New Collection

'create girt collections to be filled
Set e1Girts = New Collection
Set s2Girts = New Collection
Set e3Girts = New Collection
Set s4Girts = New Collection

'create rafter collections to be filled
Set e1Rafters = New Collection
Set intRafters = New Collection
Set e3Rafters = New Collection

'create overhang and extension members collections to be filled
Set e1OverhangMembers = New Collection
Set s2OverhangMembers = New Collection
Set e3OverhangMembers = New Collection
Set s4OverhangMembers = New Collection
Set e1ExtensionMembers = New Collection
Set s2ExtensionMembers = New Collection
Set e3ExtensionMembers = New Collection
Set s4ExtensionMembers = New Collection

'create roof purlin collection
Set RoofPurlins = New Collection

'create Weld Plate Collection
Set WeldPlates = New Collection



'''''''''''''set extension pitches
If .Range("s2_EaveExtension").Value > 0 Then
    If .Range("s2_EaveExtensionPitch").Value = "Match Roof" Then
        s2ExtensionPitch = rPitch
    Else
        s2ExtensionPitch = .Range("s2_EaveExtensionPitch").Value
    End If
End If
If .Range("s4_EaveExtension").Value > 0 Then
    If .Range("s4_EaveExtensionPitch").Value = "Match Roof" Then
        s4ExtensionPitch = rPitch
    Else
        s4ExtensionPitch = .Range("s4_EaveExtensionPitch").Value
    End If
End If

'''''''''''' generate sidewall 2 column centerlines
If .Range("BayNum").Value > 1 Then
    ''''s2 columns
    For Each BayCell In Range(.Range("Bay1_Length"), .Range("Bay12_Length"))
        If BayCell.EntireRow.Hidden = False And BayCell.Value <> 0 Then
            TotalBayLength = TotalBayLength + BayCell.Value
            If TotalBayLength = bLength Then Exit For
            'new column
            Set Column = New clsMember
            Column.CL = TotalBayLength * 12
            'add column length (building height)
            Column.Length = bHeight * 12
            Column.tEdgeHeight = Column.Length
            'add to collection
            s2Columns.Add Column
        End If
    Next BayCell
    ''''s4 columns
    TotalBayLength = 0
    For Bay = 12 To 1 Step -1
        Set BayCell = .Range("Bay" & Bay & "_Length")
        If BayCell.EntireRow.Hidden = False And BayCell.Value <> 0 Then
            TotalBayLength = TotalBayLength + BayCell.Value
            If TotalBayLength = bLength Then Exit For
            'new column
            Set Column = New clsMember
            Column.CL = TotalBayLength * 12
            'add column height (building height)
            If rShape = "Gable" Then Column.Length = bHeight * 12
            If rShape = "Single Slope" Then Column.Length = HighSideEaveHeight
            'add to collection
            Column.tEdgeHeight = Column.Length
            s4Columns.Add Column
        End If
    Next Bay
End If



''''''''''''''''''''''''''''''''''''''''''''''''' Build FO Collections  '''''''''''''''''''''''''''''''''''''''''
'pDoors
For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
        Set FO = New clsFO
        FO.FOType = "PDoor"
        FO.Height = 7 * 12
        'set width
        If FOCell.offset(0, 1).Value = "3070" Then
            FO.Width = (3 * 12)
        ElseIf FOCell.offset(0, 1).Value = "4070" Then
            FO.Width = (4 * 12)
        End If
        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
        If FOCell.offset(0, 2).Value = "Endwall 1" Or FOCell.offset(0, 2).Value = "Endwall 3" Then
            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 8) * 12
        Else
            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 8).Value * 12
        End If
        FO.Description = "pDoor #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 2).Value & ". rEdge: " & _
        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "'"
        'set wall, add to collection
        Select Case FOCell.offset(0, 2).Value
        Case "Endwall 1"
            FO.Wall = "e1"
            e1FOs.Add FO
        Case "Sidewall 2"
            FO.Wall = "s2"
            s2FOs.Add FO
        Case "Endwall 3"
            FO.Wall = "e3"
            e3FOs.Add FO
        Case "Sidewall 4"
            FO.Wall = "s4"
            s4FOs.Add FO
        Case "Field Locate"
            FO.Wall = "Field Locate"
            fieldlocateFOs.Add FO
        End Select

    End If
Next FOCell
'OHDoors
For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
    'if cell isn't hidden, door size is entered
    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
        'new FO class
        Set FO = New clsFO
        FO.FOType = "OHDoor"
        FO.Width = FOCell.offset(0, 1).Value * 12
        FO.Height = FOCell.offset(0, 2).Value * 12
        FO.bEdgeHeight = 0
        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 10) * 12
        Else
            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 10).Value * 12
        End If
        FO.Description = "OHDoor #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "' , Height: " & FO.Height / 12 & "'"
        'set wall, add to collection
        Select Case FOCell.offset(0, 3).Value
        Case "Endwall 1"
            FO.Wall = "e1"
            e1FOs.Add FO
        Case "Sidewall 2"
            FO.Wall = "s2"
            s2FOs.Add FO
        Case "Endwall 3"
            FO.Wall = "e3"
            e3FOs.Add FO
        Case "Sidewall 4"
            FO.Wall = "s4"
            s4FOs.Add FO
        End Select
    End If
Next FOCell
'Windows
For Each FOCell In Range(.Range("WindowCell1"), .Range("WindowCell12"))
    'if cell isn't hidden, door size is entered
    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
        'new FO class
        Set FO = New clsFO
        FO.FOType = "Window"
        FO.Width = FOCell.offset(0, 1).Value
        FO.Height = FOCell.offset(0, 2).Value
        FO.bEdgeHeight = FOCell.offset(0, 4).Value * 12
        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 7) * 12
        Else
            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 7).Value * 12
        End If
        FO.Description = "Window #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "', bEdge:" & FO.bEdgeHeight / 12 & "', Height: " & FO.Height / 12 & "'"
        'set wall, add to collection
        Select Case FOCell.offset(0, 3).Value
        Case "Endwall 1"
            FO.Wall = "e1"
            e1FOs.Add FO
        Case "Sidewall 2"
            FO.Wall = "s2"
            s2FOs.Add FO
        Case "Endwall 3"
            FO.Wall = "e3"
            e3FOs.Add FO
        Case "Sidewall 4"
            FO.Wall = "s4"
            s4FOs.Add FO
        Case "Field Locate"
            FO.Wall = "Field Locate"
            fieldlocateFOs.Add FO
        End Select
    End If
Next FOCell
'Misc FOs
For Each FOCell In Range(.Range("MiscFOCell1"), .Range("MiscFOCell12"))
    'if cell isn't hidden, door size is entered
    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
        'new FO class
        Set FO = New clsFO
        FO.FOType = "MiscFO"
        FO.Width = FOCell.offset(0, 1).Value * 12
        FO.Height = FOCell.offset(0, 2).Value * 12
        FO.bEdgeHeight = FOCell.offset(0, 6).Value * 12
        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 9) * 12
        Else
            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 9).Value * 12
        End If
        FO.Description = "MiscFO #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "', bEdge:" & FO.bEdgeHeight / 12 & "', Height: " & FO.Height / 12 & "'"
        'add structural steel framing selection
        FO.StructuralSteelOption = FOCell.offset(0, 10).Value
        'set wall, add to collection
        Select Case FOCell.offset(0, 3).Value
        Case "Endwall 1"
            FO.Wall = "e1"
            e1FOs.Add FO
        Case "Sidewall 2"
            FO.Wall = "s2"
            s2FOs.Add FO
        Case "Endwall 3"
            FO.Wall = "e3"
            e3FOs.Add FO
        Case "Sidewall 4"
            FO.Wall = "s4"
            s4FOs.Add FO
        End Select
    End If
Next FOCell
End With


End Sub

Private Sub Class_Terminate()

End Sub
