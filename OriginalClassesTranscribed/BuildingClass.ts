class Building {
	bLength: number
	bHeight: number
	rPitch: number
	RafterLength: number
	s2RafterSheetLength: number
	s4RafterSheetLength: number
	bWidth: number
	rShape: string
	s2Overhang: number
	s4Overhang: number
	e1Overhang: number
	e3Overhang: number
	s2Extension: number
	s4Extension: number
	e1Extension: number
	e3Extension: number
	e1ExtensionPanelQty: number
	e3ExtensionPanelQty: number
	Gutters: boolean
	BaseTri: boolean
	//endwall wall panel overlaps
	e1WallPanelOverlaps: number
	e3WallPanelOverlaps: number
	//extension pitches
	s2ExtensionPitch: number
	s4ExtensionPitch: number
	//extension Heights
	s2ExtensionHeight: number
	s4ExtensionHeight: number
	//extension widths
	s2ExtensionWidth: number
	s4ExtensionWidth: number
	//Panel Shapes
	wPanelShape: string //sidewall panel shapes
	rPanelShape: string //roof panel shapes
	//Panel Types, Colors
	rPanelType: string
	rPanelColo: string
	wPanelType: string
	wPanelColo: string
	//Trim Colors
	RakeTrimColor: string
	OutsideCorner: string
	//soffit booleans
	e1GableOverhangSoffit: boolean
	e3GableOverhangSoffit: boolean
	s2EaveOverhangSoffit: boolean
	s4EaveOverhangSoffit: boolean
	e1GableExtensionSoffi: boolean
	e3GableExtensionSoffi: boolean
	s2EaveExtensionSoffit: boolean
	s4EaveExtensionSoffit: boolean
	// this for totaling eave extension string
	EaveExtLength: number
	//roof panel overage
	bLengthRoofPanelOverage: number
	//Interior Columns Collection
	InteriorColumns: any[]
	s2ColumnWidth: number
	s4ColumnWidth: number
	//'Weld Clips
	WeldClips: number
	//'Structural Steel total cost
	SSTotalCost: number

	//''''''''''''''''''''''''''''''''''''''''''''''''' FO Collections
	e1FOs: any[]
	s2FOs: any[]
	e3FOs: any[]
	s4FOs: any[]
	fieldlocateFOs: any[]
	//''''''''''''''''''''''''''''''''''''''''''''''''' Column Collections
	e1Columns: any[]
	s2Columns: any[]
	e3Columns: any[]
	s4Columns: any[]
	//''''''''''''''''''''''''''''''''''''''''''''''''' Girt Collections
	e1Girts: any[]
	s2Girts: any[]
	e3Girts: any[]
	s4Girts: any[]
	//''''''''''''''''''''''''''''''''''''''''''''''''' Rafter Collections
	e1Rafters: any[]
	intRafters: any[]
	e3Rafters: any[]
	//''''''''''''''''''''''''''''''''''''''''''''''''' Roof Purlin Collection
	RoofPurlins = []

	//''''''''''''''''''''''''''''''''''''''''''''''''' Overhang Members
	e1OverhangMembers = []
	s2OverhangMembers = []
	e3OverhangMembers = []
	s4OverhangMembers = []
	//''''''''''''''''''''''''''''''''''''''''''''''''' Extension Members
	e1ExtensionMembers = []
	s2ExtensionMembers = []
	e3ExtensionMembers = []
	s4ExtensionMembers = []

	//''''''''''''''''''''''''''''''''''''''''''''''''' Base Angle Trim
	BaseAngleTrim = []

	//''''''''''''''''''''''''''''''''''''''''''''''''' Weld Plates
	WeldPlates = []

	constructor(
		formBHeight: number,
		formBWidth: number,
		formBLength: number,
		formRPitch: number,
		formRShape: number
	) {
		//Set Extension Pitches
		/*
		 *If .Range("s2_EaveExtension").Value > 0 Then
		 *    If .Range("s2_EaveExtensionPitch").Value = "Match Roof" Then
		 *        s2ExtensionPitch = rPitch
		 *    Else
		 *        s2ExtensionPitch = .Range("s2_EaveExtensionPitch").Value
		 *    End If
		 *End If
		 *If .Range("s4_EaveExtension").Value > 0 Then
		 *    If .Range("s4_EaveExtensionPitch").Value = "Match Roof" Then
		 *        s4ExtensionPitch = rPitch
		 *    Else
		 *        s4ExtensionPitch = .Range("s4_EaveExtensionPitch").Value
		 *    End If
		 *End If
		 */
		//Generate sidewall 2 column centerlines
		/*
		 *If .Range("BayNum").Value > 1 Then
		 *        //s2 columns
		 *    For Each BayCell In Range(.Range("Bay1_Length"), .Range("Bay12_Length"))
		 *        If BayCell.EntireRow.Hidden = False And BayCell.Value <> 0 Then
		 *            TotalBayLength = TotalBayLength + BayCell.Value
		 *            If TotalBayLength = bLength Then Exit For
		 *            //new column
		 *            Set Column = New clsMember
		 *            Column.CL = TotalBayLength * 12
		 *            //add column length (building height)
		 *            Column.Length = bHeight * 12
		 *            Column.tEdgeHeight = Column.Length
		 *            //add to collection
		 *            s2Columns.Add Column
		 *        End If
		 *    Next BayCell
		 *        //s4 columns
		 *    TotalBayLength = 0
		 *    For Bay = 12 To 1 Step -1
		 *        Set BayCell = .Range("Bay" & Bay & "_Length")
		 *        If BayCell.EntireRow.Hidden = False And BayCell.Value <> 0 Then
		 *            TotalBayLength = TotalBayLength + BayCell.Value
		 *            If TotalBayLength = bLength Then Exit For
		 *            //new column
		 *            Set Column = New clsMember
		 *            Column.CL = TotalBayLength * 12
		 *            //add column height (building height)
		 *            If rShape = "Gable" Then Column.Length = bHeight * 12
		 *            If rShape = "Single Slope" Then Column.Length = HighSideEaveHeight
		 *            //add to collection
		 *            Column.tEdgeHeight = Column.Length
		 *            s4Columns.Add Column
		 *        End If
		 *    Next Bay
		 *End If
		 */
		//Build FO Collections
		//pDoors
		/*
		 *For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
		 *    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
		 *        Set FO = New clsFO
		 *        FO.FOType = "PDoor"
		 *        FO.Height = 7 * 12
		 *        //set width
		 *        If FOCell.offset(0, 1).Value = "3070" Then
		 *            FO.Width = (3 * 12)
		 *        ElseIf FOCell.offset(0, 1).Value = "4070" Then
		 *            FO.Width = (4 * 12)
		 *        End If
		 *        //reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
		 *        If FOCell.offset(0, 2).Value = "Endwall 1" Or FOCell.offset(0, 2).Value = "Endwall 3" Then
		 *            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 8) * 12
		 *        Else
		 *            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 8).Value * 12
		 *        End If
		 *        FO.Description = "pDoor #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 2).Value & ". rEdge: " & _
		 *        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "'"
		 *        //set wall, add to collection
		 *        Select Case FOCell.offset(0, 2).Value
		 *        Case "Endwall 1"
		 *            FO.Wall = "e1"
		 *            e1FOs.Add FO
		 *        Case "Sidewall 2"
		 *            FO.Wall = "s2"
		 *            s2FOs.Add FO
		 *        Case "Endwall 3"
		 *            FO.Wall = "e3"
		 *            e3FOs.Add FO
		 *        Case "Sidewall 4"
		 *            FO.Wall = "s4"
		 *            s4FOs.Add FO
		 *        Case "Field Locate"
		 *            FO.Wall = "Field Locate"
		 *            fieldlocateFOs.Add FO
		 *        End Select
		 *
		 *    End If
		 *Next FOCell
		 *
		 * //OHDoors
		 *For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
		 *    //if cell isn't hidden, door size is entered
		 *    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
		 *        //new FO class
		 *        Set FO = New clsFO
		 *        FO.FOType = "OHDoor"
		 *        FO.Width = FOCell.offset(0, 1).Value * 12
		 *        FO.Height = FOCell.offset(0, 2).Value * 12
		 *        FO.bEdgeHeight = 0
		 *        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
		 *        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
		 *            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 10) * 12
		 *        Else
		 *            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 10).Value * 12
		 *        End If
		 *        FO.Description = "OHDoor #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
		 *        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "' , Height: " & FO.Height / 12 & "'"
		 *        //set wall, add to collection
		 *        Select Case FOCell.offset(0, 3).Value
		 *        Case "Endwall 1"
		 *            FO.Wall = "e1"
		 *            e1FOs.Add FO
		 *        Case "Sidewall 2"
		 *            FO.Wall = "s2"
		 *            s2FOs.Add FO
		 *        Case "Endwall 3"
		 *            FO.Wall = "e3"
		 *            e3FOs.Add FO
		 *        Case "Sidewall 4"
		 *            FO.Wall = "s4"
		 *            s4FOs.Add FO
		 *        End Select
		 *    End If
		 *Next FOCell
		 *
		 * //Windows
		 *For Each FOCell In Range(.Range("WindowCell1"), .Range("WindowCell12"))
		 *    //if cell isn't hidden, door size is entered
		 *    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
		 *        'new FO class
		 *        Set FO = New clsFO
		 *        FO.FOType = "Window"
		 *        FO.Width = FOCell.offset(0, 1).Value
		 *        FO.Height = FOCell.offset(0, 2).Value
		 *        FO.bEdgeHeight = FOCell.offset(0, 4).Value * 12
		 *        //reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
		 *        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
		 *            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 7) * 12
		 *        Else
		 *            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 7).Value * 12
		 *        End If
		 *        FO.Description = "Window #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
		 *        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "', bEdge:" & FO.bEdgeHeight / 12 & "', Height: " & FO.Height / 12 & "'"
		 *        //set wall, add to collection
		 *        Select Case FOCell.offset(0, 3).Value
		 *        Case "Endwall 1"
		 *            FO.Wall = "e1"
		 *            e1FOs.Add FO
		 *        Case "Sidewall 2"
		 *            FO.Wall = "s2"
		 *            s2FOs.Add FO
		 *        Case "Endwall 3"
		 *            FO.Wall = "e3"
		 *            e3FOs.Add FO
		 *        Case "Sidewall 4"
		 *            FO.Wall = "s4"
		 *            s4FOs.Add FO
		 *        Case "Field Locate"
		 *            FO.Wall = "Field Locate"
		 *            fieldlocateFOs.Add FO
		 *        End Select
		 *    End If
		 *Next FOCell
		 *
		 * //Misc FOs
		 *For Each FOCell In Range(.Range("MiscFOCell1"), .Range("MiscFOCell12"))
		 *    //if cell isn't hidden, door size is entered
		 *    If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
		 *        //new FO class
		 *        Set FO = New clsFO
		 *        FO.FOType = "MiscFO"
		 *        FO.Width = FOCell.offset(0, 1).Value * 12
		 *        FO.Height = FOCell.offset(0, 2).Value * 12
		 *        FO.bEdgeHeight = FOCell.offset(0, 6).Value * 12
		 *        'reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
		 *        If FOCell.offset(0, 3).Value = "Endwall 1" Or FOCell.offset(0, 3).Value = "Endwall 3" Then
		 *            FO.rEdgePosition = bWidth * 12 - FOCell.offset(0, 9) * 12
		 *        Else
		 *            FO.rEdgePosition = bLength * 12 - FOCell.offset(0, 9).Value * 12
		 *        End If
		 *        FO.Description = "MiscFO #" & FOCell.Value & ", an FO located on " & FOCell.offset(0, 3).Value & ". rEdge: " & _
		 *        (FO.rEdgePosition / 12) & "', lEdge: " & FO.lEdgePosition / 12 & "', bEdge:" & FO.bEdgeHeight / 12 & "', Height: " & FO.Height / 12 & "'"
		 *        //add structural steel framing selection
		 *        FO.StructuralSteelOption = FOCell.offset(0, 10).Value
		 *        //set wall, add to collection
		 *        Select Case FOCell.offset(0, 3).Value
		 *        Case "Endwall 1"
		 *            FO.Wall = "e1"
		 *            e1FOs.Add FO
		 *        Case "Sidewall 2"
		 *            FO.Wall = "s2"
		 *            s2FOs.Add FO
		 *        Case "Endwall 3"
		 *            FO.Wall = "e3"
		 *            e3FOs.Add FO
		 *        Case "Sidewall 4"
		 *            FO.Wall = "s4"
		 *            s4FOs.Add FO
		 *        End Select
		 *    End If
		 *Next FOCell
		 *End With
		 */
	}

	RoofLength() {
		return (
			this.bLength * 12 + this.e1Overhang + this.e1Extension + this.e3Overhang + this.e3Extension
		)
	}
	RoofFtLength() {
		return (
			(this.bLength * 12 +
				this.e1Overhang +
				this.e1Extension +
				this.e3Overhang +
				this.e3Extension) /
			12
		)
	}
	HighSideEaveHeight() {
		return this.bHeight * 12 + this.bWidth * this.rPitch
	}
	s2ExtensionRafterLength() {
		if (this.s2Extension === 0) {
			return 0
		} else {
			return (this.s2Extension / 12) * Math.sqrt(144 + this.s4ExtensionPitch)
		}
	}
	s4ExtensionRafterLength() {
		if (this.s4Extension === 0) {
			return 0
		} else {
			return (this.s4Extension / 12) * Math.sqrt((12 ^ 2) + (this.s4ExtensionPitch ^ 2))
		}
	}
	/*
	 * Extension Intersections
	 * Note: Intersecting extension panels are accounted for as eave extension panels
	 *
	 */
	s2e1ExtensionIntersection() {
		//If the input box for s2e1_Intersection = "N/A" or "Exclude" then
		//return false
		//If '' = "Include" then
		//return true
		//Input box from Estimation Sheet Range Key "s2e1_Intersection"
	}
	s2e3ExtensionIntersection() {
		//If the input box for s2e3_Intersection = "N/A" or "Exclude" then
		//return false
		//If '' = "Include" then
		//return true
		//Input box from Estimation Sheet Range Key "s2e3_Intersection"
	}
	s4e1ExtensionIntersection() {
		//If the input box for s4e1_Intersection = "N/A" or "Exclude" then
		//return false
		//If '' = "Include" then
		//return true
		//Input box from Estimation Sheet Range Key "s4e1_Intersection"
	}
	s4e3ExtensionIntersection() {
		//If the input box for s4e3_Intersection = "N/A" or "Exclude" then
		//return false
		//If '' = "Include" then
		//return true
		//Input box from Estimation Sheet Range Key "s4e3_Intersection"
	}
	/*
	 * Eave Extension Lengths (from endwall to endwall)
	 * The below return errors from this.fn() will be resolved when logic is resolved
	 */
	s2EaveExtensionBuildingLength() {
		this.EaveExtLength = this.bLength * 12 + this.e1Overhang + this.e3Overhang
		if (this.s2e1ExtensionIntersection()) {
			this.EaveExtLength += this.e1Extension
		}
		if (this.s2e3ExtensionIntersection()) {
			this.EaveExtLength += this.e3Extension
		}
		return this.EaveExtLength
	}
	s4EaveExtensionBuildingLength() {
		this.EaveExtLength = this.bLength * 12 + this.e1Overhang + this.e3Overhang
		if (this.s4e1ExtensionIntersection()) {
			this.EaveExtLength += this.e1Extension
		}
		if (this.s4e3ExtensionIntersection()) {
			this.EaveExtLength += this.e3Extension
		}
		return this.EaveExtLength
	}
	NetSingleRoofPanelQty() {
		return Math.round(
			(this.bLength * 12 +
				this.e1Overhang +
				this.e3Overhang +
				this.e1Extension +
				this.e3Extension) /
				12 /
				3
		)
	}
	//Wall Exclusions
	WallStatus(Wall: string) {
		/*
		 *Return the estimating sheet range(Variable("Wall") && "_WallStatus"))
		 *VBA Dev used var passed in to fn to determine which form item's value to use
		 */
	}
	LengthAboveFinishedFloor(Wall: string) {
		/*
		 *If Estimating Sheet Range of the passed in Variable("Wall") && _WallStatus
		 * = "Include" Return 0
		 * = "Partial" Return the value of the range to cells to the right....:(
		 * = "Gable Only" Return bHeight
		 */
	}
	LinerPanels(Location: string) {
		/*
		 *If Estimation Sheet Range Variable("Location") && _LinerPanels
		 * = nothing then return "None"
		 */
	}
	Wainscot(Wall: string) {
		/*
		 *If Estimation Sheet Range Variable("Wall") && Wainscot
		 * = nothing then return "None"
		 */
	}
	// Expandable Endwall
	ExpandableEndwall(eWall: string) {
		/* If Estimation Sheet Range Variable("eWall") && _Expandable
		 * <> "Yes" return False else True
		 */
	}
	/* VBA-Dev's msg here
	 *
	 *Function for height to the very top of the building (that is, the top surface, not the bottom of the rafter) at a given horizontal distance
	 * SHOULD ONLY BE CALLED AFTER INT COLUMNS ARE GENERATED
	 *
	 * Who the fuck knows how that's useful to anyone reading that in the future.
	 */
	DistanceToRoof(Wall: string, DistanceFromRightCorner: number, StartingHeight: number) {
		//Looks like a really stupid way to calculate distance coming up...
		//VBA-Dev Note:
		//ActualPitch = (((bWidth * (rPitch / 12))) / ( ))
		let DistanceFromCenter: number

		if (this.rShape === 'Gable') {
			switch (Wall) {
				case 's2' || 's4':
					return this.bHeight * 12 - StartingHeight
				case 'e1':
					if (DistanceFromRightCorner / 12 <= this.bWidth / 2) {
						return (DistanceFromRightCorner / 12) * this.rPitch + this.bHeight * 12 - StartingHeight
					} else if (DistanceFromRightCorner / 12 > this.bWidth / 2) {
						DistanceFromCenter = DistanceFromRightCorner - (this.bWidth / 2) * 12
						return (
							(this.bWidth - DistanceFromRightCorner / 12) * this.rPitch +
							this.bHeight * 12 -
							StartingHeight
						)
					}
				case 'e3':
					if (DistanceFromRightCorner / 12 <= this.bWidth / 2) {
						return (DistanceFromRightCorner / 12) * this.rPitch + this.bHeight * 12 - StartingHeight
					} else if (DistanceFromRightCorner / 12 > this.bWidth / 2) {
						DistanceFromCenter = DistanceFromRightCorner - (this.bWidth / 2) * 12
						return (
							(this.bWidth - DistanceFromRightCorner / 12) * this.rPitch +
							this.bHeight * 12 -
							StartingHeight
						)
					}
			}
		} else if (this.rShape === 'Single Slope') {
			switch (Wall) {
				case 'e1':
					/*
					 *VBA-Dev Note:
					 *Inside Distance - Distance ffrom s4 Column = Actual Distance of slope
					 * Distance of Slope * rPitch = Height above eave height
					 * Distance above eave height + eave height = distance to roof
					 */
					return (
						(this.bWidth - DistanceFromRightCorner / 12) * this.rPitch +
						this.bHeight * 12 -
						StartingHeight
					)
				case 's2':
					return this.bHeight * 12 - StartingHeight
				case 'e3':
					/*
					 *VBA-Dev Note:
					 *CL - Inside of s2 Column = Actual Distance of Slope
					 *Distance of Slope * rPitch = Height above eave height
					 *Distance above eavh height + eave height = distance to roof
					 */
					return (DistanceFromRightCorner / 12) * this.rPitch + this.bHeight * 12 - StartingHeight
				case 's4':
					return this.bHeight * 12 + this.bWidth * this.rPitch - StartingHeight
			}
		}
	}
	//Function for distance from right corner of an endwall at a given height
	DistanceFromCorner(Wall: string, HeightAlongRoof: number) {
		let DistanceFromCenter: number
		if (this.rShape === 'Gable') {
			return ((HeightAlongRoof - this.bHeight * 12) / this.rPitch) * 12
		} else if (this.rShape === 'Single Slope') {
			if (Wall === 'e1') {
				// 0 is the tallest point
				return this.bWidth * 12 - ((HeightAlongRoof - this.bHeight * 12) / this.rPitch) * 12
			} else {
				return ((HeightAlongRoof - this.bHeight * 12) / this.rPitch) * 12
			}
		}
	}
}
