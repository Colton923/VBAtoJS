// TODO: Option Explicit ... Warning!!! not translated
let bLength: number;
let bHeight: number;
let rPitch: number;
let RafterLength: number;
let s2RafterSheetLength: number;
let s4RafterSheetLength: number;
let bWidth: number;
let rShape: string;
let s2Overhang: number;
let s4Overhang: number;
let e1Overhang: number;
let e3Overhang: number;
let s2Extension: number;
let s4Extension: number;
let e1Extension: number;
let e3Extension: number;
let e1ExtensionPanelQty: number;
let e3ExtensionPanelQty: number;
let Gutters: boolean;
let BaseTrim: boolean;
let e1WallPanelOverlaps: number;
let e3WallPanelOverlaps: number;
let s2ExtensionPitch: number;
let s4ExtensionPitch: number;
let s2ExtensionHeight: number;
let s4ExtensionHeight: number;
let s2ExtensionWidth: number;
let s4ExtensionWidth: number;
let wPanelShape: string;
let rPanelShape: string;
let rPanelType: string;
let rPanelColor: string;
let wPanelType: string;
let wPanelColor: string;
let RakeTrimColor: string;
let OutsideCornerTrimColor: string;
let e1GableOverhangSoffit: boolean;
let e3GableOverhangSoffit: boolean;
let s2EaveOverhangSoffit: boolean;
let s4EaveOverhangSoffit: boolean;
let e1GableExtensionSoffit: boolean;
let e3GableExtensionSoffit: boolean;
let s2EaveExtensionSoffit: boolean;
let s4EaveExtensionSoffit: boolean;
let EaveExtLength: number;
let bLengthRoofPanelOverage: number;
let InteriorColumns: Collection;
let s2ColumnWidth: number;
let s4ColumnWidth: number;
let WeldClips: number;
let SSTotalCost: number;
let e1FOs: Collection = new Collection();
let s2FOs: Collection = new Collection();
let e3FOs: Collection = new Collection();
let s4FOs: Collection = new Collection();
let fieldlocateFOs: Collection = new Collection();
let e1Columns: Collection = new Collection();
let s2Columns: Collection = new Collection();
let e3Columns: Collection = new Collection();
let s4Columns: Collection = new Collection();
let e1Girts: Collection = new Collection();
let s2Girts: Collection = new Collection();
let e3Girts: Collection = new Collection();
let s4Girts: Collection = new Collection();
let e1Rafters: Collection = new Collection();
let intRafters: Collection = new Collection();
let e3Rafters: Collection = new Collection();
let RoofPurlins: Collection = new Collection();
let e1OverhangMembers: Collection = new Collection();
let s2OverhangMembers: Collection = new Collection();
let e3OverhangMembers: Collection = new Collection();
let s4OverhangMembers: Collection = new Collection();
let e1ExtensionMembers: Collection = new Collection();
let s2ExtensionMembers: Collection = new Collection();
let e3ExtensionMembers: Collection = new Collection();
let s4ExtensionMembers: Collection = new Collection();
let BaseAngleTrim: Collection = new Collection();
let WeldPlates: Collection = new Collection();


    public RoofLength(): number {
        return ((bLength * 12)
                    + (e1Overhang
                    + (e1Extension
                    + (e3Overhang + e3Extension))));
    }

    public RoofFtLength(): number {
        return (((bLength * 12)
                    + (e1Overhang
                    + (e1Extension
                    + (e3Overhang + e3Extension))))
                    / 12);
    }

    // high side eave height
    public HighSideEaveHeight(): number {
        // Inches
        return ((bHeight * 12)
                    + (bWidth * rPitch));
    }

    public s2ExtensionRafterLength(): number {
        if ((s2Extension == 0)) {
            s2ExtensionRafterLength = 0;
        }
        else {
            s2ExtensionRafterLength = ((s2Extension / 12)
                        * Sqr(((12 | 2)
                            + (s2ExtensionPitch | 2))));
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
        }

    }

    public s4ExtensionRafterLength(): number {
        if ((s4Extension == 0)) {
            s4ExtensionRafterLength = 0;
        }
        else {
            s4ExtensionRafterLength = ((s4Extension / 12)
                        * Sqr(((12 | 2)
                            + (s4ExtensionPitch | 2))));
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
        }

    }

    // ''''''''''''''''''''''''''''''' Extension Intersections '''''''''''''''''''''''
    // Note: Intersecting extension panels are accounted for as eave extension panels
    public s2e1ExtensionIntersection(): boolean {
        switch (EstSht.Range("s2e1_Intersection").Value) {
            case "N/A":
            case "Exclude":
                s2e1ExtensionIntersection = false;
                break;
            case "Include":
                s2e1ExtensionIntersection = true;
                break;
        }

    }

    public s2e3ExtensionIntersection(): boolean {
        switch (EstSht.Range("s2e3_Intersection").Value) {
            case "N/A":
            case "Exclude":
                s2e3ExtensionIntersection = false;
                break;
            case "Include":
                s2e3ExtensionIntersection = true;
                break;
        }

    }

    public s4e1ExtensionIntersection(): boolean {
        switch (EstSht.Range("s4e1_Intersection").Value) {
            case "N/A":
            case "Exclude":
                s4e1ExtensionIntersection = false;
                break;
            case "Include":
                s4e1ExtensionIntersection = true;
                break;
        }

    }

    public s4e3ExtensionIntersection(): boolean {
        switch (EstSht.Range("s4e3_Intersection").Value) {
            case "N/A":
            case "Exclude":
                s4e3ExtensionIntersection = false;
                break;
            case "Include":
                s4e3ExtensionIntersection = true;
                break;
        }

    }

    // '''''''''''''''''''''''''''''' Eave Extension Lengths (from endwall to endwall)
    public s2EaveExtensionBuildingLength(): number {
        EaveExtLength = ((bLength * 12)
                    + (e1Overhang + e3Overhang));
        if ((s2e1ExtensionIntersection == true)) {
            EaveExtLength = (EaveExtLength + e1Extension);
        }

        if ((s2e3ExtensionIntersection == true)) {
            EaveExtLength = (EaveExtLength + e3Extension);
        }

        return EaveExtLength;
    }

    public s4EaveExtensionBuildingLength(): number {
        EaveExtLength = ((bLength * 12)
                    + (e1Overhang + e3Overhang));
        if ((s4e1ExtensionIntersection == true)) {
            EaveExtLength = (EaveExtLength + e1Extension);
        }

        if ((s4e3ExtensionIntersection == true)) {
            EaveExtLength = (EaveExtLength + e3Extension);
        }

        return EaveExtLength;
    }

    public NetSingleRoofPanelQty(): number {
        return Application.WorksheetFunction.RoundUp(((((bLength * 12)
                        + (e1Overhang
                        + (e3Overhang
                        + (e1Extension + e3Extension))))
                        / 12)
                        / 3), 0);
    }

    // Wall Exclusions
    public WallStatus(Wall: string): string {
        return EstSht.Range((Wall + "_WallStatus")).Value;
    }

    // Partial Walls' Length Above Finished Floor
    public LengthAboveFinishedFloor(Wall: string): number {
        //  Ft
        if ((EstSht.Range((Wall + "_WallStatus")).Value == "Include")) {
            LengthAboveFinishedFloor = 0;
        }
        else if ((EstSht.Range((Wall + "_WallStatus")).Value == "Partial")) {
            LengthAboveFinishedFloor = EstSht.Range((Wall + "_WallStatus")).offset(0, 2).Value;
        }
        else if ((EstSht.Range((Wall + "_WallStatus")).Value == "Gable Only")) {
            LengthAboveFinishedFloor = bHeight;
        }

    }

    // Liner Panel Options
    public LinerPanels(Location: string): string {
        LinerPanels = EstSht.Range((Location + "_LinerPanels")).Value;
        if ((LinerPanels == "")) {
            LinerPanels = "None";
        }

        // Wainscot
    }

    Wainscot(Wall: string): string {
        Wainscot = EstSht.Range((Wall + "_Wainscot")).Value;
        if ((Wainscot == "")) {
            Wainscot = "None";
        }

        // expandable endwall
    }

    ExpandableEndwall(eWall: string): boolean {
        if ((EstSht.Range((eWall + "_Expandable")).Value != "Yes")) {
            ExpandableEndwall = false;
        }
        else {
            ExpandableEndwall = true;
        }

    }

    // function for height to the very top of the building (that is, the top surface, not the bottom of the rafter) at a given horizontal distance
    // SHOULD ONLY BE CALLED AFTER INT COLUMNS ARE GENERATED
    public DistanceToRoof(Wall: string, DistanceFromRightCorner: number, StartingHeight: number) {
        let DistanceFromCenter: number;
        // Warning!!! Optional parameters not supported
        // ActualPitch = (((bWidth * (rPitch / 12))) / (bWidth - ((s2ColumnWidth + s4ColumnWidth) / 12))) * 12
        if ((rShape == "Gable")) {
            switch (Wall) {
                case "s2":
                case "s4":
                    DistanceToRoof = ((bHeight * 12)
                                - StartingHeight);
                    break;
                case "e1":
                    if (((DistanceFromRightCorner / 12)
                                <= (bWidth / 2))) {
                        // DistanceToRoof = (((DistanceFromRightCorner - s4ColumnWidth) / 12) * ActualPitch) + bHeight * 12 - StartingHeight
                        DistanceToRoof = (((DistanceFromRightCorner / 12)
                                    * rPitch)
                                    + ((bHeight * 12)
                                    - StartingHeight));
                        // past peak
                    }
                    else if (((DistanceFromRightCorner / 12)
                                > (bWidth / 2))) {
                        DistanceFromCenter = (DistanceFromRightCorner
                                    - ((bWidth / 2)
                                    * 12));
                        // DistanceToRoof = ((bHeight * 12 + (((bWidth - s2ColumnWidth / 12) / 2) * ActualPitch)) - ((DistanceFromCenter / 12) * rPitch)) - StartingHeight
                        DistanceToRoof = (((bWidth
                                    - (DistanceFromRightCorner / 12))
                                    * rPitch)
                                    + ((bHeight * 12)
                                    - StartingHeight));
                    }

                    break;
                case "e3":
                    if (((DistanceFromRightCorner / 12)
                                <= (bWidth / 2))) {
                        // DistanceToRoof = (((DistanceFromRightCorner - s2ColumnWidth) / 12) * ActualPitch) + bHeight * 12 - StartingHeight
                        DistanceToRoof = (((DistanceFromRightCorner / 12)
                                    * rPitch)
                                    + ((bHeight * 12)
                                    - StartingHeight));
                        // past peak
                    }
                    else if (((DistanceFromRightCorner / 12)
                                > (bWidth / 2))) {
                        DistanceFromCenter = (DistanceFromRightCorner
                                    - ((bWidth / 2)
                                    * 12));
                        // DistanceToRoof = ((bHeight * 12 + (((bWidth - s4ColumnWidth / 12) / 2) * ActualPitch)) - ((DistanceFromCenter / 12) * ActualPitch)) - StartingHeight
                        DistanceToRoof = (((bWidth
                                    - (DistanceFromRightCorner / 12))
                                    * rPitch)
                                    + ((bHeight * 12)
                                    - StartingHeight));
                    }

                    break;
            }

        }
        else if ((rShape == "Single Slope")) {
            switch (Wall) {
                case "e1":
                    DistanceToRoof = (((bWidth
                                - (DistanceFromRightCorner / 12))
                                * rPitch)
                                + ((bHeight * 12)
                                - StartingHeight));
                    break;
                case "s2":
                    DistanceToRoof = ((bHeight * 12)
                                - StartingHeight);
                    break;
                case "e3":
                    DistanceToRoof = (((DistanceFromRightCorner / 12)
                                * rPitch)
                                + ((bHeight * 12)
                                - StartingHeight));
                    break;
                case "s4":
                    DistanceToRoof = ((bHeight * 12)
                                + ((bWidth * rPitch)
                                - StartingHeight));
                    break;
            }

        }

    }

    // function for distance from right corner of an endwall at a given height
    public DistanceFromCorner(Wall: string, HeightAlongRoof: number) {
        let DistanceFromCenter: number;
        if ((rShape == "Gable")) {
            if ((HeightAlongRoof
                        < (bWidth * (12 / 2)))) {
                if ((Wall == "e1")) {
                    DistanceFromCorner = (((HeightAlongRoof
                                - (bHeight * 12))
                                / rPitch)
                                * 12);
                }
                else {
                    DistanceFromCorner = (((HeightAlongRoof
                                - (bHeight * 12))
                                / rPitch)
                                * 12);
                }

            }
            else {
                // right now these are the same
                if ((Wall == "e3")) {
                    DistanceFromCorner = (((HeightAlongRoof
                                - (bHeight * 12))
                                / rPitch)
                                * 12);
                }
                else {
                    DistanceFromCorner = (((HeightAlongRoof
                                - (bHeight * 12))
                                / rPitch)
                                * 12);
                }

            }

        }
        else if ((rShape == "Single Slope")) {
            if ((Wall == "e1")) {
                // 0 is the tallest point
                DistanceFromCorner = ((bWidth * 12)
                            - (((HeightAlongRoof
                            - (bHeight * 12))
                            / rPitch)
                            * 12));
            }
            else {
                // for e3 0 is the lowest point
                DistanceFromCorner = (((HeightAlongRoof
                            - (bHeight * 12))
                            / rPitch)
                            * 12);
            }

        }

    }

    private Class_Initialize() {
        let FOCell: Range;
        let FO: clsFO;
        let BayCell: Range;
        let TotalBayLength: number;
        let Column: clsMember;
        let Bay: number;
        // set basic building parameters
        // With...
        bHeight = EstSht.Range;
        "Building_Height".Value;
        bWidth = EstSht.Range;
        "Building_Width".Value;
        bLength = EstSht.Range;
        "Building_Length".Value;
        rPitch = EstSht.Range;
        "Roof_Pitch".Value;
        rShape = EstSht.Range;
        "Roof_Shape".Value;
        // create Int Columns collection
        InteriorColumns = new Collection();
        // create girt collections to be filled
        e1Girts = new Collection();
        s2Girts = new Collection();
        e3Girts = new Collection();
        s4Girts = new Collection();
        // create rafter collections to be filled
        e1Rafters = new Collection();
        intRafters = new Collection();
        e3Rafters = new Collection();
        // create overhang and extension members collections to be filled
        e1OverhangMembers = new Collection();
        s2OverhangMembers = new Collection();
        e3OverhangMembers = new Collection();
        s4OverhangMembers = new Collection();
        e1ExtensionMembers = new Collection();
        s2ExtensionMembers = new Collection();
        e3ExtensionMembers = new Collection();
        s4ExtensionMembers = new Collection();
        // create roof purlin collection
        RoofPurlins = new Collection();
        // create Weld Plate Collection
        WeldPlates = new Collection();
        // ''''''''''''set extension pitches
        if (EstSht.Range) {
            ("s2_EaveExtension".Value > 0);
            if (EstSht.Range) {
                "s2_EaveExtensionPitch".Value = "Match Roof";
                s2ExtensionPitch = rPitch;
            }
            else {
                s2ExtensionPitch = EstSht.Range;
                "s2_EaveExtensionPitch".Value;
            }

        }

        if (EstSht.Range) {
            ("s4_EaveExtension".Value > 0);
            if (EstSht.Range) {
                "s4_EaveExtensionPitch".Value = "Match Roof";
                s4ExtensionPitch = rPitch;
            }
            else {
                s4ExtensionPitch = EstSht.Range;
                "s4_EaveExtensionPitch".Value;
            }

        }

        // ''''''''''' generate sidewall 2 column centerlines
        if (EstSht.Range) {
            ("BayNum".Value > 1);
            // '''s2 columns
            for (BayCell in Range(EstSht.Range, "Bay1_Length", EstSht.Range, "Bay12_Length")) {
                if (((BayCell.EntireRow.Hidden == false)
                            && (BayCell.Value != 0))) {
                    TotalBayLength = (TotalBayLength + BayCell.Value);
                    if ((TotalBayLength == bLength)) {
                        break;
                    }

                    // new column
                    Column = new clsMember();
                    Column.CL = (TotalBayLength * 12);
                    // add column length (building height)
                    Column.Length = (bHeight * 12);
                    Column.tEdgeHeight = Column.Length;
                    // add to collection
                    s2Columns.Add;
                    Column;
                }

            }

            // '''s4 columns
            TotalBayLength = 0;
            for (Bay = 12; (Bay <= 1); Bay = (Bay + -1)) {
                BayCell = EstSht.Range;
                ("Bay"
                            + (Bay + "_Length"));
                if (((BayCell.EntireRow.Hidden == false)
                            && (BayCell.Value != 0))) {
                    TotalBayLength = (TotalBayLength + BayCell.Value);
                    if ((TotalBayLength == bLength)) {
                        break;
                    }

                    // new column
                    Column = new clsMember();
                    Column.CL = (TotalBayLength * 12);
                    // add column height (building height)
                    if ((rShape == "Gable")) {
                        Column.Length = (bHeight * 12);
                    }

                    if ((rShape == "Single Slope")) {
                        Column.Length = HighSideEaveHeight;
                    }

                    // add to collection
                    Column.tEdgeHeight = Column.Length;
                    s4Columns.Add;
                    Column;
                }

            }

        }

        // '''''''''''''''''''''''''''''''''''''''''''''''' Build FO Collections  '''''''''''''''''''''''''''''''''''''''''
        // pDoors
        for (FOCell in Range(EstSht.Range, "pDoorCell1", EstSht.Range, "pDoorCell12")) {
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                FO = new clsFO();
                FO.FOType = "PDoor";
                FO.Height = (7 * 12);
                // set width
                if ((FOCell.offset(0, 1).Value == "3070")) {
                    FO.Width = (3 * 12);
                }
                else if ((FOCell.offset(0, 1).Value == "4070")) {
                    FO.Width = (4 * 12);
                }

                // reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
                if (((FOCell.offset(0, 2).Value == "Endwall 1")
                            || (FOCell.offset(0, 2).Value == "Endwall 3"))) {
                    FO.rEdgePosition = ((bWidth * 12)
                                - (FOCell.offset(0, 8) * 12));
                }
                else {
                    FO.rEdgePosition = ((bLength * 12)
                                - (FOCell.offset(0, 8).Value * 12));
                }

                FO.Description = ("pDoor #"
                            + (FOCell.Value + (", an FO located on "
                            + (FOCell.offset(0, 2).Value + (". rEdge: "
                            + ((FO.rEdgePosition / 12) + ("', lEdge: "
                            + ((FO.lEdgePosition / 12)
                            + "'"))))))));
                switch (FOCell.offset(0, 2).Value) {
                    case "Endwall 1":
                        FO.Wall = "e1";
                        e1FOs.Add;
                        FO;
                        break;
                    case "Sidewall 2":
                        FO.Wall = "s2";
                        s2FOs.Add;
                        FO;
                        break;
                    case "Endwall 3":
                        FO.Wall = "e3";
                        e3FOs.Add;
                        FO;
                        break;
                    case "Sidewall 4":
                        FO.Wall = "s4";
                        s4FOs.Add;
                        FO;
                        break;
                    case "Field Locate":
                        FO.Wall = "Field Locate";
                        fieldlocateFOs.Add;
                        FO;
                        break;
                }

            }

        }

        // OHDoors
        for (FOCell in Range(EstSht.Range, "OHDoorCell1", EstSht.Range, "OHDoorCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                // new FO class
                FO = new clsFO();
                FO.FOType = "OHDoor";
                FO.Width = (FOCell.offset(0, 1).Value * 12);
                FO.Height = (FOCell.offset(0, 2).Value * 12);
                FO.bEdgeHeight = 0;
                // reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
                if (((FOCell.offset(0, 3).Value == "Endwall 1")
                            || (FOCell.offset(0, 3).Value == "Endwall 3"))) {
                    FO.rEdgePosition = ((bWidth * 12)
                                - (FOCell.offset(0, 10) * 12));
                }
                else {
                    FO.rEdgePosition = ((bLength * 12)
                                - (FOCell.offset(0, 10).Value * 12));
                }

                FO.Description = ("OHDoor #"
                            + (FOCell.Value + (", an FO located on "
                            + (FOCell.offset(0, 3).Value + (". rEdge: "
                            + ((FO.rEdgePosition / 12) + ("', lEdge: "
                            + ((FO.lEdgePosition / 12) + ("' , Height: "
                            + ((FO.Height / 12)
                            + "'"))))))))));
                switch (FOCell.offset(0, 3).Value) {
                    case "Endwall 1":
                        FO.Wall = "e1";
                        e1FOs.Add;
                        FO;
                        break;
                    case "Sidewall 2":
                        FO.Wall = "s2";
                        s2FOs.Add;
                        FO;
                        break;
                    case "Endwall 3":
                        FO.Wall = "e3";
                        e3FOs.Add;
                        FO;
                        break;
                    case "Sidewall 4":
                        FO.Wall = "s4";
                        s4FOs.Add;
                        FO;
                        break;
                }

            }

        }

        // Windows
        for (FOCell in Range(EstSht.Range, "WindowCell1", EstSht.Range, "WindowCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                // new FO class
                FO = new clsFO();
                FO.FOType = "Window";
                FO.Width = FOCell.offset(0, 1).Value;
                FO.Height = FOCell.offset(0, 2).Value;
                FO.bEdgeHeight = (FOCell.offset(0, 4).Value * 12);
                // reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
                if (((FOCell.offset(0, 3).Value == "Endwall 1")
                            || (FOCell.offset(0, 3).Value == "Endwall 3"))) {
                    FO.rEdgePosition = ((bWidth * 12)
                                - (FOCell.offset(0, 7) * 12));
                }
                else {
                    FO.rEdgePosition = ((bLength * 12)
                                - (FOCell.offset(0, 7).Value * 12));
                }

                FO.Description = ("Window #"
                            + (FOCell.Value + (", an FO located on "
                            + (FOCell.offset(0, 3).Value + (". rEdge: "
                            + ((FO.rEdgePosition / 12) + ("', lEdge: "
                            + ((FO.lEdgePosition / 12) + ("', bEdge:"
                            + ((FO.bEdgeHeight / 12) + ("', Height: "
                            + ((FO.Height / 12)
                            + "'"))))))))))));
                switch (FOCell.offset(0, 3).Value) {
                    case "Endwall 1":
                        FO.Wall = "e1";
                        e1FOs.Add;
                        FO;
                        break;
                    case "Sidewall 2":
                        FO.Wall = "s2";
                        s2FOs.Add;
                        FO;
                        break;
                    case "Endwall 3":
                        FO.Wall = "e3";
                        e3FOs.Add;
                        FO;
                        break;
                    case "Sidewall 4":
                        FO.Wall = "s4";
                        s4FOs.Add;
                        FO;
                        break;
                    case "Field Locate":
                        FO.Wall = "Field Locate";
                        fieldlocateFOs.Add;
                        FO;
                        break;
                }

            }

        }

        // Misc FOs
        for (FOCell in Range(EstSht.Range, "MiscFOCell1", EstSht.Range, "MiscFOCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                // new FO class
                FO = new clsFO();
                FO.FOType = "MiscFO";
                FO.Width = (FOCell.offset(0, 1).Value * 12);
                FO.Height = (FOCell.offset(0, 2).Value * 12);
                FO.bEdgeHeight = (FOCell.offset(0, 6).Value * 12);
                // reverse coordinates as listed on Project Details page; User inputs coordinates from left to right as opposed to left to right (standard within the code)
                if (((FOCell.offset(0, 3).Value == "Endwall 1")
                            || (FOCell.offset(0, 3).Value == "Endwall 3"))) {
                    FO.rEdgePosition = ((bWidth * 12)
                                - (FOCell.offset(0, 9) * 12));
                }
                else {
                    FO.rEdgePosition = ((bLength * 12)
                                - (FOCell.offset(0, 9).Value * 12));
                }

                FO.Description = ("MiscFO #"
                            + (FOCell.Value + (", an FO located on "
                            + (FOCell.offset(0, 3).Value + (". rEdge: "
                            + ((FO.rEdgePosition / 12) + ("', lEdge: "
                            + ((FO.lEdgePosition / 12) + ("', bEdge:"
                            + ((FO.bEdgeHeight / 12) + ("', Height: "
                            + ((FO.Height / 12)
                            + "'"))))))))))));
                FO.StructuralSteelOption = FOCell.offset(0, 10).Value;
                // set wall, add to collection
                switch (FOCell.offset(0, 3).Value) {
                    case "Endwall 1":
                        FO.Wall = "e1";
                        e1FOs.Add;
                        FO;
                        break;
                    case "Sidewall 2":
                        FO.Wall = "s2";
                        s2FOs.Add;
                        FO;
                        break;
                    case "Endwall 3":
                        FO.Wall = "e3";
                        e3FOs.Add;
                        FO;
                        break;
                    case "Sidewall 4":
                        FO.Wall = "s4";
                        s4FOs.Add;
                        FO;
                        break;
                }

            }

        }

    }
