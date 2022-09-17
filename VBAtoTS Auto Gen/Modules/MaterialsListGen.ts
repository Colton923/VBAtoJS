// TODO: Option Explicit ... Warning!!! not translated


MaterialsListGen() {
    let Qty: number;
    let rShape: string;
    let pShape: string;
    let pType: string;
    let rColor: string;
    let bLength: number;
    let bWidth: number;
    let bHeight: number;
    let rPitch: string;
    // '' Roof Pitch Rise
    let RafterLength: number;
    let s2RafterSheetLength: number;
    let s4RafterSheetLength: number;
    let RoofPitchHypot: number;
    // '' Inches Rise per Ft Roof Span
    let s2EaveOverhang: number;
    let s4EaveOverhang: number;
    let e1GableOverhang: number;
    let e3GableOverhang: number;
    // new roof panel calculations
    let s2RoofPanels: Collection;
    let s4RoofPanels: Collection;
    let e1GableExtensionPanels: Collection;
    let e3GableExtensionPanels: Collection;
    let RoofLength: number;
    let RoofPanel: clsPanel;
    // Standard Overhang
    let StandardEaveOverhang: number;
    let Feet: number;
    let Inches: number;
    let InchFraction: number;
    let Undercut: number;
    // ''''''''' ridge cap
    let RidgeCapQty: number;
    let PitchString: string;
    // '' sidewall
    let wShape: string;
    let wType: string;
    let wColor: string;
    // '' High sidewall height
    let HighSideEaveHeight: number;
    // sidewall panels
    let s2SidewallPanels: Collection = new Collection();
    let s4SidewallPanels: Collection = new Collection();
    // '' endwalls
    let e1EndwallPanels: Collection = new Collection();
    let e3EndwallPanels: Collection = new Collection();
    let EndwallPanelCount: number;
    let e1PanelQty: number;
    let e1PanelLength: string;
    let e3PanelQty: number;
    let e3PanelLength: string;
    let PanelNumber: number;
    let pLength: number;
    let ePanel: clsPanel;
    let p: number;
    let MaxHeight: number;
    // wall panel class for vendor material list
    let WallPanel: clsPanel;
    // ''''' Rake Trim
    let RakeTrimPieces: Collection;
    let NetRafterLength: number;
    let RakeTrimColor: string;
    let TrimPiece: clsTrim;
    // ''''' Eave Trim
    let EaveTrimPieces: Collection;
    let s2EaveTrimLength: number;
    let s4EaveTrimLength: number;
    let EaveTrimColor: string;
    // ''''' Outside Corner Trim
    let OutsideCornerTrimPieces: Collection;
    let NetCornerLength: number;
    let OutsideCornerTrimColor: string;
    // ''''' Base Trim
    let BaseTrimPieces: Collection;
    let NetBaseTrimLength: number;
    let NetPDoorWidth: number;
    // ' For net personelle door
    let NetOHDoorWidth: number;
    // ' For net overhead door
    let BaseTrimColor: string;
    // ''''Wainscot Trim
    let WainscotTrimPieces: Collection;
    let TempDoorWidth: number;
    // 'For Wainscot Trim
    let NetWainscotTrimLength: number;
    // ''''' Framed Openings
    let FOCell: Range;
    // ''''' Gutters & Downspouts
    let GutterPieces: Collection;
    let Gutters: boolean;
    let NetGutterLength: number;
    let GutterEndCapQty: number;
    let GutterStrapQty: number;
    let GutterPiece: clsTrim;
    // trim class since available in the same sizes
    let GutterColor: string;
    let DownspoutColor: string;
    let DownspoutQty: number;
    let DownspoutPieces: Collection;
    let DownspoutPiece: clsTrim;
    // trim class since available in the same sizes
    let RemainingHeight: number;
    let DownspoutStrapQty: number;
    let h: number;
    let PopRivitQty: number;
    // Bays
    let BayCount: number;
    //  Translucent Wall Panels & Skylights
    let SkylightPanelQty: number;
    let SkylightPanel: clsPanel;
    // ''''' Soffits
    let e1GableOverhangSoffit: boolean;
    let e1GableExtensionSoffit: boolean;
    let s2EaveExtensionSoffit: boolean;
    let s2EaveOverhangSoffit: boolean;
    let e3GableOverhangSoffit: boolean;
    let e3GableExtensionSoffit: boolean;
    let s4EaveOverhangSoffit: boolean;
    let s4EaveExtensionSoffit: boolean;
    let NetOutsideAngleLength: number;
    let SoffitPanel: clsPanel;
    let SoffitTrim: clsTrim;
    let SoffitPiece: Object;
    let SoffitQty: number;
    // '' Extensions
    let s2EaveExtension: number;
    let s4EaveExtension: number;
    let e1GableExtension: number;
    let e3GableExtension: number;
    let PanelQty: number;
    let e1ExtensionPanels: Collection;
    let s2ExtensionPanels: Collection;
    let e3ExtensionPanels: Collection;
    let s4ExtensionPanels: Collection;
    let ExtensionPanel: clsPanel;
    // '' Overhangs
    let e1GableOverhangSection: boolean;
    let s2EaveOverhangSection: boolean;
    let e3GableOverhangSection: boolean;
    let s4EaveOverhangSection: boolean;
    let e1GableExtensionSection: boolean;
    let s2EaveExtensionSection: boolean;
    let e3GableExtensionSection: boolean;
    let s4EaveExtensionSection: boolean;
    // overhang or extension collections
    //  soffit collections
    let e1SoffitPanels: Collection;
    let e1SoffitTrim: Collection;
    let s2SoffitPanels: Collection;
    let s2SoffitTrim: Collection;
    let e3SoffitPanels: Collection;
    let e3SoffitTrim: Collection;
    let s4SoffitPanels: Collection;
    let s4SoffitTrim: Collection;
    // 2x8 inside angle
    let e1InsideAngleTrim: Collection;
    let e3InsideAngleTrim: Collection;
    let NetInsideAngleLength: number;
    // '' Fasteners
    let rTekScrewQty: number;
    let rLapScrewQty: number;
    let wTekScrewQty: number;
    let wLapScrewQty: number;
    let rPurlins: number;
    let pTypeCount: number;
    let rOverlaps: number;
    let sOverlaps: number;
    let eOverlaps: number;
    let TrimScrewQty: number;
    let NetRakeTrimLength: number;
    let SoffitScrewQty: number;
    let SoffitScrewColor: string;
    let Screw: clsFastener;
    let TrimScrews: Collection;
    // Liner Panels
    let e1LinerPanels: Collection = new Collection();
    let e3LinerPanels: Collection = new Collection();
    let s2LinerPanels: Collection = new Collection();
    let s4LinerPanels: Collection = new Collection();
    let RoofLinerPanels: Collection = new Collection();
    let LinerPanelsSection: boolean;
    // clsPanel
    let Panel: clsPanel;
    // '' Miscellaneous
    let ButylTapeQty: number;
    let InsideClosureQty: number;
    let OutsideClosureQty: number;
    // '' Vendor Materials List
    let PanelCollection: Collection;
    let TrimCollection: Collection;
    let MiscCollection: Collection;
    let item: clsMiscItem;
    // '' Building Class '''
    let b: clsBuilding;
    // Misc Variables
    let MatSht: Worksheet;
    let N: number;
    let WriteCell: Range;
    // '''''''' Generate Roofing Section
    // '' Read Information
    // With...
    // building width, building length, roof pitch
    bWidth = EstSht.Range;
    "Building_Width".Value;
    bHeight = EstSht.Range;
    "Building_Height".Value;
    bLength = EstSht.Range;
    "Building_Length".Value;
    rPitch = EstSht.Range;
    "Roof_Pitch".Value;
    // single slope or gable
    rShape = EstSht.Range;
    "Roof_Shape".Value;
    // roof panel info
    pShape = EstSht.Range;
    "Roof_pShape".Value;
    pType = EstSht.Range;
    "Roof_pType".Value;
    rColor = EstSht.Range;
    "Roof_Color".Value;
    //  wall panel info
    wShape = EstSht.Range;
    "Wall_pShape".Value;
    wType = EstSht.Range;
    "Wall_pType".Value;
    wColor = EstSht.Range;
    "Wall_Color".Value;
    // check for invalid building height
    if ((bHeight > 80)) {
        MsgBox;
        "Buildings cannot be taller than 80'. Please correct the data before generating a materials list.";
        vbExclamation;
        "Building Height Error";
        return;
    }
    else if ((rShape == "Single Slope")) {
        if (((bHeight
                    + ((bWidth * rPitch)
                    / 12))
                    > 100)) {
            MsgBox;
            "The high side eave cannot be greater than 100'. Please correct the data before generating a materials list.";
            vbExclamation;
            "High Side Eave Error";
            return;
        }

    }

    // '' overhang info
    // convert to inches
    e1GableOverhang = EstSht.Range;
    ("e1_GableOverhang".Value * 12);
    s2EaveOverhang = EstSht.Range;
    ("s2_EaveOverhang".Value * 12);
    e3GableOverhang = EstSht.Range;
    ("e3_GableOverhang".Value * 12);
    s4EaveOverhang = EstSht.Range;
    ("s4_EaveOverhang".Value * 12);
    // '' Extensions
    e1GableExtension = EstSht.Range;
    ("e1_GableExtension".Value * 12);
    s2EaveExtension = EstSht.Range;
    ("s2_EaveExtension".Value * 12);
    e3GableExtension = EstSht.Range;
    ("e3_GableExtension".Value * 12);
    s4EaveExtension = EstSht.Range;
    ("s4_EaveExtension".Value * 12);
    // '' Trim
    RakeTrimColor = EstSht.Range;
    "Rake_tColor".Value;
    EaveTrimColor = EstSht.Range;
    "Eave_tColor".Value;
    OutsideCornerTrimColor = EstSht.Range;
    "OutsideCorner_tColor".Value;
    BaseTrimColor = EstSht.Range;
    "Base_tColor".Value;
    // '' Gutters
    if (EstSht.Range) {
        "GutterAndDownspouts".Value = "Yes";
        Gutters = true;
        GutterColor = EstSht.Range;
        "GutterColor".Value;
        DownspoutColor = EstSht.Range;
        "DownspoutColor".Value;
        // '' Soffits
        // check if soffits
        if (EstSht.Range) {
            "e1_GableOverhangSoffit".Value = "Yes";
            e1GableOverhangSoffit = true;
            if (EstSht.Range) {
                "e1_GableExtensionSoffit".Value = "Yes";
                e1GableExtensionSoffit = true;
                if (EstSht.Range) {
                    "s2_EaveOverhangSoffit".Value = "Yes";
                    s2EaveOverhangSoffit = true;
                    if (EstSht.Range) {
                        "s2_EaveExtensionSoffit".Value = "Yes";
                        s2EaveExtensionSoffit = true;
                        if (EstSht.Range) {
                            "e3_GableOverhangSoffit".Value = "Yes";
                            e3GableOverhangSoffit = true;
                            if (EstSht.Range) {
                                "e3_GableExtensionSoffit".Value = "Yes";
                                e3GableExtensionSoffit = true;
                                if (EstSht.Range) {
                                    "s4_EaveOverhangSoffit".Value = "Yes";
                                    s4EaveOverhangSoffit = true;
                                    if (EstSht.Range) {
                                        "s4_EaveExtensionSoffit".Value = "Yes";
                                        s4EaveExtensionSoffit = true;
                                        if (EstSht.Range) {
                                            (("e1_GableOverhang".Value > 0)
                                                        & e1GableOverhangSoffit) = true;
                                            e1GableOverhangSection = true;
                                            if (EstSht.Range) {
                                                (("s2_EaveOverhang".Value > 0)
                                                            & s2EaveOverhangSoffit) = true;
                                                s2EaveOverhangSection = true;
                                                if (EstSht.Range) {
                                                    (("e3_GableOverhang".Value > 0)
                                                                & e3GableOverhangSoffit) = true;
                                                    e3GableOverhangSection = true;
                                                    if (EstSht.Range) {
                                                        (("s4_EaveOverhang".Value > 0)
                                                                    & s4EaveOverhangSoffit) = true;
                                                        s4EaveOverhangSection = true;
                                                        if (EstSht.Range) {
                                                            ("e1_GableExtension".Value > 0);
                                                            e1GableExtensionSection = true;
                                                            if (EstSht.Range) {
                                                                ("s2_EaveExtension".Value > 0);
                                                                s2EaveExtensionSection = true;
                                                                if (EstSht.Range) {
                                                                    ("e3_GableExtension".Value > 0);
                                                                    e3GableExtensionSection = true;
                                                                    if (EstSht.Range) {
                                                                        ("s4_EaveExtension".Value > 0);
                                                                        s4EaveExtensionSection = true;
                                                                        s2EaveOverhang = (s2EaveOverhang + 4.25);
                                                                        // ''s4
                                                                        // always additional 4.25 overhang for gable or if an s4 eave extension
                                                                        if (((rShape == "Gable")
                                                                                    || (s4EaveExtensionSection == true))) {
                                                                            s4EaveOverhang = (s4EaveOverhang + 4.25);
                                                                        }

                                                                        // for single slope, no additional 4.25" s4 overhang as long as there's no extension
                                                                        // ' building class setup '''
                                                                        b = new clsBuilding();
                                                                        b.bHeight = bHeight;
                                                                        b.bLength = bLength;
                                                                        b.bWidth = bWidth;
                                                                        b.e1Overhang = e1GableOverhang;
                                                                        b.e3Overhang = e3GableOverhang;
                                                                        b.s2Overhang = s2EaveOverhang;
                                                                        b.s4Overhang = s4EaveOverhang;
                                                                        b.e1Extension = e1GableExtension;
                                                                        b.e3Extension = e3GableExtension;
                                                                        b.s2Extension = s2EaveExtension;
                                                                        b.s4Extension = s4EaveExtension;
                                                                        b.rPitch = rPitch;
                                                                        b.rShape = rShape;
                                                                        b.Gutters = Gutters;
                                                                        b.wPanelShape = wShape;
                                                                        b.wPanelColor = wColor;
                                                                        b.wPanelType = wType;
                                                                        b.rPanelShape = pShape;
                                                                        b.rPanelType = pType;
                                                                        b.rPanelColor = rColor;
                                                                        b.RakeTrimColor = RakeTrimColor;
                                                                        b.OutsideCornerTrimColor = OutsideCornerTrimColor;
                                                                        // check for base trim
                                                                        if (((BaseTrimColor == "None")
                                                                                    || (BaseTrimColor == ""))) {
                                                                            b.BaseTrim = false;
                                                                        }
                                                                        else {
                                                                            b.BaseTrim = true;
                                                                        }

                                                                        // set panel overage along building length
                                                                        b.bLengthRoofPanelOverage = ((Application.WorksheetFunction.RoundUp((((bLength * 12)
                                                                                        + (e1GableOverhang + e3GableOverhang)) / (3 * 12)), 0) * (3 * 12))
                                                                                    - ((bLength * 12)
                                                                                    + (e1GableOverhang + e3GableOverhang)));
                                                                        if ((e1GableOverhangSoffit == true)) {
                                                                            b.e1GableOverhangSoffit = true;
                                                                        }

                                                                        if ((e3GableOverhangSoffit == true)) {
                                                                            b.e3GableOverhangSoffit = true;
                                                                        }

                                                                        if ((s2EaveOverhangSoffit == true)) {
                                                                            b.s2EaveOverhangSoffit = true;
                                                                        }

                                                                        if ((s4EaveOverhangSoffit == true)) {
                                                                            b.s4EaveOverhangSoffit = true;
                                                                        }

                                                                        if ((e1GableExtensionSoffit == true)) {
                                                                            b.e1GableExtensionSoffit = true;
                                                                        }

                                                                        if ((e3GableExtensionSoffit == true)) {
                                                                            b.e3GableExtensionSoffit = true;
                                                                        }

                                                                        if ((s2EaveExtensionSoffit == true)) {
                                                                            b.s2EaveExtensionSoffit = true;
                                                                        }

                                                                        if ((s4EaveExtensionSoffit == true)) {
                                                                            b.s4EaveExtensionSoffit = true;
                                                                        }

                                                                        // With...
                                                                        // s2 eave extension
                                                                        if (EstSht.Range) {
                                                                            "s2_EaveExtensionPitch".Value = "Match Roof";
                                                                            b.s2ExtensionPitch = rPitch;
                                                                        }
                                                                        else {
                                                                            b.s2ExtensionPitch = EstSht.Range;
                                                                            "s2_EaveExtensionPitch".Value;
                                                                        }

                                                                        // s4 eave extension
                                                                        if (EstSht.Range) {
                                                                            "s4_EaveExtensionPitch".Value = "Match Roof";
                                                                            b.s4ExtensionPitch = rPitch;
                                                                        }
                                                                        else {
                                                                            b.s4ExtensionPitch = EstSht.Range;
                                                                            "s4_EaveExtensionPitch".Value;
                                                                        }

                                                                        // '' Liner Panels Section
                                                                        if (((b.LinerPanels("e1") == "8'")
                                                                                    || ((b.LinerPanels("e1") == "Full Height")
                                                                                    || ((b.LinerPanels("e3") == "8'")
                                                                                    || ((b.LinerPanels("e3") == "Full Height")
                                                                                    || ((b.LinerPanels("s2") == "8'")
                                                                                    || ((b.LinerPanels("s2") == "Full Height")
                                                                                    || ((b.LinerPanels("s4") == "8'")
                                                                                    || (b.LinerPanels("s4") == "Full Height"))))))))) {
                                                                            LinerPanelsSection = true;
                                                                        }

                                                                        for (FOCell in Range(EstSht.Range, "pDoorCell1", EstSht.Range, "pDoorCell12")) {
                                                                            // if cell isn't hidden, door size is entered
                                                                            if (((FOCell.EntireRow.Hidden == false)
                                                                                        && (FOCell.offset(0, 1).Value != ""))) {
                                                                                // add size to perimeter
                                                                                if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                    NetPDoorWidth = (NetPDoorWidth + 3);
                                                                                }
                                                                                else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                    NetPDoorWidth = (NetPDoorWidth + 4);
                                                                                }

                                                                            }

                                                                        }

                                                                        // Overhead Doors
                                                                        for (FOCell in Range(EstSht.Range, "OHDoorCell1", EstSht.Range, "OHDoorCell12")) {
                                                                            // if cell isn't hidden, door width is entered
                                                                            if (((FOCell.EntireRow.Hidden == false)
                                                                                        && (FOCell.offset(0, 1).Value != ""))) {
                                                                                // add size to perimeter
                                                                                NetOHDoorWidth = (NetOHDoorWidth + FOCell.offset(0, 1).Value);
                                                                            }

                                                                        }

                                                                        // '' Bays
                                                                        BayCount = EstSht.Range;
                                                                        "BayNum".Value;
                                                                    }

                                                                    // standard undercut
                                                                    Undercut = 4.25;
                                                                    // roof pitch string for product names
                                                                    PitchString = (rPitch + ":12");
                                                                    if (((bWidth == 0)
                                                                                || ((rPitch == 0)
                                                                                || ((rShape == "")
                                                                                || ((pShape == "")
                                                                                || ((pType == "")
                                                                                || ((pType == "")
                                                                                || (rColor == "")))))))) {
                                                                        /* Warning! GOTO is not Implemented */}

                                                                    // '''' Roof Pitch Hypotenuse (inches per ft width)
                                                                    RoofPitchHypot = Sqr((rPitch
                                                                                    | ((2 + 12)
                                                                                    | 2)));
                                                                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                                                                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                                                                    // roof length
                                                                    RoofLength = (bLength
                                                                                + ((e1GableOverhang / 12)
                                                                                + (e3GableOverhang / 12)));
                                                                    if ((rShape == "Gable")) {
                                                                        // '''''''''''' Panel Length '''''''''''''''''
                                                                        // normal roof rafter length (inches)
                                                                        RafterLength = ((bWidth / 2)
                                                                                    * RoofPitchHypot);
                                                                        b.RafterLength = RafterLength;
                                                                        // B.RafterLength = RafterLength - Undercut
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' sidewall 2 roof panels'''''
                                                                        // 'add overhang and undercut to rafter length
                                                                        s2RafterSheetLength = (RafterLength
                                                                                    + (s2EaveOverhang - Undercut));
                                                                        b.s2RafterSheetLength = s2RafterSheetLength;
                                                                        // s2RafterLength = RafterLength + s2EaveOverhang - Undercut+standardeaveoverhang
                                                                        // generate sidewall 2 panels
                                                                        s2RoofPanels = new Collection();
                                                                        RoofPanelGen(s2RoofPanels, s2RafterSheetLength, s2EaveOverhang, RoofLength, rShape);
                                                                        // add qualities
                                                                        for (RoofPanel in s2RoofPanels) {
                                                                            RoofPanel.PanelShape = pShape;
                                                                            RoofPanel.PanelType = pType;
                                                                            RoofPanel.PanelColor = rColor;
                                                                        }

                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' sidewall 4 roof panels'''''
                                                                        // add overhang to rafter length
                                                                        // 'add overhang and undercut to rafter length
                                                                        s4RafterSheetLength = (RafterLength
                                                                                    + (s4EaveOverhang - Undercut));
                                                                        b.s4RafterSheetLength = s4RafterSheetLength;
                                                                        // s4RafterLength = RafterLength + s4EaveOverhang - Undercut+standardeaveoverhang
                                                                        // generate sidewall 2 panels
                                                                        s4RoofPanels = new Collection();
                                                                        RoofPanelGen(s4RoofPanels, s4RafterSheetLength, s4EaveOverhang, RoofLength, rShape);
                                                                        // add qualities
                                                                        for (RoofPanel in s4RoofPanels) {
                                                                            RoofPanel.PanelShape = pShape;
                                                                            RoofPanel.PanelType = pType;
                                                                            RoofPanel.PanelColor = rColor;
                                                                        }

                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate ridge cap qty
                                                                        RidgeCapQty = Application.WorksheetFunction.RoundUp(((bLength
                                                                                        + ((b.e1Overhang / 12)
                                                                                        + ((b.e3Overhang / 12)
                                                                                        + ((b.e1Extension / 12)
                                                                                        + (b.e3Extension / 12)))))
                                                                                        / 3), 0);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Panels
                                                                        SidewallPanelGen(s2SidewallPanels, "s2", b);
                                                                        SidewallPanelGen(s4SidewallPanels, "s4", b);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Endwall Panels
                                                                        EndwallPanelGen(e1EndwallPanels, "e1", b);
                                                                        EndwallPanelGen(e3EndwallPanels, "e3", b);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' For Single Slope
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                    }
                                                                    else if ((rShape == "Single Slope")) {
                                                                        // '''''''''''' Panel Length '''''''''''''''''
                                                                        // '' Note: This calculation was preformed differently in version 1.5 and earlier
                                                                        // normal roof rafter length (in)
                                                                        RafterLength = (bWidth * RoofPitchHypot);
                                                                        b.RafterLength = RafterLength;
                                                                        // '''' sidewall 2 '''''
                                                                        // 'add overhang rafter length
                                                                        // also add in s4 overhang since it just extends the rafter length
                                                                        s2RafterSheetLength = (RafterLength
                                                                                    + (s2EaveOverhang + s4EaveOverhang));
                                                                        b.s2RafterSheetLength = s2RafterSheetLength;
                                                                        // generate sidewall 2 panels
                                                                        s2RoofPanels = new Collection();
                                                                        RoofPanelGen(s2RoofPanels, s2RafterSheetLength, s2EaveOverhang, RoofLength, rShape);
                                                                        // add qualities
                                                                        for (RoofPanel in s2RoofPanels) {
                                                                            RoofPanel.PanelShape = pShape;
                                                                            RoofPanel.PanelType = pType;
                                                                            RoofPanel.PanelColor = rColor;
                                                                        }

                                                                        // blank sidewall 4 collection
                                                                        s4RoofPanels = new Collection();
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Panels
                                                                        SidewallPanelGen(s2SidewallPanels, "s2", b);
                                                                        SidewallPanelGen(s4SidewallPanels, "s4", b);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Endwall Panels
                                                                        EndwallPanelGen(e1EndwallPanels, "e1", b);
                                                                        EndwallPanelGen(e3EndwallPanels, "e3", b);
                                                                    }

                                                                    // '''''''''''''''''''''''''''''''''''''''''Liner Panels
                                                                    LinerPanelGen(e1LinerPanels, b, "e1");
                                                                    LinerPanelGen(e3LinerPanels, b, "e3");
                                                                    LinerPanelGen(s2LinerPanels, b, "s2");
                                                                    LinerPanelGen(s4LinerPanels, b, "s4");
                                                                    LinerPanelGen(RoofLinerPanels, b, "Roof");
                                                                    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim
                                                                    // '''
                                                                    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Rake Trim (For Either roof Type)
                                                                    // '''
                                                                    RakeTrimPieces = new Collection();
                                                                    // calculate net rafter length (without factoring in undercut)
                                                                    if ((rShape == "Single Slope")) {
                                                                        NetRafterLength = b.RafterLength;
                                                                    }
                                                                    else if ((rShape == "Gable")) {
                                                                        NetRafterLength = (b.RafterLength * 2);
                                                                    }

                                                                    // add extension/overhang
                                                                    NetRafterLength = (NetRafterLength
                                                                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength));
                                                                    // two endwalls, two trim lengths
                                                                    NetRafterLength = (NetRafterLength * 2);
                                                                    // pass to calc
                                                                    TrimPieceCalc(RakeTrimPieces, NetRafterLength, "Rake", ,, b);
                                                                    // '''
                                                                    // add qualities
                                                                    for (TrimPiece in RakeTrimPieces) {
                                                                        TrimPiece.tShape = pShape;
                                                                        // '' rake trim is always roof panel shape
                                                                        TrimPiece.Color = RakeTrimColor;
                                                                        // increase 20'4" pieces to 21'
                                                                        if ((TrimPiece.tLength == 244)) {
                                                                            TrimPiece.tLength = (21 * 12);
                                                                            TrimPiece.tMeasurement = ImperialMeasurementFormat(TrimPiece.tLength);
                                                                        }

                                                                    }

                                                                    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Eave Trim
                                                                    // '''
                                                                    EaveTrimPieces = new Collection();
                                                                    // ''''''''''''''''''''''''' Sidewall 2
                                                                    // '' Outside Eave Trim (Along outside of Eave)
                                                                    // start with building length
                                                                    s2EaveTrimLength = (bLength * 12);
                                                                    // add endwall overhangs or extensions
                                                                    s2EaveTrimLength = (s2EaveTrimLength
                                                                                + (e1GableExtension
                                                                                + (e1GableOverhang
                                                                                + (e3GableExtension + e3GableOverhang))));
                                                                    // '' Inside Eave Trim (Where eave meets sidewalls)
                                                                    // ' Check if inside eave trim is needed
                                                                    // First check if there's an s2 extension/overhang
                                                                    if (((s2EaveExtension != 0)
                                                                                || (s2EaveOverhang != 4.25))) {
                                                                        // '' s2 eave extension/overhang present                                               '''
                                                                        // Add inside Trim if no eave soffit
                                                                        if (((s2EaveOverhangSoffit == false)
                                                                                    && (s2EaveExtensionSoffit == false))) {
                                                                            //  add additional trim along building length
                                                                            s2EaveTrimLength = (s2EaveTrimLength
                                                                                        + (bLength * 12));
                                                                        }

                                                                    }

                                                                    // ''''''''''''''''''''''''' Sidewall 4, generate trim collections
                                                                    // '' Outside Eave Trim (Along outside of Eave)
                                                                    // start with building length
                                                                    s4EaveTrimLength = (bLength * 12);
                                                                    // add endwall overhangs or extensions
                                                                    s4EaveTrimLength = (s4EaveTrimLength
                                                                                + (e1GableExtension
                                                                                + (e1GableOverhang
                                                                                + (e3GableExtension + e3GableOverhang))));
                                                                    // '' Inside Eave Trim If Needed (Where eave meets sidewalls)
                                                                    // Additional condition due 4.25 s4 standard overhang on gable and 0 s4 standard overhang on single slope
                                                                    if (((s4EaveExtension != 0)
                                                                                || ((s4EaveOverhang != 4.25)
                                                                                && (s4EaveOverhang != 0)))) {
                                                                        // '' s4 eave extension/overhang present                                               '''
                                                                        // Add inside Trim if no eave soffit
                                                                        if (((s4EaveExtensionSoffit == false)
                                                                                    && (s4EaveOverhangSoffit == false))) {
                                                                            //  add additional trim along building length
                                                                            s4EaveTrimLength = (s4EaveTrimLength
                                                                                        + (bLength * 12));
                                                                        }

                                                                    }

                                                                    if ((rShape == "Gable")) {
                                                                        // '''generate trim piece collection
                                                                        TrimPieceCalc(EaveTrimPieces, (s2EaveTrimLength + s4EaveTrimLength), "Short Eave", PitchString);
                                                                    }
                                                                    else if ((rShape == "Single Slope")) {
                                                                        // seperate high side and short side eave trim collections
                                                                        TrimPieceCalc(EaveTrimPieces, s2EaveTrimLength, "Short Eave", PitchString);
                                                                        TrimPieceCalc(EaveTrimPieces, s4EaveTrimLength, "High Eave", PitchString);
                                                                    }

                                                                    // add qualities
                                                                    for (TrimPiece in EaveTrimPieces) {
                                                                        TrimPiece.tShape = "R-Loc";
                                                                        TrimPiece.Color = EaveTrimColor;
                                                                    }

                                                                    // '''
                                                                    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Outside Corner Trim
                                                                    // '''
                                                                    OutsideCornerTrimPieces = new Collection();
                                                                    if ((b.rShape == "Gable")) {
                                                                        // assume complete, exclude intersections if needed
                                                                        NetCornerLength = (b.bHeight * (4 * 12));
                                                                        if (((b.WallStatus("e1") != "Include")
                                                                                    && (b.WallStatus("s2") != "Include"))) {
                                                                            NetCornerLength = (NetCornerLength
                                                                                        - (b.bHeight * 12));
                                                                        }

                                                                        if (((b.WallStatus("e1") != "Include")
                                                                                    && (b.WallStatus("s4") != "Include"))) {
                                                                            NetCornerLength = (NetCornerLength
                                                                                        - (b.bHeight * 12));
                                                                        }

                                                                        if (((b.WallStatus("e3") != "Include")
                                                                                    && (b.WallStatus("s2") != "Include"))) {
                                                                            NetCornerLength = (NetCornerLength
                                                                                        - (b.bHeight * 12));
                                                                        }

                                                                        if (((b.WallStatus("e3") != "Include")
                                                                                    && (b.WallStatus("s4") != "Include"))) {
                                                                            NetCornerLength = (NetCornerLength
                                                                                        - (b.bHeight * 12));
                                                                        }
                                                                        else if ((rShape == "Single Slope")) {
                                                                            // sidewall 2 corners + s4 corners
                                                                            NetCornerLength = ((b.bHeight * (12 * 2))
                                                                                        + (b.HighSideEaveHeight * 2));
                                                                            if (((b.WallStatus("s2") != "Include")
                                                                                        && (b.WallStatus("e1") != "Include"))) {
                                                                                NetCornerLength = (NetCornerLength
                                                                                            - (b.bHeight * 12));
                                                                            }

                                                                            if (((b.WallStatus("s2") != "Include")
                                                                                        && (b.WallStatus("e3") != "Include"))) {
                                                                                NetCornerLength = (NetCornerLength
                                                                                            - (b.bHeight * 12));
                                                                            }

                                                                            if (((b.WallStatus("s4") != "Include")
                                                                                        && (b.WallStatus("e1") != "Include"))) {
                                                                                NetCornerLength = (NetCornerLength - b.HighSideEaveHeight);
                                                                            }

                                                                            if (((b.WallStatus("s4") != "Include")
                                                                                        && (b.WallStatus("e3") != "Include"))) {
                                                                                NetCornerLength = (NetCornerLength - b.HighSideEaveHeight);
                                                                            }

                                                                            // generate trim collection
                                                                            TrimPieceCalc(OutsideCornerTrimPieces, NetCornerLength, "Outside Corner", ,, b);
                                                                            // add qualities
                                                                            for (TrimPiece in OutsideCornerTrimPieces) {
                                                                                TrimPiece.tShape = "R-Loc";
                                                                                TrimPiece.Color = OutsideCornerTrimColor;
                                                                            }

                                                                            // '''
                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Base Trim
                                                                            // '''
                                                                            if ((b.BaseTrim == true)) {
                                                                                BaseTrimPieces = new Collection();
                                                                                // Perimeter
                                                                                if ((b.WallStatus("e1") == "Include")) {
                                                                                    NetBaseTrimLength = b.bWidth;
                                                                                }

                                                                                if ((b.WallStatus("e3") == "Include")) {
                                                                                    NetBaseTrimLength = (NetBaseTrimLength + b.bWidth);
                                                                                }

                                                                                if ((b.WallStatus("s2") == "Include")) {
                                                                                    NetBaseTrimLength = (NetBaseTrimLength + b.bLength);
                                                                                }

                                                                                if ((b.WallStatus("s4") == "Include")) {
                                                                                    NetBaseTrimLength = (NetBaseTrimLength + b.bLength);
                                                                                }

                                                                                // subtract width of OH doors, P doors
                                                                                NetBaseTrimLength = (NetBaseTrimLength
                                                                                            - (NetOHDoorWidth + NetPDoorWidth));
                                                                                NetBaseTrimLength = (NetBaseTrimLength * 12);
                                                                                // generate trim collection
                                                                                TrimPieceCalc(BaseTrimPieces, NetBaseTrimLength, "Base");
                                                                                // add qualities
                                                                                for (TrimPiece in BaseTrimPieces) {
                                                                                    TrimPiece.tShape = "R-Loc";
                                                                                    TrimPiece.Color = BaseTrimColor;
                                                                                }

                                                                            }

                                                                            // '''
                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Wainscot Trim
                                                                            // '''
                                                                            if ((EstSht.Range("Wainscot").Value == "Yes")) {
                                                                                WainscotTrimPieces = new Collection();
                                                                                // ''''''''''''''''''''''''''''''''''''''''Standard Wainscot Trim'''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                // Endwall 1 - check if standard
                                                                                if (((b.Wainscot("e1") != "None")
                                                                                            && (b.Wainscot("e1").IndexOf("Standard", 0) + 1))) {
                                                                                    // loop through Pdoors for door widths on Endwall 1
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Endwall 1")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    // loop through OHdoors for door widths on Endwall 1
                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Endwall 1")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = b.bWidth;
                                                                                }

                                                                                // Endwall 3 - check if standard
                                                                                if (((b.Wainscot("e3") != "None")
                                                                                            && (b.Wainscot("e3").IndexOf("Standard", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Endwall 3")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Endwall 3")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bWidth + NetWainscotTrimLength);
                                                                                }

                                                                                // Sidewall 2 - check if standard
                                                                                if (((b.Wainscot("s2") != "None")
                                                                                            && (b.Wainscot("s2").IndexOf("Standard", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Sidewall 2")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Sidewall 2")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bLength + NetWainscotTrimLength);
                                                                                }

                                                                                // Sidewall 4 - check if standard
                                                                                if (((b.Wainscot("s4") != "None")
                                                                                            && (b.Wainscot("s4").IndexOf("Standard", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Sidewall 4")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Sidewall 4")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bLength + NetWainscotTrimLength);
                                                                                }

                                                                                // subtract width of OH doors and Pdoors where there is standard wainscot
                                                                                NetWainscotTrimLength = (NetWainscotTrimLength - TempDoorWidth);
                                                                                // convert
                                                                                NetWainscotTrimLength = (NetWainscotTrimLength * 12);
                                                                                // generate trim collection
                                                                                TrimPieceCalc(WainscotTrimPieces, NetWainscotTrimLength, "Standard Wainscot");
                                                                                // Reset Variables
                                                                                TempDoorWidth = 0;
                                                                                NetWainscotTrimLength = 0;
                                                                                // ''''''''''''''''''''''''''''''''''''''''Masonry Wainscot Trim'''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                // Endwall 1 - check if standard
                                                                                if (((b.Wainscot("e1") != "None")
                                                                                            && (b.Wainscot("e1").IndexOf("Masonry", 0) + 1))) {
                                                                                    // loop through Pdoors for door widths on Endwall 1
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Endwall 1")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    // loop through OHdoors for door widths on Endwall 1
                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Endwall 1")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = b.bWidth;
                                                                                }

                                                                                // Endwall 3 - check if standard
                                                                                if (((b.Wainscot("e3") != "None")
                                                                                            && (b.Wainscot("e3").IndexOf("Masonry", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Endwall 3")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Endwall 3")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bWidth + NetWainscotTrimLength);
                                                                                }

                                                                                // Sidewall 2 - check if standard
                                                                                if (((b.Wainscot("s2") != "None")
                                                                                            && (b.Wainscot("s2").IndexOf("Masonry", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Sidewall 2")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Sidewall 2")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bLength + NetWainscotTrimLength);
                                                                                }

                                                                                // Sidewall 4 - check if standard
                                                                                if (((b.Wainscot("s4") != "None")
                                                                                            && (b.Wainscot("s4").IndexOf("Masonry", 0) + 1))) {
                                                                                    for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                                                        // if cell isn't hidden, door size is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 2).Value == "Sidewall 4")))) {
                                                                                            // add size to perimeter
                                                                                            if ((FOCell.offset(0, 1).Value == "3070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 3);
                                                                                            }
                                                                                            else if ((FOCell.offset(0, 1).Value == "4070")) {
                                                                                                TempDoorWidth = (TempDoorWidth + 4);
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                                                        // if cell isn't hidden, door width is entered
                                                                                        if (((FOCell.EntireRow.Hidden == false)
                                                                                                    && ((FOCell.offset(0, 1).Value != "")
                                                                                                    && (FOCell.offset(0, 3).Value == "Sidewall 4")))) {
                                                                                            // add size to perimeter
                                                                                            TempDoorWidth = (TempDoorWidth + FOCell.offset(0, 1).Value);
                                                                                        }

                                                                                    }

                                                                                    NetWainscotTrimLength = (b.bLength + NetWainscotTrimLength);
                                                                                }

                                                                                // subtract width of OH doors and Pdoors where there is Masonry Wainscot
                                                                                NetWainscotTrimLength = (NetWainscotTrimLength - TempDoorWidth);
                                                                                // convert
                                                                                NetWainscotTrimLength = (NetWainscotTrimLength * 12);
                                                                                // generate trim collection
                                                                                TrimPieceCalc(WainscotTrimPieces, NetWainscotTrimLength, "Masonry Wainscot");
                                                                                // Add qualities
                                                                                for (TrimPiece in WainscotTrimPieces) {
                                                                                    TrimPiece.tShape = "R-Loc";
                                                                                    TrimPiece.Color = EstSht.Range("Wainscot_tColor").Value;
                                                                                }

                                                                            }

                                                                            // '''
                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  Gutters & Downspouts
                                                                            // '''
                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Gutters
                                                                            if ((b.Gutters == true)) {
                                                                                GutterPieces = new Collection();
                                                                                // ''calculate gutter length
                                                                                // start with building length
                                                                                NetGutterLength = (bLength * 12);
                                                                                // add endwall overhangs or extensions
                                                                                NetGutterLength = (NetGutterLength
                                                                                            + (e1GableExtension
                                                                                            + (e1GableOverhang
                                                                                            + (e3GableExtension + e3GableOverhang))));
                                                                                // if a gable roof, multiply by 2 to account for gutter along both sidewalls
                                                                                if ((rShape == "Gable")) {
                                                                                    NetGutterLength = (NetGutterLength * 2);
                                                                                }

                                                                                // generate gutter piece collection (done in the same way as trim, so using the trim piece sub)
                                                                                TrimPieceCalc(GutterPieces, NetGutterLength, "Gutter", PitchString);
                                                                                // add qualities
                                                                                for (GutterPiece in GutterPieces) {
                                                                                    GutterPiece.tShape = "R-Loc";
                                                                                    GutterPiece.Color = GutterColor;
                                                                                    if ((GutterPiece.tLength == 244)) {
                                                                                        GutterPiece.tLength = (21 * 12);
                                                                                        GutterPiece.tMeasurement = ImperialMeasurementFormat(GutterPiece.tLength);
                                                                                    }

                                                                                }

                                                                                // end caps
                                                                                if ((rShape == "Gable")) {
                                                                                    GutterEndCapQty = 4;
                                                                                }
                                                                                else if ((rShape == "Single Slope")) {
                                                                                    GutterEndCapQty = 2;
                                                                                }

                                                                                // 'straps (same as the qty of bottom roof sheets)
                                                                                // first, building length in ft
                                                                                GutterStrapQty = (((bLength * 12)
                                                                                            + (e1GableExtension
                                                                                            + (e1GableOverhang
                                                                                            + (e3GableExtension + e3GableOverhang))))
                                                                                            / 12);
                                                                                // divide by 3', round up
                                                                                GutterStrapQty = Application.WorksheetFunction.RoundUp((GutterStrapQty / 3), 0);
                                                                                // multiply by 2 if a gable roof
                                                                                if ((rShape == "Gable")) {
                                                                                    GutterStrapQty = (GutterStrapQty * 2);
                                                                                }

                                                                                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Downspouts
                                                                                DownspoutPieces = new Collection();
                                                                                // find downspout quantity
                                                                                if ((rShape == "Gable")) {
                                                                                    DownspoutQty = ((BayCount + 1)
                                                                                                * 2);
                                                                                }
                                                                                else if ((rShape == "Single Slope")) {
                                                                                    DownspoutQty = (BayCount + 1);
                                                                                }

                                                                                // find first kickout piece height
                                                                                DownspoutPiece = new clsTrim();
                                                                                DownspoutPiece.tType = "Square Downspout W/ Kickout";
                                                                                switch (bHeight) {
                                                                                        break;
                                                                                }

                                                                                122;
                                                                                DownspoutPiece.tMeasurement = "10'2""";
                                                                                DownspoutPiece.tLength = 122;
                                                                                146;
                                                                                DownspoutPiece.tMeasurement = "12'2""";
                                                                                DownspoutPiece.tLength = 146;
                                                                                170;
                                                                                DownspoutPiece.tMeasurement = "14'2""";
                                                                                DownspoutPiece.tLength = 170;
                                                                                194;
                                                                                DownspoutPiece.tMeasurement = "16'2""";
                                                                                DownspoutPiece.tLength = 194;
                                                                                218;
                                                                                DownspoutPiece.tMeasurement = "18'2""";
                                                                                DownspoutPiece.tLength = 218;
                                                                                244;
                                                                                DownspoutPiece.tMeasurement = "20'4""";
                                                                                DownspoutPiece.tLength = 244;
                                                                            }
                                                                            else {
                                                                                // greater than 20'4
                                                                                RemainingHeight = ((bHeight * 12)
                                                                                            - 242);
                                                                                DownspoutPiece.tMeasurement = "20'4""";
                                                                                DownspoutPiece.tLength = 244;
                                                                            }

                                                                            // set quantity, shape, color
                                                                            DownspoutPiece.Quantity = DownspoutQty;
                                                                            DownspoutPiece.tShape = "R-Loc";
                                                                            DownspoutPiece.Color = DownspoutColor;
                                                                            DownspoutPieces.Add;
                                                                            DownspoutPiece;
                                                                            // find the rest of pieces
                                                                            if ((RemainingHeight != 0)) {
                                                                                TrimPieceCalc(DownspoutPieces, RemainingHeight, "Downspout", ,, DownspoutQty);
                                                                            }

                                                                            // update downspout without kickout shape
                                                                            for (DownspoutPiece in DownspoutPieces) {
                                                                                if ((DownspoutPiece.tType == "Square Downspout W/O Kickout")) {
                                                                                    DownspoutPiece.tShape = "R-Loc";
                                                                                    DownspoutPiece.Color = DownspoutColor;
                                                                                }

                                                                            }

                                                                            // ''' Straps
                                                                            //  find straps per downspout
                                                                            // reset building height. first strap as at 12'
                                                                            RemainingHeight = (bHeight - 12);
                                                                            DownspoutStrapQty = 1;
                                                                            while ((RemainingHeight > 0)) {
                                                                                // strap every 7'
                                                                                RemainingHeight = (RemainingHeight - 7);
                                                                                DownspoutStrapQty = (DownspoutStrapQty + 1);
                                                                            }

                                                                            DownspoutStrapQty = (DownspoutStrapQty * DownspoutQty);
                                                                            // ''' Pop Rivits
                                                                            // # of rivits = (gutter piece qty *2*10) rounded up to 100
                                                                            for (GutterPiece in GutterPieces) {
                                                                                PopRivitQty = (PopRivitQty + GutterPiece.Quantity);
                                                                            }

                                                                            PopRivitQty = (PopRivitQty * (2 * 10));
                                                                            // round up to nearest 100
                                                                            PopRivitQty = (Application.WorksheetFunction.RoundUp((PopRivitQty / 100), 0) * 100);
                                                                        }

                                                                        // '''
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Translucent Wall Panels & Skylights
                                                                        // '''
                                                                        // '' skylight qty
                                                                        // 1 per skylight
                                                                        if ((EstSht.Range("SkylightQty").Value > 0)) {
                                                                            SkylightPanelQty = EstSht.Range("SkylightQty").Value;
                                                                        }

                                                                        // 'translucent wall panel qty
                                                                        // half per translucent panel
                                                                        SkylightPanelQty = (SkylightPanelQty + Application.WorksheetFunction.RoundUp((EstSht.Range("TranslucentWallPanelQty").Value / 2), 0));
                                                                        // '''
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhang Soffits
                                                                        // '''
                                                                        e1SoffitPanels = new Collection();
                                                                        e1SoffitTrim = new Collection();
                                                                        s2SoffitPanels = new Collection();
                                                                        s2SoffitTrim = new Collection();
                                                                        e3SoffitPanels = new Collection();
                                                                        e3SoffitTrim = new Collection();
                                                                        s4SoffitPanels = new Collection();
                                                                        s4SoffitTrim = new Collection();
                                                                        // e1 Gable overhang
                                                                        if ((e1GableOverhangSoffit == true)) {
                                                                            SoffitGen(e1SoffitPanels, e1SoffitTrim, "e1_GableOverhang", b, s2RoofPanels, s4RoofPanels);
                                                                            // update soffit trim color, shape
                                                                            for (SoffitTrim in e1SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("e1_GableOverhangSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("e1_GableOverhangSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // e3 Gable overhang
                                                                        if ((e3GableOverhangSoffit == true)) {
                                                                            SoffitGen(e3SoffitPanels, e3SoffitTrim, "e3_GableOverhang", b, s2RoofPanels, s4RoofPanels);
                                                                            // update soffit trim color, shape
                                                                            for (SoffitTrim in e3SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("e3_GableOverhangSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("e3_GableOverhangSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // s2 eave overhang
                                                                        if ((s2EaveOverhangSoffit == true)) {
                                                                            SoffitGen(s2SoffitPanels, s2SoffitTrim, "s2_EaveOverhang", b, s2RoofPanels, s4RoofPanels);
                                                                            // update soffit trim color, shape
                                                                            for (SoffitTrim in s2SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("s2_EaveOverhangSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("s2_EaveOverhangSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // s4 eave overhang
                                                                        if ((s4EaveOverhangSoffit == true)) {
                                                                            SoffitGen(s4SoffitPanels, s4SoffitTrim, "s4_EaveOverhang", b, s2RoofPanels, s4RoofPanels);
                                                                            // update soffit trim color, shape
                                                                            for (SoffitTrim in s4SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("s4_EaveOverhangSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("s4_EaveOverhangSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // '''
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Extensions
                                                                        // '''
                                                                        e1ExtensionPanels = new Collection();
                                                                        s2ExtensionPanels = new Collection();
                                                                        e3ExtensionPanels = new Collection();
                                                                        s4ExtensionPanels = new Collection();
                                                                        // ''' e1 Gable Extension
                                                                        if ((e1GableExtensionSection == true)) {
                                                                            ExtensionPanelGen(e1ExtensionPanels, b, "e1_GableExtension", s2RoofPanels, s4RoofPanels);
                                                                        }

                                                                        // 2x8 inside angle
                                                                        if ((e1GableExtensionSoffit == true)) {
                                                                            // add soffit panels to extension panel collection
                                                                            SoffitGen(e1SoffitPanels, e1SoffitTrim, "e1_GableExtension", b, s2RoofPanels, s4RoofPanels);
                                                                            for (SoffitPanel in e1SoffitPanels) {
                                                                                // add to extension panels
                                                                                e1ExtensionPanels.Add;
                                                                                SoffitPanel;
                                                                            }

                                                                            // consolodate duplicate panels
                                                                            DuplicateMaterialRemoval(e1ExtensionPanels, "Panel");
                                                                            // correct trim color, shape
                                                                            for (SoffitTrim in e1SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("e1_GableExtensionSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("e1_GableExtensionSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // ''' e3 Gable Extension
                                                                        if ((e3GableExtensionSection == true)) {
                                                                            ExtensionPanelGen(e3ExtensionPanels, b, "e3_GableExtension", s2RoofPanels, s4RoofPanels);
                                                                        }

                                                                        if ((e3GableExtensionSoffit == true)) {
                                                                            SoffitGen(e3SoffitPanels, e3SoffitTrim, "e3_GableExtension", b, s2RoofPanels, s4RoofPanels);
                                                                            // add soffit panels to extension panel collection
                                                                            for (SoffitPanel in e3SoffitPanels) {
                                                                                // add to extension panels
                                                                                e3ExtensionPanels.Add;
                                                                                SoffitPanel;
                                                                            }

                                                                            // consolodate duplicate panels
                                                                            DuplicateMaterialRemoval(e3ExtensionPanels, "Panel");
                                                                            // correct trim color, shape
                                                                            for (SoffitTrim in e3SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("e3_GableExtensionSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("e3_GableExtensionSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // s2 eave Extension
                                                                        if ((s2EaveExtensionSection == true)) {
                                                                            ExtensionPanelGen(s2ExtensionPanels, b, "s2_EaveExtension");
                                                                        }

                                                                        if ((s2EaveExtensionSoffit == true)) {
                                                                            SoffitGen(s2SoffitPanels, s2SoffitTrim, "s2_EaveExtension", b);
                                                                            // add soffit panels to extension panel collection
                                                                            for (SoffitPanel in s2SoffitPanels) {
                                                                                // add to extension panels
                                                                                s2ExtensionPanels.Add;
                                                                                SoffitPanel;
                                                                            }

                                                                            // consolodate duplicate panels
                                                                            DuplicateMaterialRemoval(s2ExtensionPanels, "Panel");
                                                                            // correct trim color, shape
                                                                            for (SoffitTrim in s2SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("s2_EaveExtensionSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("s2_EaveExtensionSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // s4 eave Extension
                                                                        if ((s4EaveExtensionSection == true)) {
                                                                            ExtensionPanelGen(s4ExtensionPanels, b, "s4_EaveExtension");
                                                                        }

                                                                        if ((s4EaveExtensionSoffit == true)) {
                                                                            SoffitGen(s4SoffitPanels, s4SoffitTrim, "s4_EaveExtension", b);
                                                                            // add soffit panels to extension panel collection
                                                                            for (SoffitPanel in s4SoffitPanels) {
                                                                                // add to extension panels
                                                                                s4ExtensionPanels.Add;
                                                                                SoffitPanel;
                                                                            }

                                                                            // consolodate duplicate panels
                                                                            DuplicateMaterialRemoval(s4ExtensionPanels, "Panel");
                                                                            // correct trim color, shape
                                                                            for (SoffitTrim in s4SoffitTrim) {
                                                                                if ((SoffitTrim.tType != "2x6 Outside Angle Trim")) {
                                                                                    SoffitTrim.tShape = EstSht.Range("s4_EaveExtensionSoffit").offset(0, 1).Value;
                                                                                }

                                                                                SoffitTrim.Color = EstSht.Range("s4_EaveExtensionSoffit").offset(0, 4).Value;
                                                                            }

                                                                        }

                                                                        // ''''''''''''''''''' 2x8 Outside Angle trim for gable extensions without soffit
                                                                        // With...
                                                                        // Generate 2x8 Inside angle trim
                                                                        if ((b.rShape == "Single Slope")) {
                                                                            if (((e1GableExtensionSection == true)
                                                                                        && (e1GableExtensionSoffit == false))) {
                                                                                e1InsideAngleTrim = new Collection();
                                                                                TrimPieceCalc(e1InsideAngleTrim, (b.s2RafterSheetLength
                                                                                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength)), "Inside Angle", ,, b);
                                                                            }

                                                                            if (((e3GableExtensionSection == true)
                                                                                        && (e3GableExtensionSoffit == false))) {
                                                                                e3InsideAngleTrim = new Collection();
                                                                                TrimPieceCalc(e3InsideAngleTrim, (b.s2RafterSheetLength
                                                                                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength)), "Inside Angle", ,, b);
                                                                            }

                                                                        }
                                                                        else if ((b.rShape == "Gable")) {
                                                                            if (((e1GableExtensionSection == true)
                                                                                        && (e1GableExtensionSoffit == false))) {
                                                                                e1InsideAngleTrim = new Collection();
                                                                                TrimPieceCalc(e1InsideAngleTrim, (b.s2RafterSheetLength
                                                                                                + (b.s4RafterSheetLength
                                                                                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength))), "Inside Angle", ,, b);
                                                                            }

                                                                            if (((e3GableExtensionSection == true)
                                                                                        && (e3GableExtensionSoffit == false))) {
                                                                                e3InsideAngleTrim = new Collection();
                                                                                TrimPieceCalc(e3InsideAngleTrim, (b.s2RafterSheetLength
                                                                                                + (b.s4RafterSheetLength
                                                                                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength))), "Inside Angle", ,, b);
                                                                            }

                                                                        }

                                                                        if (!(e1InsideAngleTrim == null)) {
                                                                            for (TrimPiece in e1InsideAngleTrim) {
                                                                                TrimPiece.Color = b.rPanelColor;
                                                                            }

                                                                        }

                                                                        if (!(e3InsideAngleTrim == null)) {
                                                                            for (TrimPiece in e3InsideAngleTrim) {
                                                                                TrimPiece.Color = b.rPanelColor;
                                                                            }

                                                                        }

                                                                        // '''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Fasteners
                                                                        // '''
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Roof Screws
                                                                        // calculate panel overlaps
                                                                        if ((rShape == "Single Slope")) {
                                                                            rOverlaps = (s2RoofPanels.Count - 1);
                                                                        }
                                                                        else if ((rShape == "Gable")) {
                                                                            rOverlaps = ((s2RoofPanels.Count - 1)
                                                                                        + (s4RoofPanels.Count - 1));
                                                                        }

                                                                        // ''extension Overlaps
                                                                        // rOverlaps = rOverlaps + (s2ExtensionPanels.Count - 1) + (s4ExtensionPanels.Count - 1)
                                                                        // generate roof screws
                                                                        RoofScrewGen(rTekScrewQty, rLapScrewQty, b, rOverlaps);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sidewall Screws
                                                                        // sidewall panel overlaps
                                                                        if ((s2SidewallPanels.Count > 0)) {
                                                                            sOverlaps = (s2SidewallPanels.Count - 1);
                                                                        }

                                                                        if ((s4SidewallPanels.Count > 0)) {
                                                                            sOverlaps = (sOverlaps
                                                                                        + (s4SidewallPanels.Count - 1));
                                                                        }

                                                                        // endwall overlaps
                                                                        eOverlaps = (b.e1WallPanelOverlaps + b.e3WallPanelOverlaps);
                                                                        // generate screws
                                                                        WallScrewGen(wTekScrewQty, wLapScrewQty, b, sOverlaps, eOverlaps);
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim Screws
                                                                        TrimScrews = new Collection();
                                                                        TrimScrewCalc(TrimScrews, RakeTrimPieces, b);
                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Screws
                                                                        if ((e1GableOverhangSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e1_GableOverhang", b);
                                                                        }

                                                                        if ((s2EaveOverhangSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s2_EaveOverhang", b);
                                                                        }

                                                                        if ((e3GableOverhangSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e3_GableOverhang", b);
                                                                        }

                                                                        if ((s4EaveOverhangSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s4_EaveOverhang", b);
                                                                        }

                                                                        if ((e1GableExtensionSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e1_GableExtension", b);
                                                                        }

                                                                        if ((s2EaveExtensionSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s2_EaveExtension", b);
                                                                        }

                                                                        if ((e3GableExtensionSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "e3_GableExtension", b);
                                                                        }

                                                                        if ((s4EaveExtensionSoffit == true)) {
                                                                            SoffitScrewCalc(SoffitScrewQty, SoffitScrewColor, "s4_EaveExtension", b);
                                                                        }

                                                                        // round up to nearest 250
                                                                        SoffitScrewQty = (Application.WorksheetFunction.RoundUp((SoffitScrewQty / 250), 0) * 250);
                                                                        // '''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Miscellaneous
                                                                        // '''
                                                                        // Butyl tape, inside closures, and outside closures
                                                                        MiscMaterialCalc(ButylTapeQty, InsideClosureQty, OutsideClosureQty, b, rOverlaps);
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''     MATERIALS LIST OUTPUT
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        TestingSub2(b);
                                                                        // delete old output sheet
                                                                        Application.DisplayAlerts = false;
                                                                        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
                                                                            if ((ThisWorkbook.Sheets(N).Name == "Employee Materials List")) {
                                                                                ThisWorkbook.Sheets(N).Delete;
                                                                                break;
                                                                            }

                                                                        }

                                                                        Application.DisplayAlerts = true;
                                                                        MatShtTmp.Copy;
                                                                        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Wall Drawings");
                                                                        MatSht = ThisWorkbook.Sheets("MaterialsListTmp (2)");
                                                                        // rename
                                                                        MatSht.Name = "Employee Materials List";
                                                                        MatSht.Visible = xlSheetVisible;
                                                                        // combined material collections
                                                                        PanelCollection = new Collection();
                                                                        TrimCollection = new Collection();
                                                                        MiscCollection = new Collection();
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' OUTPUT '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                        // With...
                                                                        // ''''''''''''''''''''''''''''''''''''' Roof Panels ''''''''''''''''''''''''''''''''''
                                                                        // '''''' Sidewall 2 Roof Panels ''''''''
                                                                        MatListSectionWrite(MatSht, MatSht.Range, "s2_RoofSheetQtyCell1", s2RoofPanels, "Panel");
                                                                        // '''''' Sidewall 4  Roof Panels ''''''''
                                                                        // delete if a single slope
                                                                        if ((rShape == "Single Slope")) {
                                                                            MatSht.Range;
                                                                            "s4_RoofSheetQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                            // output for gable roof
                                                                        }
                                                                        else if ((rShape == "Gable")) {
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "s4_RoofSheetQtyCell1", s4RoofPanels, "Panel");
                                                                        }

                                                                        // '''''''''''''''''''''''''''''''''''''''' ridge caps
                                                                        // delete if a single slope
                                                                        if ((rShape == "Single Slope")) {
                                                                            MatSht.Range;
                                                                            "Roof_RidgeCapQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                            // output for gable roof
                                                                        }
                                                                        else if ((rShape == "Gable")) {
                                                                            WriteCell = MatSht.Range;
                                                                            "Roof_RidgeCapQtyCell1";
                                                                            WriteCell.Value = RidgeCapQty;
                                                                            WriteCell.offset(0, 1).Value = ("Formed Ridge Cap " + PitchString);
                                                                            WriteCell.offset(0, 3).Value = "3'";
                                                                            WriteCell.offset(0, 4).Value = rColor;
                                                                        }

                                                                        // '''''''''''''''''''''''''''''''''''''''''''''' sidewall panels
                                                                        // ' sidewall 2
                                                                        MatListSectionWrite(MatSht, MatSht.Range, "s2_SidewallSheetQtyCell1", s2SidewallPanels, "Panel");
                                                                        // ' sidewall 4
                                                                        MatListSectionWrite(MatSht, MatSht.Range, "s4_SidewallSheetQtyCell1", s4SidewallPanels, "Panel");
                                                                        // '''''''''''''''''''''''''''''''''''' endwall panels
                                                                        // '' endwall #1
                                                                        MatListSectionWrite(MatSht, MatSht.Range, "e1_EndwallSheetQtyCell1", e1EndwallPanels, "Panel");
                                                                        // '' endwall #3
                                                                        MatListSectionWrite(MatSht, MatSht.Range, "e3_EndwallSheetQtyCell1", e3EndwallPanels, "Panel");
                                                                        // '''''''''''''''''''''''''''''''''''''''' Liner Panels
                                                                        if ((LinerPanelsSection == false)) {
                                                                            Range(MatSht.Range, "e1_LinerPanelsQtyCell1".offset(-5, 0), MatSht.Range, "Roof_LinerPanelsQtyCell1".offset(1, 0)).EntireRow.Delete;
                                                                        }
                                                                        else {
                                                                            // ''''''''''write liner panels
                                                                            WriteCell = MatSht.Range;
                                                                            "e1_LinerPanelsQtyCell1";
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "e1_LinerPanelsQtyCell1", e1LinerPanels, "Panel");
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "e3_LinerPanelsQtyCell1", e3LinerPanels, "Panel");
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "s2_LinerPanelsQtyCell1", s2LinerPanels, "Panel");
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "s4_LinerPanelsQtyCell1", s4LinerPanels, "Panel");
                                                                            MatListSectionWrite(MatSht, MatSht.Range, "Roof_LinerPanelsQtyCell1", RoofLinerPanels, "Panel");
                                                                            // clean up unused sections
                                                                            if (MatSht.Range) {
                                                                                "e1_LinerPanelsQtyCell1".Value = "";
                                                                                MatSht.Range;
                                                                                "e1_LinerPanelsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                if (MatSht.Range) {
                                                                                    "e3_LinerPanelsQtyCell1".Value = "";
                                                                                    MatSht.Range;
                                                                                    "e3_LinerPanelsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                    if (MatSht.Range) {
                                                                                        "s2_LinerPanelsQtyCell1".Value = "";
                                                                                        MatSht.Range;
                                                                                        "s2_LinerPanelsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                        if (MatSht.Range) {
                                                                                            "s4_LinerPanelsQtyCell1".Value = "";
                                                                                            MatSht.Range;
                                                                                            "s4_LinerPanelsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                            if (MatSht.Range) {
                                                                                                "Roof_LinerPanelsQtyCell1".Value = "";
                                                                                                MatSht.Range;
                                                                                                "Roof_LinerPanelsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                            }

                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''' Trim
                                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                            // '' Rake
                                                                                            MatListSectionWrite(MatSht, MatSht.Range, "RakeTrimQtyCell1", RakeTrimPieces, "Trim");
                                                                                            // '' Eave
                                                                                            MatListSectionWrite(MatSht, MatSht.Range, "EaveTrimQtyCell1", EaveTrimPieces, "Trim");
                                                                                            // '' Outside Corner
                                                                                            MatListSectionWrite(MatSht, MatSht.Range, "OutsideCornerTrimQtyCell1", OutsideCornerTrimPieces, "Trim");
                                                                                            // '' Base Trim
                                                                                            if ((b.BaseTrim == false)) {
                                                                                                MatSht.Range;
                                                                                                "BaseTrimQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                            }
                                                                                            else {
                                                                                                MatListSectionWrite(MatSht, MatSht.Range, "BaseTrimQtyCell1", BaseTrimPieces, "Trim");
                                                                                            }

                                                                                            // '' Wainscot Trim
                                                                                            if ((EstSht.Range("Wainscot").Value == "No")) {
                                                                                                MatSht.Range;
                                                                                                "WainscotTrimQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                            }
                                                                                            else {
                                                                                                MatListSectionWrite(MatSht, MatSht.Range, "WainscotTrimQtyCell1", WainscotTrimPieces, "Trim");
                                                                                            }

                                                                                            // '''
                                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FO Trim
                                                                                            // '''
                                                                                            FOMaterialGen(MatSht, TrimCollection, MiscCollection);
                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''' Gutters and Downspouts
                                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                            // '' gutters
                                                                                            // delete entier gutters section if no gutters
                                                                                            if ((Gutters == false)) {
                                                                                                Range(MatSht.Range, "GutterQtyCell1".offset(-5, 0), MatSht.Range, "DownspoutQtyCell1".offset(2, 0)).EntireRow.Delete;
                                                                                            }
                                                                                            else {
                                                                                                WriteCell = MatSht.Range;
                                                                                                "GutterQtyCell1";
                                                                                                for (GutterPiece in GutterPieces) {
                                                                                                    // insert new row if not the first write cell in the section
                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                        "GutterQtyCell1";
                                                                                                        MatSht.Rows;
                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                        // add piece
                                                                                                        WriteCell.Value = GutterPiece.Quantity;
                                                                                                        WriteCell.offset(0, 1).Value = GutterPiece.tShape;
                                                                                                        WriteCell.offset(0, 2).Value = GutterPiece.tType;
                                                                                                        WriteCell.offset(0, 3).Value = GutterPiece.tMeasurement;
                                                                                                        WriteCell.offset(0, 4).Value = GutterPiece.Color;
                                                                                                        // update write cell
                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                        GutterPiece;
                                                                                                        // ''end caps
                                                                                                        MatSht.Rows;
                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                        WriteCell.Value = GutterEndCapQty;
                                                                                                        WriteCell.offset(0, 1).Value = "R-Loc";
                                                                                                        WriteCell.offset(0, 2).Value = "Sculptured Gutter End Cap";
                                                                                                        WriteCell.offset(0, 3).Value = "N/A";
                                                                                                        WriteCell.offset(0, 4).Value = GutterColor;
                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                        // ''straps
                                                                                                        MatSht.Rows;
                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                        WriteCell.Value = GutterStrapQty;
                                                                                                        WriteCell.offset(0, 1).Value = "R-Loc";
                                                                                                        WriteCell.offset(0, 2).Value = "Gutter Strap 9""";
                                                                                                        WriteCell.offset(0, 3).Value = "N/A";
                                                                                                        WriteCell.offset(0, 4).Value = GutterColor;
                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                        // ''' Downspouts
                                                                                                        WriteCell = MatSht.Range;
                                                                                                        "DownspoutQtyCell1";
                                                                                                        for (DownspoutPiece in DownspoutPieces) {
                                                                                                            // insert new row if not the first write cell in the section
                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                "DownspoutQtyCell1";
                                                                                                                MatSht.Rows;
                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                // add piece
                                                                                                                WriteCell.Value = DownspoutPiece.Quantity;
                                                                                                                WriteCell.offset(0, 1).Value = DownspoutPiece.tShape;
                                                                                                                WriteCell.offset(0, 2).Value = DownspoutPiece.tType;
                                                                                                                WriteCell.offset(0, 3).Value = DownspoutPiece.tMeasurement;
                                                                                                                WriteCell.offset(0, 4).Value = DownspoutPiece.Color;
                                                                                                                // update write cell
                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                DownspoutPiece;
                                                                                                                // '' downspout straps
                                                                                                                MatSht.Rows;
                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                WriteCell.Value = DownspoutStrapQty;
                                                                                                                WriteCell.offset(0, 1).Value = "N/A";
                                                                                                                WriteCell.offset(0, 2).Value = "Downspout Strap";
                                                                                                                WriteCell.offset(0, 3).Value = "N/A";
                                                                                                                WriteCell.offset(0, 4).Value = DownspoutColor;
                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                // '' pop rivits
                                                                                                                MatSht.Rows;
                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                WriteCell.Value = PopRivitQty;
                                                                                                                WriteCell.offset(0, 1).Value = "N/A";
                                                                                                                WriteCell.offset(0, 2).Value = "Pop Rivets";
                                                                                                                WriteCell.offset(0, 3).Value = "1""";
                                                                                                                WriteCell.offset(0, 4).Value = DownspoutColor;
                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                            }

                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''' Additional Options
                                                                                                            // ''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                            // Skylights & Translucent Wall Panels
                                                                                                            if ((SkylightPanelQty == 0)) {
                                                                                                                // check if deleting entire section
                                                                                                                if (((e1GableOverhangSection == false)
                                                                                                                            && ((s2EaveOverhangSection == false)
                                                                                                                            && ((e3GableOverhangSection == false)
                                                                                                                            && ((s4EaveOverhangSection == false)
                                                                                                                            && ((e1GableExtensionSection == false)
                                                                                                                            && ((s2EaveExtensionSection == false)
                                                                                                                            && ((e3GableExtensionSection == false)
                                                                                                                            && (s4EaveExtensionSection == false))))))))) {
                                                                                                                    // delete "Additional Options" Section heading as well
                                                                                                                    MatSht.Range;
                                                                                                                    "skylightPanelQtyCell1".offset(-4, 0).Resize(6, 1).EntireRow.Delete;
                                                                                                                }
                                                                                                                else {
                                                                                                                    // just delete skylights and translucent wall panels heading
                                                                                                                    MatSht.Range;
                                                                                                                    "SkylightPanelQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                }

                                                                                                            }
                                                                                                            else {
                                                                                                                WriteCell = MatSht.Range;
                                                                                                                "SkylightPanelQtyCell1";
                                                                                                                WriteCell.Value = SkylightPanelQty;
                                                                                                                WriteCell.offset(0, 1).Value = "R-Loc";
                                                                                                                WriteCell.offset(0, 2).Value = "Skylights, Fiberglass, White";
                                                                                                                WriteCell.offset(0, 3).Value = "12'";
                                                                                                                WriteCell.offset(0, 4).Value = "N/A";
                                                                                                            }

                                                                                                            // e1 Gable overhang Soffit
                                                                                                            if ((e1GableOverhangSection == false)) {
                                                                                                                MatSht.Range;
                                                                                                                "e1_GableOverhangMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                            }
                                                                                                            else {
                                                                                                                // Call MatListSectionWrite(MatSht, .Range("e1_GableOverhangMatQtyCell1"), e1SoffitPanels, "Panel")
                                                                                                                WriteCell = MatSht.Range;
                                                                                                                "e1_GableOverhangMatQtyCell1";
                                                                                                                for (SoffitPanel in e1SoffitPanels) {
                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                        "e1_GableOverhangMatQtyCell1";
                                                                                                                        MatSht.Rows;
                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                        WriteCell.Value = SoffitPanel.Quantity;
                                                                                                                        WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape;
                                                                                                                        WriteCell.offset(0, 2).Value = SoffitPanel.PanelType;
                                                                                                                        WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement;
                                                                                                                        WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor;
                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                        SoffitPanel;
                                                                                                                        //  Soffit Trim
                                                                                                                        for (SoffitTrim in e1SoffitTrim) {
                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                "e1_GableOverhangMatQtyCell1";
                                                                                                                                MatSht.Rows;
                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                SoffitTrim;
                                                                                                                            }

                                                                                                                            // s2 eave overhang Soffit
                                                                                                                            if ((s2EaveOverhangSection == false)) {
                                                                                                                                MatSht.Range;
                                                                                                                                "s2_EaveOverhangMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                            }
                                                                                                                            else {
                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                "s2_EaveOverhangMatQtyCell1";
                                                                                                                                for (SoffitPanel in s2SoffitPanels) {
                                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                                        "s2_EaveOverhangMatQtyCell1";
                                                                                                                                        MatSht.Rows;
                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                        WriteCell.Value = SoffitPanel.Quantity;
                                                                                                                                        WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape;
                                                                                                                                        WriteCell.offset(0, 2).Value = SoffitPanel.PanelType;
                                                                                                                                        WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement;
                                                                                                                                        WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor;
                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                        SoffitPanel;
                                                                                                                                        //  Soffit Trim
                                                                                                                                        for (SoffitTrim in s2SoffitTrim) {
                                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                                "s2_EaveOverhangMatQtyCell1";
                                                                                                                                                MatSht.Rows;
                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                SoffitTrim;
                                                                                                                                            }

                                                                                                                                            // e3 Gable overhang Soffit
                                                                                                                                            if ((e3GableOverhangSection == false)) {
                                                                                                                                                MatSht.Range;
                                                                                                                                                "e3_GableOverhangMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                            }
                                                                                                                                            else {
                                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                                "e3_GableOverhangMatQtyCell1";
                                                                                                                                                for (SoffitPanel in e3SoffitPanels) {
                                                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                                                        "e3_GableOverhangMatQtyCell1";
                                                                                                                                                        MatSht.Rows;
                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                        WriteCell.Value = SoffitPanel.Quantity;
                                                                                                                                                        WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape;
                                                                                                                                                        WriteCell.offset(0, 2).Value = SoffitPanel.PanelType;
                                                                                                                                                        WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement;
                                                                                                                                                        WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor;
                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                        SoffitPanel;
                                                                                                                                                        //  Soffit Trim
                                                                                                                                                        for (SoffitTrim in e3SoffitTrim) {
                                                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                "e3_GableOverhangMatQtyCell1";
                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                SoffitTrim;
                                                                                                                                                            }

                                                                                                                                                            // s4 eave overhang Soffit
                                                                                                                                                            if ((s4EaveOverhangSection == false)) {
                                                                                                                                                                MatSht.Range;
                                                                                                                                                                "s4_EaveOverhangMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                                                "s4_EaveOverhangMatQtyCell1";
                                                                                                                                                                for (SoffitPanel in s4SoffitPanels) {
                                                                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                        "s4_EaveOverhangMatQtyCell1";
                                                                                                                                                                        MatSht.Rows;
                                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                                        WriteCell.Value = SoffitPanel.Quantity;
                                                                                                                                                                        WriteCell.offset(0, 1).Value = SoffitPanel.PanelShape;
                                                                                                                                                                        WriteCell.offset(0, 2).Value = SoffitPanel.PanelType;
                                                                                                                                                                        WriteCell.offset(0, 3).Value = SoffitPanel.PanelMeasurement;
                                                                                                                                                                        WriteCell.offset(0, 4).Value = SoffitPanel.PanelColor;
                                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                        SoffitPanel;
                                                                                                                                                                        //  Soffit Trim
                                                                                                                                                                        for (SoffitTrim in s4SoffitTrim) {
                                                                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                "s4_EaveOverhangMatQtyCell1";
                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                SoffitTrim;
                                                                                                                                                                            }

                                                                                                                                                                            // e1 Gable Extension
                                                                                                                                                                            if ((e1GableExtensionSection == false)) {
                                                                                                                                                                                MatSht.Range;
                                                                                                                                                                                "e1_GableExtensionMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                                            }
                                                                                                                                                                            else {
                                                                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                                                                "e1_GableExtensionMatQtyCell1";
                                                                                                                                                                                for (ExtensionPanel in e1ExtensionPanels) {
                                                                                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                        "e1_GableExtensionMatQtyCell1";
                                                                                                                                                                                        MatSht.Rows;
                                                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                                                        WriteCell.Value = ExtensionPanel.Quantity;
                                                                                                                                                                                        WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape;
                                                                                                                                                                                        WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType;
                                                                                                                                                                                        WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement;
                                                                                                                                                                                        WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor;
                                                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                        ExtensionPanel;
                                                                                                                                                                                        // 2x8 inside angle
                                                                                                                                                                                        if (!(e1InsideAngleTrim == null)) {
                                                                                                                                                                                            for (TrimPiece in e1InsideAngleTrim) {
                                                                                                                                                                                                if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                                    "e1_GableExtensionMatQtyCell1";
                                                                                                                                                                                                    MatSht.Rows;
                                                                                                                                                                                                    (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                    WriteCell.Value = TrimPiece.Quantity;
                                                                                                                                                                                                    WriteCell.offset(0, 1).Value = TrimPiece.tShape;
                                                                                                                                                                                                    WriteCell.offset(0, 2).Value = TrimPiece.tType;
                                                                                                                                                                                                    WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement;
                                                                                                                                                                                                    WriteCell.offset(0, 4).Value = TrimPiece.Color;
                                                                                                                                                                                                    WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                    TrimPiece;
                                                                                                                                                                                                }

                                                                                                                                                                                                //  Soffit Trim
                                                                                                                                                                                                for (SoffitTrim in e1SoffitTrim) {
                                                                                                                                                                                                    if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                                        "e1_GableExtensionMatQtyCell1";
                                                                                                                                                                                                        MatSht.Rows;
                                                                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                        WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                                                        WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                                                        WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                                                        WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                                                        WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                        SoffitTrim;
                                                                                                                                                                                                    }

                                                                                                                                                                                                    // s2 eave Extension
                                                                                                                                                                                                    if ((s2EaveExtensionSection == false)) {
                                                                                                                                                                                                        MatSht.Range;
                                                                                                                                                                                                        "s2_EaveExtensionMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                                                                    }
                                                                                                                                                                                                    else {
                                                                                                                                                                                                        WriteCell = MatSht.Range;
                                                                                                                                                                                                        "s2_EaveExtensionMatQtyCell1";
                                                                                                                                                                                                        for (ExtensionPanel in s2ExtensionPanels) {
                                                                                                                                                                                                            if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                "s2_EaveExtensionMatQtyCell1".Address;
                                                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                WriteCell.Value = ExtensionPanel.Quantity;
                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape;
                                                                                                                                                                                                                WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType;
                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement;
                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor;
                                                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                ExtensionPanel;
                                                                                                                                                                                                                //  Soffit Trim
                                                                                                                                                                                                                for (SoffitTrim in s2SoffitTrim) {
                                                                                                                                                                                                                    if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                        "s2_EaveExtensionMatQtyCell1".Address;
                                                                                                                                                                                                                        MatSht.Rows;
                                                                                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                        WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                                                                        WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                                                                        WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                                                                        WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                                                                        WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                        SoffitTrim;
                                                                                                                                                                                                                    }

                                                                                                                                                                                                                    // e3 Gable Extension
                                                                                                                                                                                                                    if ((e3GableExtensionSection == false)) {
                                                                                                                                                                                                                        MatSht.Range;
                                                                                                                                                                                                                        "e3_GableExtensionMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                                                                                    }
                                                                                                                                                                                                                    else {
                                                                                                                                                                                                                        WriteCell = MatSht.Range;
                                                                                                                                                                                                                        "e3_GableExtensionMatQtyCell1";
                                                                                                                                                                                                                        for (ExtensionPanel in e3ExtensionPanels) {
                                                                                                                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                                                                "e3_GableExtensionMatQtyCell1";
                                                                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                WriteCell.Value = ExtensionPanel.Quantity;
                                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape;
                                                                                                                                                                                                                                WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType;
                                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement;
                                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor;
                                                                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                ExtensionPanel;
                                                                                                                                                                                                                                // 2x8 inside angle
                                                                                                                                                                                                                                if (!(e3InsideAngleTrim == null)) {
                                                                                                                                                                                                                                    for (TrimPiece in e3InsideAngleTrim) {
                                                                                                                                                                                                                                        if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                                                                            "e3_GableExtensionMatQtyCell1";
                                                                                                                                                                                                                                            MatSht.Rows;
                                                                                                                                                                                                                                            (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                            WriteCell.Value = TrimPiece.Quantity;
                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = TrimPiece.tShape;
                                                                                                                                                                                                                                            WriteCell.offset(0, 2).Value = TrimPiece.tType;
                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement;
                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = TrimPiece.Color;
                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                        //  Soffit Trim
                                                                                                                                                                                                                                        for (SoffitTrim in e3SoffitTrim) {
                                                                                                                                                                                                                                            if ((WriteCell != MatSht.Range)) {
                                                                                                                                                                                                                                                "e3_GableExtensionMatQtyCell1";
                                                                                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                SoffitTrim;
                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                            // s4 eave Extension
                                                                                                                                                                                                                                            if ((s4EaveExtensionSection == false)) {
                                                                                                                                                                                                                                                MatSht.Range;
                                                                                                                                                                                                                                                "s4_EaveExtensionMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                            else {
                                                                                                                                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                "s4_EaveExtensionMatQtyCell1";
                                                                                                                                                                                                                                                for (ExtensionPanel in s4ExtensionPanels) {
                                                                                                                                                                                                                                                    if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                                                        "s4_EaveExtensionMatQtyCell1".Address;
                                                                                                                                                                                                                                                        MatSht.Rows;
                                                                                                                                                                                                                                                        (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                        WriteCell.Value = ExtensionPanel.Quantity;
                                                                                                                                                                                                                                                        WriteCell.offset(0, 1).Value = ExtensionPanel.PanelShape;
                                                                                                                                                                                                                                                        WriteCell.offset(0, 2).Value = ExtensionPanel.PanelType;
                                                                                                                                                                                                                                                        WriteCell.offset(0, 3).Value = ExtensionPanel.PanelMeasurement;
                                                                                                                                                                                                                                                        WriteCell.offset(0, 4).Value = ExtensionPanel.PanelColor;
                                                                                                                                                                                                                                                        WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                        ExtensionPanel;
                                                                                                                                                                                                                                                        //  Soffit Trim
                                                                                                                                                                                                                                                        for (SoffitTrim in s4SoffitTrim) {
                                                                                                                                                                                                                                                            if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                                                                "s4_EaveExtensionMatQtyCell1".Address;
                                                                                                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                                WriteCell.Value = SoffitTrim.Quantity;
                                                                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = SoffitTrim.tShape;
                                                                                                                                                                                                                                                                WriteCell.offset(0, 2).Value = SoffitTrim.tType;
                                                                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = SoffitTrim.tMeasurement;
                                                                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitTrim.Color;
                                                                                                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                SoffitTrim;
                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Fasteners
                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                                                                                                                            // ''''''''''''''''''''''''''''''''''''''''''''' roof screws
                                                                                                                                                                                                                                                            WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                            "RoofScrewsQtyCell1";
                                                                                                                                                                                                                                                            WriteCell.Value = rTekScrewQty;
                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Tek Screws";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = "1.25""";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = rColor;
                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                            // lap screws
                                                                                                                                                                                                                                                            MatSht.Rows;
                                                                                                                                                                                                                                                            (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                            WriteCell.Value = rLapScrewQty;
                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Lap Screws";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = ".875""";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = rColor;
                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''' wall screws
                                                                                                                                                                                                                                                            WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                            "WallScrewsQtyCell1";
                                                                                                                                                                                                                                                            WriteCell.Value = wTekScrewQty;
                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Tek Screws";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = "1.25""";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = wColor;
                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                            // lap screws
                                                                                                                                                                                                                                                            MatSht.Rows;
                                                                                                                                                                                                                                                            (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                            WriteCell.Value = wLapScrewQty;
                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Lap Screws";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = ".875""";
                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = wColor;
                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''' Trim screws
                                                                                                                                                                                                                                                            WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                            "TrimScrewsQtyCell1";
                                                                                                                                                                                                                                                            for (Screw in TrimScrews) {
                                                                                                                                                                                                                                                                if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                                                                    "TrimScrewsQtyCell1".Address;
                                                                                                                                                                                                                                                                    MatSht.Rows;
                                                                                                                                                                                                                                                                    (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                                    WriteCell.Value = Screw.Quantity;
                                                                                                                                                                                                                                                                    WriteCell.offset(0, 1).Value = "Tek Screws";
                                                                                                                                                                                                                                                                    WriteCell.offset(0, 3).Value = "1.25""";
                                                                                                                                                                                                                                                                    WriteCell.offset(0, 4).Value = Screw.Color;
                                                                                                                                                                                                                                                                    WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                    Screw;
                                                                                                                                                                                                                                                                    // lap screws (duplicate colors/quantities)
                                                                                                                                                                                                                                                                    for (Screw in TrimScrews) {
                                                                                                                                                                                                                                                                        if ((WriteCell.Address != MatSht.Range)) {
                                                                                                                                                                                                                                                                            "TrimScrewsQtyCell1".Address;
                                                                                                                                                                                                                                                                            MatSht.Rows;
                                                                                                                                                                                                                                                                            (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                                            WriteCell.Value = Screw.Quantity;
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Lap Screws";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = ".875""";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = Screw.Color;
                                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                            Screw;
                                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''' Soffit Screws
                                                                                                                                                                                                                                                                            if ((SoffitScrewQty == 0)) {
                                                                                                                                                                                                                                                                                MatSht.Range;
                                                                                                                                                                                                                                                                                "SoffitScrewsQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                                                            else {
                                                                                                                                                                                                                                                                                WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                                                "SoffitScrewsQtyCell1";
                                                                                                                                                                                                                                                                                WriteCell.Value = SoffitScrewQty;
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = "Tek Screws";
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = "1.25""";
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitScrewColor;
                                                                                                                                                                                                                                                                                WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                                // lap screws
                                                                                                                                                                                                                                                                                MatSht.Rows;
                                                                                                                                                                                                                                                                                (WriteCell.Row + 1).Insert;
                                                                                                                                                                                                                                                                                WriteCell.Value = SoffitScrewQty;
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 1).Value = "Lap Screws";
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 3).Value = ".875""";
                                                                                                                                                                                                                                                                                WriteCell.offset(0, 4).Value = SoffitScrewColor;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Miscellaneous
                                                                                                                                                                                                                                                                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                                                                                                                                            // Butyl Tape
                                                                                                                                                                                                                                                                            WriteCell = MatSht.Range;
                                                                                                                                                                                                                                                                            "MiscMaterialsQtyCell1";
                                                                                                                                                                                                                                                                            WriteCell.Value = ButylTapeQty;
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Butyl Tape";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = "44'";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = "N/A";
                                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                            // Inside Closures
                                                                                                                                                                                                                                                                            WriteCell.Value = InsideClosureQty;
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Inside Closures";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = "3'";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = "N/A";
                                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                            // Outside Closures
                                                                                                                                                                                                                                                                            WriteCell.Value = OutsideClosureQty;
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 1).Value = "Outside Closures";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 3).Value = "3'";
                                                                                                                                                                                                                                                                            WriteCell.offset(0, 4).Value = "N/A";
                                                                                                                                                                                                                                                                            WriteCell = WriteCell.offset(1, 0);
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // autofit columns
                                                                                                                                                                                                                                                                        MatSht.Columns.AutoFit;
                                                                                                                                                                                                                                                                        // '''''''''''''''''''''''''''''''''''' Vendor Material List
                                                                                                                                                                                                                                                                        // ''''Roof Panels
                                                                                                                                                                                                                                                                        for (RoofPanel in s2RoofPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            RoofPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (RoofPanel in s4RoofPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            RoofPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        //  Ridge Caps
                                                                                                                                                                                                                                                                        if ((RidgeCapQty != 0)) {
                                                                                                                                                                                                                                                                            item = new clsMiscItem();
                                                                                                                                                                                                                                                                            item.Quantity = RidgeCapQty;
                                                                                                                                                                                                                                                                            item.Name = ("Formed Ridge Cap " + PitchString);
                                                                                                                                                                                                                                                                            item.Measurement = "3'";
                                                                                                                                                                                                                                                                            item.Color = rColor;
                                                                                                                                                                                                                                                                            MiscCollection.Add;
                                                                                                                                                                                                                                                                            item;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Sidewall Panels
                                                                                                                                                                                                                                                                        for (Panel in s2SidewallPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in s4SidewallPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Endwall panels
                                                                                                                                                                                                                                                                        for (Panel in e1EndwallPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in e3EndwallPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // liner panels
                                                                                                                                                                                                                                                                        for (Panel in e1LinerPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in e3LinerPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in s2LinerPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in s4LinerPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (Panel in RoofLinerPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            Panel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Trim
                                                                                                                                                                                                                                                                        for (TrimPiece in RakeTrimPieces) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (TrimPiece in EaveTrimPieces) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (TrimPiece in OutsideCornerTrimPieces) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        if ((b.BaseTrim == true)) {
                                                                                                                                                                                                                                                                            for (TrimPiece in BaseTrimPieces) {
                                                                                                                                                                                                                                                                                TrimCollection.Add;
                                                                                                                                                                                                                                                                                TrimPiece;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Gutters
                                                                                                                                                                                                                                                                        if ((b.Gutters == true)) {
                                                                                                                                                                                                                                                                            for (GutterPiece in GutterPieces) {
                                                                                                                                                                                                                                                                                // trim item used for gutters
                                                                                                                                                                                                                                                                                TrimCollection.Add;
                                                                                                                                                                                                                                                                                GutterPiece;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            //  Gutter End Caps
                                                                                                                                                                                                                                                                            if ((GutterEndCapQty != 0)) {
                                                                                                                                                                                                                                                                                item = new clsMiscItem();
                                                                                                                                                                                                                                                                                item.Quantity = GutterEndCapQty;
                                                                                                                                                                                                                                                                                item.Shape = "R-Loc";
                                                                                                                                                                                                                                                                                item.Name = "Sculptured Gutter End Cap";
                                                                                                                                                                                                                                                                                item.Measurement = "N/A";
                                                                                                                                                                                                                                                                                item.Color = GutterColor;
                                                                                                                                                                                                                                                                                MiscCollection.Add;
                                                                                                                                                                                                                                                                                item;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            //  Gutter Straps
                                                                                                                                                                                                                                                                            if ((GutterStrapQty != 0)) {
                                                                                                                                                                                                                                                                                item = new clsMiscItem();
                                                                                                                                                                                                                                                                                item.Quantity = GutterStrapQty;
                                                                                                                                                                                                                                                                                item.Shape = "R-Loc";
                                                                                                                                                                                                                                                                                item.Name = "Gutter Strap";
                                                                                                                                                                                                                                                                                item.Measurement = "9""";
                                                                                                                                                                                                                                                                                item.Color = GutterColor;
                                                                                                                                                                                                                                                                                MiscCollection.Add;
                                                                                                                                                                                                                                                                                item;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            for (DownspoutPiece in DownspoutPieces) {
                                                                                                                                                                                                                                                                                // trim item used for downspout pieces
                                                                                                                                                                                                                                                                                TrimCollection.Add;
                                                                                                                                                                                                                                                                                DownspoutPiece;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            //  Downspout Straps
                                                                                                                                                                                                                                                                            if ((DownspoutStrapQty != 0)) {
                                                                                                                                                                                                                                                                                item = new clsMiscItem();
                                                                                                                                                                                                                                                                                item.Quantity = DownspoutStrapQty;
                                                                                                                                                                                                                                                                                item.Shape = "N/A";
                                                                                                                                                                                                                                                                                item.Name = "Downspout Strap";
                                                                                                                                                                                                                                                                                item.Measurement = "N/A";
                                                                                                                                                                                                                                                                                item.Color = DownspoutColor;
                                                                                                                                                                                                                                                                                MiscCollection.Add;
                                                                                                                                                                                                                                                                                item;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            //  Pop Rivets
                                                                                                                                                                                                                                                                            if ((PopRivitQty != 0)) {
                                                                                                                                                                                                                                                                                item = new clsMiscItem();
                                                                                                                                                                                                                                                                                item.Quantity = PopRivitQty;
                                                                                                                                                                                                                                                                                item.Shape = "N/A";
                                                                                                                                                                                                                                                                                item.Name = "Pop Rivets";
                                                                                                                                                                                                                                                                                item.Measurement = "1""";
                                                                                                                                                                                                                                                                                item.Color = DownspoutColor;
                                                                                                                                                                                                                                                                                MiscCollection.Add;
                                                                                                                                                                                                                                                                                item;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Skylights and Translucent Wall Panels
                                                                                                                                                                                                                                                                        if ((SkylightPanelQty != 0)) {
                                                                                                                                                                                                                                                                            SkylightPanel = new clsPanel();
                                                                                                                                                                                                                                                                            SkylightPanel.PanelShape = "R-Loc";
                                                                                                                                                                                                                                                                            SkylightPanel.PanelType = "Skylights, Fiberglass, White";
                                                                                                                                                                                                                                                                            SkylightPanel.PanelMeasurement = "12'";
                                                                                                                                                                                                                                                                            SkylightPanel.PanelLength = (12 * 12);
                                                                                                                                                                                                                                                                            SkylightPanel.PanelColor = "N/A";
                                                                                                                                                                                                                                                                            SkylightPanel.Quantity = SkylightPanelQty;
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            SkylightPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Overhangs, Extensions, and Soffits
                                                                                                                                                                                                                                                                        // extension panels
                                                                                                                                                                                                                                                                        for (ExtensionPanel in e1ExtensionPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            ExtensionPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (ExtensionPanel in s2ExtensionPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            ExtensionPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (ExtensionPanel in e3ExtensionPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            ExtensionPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (ExtensionPanel in s4ExtensionPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            ExtensionPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // soffit panels
                                                                                                                                                                                                                                                                        for (SoffitPanel in e1SoffitPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            SoffitPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (SoffitPanel in s2SoffitPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            SoffitPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (SoffitPanel in e3SoffitPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            SoffitPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (SoffitPanel in s4SoffitPanels) {
                                                                                                                                                                                                                                                                            PanelCollection.Add;
                                                                                                                                                                                                                                                                            SoffitPanel;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // soffit trim
                                                                                                                                                                                                                                                                        for (TrimPiece in e1SoffitTrim) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (TrimPiece in s2SoffitTrim) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (TrimPiece in e3SoffitTrim) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        for (TrimPiece in s4SoffitTrim) {
                                                                                                                                                                                                                                                                            TrimCollection.Add;
                                                                                                                                                                                                                                                                            TrimPiece;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Fasteners and Miscelaneous
                                                                                                                                                                                                                                                                        //  Roof Screws
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = rTekScrewQty;
                                                                                                                                                                                                                                                                        item.Name = "Tek Screws";
                                                                                                                                                                                                                                                                        item.Measurement = "1.25""";
                                                                                                                                                                                                                                                                        item.Color = rColor;
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = rLapScrewQty;
                                                                                                                                                                                                                                                                        item.Name = "Lap Screws";
                                                                                                                                                                                                                                                                        item.Measurement = ".875""";
                                                                                                                                                                                                                                                                        item.Color = rColor;
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        //  Wall Screws
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = wTekScrewQty;
                                                                                                                                                                                                                                                                        item.Name = "Tek Screws";
                                                                                                                                                                                                                                                                        item.Measurement = "1.25""";
                                                                                                                                                                                                                                                                        item.Color = wColor;
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = wLapScrewQty;
                                                                                                                                                                                                                                                                        item.Name = "Lap Screws";
                                                                                                                                                                                                                                                                        item.Measurement = ".875""";
                                                                                                                                                                                                                                                                        item.Color = wColor;
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        // '' Trim Screws
                                                                                                                                                                                                                                                                        // Tek
                                                                                                                                                                                                                                                                        for (Screw in TrimScrews) {
                                                                                                                                                                                                                                                                            item = new clsMiscItem();
                                                                                                                                                                                                                                                                            item.Quantity = Screw.Quantity;
                                                                                                                                                                                                                                                                            item.Name = "Tek Screws";
                                                                                                                                                                                                                                                                            item.Measurement = "1.25""";
                                                                                                                                                                                                                                                                            item.Color = Screw.Color;
                                                                                                                                                                                                                                                                            MiscCollection.Add;
                                                                                                                                                                                                                                                                            item;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // Duplicate Lap
                                                                                                                                                                                                                                                                        for (Screw in TrimScrews) {
                                                                                                                                                                                                                                                                            item = new clsMiscItem();
                                                                                                                                                                                                                                                                            item.Quantity = Screw.Quantity;
                                                                                                                                                                                                                                                                            item.Name = "Lap Screws";
                                                                                                                                                                                                                                                                            item.Measurement = ".875""";
                                                                                                                                                                                                                                                                            item.Color = Screw.Color;
                                                                                                                                                                                                                                                                            MiscCollection.Add;
                                                                                                                                                                                                                                                                            item;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '' Soffit Screws
                                                                                                                                                                                                                                                                        if ((SoffitScrewQty != 0)) {
                                                                                                                                                                                                                                                                            // Tek
                                                                                                                                                                                                                                                                            item = new clsMiscItem();
                                                                                                                                                                                                                                                                            item.Quantity = SoffitScrewQty;
                                                                                                                                                                                                                                                                            item.Name = "Tek Screws";
                                                                                                                                                                                                                                                                            item.Measurement = "1.25""";
                                                                                                                                                                                                                                                                            item.Color = SoffitScrewColor;
                                                                                                                                                                                                                                                                            MiscCollection.Add;
                                                                                                                                                                                                                                                                            item;
                                                                                                                                                                                                                                                                            // Duplicate Lap
                                                                                                                                                                                                                                                                            item = new clsMiscItem();
                                                                                                                                                                                                                                                                            item.Quantity = SoffitScrewQty;
                                                                                                                                                                                                                                                                            item.Name = "Lap Screws";
                                                                                                                                                                                                                                                                            item.Measurement = ".875""";
                                                                                                                                                                                                                                                                            item.Color = SoffitScrewColor;
                                                                                                                                                                                                                                                                            MiscCollection.Add;
                                                                                                                                                                                                                                                                            item;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // '''' Miscellaneous
                                                                                                                                                                                                                                                                        // Butyl Tape
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = ButylTapeQty;
                                                                                                                                                                                                                                                                        item.Name = "Butyl Tape";
                                                                                                                                                                                                                                                                        item.Measurement = "44'";
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        // Inside Closures
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = InsideClosureQty;
                                                                                                                                                                                                                                                                        item.Name = "Inside Closures";
                                                                                                                                                                                                                                                                        item.Measurement = "3'";
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        // Outside Closures
                                                                                                                                                                                                                                                                        item = new clsMiscItem();
                                                                                                                                                                                                                                                                        item.Quantity = OutsideClosureQty;
                                                                                                                                                                                                                                                                        item.Name = "Outside Closures";
                                                                                                                                                                                                                                                                        item.Measurement = "3'";
                                                                                                                                                                                                                                                                        MiscCollection.Add;
                                                                                                                                                                                                                                                                        item;
                                                                                                                                                                                                                                                                        //  generate the rest of the misc. materials
                                                                                                                                                                                                                                                                        MiscMaterialsGen.MiscMaterialCalc(MiscCollection, WriteCell, b);
                                                                                                                                                                                                                                                                        // ''''generate vendor material list
                                                                                                                                                                                                                                                                        // remove duplicates
                                                                                                                                                                                                                                                                        DuplicateMaterialRemoval(PanelCollection, "Panel");
                                                                                                                                                                                                                                                                        DuplicateMaterialRemoval(TrimCollection, "Trim");
                                                                                                                                                                                                                                                                        DuplicateMaterialRemoval(MiscCollection, "Misc");
                                                                                                                                                                                                                                                                        VendorAndPriceLists.VendorMaterialListsGen(PanelCollection, TrimCollection, MiscCollection);
                                                                                                                                                                                                                                                                        VendorAndPriceLists.PriceListGen(PanelCollection, TrimCollection, MiscCollection);
                                                                                                                                                                                                                                                                        VendorAndPriceLists.CostEstimateGen(PanelCollection, TrimCollection, MiscCollection, b);
                                                                                                                                                                                                                                                                        //  generate description
                                                                                                                                                                                                                                                                        VendorAndPriceLists.DescriptionGen(b);
                                                                                                                                                                                                                                                                        let LastRow: number;
                                                                                                                                                                                                                                                                        let mCell: Range;
                                                                                                                                                                                                                                                                        let MissingPrice: boolean;
                                                                                                                                                                                                                                                                        // Check SS Price List for missing prices
                                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                                        LastRow = ThisWorkbook.Sheets("Structural Steel Price List").Cells;
                                                                                                                                                                                                                                                                        ThisWorkbook.Sheets("Structural Steel Price List").Rows.Count;
                                                                                                                                                                                                                                                                        "H";
                                                                                                                                                                                                                                                                        ThisWorkbook.Sheets("Structural Steel Price List").End;
                                                                                                                                                                                                                                                                        xlUp.Row;
                                                                                                                                                                                                                                                                        for (mCell in ThisWorkbook.Sheets("Structural Steel Price List").Range) {
                                                                                                                                                                                                                                                                            ("H4:H" + LastRow);
                                                                                                                                                                                                                                                                            if ((IsNumeric(mCell.Value) == false)) {
                                                                                                                                                                                                                                                                                MissingPrice = true;
                                                                                                                                                                                                                                                                                break;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                                        LastRow = ThisWorkbook.Sheets("Materials Price List").Cells;
                                                                                                                                                                                                                                                                        ThisWorkbook.Sheets("Materials Price List").Rows.Count;
                                                                                                                                                                                                                                                                        "H";
                                                                                                                                                                                                                                                                        ThisWorkbook.Sheets("Materials Price List").End;
                                                                                                                                                                                                                                                                        xlUp.Row;
                                                                                                                                                                                                                                                                        for (mCell in ThisWorkbook.Sheets("Materials Price List").Range) {
                                                                                                                                                                                                                                                                            ("H4:H" + LastRow);
                                                                                                                                                                                                                                                                            if ((IsNumeric(mCell.Value) == false)) {
                                                                                                                                                                                                                                                                                MissingPrice = true;
                                                                                                                                                                                                                                                                                break;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        if (MissingPrice) {
                                                                                                                                                                                                                                                                            MsgBox;
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        "At least one price could  not be found on the structrual steel or materials price lists.";
                                                                                                                                                                                                                                                                        vbExclamation;
                                                                                                                                                                                                                                                                        "Missing Price";
                                                                                                                                                                                                                                                                        ThisWorkbook.Sheets("Employee Materials List").Select;
                                                                                                                                                                                                                                                                        Application.ScreenUpdating = true;
                                                                                                                                                                                                                                                                        MsgBox;
                                                                                                                                                                                                                                                                        "Materials list generation complete!";
                                                                                                                                                                                                                                                                        System.Windows.Forms.MessageBoxIcon.Information;
                                                                                                                                                                                                                                                                        "Generation Complete";
                                                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */MsgBox;
                                                                                                                                                                                                                                                                        "Key information to determine the roofing materials is missing! Please check the template for missing data and try again.";
                                                                                                                                                                                                                                                                        vbExclamation;
                                                                                                                                                                                                                                                                        "Missing Data";
                                                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */MsgBox;
                                                                                                                                                                                                                                                                        "It has been calculated that more than 5 seperate panels will be needed to cover the rafter length of the roof. Please perform this calculation manually.";
                                                                                                                                                                                                                                                                        vbExclamation;
                                                                                                                                                                                                                                                                        "Program Design Exceeded";
                                                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */MsgBox;
                                                                                                                                                                                                                                                                        "It has been calculated that 4+ seperate panels will be needed to cover the rafter length of the roof. This calculation is currently disabled.";
                                                                                                                                                                                                                                                                        System.Windows.Forms.MessageBoxIcon.Information;
                                                                                                                                                                                                                                                                        "Currently Disabled";
                                                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */MsgBox;
                                                                                                                                                                                                                                                                        "Single slope calculations are currently disabled.";
                                                                                                                                                                                                                                                                        System.Windows.Forms.MessageBoxIcon.Information;
                                                                                                                                                                                                                                                                        "Feature Disabled";
                                                                                                                                                                                                                                                                        return;
                                                                                                                                                                                                                                                                        (<string>(ImperialMeasurementFormat((<number>(TotalInches)))));
                                                                                                                                                                                                                                                                        let Feet: number;
                                                                                                                                                                                                                                                                        let Inches: number;
                                                                                                                                                                                                                                                                        let InchFraction: number;
                                                                                                                                                                                                                                                                        let InchFractString: Object;
                                                                                                                                                                                                                                                                        Feet = Application.WorksheetFunction.RoundDown((TotalInches / 12), 0);
                                                                                                                                                                                                                                                                        Inches = Application.WorksheetFunction.RoundDown((XLMod((TotalInches / 12), 1) * 12), 0);
                                                                                                                                                                                                                                                                        InchFraction = Application.WorksheetFunction.MRound(XLMod((XLMod((TotalInches / 12), 1) * 12), 1), (1 / 16));
                                                                                                                                                                                                                                                                        // add to inches if inch fraction = 1
                                                                                                                                                                                                                                                                        if ((InchFraction == 1)) {
                                                                                                                                                                                                                                                                            Inches = (Inches + 1);
                                                                                                                                                                                                                                                                            InchFraction = 0;
                                                                                                                                                                                                                                                                            // check if 12 inches
                                                                                                                                                                                                                                                                            if ((Inches == 12)) {
                                                                                                                                                                                                                                                                                Feet = (Feet + 1);
                                                                                                                                                                                                                                                                                Inches = 0;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        // 'write values to formatting cell
                                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                                        // write string
                                                                                                                                                                                                                                                                        InchFraction.Range("Inch_Format").Value = Feet;
                                                                                                                                                                                                                                                                        HiddenSht.Range("Inch_Fraction_Format").Value = Feet;
                                                                                                                                                                                                                                                                        if (((Inches == 0)
                                                                                                                                                                                                                                                                                    && (InchFraction == 0))) {
                                                                                                                                                                                                                                                                            (Range("Feet_Format").Text + "'");
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else if ((InchFraction == 0)) {
                                                                                                                                                                                                                                                                            // ImperialMeasurementFormat = .Range("Feet_Format").Text & "'" & " " & .Range("Inch_Format").Text & "''"
                                                                                                                                                                                                                                                                            (Range("Inch_Format").Text + "''");
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else if ((Inches == 0)) {
                                                                                                                                                                                                                                                                            (Range("Feet_Format").Text + ("'" + (" "
                                                                                                                                                                                                                                                                                        + (Trim(., Range("Inch_Fraction_Format").Text) + "''"))));
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else {
                                                                                                                                                                                                                                                                            (Range("Inch_Format").Text + (" "
                                                                                                                                                                                                                                                                                        + (Trim(., Range("Inch_Fraction_Format").Text) + "''")));
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        XLMod(a, b);
                                                                                                                                                                                                                                                                        //  This replicates the Excel MOD function
                                                                                                                                                                                                                                                                        XLMod = (a
                                                                                                                                                                                                                                                                                    - (b * Int((a / b))));
                                                                                                                                                                                                                                                                        (<number>(ClosestWallPurlin((<void>(Height)), Variant, Optional, (<number>(Direction)), Optional, (<boolean>(NonstandardFloorPurlin)))));
                                                                                                                                                                                                                                                                        // DEVELOPER: Ryan Wells (wellsr.com)
                                                                                                                                                                                                                                                                        // DESCRIPTION: Function returns the nearest value to a target
                                                                                                                                                                                                                                                                        // INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
                                                                                                                                                                                                                                                                        //  and an optional direction variable described below.
                                                                                                                                                                                                                                                                        // OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
                                                                                                                                                                                                                                                                        //  Set equal to -1 to find the closest value below your target
                                                                                                                                                                                                                                                                        //  set equal to 1 to find the closest value above your target
                                                                                                                                                                                                                                                                        // OUTPUT: The output is the number in the range closest to your target value.
                                                                                                                                                                                                                                                                        //  Because the output is a variant, the address of the closest number can also be returned when
                                                                                                                                                                                                                                                                        //  calling this function from another VBA macro.
                                                                                                                                                                                                                                                                        let t: Object;
                                                                                                                                                                                                                                                                        let u: Object;
                                                                                                                                                                                                                                                                        let Purlins: Object;
                                                                                                                                                                                                                                                                        let Purlin: Object;
                                                                                                                                                                                                                                                                        let pAbove: number;
                                                                                                                                                                                                                                                                        let pBelow: number;
                                                                                                                                                                                                                                                                        Purlins = Array(87.5, 147.5, 207.5, 267.5, 327.5, 387.5, 447.5, 507.5, 567.5, 627.5, 687.5, 747.5, 807.5, 867.5, 927.5, 987.5, 1047.5, 1107.5, 1167.5);
                                                                                                                                                                                                                                                                        // Normal (Floor to Eave)
                                                                                                                                                                                                                                                                        if ((NonstandardFloorPurlin == false)) {
                                                                                                                                                                                                                                                                            t = 1.79769313486231E+308;
                                                                                                                                                                                                                                                                            // initialize
                                                                                                                                                                                                                                                                            // ClosestWallPurlin = "No value found"
                                                                                                                                                                                                                                                                            for (Purlin in Purlins) {
                                                                                                                                                                                                                                                                                if (IsNumeric(Purlin)) {
                                                                                                                                                                                                                                                                                    u = Abs((Purlin - Height));
                                                                                                                                                                                                                                                                                    if (((Direction > 0)
                                                                                                                                                                                                                                                                                                && (Purlin >= Height))) {
                                                                                                                                                                                                                                                                                        // only report if closer number is greater than the target
                                                                                                                                                                                                                                                                                        if ((u < t)) {
                                                                                                                                                                                                                                                                                            t = u;
                                                                                                                                                                                                                                                                                            ClosestWallPurlin = Purlin;
                                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                    else if (((Direction < 0)
                                                                                                                                                                                                                                                                                                && (Purlin <= Height))) {
                                                                                                                                                                                                                                                                                        // only report if closer number is less than the target
                                                                                                                                                                                                                                                                                        if ((u < t)) {
                                                                                                                                                                                                                                                                                            t = u;
                                                                                                                                                                                                                                                                                            ClosestWallPurlin = Purlin;
                                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                    else if ((Direction == 0)) {
                                                                                                                                                                                                                                                                                        if ((u < t)) {
                                                                                                                                                                                                                                                                                            t = u;
                                                                                                                                                                                                                                                                                            ClosestWallPurlin = Purlin;
                                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                            // starting at bottom of partial wall instead of ground
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else if ((NonstandardFloorPurlin == true)) {
                                                                                                                                                                                                                                                                            pBelow = (Application.WorksheetFunction.RoundDown((Height / 5), 0) * 5);
                                                                                                                                                                                                                                                                            pAbove = (Application.WorksheetFunction.RoundUp((Height / 5), 0) * 5);
                                                                                                                                                                                                                                                                            switch (Direction) {
                                                                                                                                                                                                                                                                                case -1:
                                                                                                                                                                                                                                                                                    ClosestWallPurlin = pBelow;
                                                                                                                                                                                                                                                                                    break;
                                                                                                                                                                                                                                                                                case 1:
                                                                                                                                                                                                                                                                                    ClosestWallPurlin = pAbove;
                                                                                                                                                                                                                                                                                    break;
                                                                                                                                                                                                                                                                                case 0:
                                                                                                                                                                                                                                                                                    // report the closest purlin
                                                                                                                                                                                                                                                                                    if ((Abs((pAbove - Height)) < Abs((pBelow - Height)))) {
                                                                                                                                                                                                                                                                                        ClosestWallPurlin = pAbove;
                                                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                                                    else {
                                                                                                                                                                                                                                                                                        ClosestWallPurlin = pBelow;
                                                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                                                    break;
                                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                        (<number>(ClosestRoofPurlin((<void>(RafterLength)), Variant, Optional, (<number>(Direction)))));
                                                                                                                                                                                                                                                                        if ((Direction == 1)) {
                                                                                                                                                                                                                                                                            // closest rounding up
                                                                                                                                                                                                                                                                            ClosestRoofPurlin = (Application.WorksheetFunction.RoundUp((RafterLength / 60), 0) * 60);
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else if ((Direction == -1)) {
                                                                                                                                                                                                                                                                            // closest rounding down
                                                                                                                                                                                                                                                                            ClosestRoofPurlin = (Application.WorksheetFunction.RoundDown((RafterLength / 60), 0) * 60);
                                                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                                                        else {
                                                                                                                                                                                                                                                                            //  closest without caring
                                                                                                                                                                                                                                                                            ClosestRoofPurlin = (Application.WorksheetFunction.Round((RafterLength / 60), 0) * 60);
                                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                }

                                                                                                                                                                                                                            }

                                                                                                                                                                                                                        }

                                                                                                                                                                                                                    }

                                                                                                                                                                                                                }

                                                                                                                                                                                                            }

                                                                                                                                                                                                        }

                                                                                                                                                                                                    }

                                                                                                                                                                                                }

                                                                                                                                                                                            }

                                                                                                                                                                                        }

                                                                                                                                                                                    }

                                                                                                                                                                                }

                                                                                                                                                                            }

                                                                                                                                                                        }

                                                                                                                                                                    }

                                                                                                                                                                }

                                                                                                                                                            }

                                                                                                                                                        }

                                                                                                                                                    }

                                                                                                                                                }

                                                                                                                                            }

                                                                                                                                        }

                                                                                                                                    }

                                                                                                                                }

                                                                                                                            }

                                                                                                                        }

                                                                                                                    }

                                                                                                                }

                                                                                                            }

                                                                                                        }

                                                                                                    }

                                                                                                }

                                                                                            }

                                                                                        }

                                                                                    }

                                                                                }

                                                                            }

                                                                        }

                                                                    }

                                                                }

                                                            }

                                                        }

                                                    }

                                                }

                                            }

                                        }

                                    }

                                }

                            }

                        }

                    }

                }

            }

        }

    }

}
private IsEven(PanelCount: number): boolean {
    if ((((PanelCount % 2)
                == 0)
                == true)) {
        IsEven = true;
    }
    else {
        IsEven = false;
    }

}

private EndwallPanelGen(EndwallPanels: Collection, eWall: string, b: clsBuilding, FullHeightLinerPanels: boolean) {
    let eP1: clsPanel;
    // Warning!!! Optional parameters not supported
    let eP2: clsPanel;
    let eP3: clsPanel;
    let WainscotPanel: clsPanel;
    let WainscotFtLength: number;
    let ePanel: clsPanel;
    let ePanelCount: number;
    let pLengthMax: number;
    // in
    let pNum: number;
    let pLength: number;
    let SpecialBottomPurlin: boolean;
    // '' this boolean applies when the endwall is marked as partial or as gable only
    let UnsplicedPanels: Collection = new Collection();
    let MaxSegments: number;
    let Segment1Length: number;
    let Segment2Length: number;
    let TopPanelLengths: number[];
    let FOCollection: Collection;
    let FO: clsFO;
    // Note: Account for Max Height term in endwall panel segment?
    // With...
    // determine number of panels
    ePanelCount = Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0);
    // top down purlins
    if (b.WallStatus) {
        eWall = ("Partial" | b.WallStatus);
        eWall = "Gable Only";
        SpecialBottomPurlin = true;
        if (b.Wainscot) {
            (eWall != "None");
            WainscotPanel = new clsPanel();
            WainscotPanel.PanelLength = number.Parse(Left(b.Wainscot, eWall, 2));
            if ((FullHeightLinerPanels == false)) {
                WainscotFtLength = (WainscotPanel.PanelLength / 12);
            }

            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Gable Roofs ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            if ((b.rShape == "Gable")) {
                switch (b.WallStatus) {
                    case eWall:
                        break;
                    case "Exclude":
                        return;
                        break;
                    case "Include":
                    case "Partial":
                        if ((IsEven(ePanelCount) == true)) {
                            // add lengths symetrically to all panels
                            for (pNum = 1; (pNum
                                        <= (ePanelCount / 2)); pNum++) {
                                eP1 = new clsPanel();
                                // Check if adding length to first panel or not
                                if ((pNum == 1)) {
                                    // '''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                                    if ((b.rPitch == 1)) {
                                        eP1.PanelLength = ((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength);
                                        12;
                                        //  only bHeight for rPitch 1
                                    }
                                    else {
                                        eP1.PanelLength = (((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength)
                                                    * 12);
                                        (b.rPitch * 3);
                                    }

                                }
                                else if ((pNum != 1)) {
                                    eP1.PanelLength = (pLengthMax
                                                + (b.rPitch * 3));
                                }

                                // one panel for each side of the endwall
                                eP1.Quantity = 1;
                                eP1.rEdgePosition = ((pNum - 1) * (3 * 12));
                                // add panel to collection
                                UnsplicedPanels.Add;
                                eP1;
                                // create, add the duplicate panel for the other side of the endwall to the collection
                                eP1 = new clsPanel();
                                eP1.Quantity = 1;
                                eP1.PanelLength = UnsplicedPanels[UnsplicedPanels.Count].PanelLength;
                                eP1.rEdgePosition = (((Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0) * 3)
                                            - (UnsplicedPanels.Count * 3))
                                            * 12);
                                UnsplicedPanels.Add;
                                eP1;
                                // update running panel length
                                pLengthMax = eP1.PanelLength;
                            }

                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If odd # of panels
                        }
                        else if ((IsEven(ePanelCount) == false)) {
                            // add lengths symetrically to all panels except the long middle one
                            for (pNum = 1; (pNum
                                        <= ((ePanelCount - 1)
                                        / 2)); pNum++) {
                                eP1 = new clsPanel();
                                // Check if adding length to first panel or not
                                if ((pNum == 1)) {
                                    // '''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                                    if ((b.rPitch == 1)) {
                                        eP1.PanelLength = ((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength);
                                        12;
                                        //  only bHeight for rPitch 1
                                    }
                                    else {
                                        eP1.PanelLength = (((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength)
                                                    * 12);
                                        (b.rPitch * 3);
                                    }

                                }
                                else if ((pNum != 1)) {
                                    eP1.PanelLength = (pLengthMax
                                                + (b.rPitch * 3));
                                }

                                // one panel for each side of the endwall
                                eP1.Quantity = 1;
                                eP1.rEdgePosition = ((pNum - 1) * (3 * 12));
                                // add panel to collection
                                UnsplicedPanels.Add;
                                eP1;
                                // create, add the duplicate panel for the other side of the endwall to the collection
                                eP1 = new clsPanel();
                                eP1.Quantity = 1;
                                eP1.PanelLength = UnsplicedPanels[UnsplicedPanels.Count].PanelLength;
                                eP1.rEdgePosition = (((Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0) * 3)
                                            - (UnsplicedPanels.Count * 3))
                                            * 12);
                                UnsplicedPanels.Add;
                                eP1;
                                // update running panel length
                                pLengthMax = eP1.PanelLength;
                            }

                            //  add roof pitch contribution again one more time for middle panel
                            eP1 = new clsPanel();
                            eP1.PanelLength = (pLengthMax
                                        + (b.rPitch * 3));
                            eP1.Quantity = 1;
                            eP1.rEdgePosition = (pNum * 3);
                            // add panels to collection
                            UnsplicedPanels.Add;
                            eP1;
                            // update running panel length
                            pLengthMax = eP1.PanelLength;
                        }

                        break;
                    case "Gable Only":
                        if ((IsEven(ePanelCount) == true)) {
                            // add lengths symetrically to all panels
                            for (pNum = 1; (pNum
                                        <= (ePanelCount / 2)); pNum++) {
                                if (((pNum == 1)
                                            && (b.rPitch != 1))) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = (b.rPitch * 3);
                                    // rPitch contribution
                                }
                                else if ((pNum != 1)) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = (pLengthMax
                                                + (b.rPitch * 3));
                                }

                                // add panels to collection
                                if (!(eP1 == null)) {
                                    eP1.Quantity = 2;
                                    UnsplicedPanels.Add;
                                    eP1;
                                    pLengthMax = eP1.PanelLength;
                                }

                            }

                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' If odd # of panels
                        }
                        else if ((IsEven(ePanelCount) == false)) {
                            // add lengths symetrically to all panels except the long middle one
                            for (pNum = 1; (pNum
                                        <= ((ePanelCount - 1)
                                        / 2)); pNum++) {
                                if (((pNum == 1)
                                            && (b.rPitch != 1))) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = ((b.bHeight * 12)
                                                + (b.rPitch * 3));
                                }
                                else if ((pNum != 1)) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = (pLengthMax
                                                + (b.rPitch * 3));
                                }

                                // add panels to collection, update plength
                                if (!(eP1 == null)) {
                                    eP1.Quantity = 2;
                                    UnsplicedPanels.Add;
                                    eP1;
                                    pLengthMax = eP1.PanelLength;
                                }

                            }

                            //  add roof pitch contribution again one more time for middle panel
                            eP1 = new clsPanel();
                            eP1.PanelLength = (pLengthMax
                                        + (b.rPitch * 3));
                            eP1.Quantity = 1;
                            // add panels to collection
                            UnsplicedPanels.Add;
                            eP1;
                            // update running panel length
                            pLengthMax = eP1.PanelLength;
                        }

                        break;
                }

                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Single Slope Roofs ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            }
            else if ((b.rShape == "Single Slope")) {
                switch (b.WallStatus) {
                    case eWall:
                        break;
                    case "Exclude":
                        return;
                        break;
                    case "Include":
                    case "Partial":
                        for (pNum = 1; (pNum <= ePanelCount); pNum++) {
                            eP1 = new clsPanel();
                            // Check if adding length to first panel or not
                            if ((pNum == 1)) {
                                // '''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                                if ((b.rPitch == 1)) {
                                    eP1.PanelLength = ((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength);
                                    12;
                                    //  only bHeight for rPitch 1
                                }
                                else {
                                    eP1.PanelLength = (((b.bHeight - b.LengthAboveFinishedFloor)[eWall] - WainscotFtLength)
                                                * 12);
                                    (b.rPitch * 3);
                                }

                            }
                            else if ((pNum != 1)) {
                                eP1.PanelLength = (pLengthMax
                                            + (b.rPitch * 3));
                            }

                            eP1.Quantity = 1;
                            eP1.rEdgePosition = ((pNum - 1) * (3 * 12));
                            // add panels to collection
                            UnsplicedPanels.Add;
                            eP1;
                            pLengthMax = eP1.PanelLength;
                        }

                        break;
                    case "Gable Only":
                        for (pNum = 1; (pNum <= ePanelCount); pNum++) {
                            // '''''''' check roof pitch. If 1, don't add pitch contribution to first panel.
                            if (((pNum == 1)
                                        && (b.rPitch != 1))) {
                                eP1 = new clsPanel();
                                eP1.PanelLength = (b.rPitch * 3);
                                // '' bHeight + rPitch contribution
                            }
                            else if ((pNum != 1)) {
                                eP1 = new clsPanel();
                                eP1.PanelLength = (pLengthMax
                                            + (b.rPitch * 3));
                            }

                            // add panel to collection
                            if (!(eP1 == null)) {
                                eP1.Quantity = 1;
                                UnsplicedPanels.Add;
                                eP1;
                                pLengthMax = eP1.PanelLength;
                            }

                        }

                        break;
                }

            }

        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''' Segment Panels & Account for FO Cutouts''''''''''''''''''''''''''''''''''''''''
        // set FO Collection
        if ((eWall == "e1")) {

        }

        FOCollection = b.e1FOs;
    }
    else {

    }

    FOCollection = b.e3FOs;
    // '''''''''''''''''''''''''''''''''''''''''''''''''' Using the max length of the unsegmented panels, first calculate the evenly porportioned segment lengths ''''''''''''
    // '' Max Segments (Not Factoring FO Cutouts) = 1
    if ((pLengthMax <= (42 * 12))) {
        MaxSegments = 1;
        //  Only 1 segment and panels don't need splicing. Add panels, and exit sub
        for (ePanel in UnsplicedPanels) {
            for (FO in FOCollection) {
                if ((((FO.FOType == "OHDoor")
                            || (FO.FOType == "MiscFO"))
                            && ((FO.Width >= (7 * 12))
                            && (FO.bEdgeHeight == 0)))) {
                    if (((ePanel.rEdgePosition
                                >= (FO.rEdgePosition + (3 * 12)))
                                && (ePanel.lEdgePosition
                                <= (FO.lEdgePosition - (3 * 12))))) {
                        ePanel.PanelLength = (ePanel.PanelLength
                                    - (FO.Height
                                    - (WainscotFtLength * 12)));
                    }

                }

            }

            // deduct 8" from full height liners
            if ((FullHeightLinerPanels == true)) {
                ePanel.PanelLength = (ePanel.PanelLength - 8);
            }

            // make sure panel length is > 0 before adding. this can happen when full height liners are less than 8" from the ceiling
            if ((ePanel.PanelLength > 0)) {
                EndwallPanels.Add;
            }

            ePanel;
        }

        // save overlaps to building class
        if ((eWall == "e1")) {
            b.e1WallPanelOverlaps = 0;
        }
        else if ((eWall == "e3")) {
            b.e3WallPanelOverlaps = 0;
        }

        /* Warning! GOTO is not Implemented */// ' <--------- Exit and finish collection at the end of 1 segment condition
        // '' Max Segments (Not Factoring FO Cutouts) = 2
    }
    else if ((pLengthMax
                <= ((79 * 12)
                + 3.5))) {
        MaxSegments = 2;
        Segment1Length = ClosestWallPurlin((pLengthMax / 2), 0, SpecialBottomPurlin);
        // correct if greater than 37'3.5" for purlins that go from the bottom up
        if ((SpecialBottomPurlin == false)) {
            if ((Segment1Length
                        > ((37 * 12)
                        + 3.5))) {
                Segment1Length = ((37 * 12)
                            + 3.5);
            }

        }

        if ((eWall == "e1")) {
            b.e1WallPanelOverlaps = 1;
        }
        else if ((eWall == "e3")) {
            b.e3WallPanelOverlaps = 1;
        }

        // '' Max Segments (Not Factoring FO Cutouts) = 3
    }
    else {
        MaxSegments = 3;
        Segment1Length = ClosestWallPurlin((pLengthMax / 3), 0, SpecialBottomPurlin);
        Segment2Length = (ClosestWallPurlin((pLengthMax / (3 * 2)), 0, SpecialBottomPurlin) - Segment1Length);
        // add overlaps to building class
        if ((eWall == "e1")) {
            b.e1WallPanelOverlaps = 2;
        }
        else if ((eWall == "e3")) {
            b.e3WallPanelOverlaps = 2;
        }

    }

    // '''''''' Splice using the segment lengths calculated above ''''''''''''''
    // '' Note: This occurs when we have an unspliced panel collection which we know must include panels which need to be segmented
    for (ePanel in UnsplicedPanels) {
        if ((ePanel.PanelLength <= Segment1Length)) {
            // '' check for intersecting FOs
            for (FO in FOCollection) {
                if ((((FO.FOType == "OHDoor")
                            || (FO.FOType == "MiscFO"))
                            && ((FO.Width >= (7 * 12))
                            && (FO.bEdgeHeight == 0)))) {
                    if (((ePanel.rEdgePosition
                                >= (FO.rEdgePosition + (3 * 12)))
                                && (ePanel.lEdgePosition
                                <= (FO.lEdgePosition - (3 * 12))))) {
                        ePanel.PanelLength = (ePanel.PanelLength
                                    - (FO.Height
                                    - (WainscotFtLength * 12)));
                    }

                }

            }

            eP1 = new clsPanel();
            eP1.PanelLength = ePanel.PanelLength;
            // deduct 8" from full height liners
            if ((FullHeightLinerPanels == true)) {
                eP1.PanelLength = (eP1.PanelLength - 8);
            }

            // EndwallPanels.Add ePanel
        }
        else if ((MaxSegments == 2)) {
            if ((ePanel.PanelLength > Segment1Length)) {
                // '' check for intersecting FOs
                for (FO in FOCollection) {
                    if ((((FO.FOType == "OHDoor")
                                || (FO.FOType == "MiscFO"))
                                && ((FO.Width >= (7 * 12))
                                && (FO.bEdgeHeight == 0)))) {
                        if (((ePanel.rEdgePosition
                                    >= (FO.rEdgePosition + (3 * 12)))
                                    && (ePanel.lEdgePosition
                                    <= (FO.lEdgePosition - (3 * 12))))) {
                            // If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                            if (((Segment1Length
                                        - (FO.Height
                                        - (WainscotFtLength * 12)))
                                        > 0)) {
                                eP1 = new clsPanel();
                                eP1.PanelLength = ((Segment1Length
                                            - (FO.Height
                                            - (WainscotFtLength * 12)))
                                            + 1.5);
                                eP2 = new clsPanel();
                                eP2.PanelLength = ((ePanel.PanelLength - Segment1Length)
                                            + 1.5);
                                if ((FullHeightLinerPanels == true)) {
                                    eP2.PanelLength = (eP2.PanelLength - 8);
                                }

                                /* Warning! GOTO is not Implemented */// '' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                            }
                            else if (((Segment1Length
                                        - (FO.Height
                                        - (WainscotFtLength * 12)))
                                        <= 0)) {
                                eP2 = new clsPanel();
                                eP2.PanelLength = (ePanel.PanelLength
                                            - (FO.Height
                                            - (WainscotFtLength * 12)));
                                if ((FullHeightLinerPanels == true)) {
                                    eP2.PanelLength = (eP2.PanelLength - 8);
                                }

                                /* Warning! GOTO is not Implemented */}

                        }

                    }

                }

                // ''''''''''' no intersecting FOs
                eP1 = new clsPanel();
                eP1.PanelLength = (Segment1Length + 1.5);
                eP2 = new clsPanel();
                eP2.PanelLength = ((ePanel.PanelLength - Segment1Length)
                            + 1.5);
                if ((FullHeightLinerPanels == true)) {
                    eP2.PanelLength = (eP2.PanelLength - 8);
                }

            }
            else if ((MaxSegments == 3)) {
                if ((ePanel.PanelLength
                            <= (Segment1Length + Segment2Length))) {
                    // '' check for intersecting FOs
                    for (FO in FOCollection) {
                        if ((((FO.FOType == "OHDoor")
                                    || (FO.FOType == "MiscFO"))
                                    && ((FO.Width >= (7 * 12))
                                    && (FO.bEdgeHeight == 0)))) {
                            if (((ePanel.rEdgePosition
                                        >= (FO.rEdgePosition + (3 * 12)))
                                        && (ePanel.lEdgePosition
                                        <= (FO.lEdgePosition - (3 * 12))))) {
                                // If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                                if (((Segment1Length
                                            - (FO.Height
                                            - (WainscotFtLength * 12)))
                                            > 0)) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = ((Segment1Length
                                                - (FO.Height
                                                - (WainscotFtLength * 12)))
                                                + 1.5);
                                    eP2 = new clsPanel();
                                    eP2.PanelLength = ((ePanel.PanelLength - Segment1Length)
                                                + 1.5);
                                    /* Warning! GOTO is not Implemented */// '' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                                }
                                else if (((Segment1Length
                                            - (FO.Height
                                            - (WainscotFtLength * 12)))
                                            <= 0)) {
                                    eP2 = new clsPanel();
                                    eP2.PanelLength = (((Segment2Length + Segment1Length)
                                                - (FO.Height
                                                - (WainscotFtLength * 12)))
                                                + 1.5);
                                    /* Warning! GOTO is not Implemented */}

                            }

                        }

                    }

                    // ''''''''''' no intersecting FOs
                    eP1 = new clsPanel();
                    eP1.PanelLength = (Segment1Length + 1.5);
                    eP2 = new clsPanel();
                    eP2.PanelLength = ((ePanel.PanelLength - Segment1Length)
                                + 1.5);
                }
                else if ((ePanel.PanelLength
                            > (Segment1Length + Segment2Length))) {
                    // '' check for intersecting FOs
                    for (FO in FOCollection) {
                        if ((((FO.FOType == "OHDoor")
                                    || (FO.FOType == "MiscFO"))
                                    && ((FO.Width >= (7 * 12))
                                    && (FO.bEdgeHeight == 0)))) {
                            if (((ePanel.rEdgePosition
                                        >= (FO.rEdgePosition + (3 * 12)))
                                        && (ePanel.lEdgePosition
                                        <= (FO.lEdgePosition - (3 * 12))))) {
                                // If FO takes up less than segment 1, create first panel from height remaining after cutout and create segment 2 normally (since it's above the FO and not effected)
                                if (((Segment1Length
                                            - (FO.Height
                                            - (WainscotFtLength * 12)))
                                            > 0)) {
                                    eP1 = new clsPanel();
                                    eP1.PanelLength = ((Segment1Length
                                                - (FO.Height
                                                - (WainscotFtLength * 12)))
                                                + 1.5);
                                    eP2 = new clsPanel();
                                    eP2.PanelLength = (Segment2Length + 3);
                                    eP3 = new clsPanel();
                                    eP3.PanelLength = ((ePanel.PanelLength
                                                - (Segment1Length - Segment2Length))
                                                + 1.5);
                                    if ((FullHeightLinerPanels == true)) {
                                        eP3.PanelLength = (eP3.PanelLength - 8);
                                    }

                                    /* Warning! GOTO is not Implemented */// '' If FO takes up the entirity of segment 1 or more, add segment 2 without overlap and subtract the height remaining after the FO cutout
                                }
                                else if (((Segment1Length
                                            - (FO.Height
                                            - (WainscotFtLength * 12)))
                                            <= 0)) {
                                    eP2 = new clsPanel();
                                    eP2.PanelLength = (((Segment2Length + Segment1Length)
                                                - (FO.Height
                                                - (WainscotFtLength * 12)))
                                                + 1.5);
                                    eP3 = new clsPanel();
                                    eP3.PanelLength = ((ePanel.PanelLength
                                                - (Segment1Length - Segment2Length))
                                                + 1.5);
                                    if ((FullHeightLinerPanels == true)) {
                                        eP3.PanelLength = (eP3.PanelLength - 8);
                                    }

                                    /* Warning! GOTO is not Implemented */}

                            }

                        }

                    }

                    // ''''''''''' no intersecting FOs
                    eP1 = new clsPanel();
                    eP1.PanelLength = (Segment1Length + 1.5);
                    eP2 = new clsPanel();
                    eP2.PanelLength = (Segment2Length + 3);
                    eP3 = new clsPanel();
                    eP3.PanelLength = ((ePanel.PanelLength
                                - (Segment1Length - Segment2Length))
                                + 1.5);
                    // deduct 8" from full height liners
                    if ((FullHeightLinerPanels == true)) {
                        eP3.PanelLength = (eP3.PanelLength - 8);
                    }

                }

                /* Warning! Labeled Statements are not Implemented */if (!(eP1 == null)) {
                    eP1.Quantity = ePanel.Quantity;
                    if ((eP1.PanelLength > 0)) {
                        EndwallPanels.Add;
                    }

                    eP1;
                    eP1 = null;
                }

                if (!(eP2 == null)) {
                    eP2.Quantity = ePanel.Quantity;
                    if ((eP2.PanelLength > 0)) {
                        EndwallPanels.Add;
                    }

                    eP2;
                    eP2 = null;
                }

                if (!(eP3 == null)) {
                    eP3.Quantity = ePanel.Quantity;
                    if ((eP3.PanelLength > 0)) {
                        EndwallPanels.Add;
                    }

                    eP3;
                    eP3 = null;
                }

                ePanel;
                // ''
                /* Warning! Labeled Statements are not Implemented */for (ePanel in EndwallPanels) {
                    ePanel.PanelShape = b.wPanelShape;
                    ePanel.PanelType = b.wPanelType;
                    ePanel.PanelColor = b.wPanelColor;
                    ePanel.PanelMeasurement = ImperialMeasurementFormat(ePanel.PanelLength);
                }

                if ((!(WainscotPanel == null)
                            && (FullHeightLinerPanels == false))) {
                    WainscotPanel.PanelMeasurement = ImperialMeasurementFormat(WainscotPanel.PanelLength);
                    WainscotPanel.Quantity = Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0);
                    WainscotPanel.PanelColor = EstSht.Range((eWall + "_Wainscot")).offset(0, 2).Value;
                    WainscotPanel.PanelType = EstSht.Range((eWall + "_Wainscot")).offset(0, 1).Value;
                    WainscotPanel.PanelShape = b.wPanelShape;
                    EndwallPanels.Add;
                    WainscotPanel;
                }

                DuplicateMaterialRemoval(EndwallPanels, "Panel");
            }

            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Trim calculation (and downspouts and gutters)
            TrimPieceCalc(/* ref */(<Collection>(TrimCollection)), (<number>(NetTrimLength)), (<string>(TrimType)), Optional, (<string>(rPitchString)), Optional, (<number>(DownspoutQty)), Optional, (<clsBuilding>(b)));
            let Trim: clsTrim;
            let ExistingTrim: clsTrim;
            // vars to fill in Trim class
            let tQty: number;
            let tLength: number;
            let tTypeString: string;
            let RemainingLength: number;
            let t: number;
            let DuplicateFound: boolean;
            let LargestTrimDivisor: number;
            let IdealTrimSize: number;
            let TrimSegmentsRequired: number;
            let N: number;
            let tLengthRemaining: number;
            let CurrentLength: number;
            // trim Type
            switch (TrimType) {
                case "Rake":
                    tTypeString = "Rake Trim";
                    // With...
                    // ''single slope
                    if ((b.rShape == "Single Slope")) {
                        // sidewall 2
                        if ((b.s2RafterSheetLength
                                    + (b.s2ExtensionRafterLength
                                    + (b.s4ExtensionRafterLength <= 244)))) {
                            // add trim for single side
                            Trim = new clsTrim();
                            Trim.tLength = NearestTrimSize((b.s2RafterSheetLength
                                            + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength)), 1, ,, true);
                            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                            Trim.Quantity = 2;
                            Trim.tType = tTypeString;
                            TrimCollection.Add;
                            Trim;
                            // decrease needed trim length
                            NetTrimLength = (NetTrimLength
                                        - (Trim.tLength * Trim.Quantity));
                        }

                        // '''Gable
                    }
                    else if ((b.rShape == "Gable")) {
                        // sidewall 2
                        if ((b.s2RafterSheetLength
                                    + (b.s2ExtensionRafterLength <= 244))) {
                            // add trim for single side
                            Trim = new clsTrim();
                            Trim.tLength = NearestTrimSize((b.s2RafterSheetLength + b.s2ExtensionRafterLength), 1, ,, true);
                            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                            Trim.Quantity = 2;
                            Trim.tType = tTypeString;
                            TrimCollection.Add;
                            Trim;
                            // decrease needed trim length
                            NetTrimLength = (NetTrimLength
                                        - (Trim.tLength * Trim.Quantity));
                        }

                        if ((b.s4RafterSheetLength
                                    + (b.s4ExtensionRafterLength <= 244))) {
                            // add trim for single side
                            Trim = new clsTrim();
                            Trim.tLength = NearestTrimSize((b.s4RafterSheetLength + b.s4ExtensionRafterLength), 1, ,, true);
                            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                            Trim.Quantity = 2;
                            Trim.tType = tTypeString;
                            TrimCollection.Add;
                            Trim;
                            // decrease needed trim length
                            NetTrimLength = (NetTrimLength
                                        - (Trim.tLength * Trim.Quantity));
                        }

                    }

                    break;
                case "Short Eave":
                    tTypeString = ("Short Eave" + (" " + rPitchString));
                    break;
                case "High Eave":
                    tTypeString = ("High-Side Eave" + (" " + rPitchString));
                    break;
                case "Outside Corner":
                    tTypeString = "Outside Corner Trim";
                    // With...
                    if (((b.bHeight * 12)
                                <= 244)) {
                        // add trim for single side
                        Trim = new clsTrim();
                        Trim.tLength = NearestTrimSize((b.bHeight * 12), 1, ,, true);
                        Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                        if ((b.rShape == "Gable")) {
                            Trim.Quantity = 4;
                        }
                        else if ((b.rShape == "Single Slope")) {
                            Trim.Quantity = 2;
                        }

                        Trim.tType = tTypeString;
                        TrimCollection.Add;
                        Trim;
                        // decrease needed trim length
                        NetTrimLength = (NetTrimLength
                                    - (Trim.tLength * Trim.Quantity));
                    }

                    // check high side height if single slope
                    if ((b.rShape == "Single Slope")) {
                        if ((b.HighSideEaveHeight <= 244)) {
                            // add trim for single side
                            Trim = new clsTrim();
                            Trim.tLength = NearestTrimSize(b.HighSideEaveHeight, 1, ,, true);
                            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                            Trim.Quantity = 2;
                            Trim.tType = tTypeString;
                            TrimCollection.Add;
                            Trim;
                            // decrease needed trim length
                            NetTrimLength = (NetTrimLength
                                        - (Trim.tLength * Trim.Quantity));
                        }

                    }

                    break;
                case "Base":
                    tTypeString = "Base Trim";
                    break;
                case "Gutter":
                    tTypeString = ("Sculptured Gutter Hang-On" + (" " + rPitchString));
                    break;
                case "Downspout":
                    tTypeString = "Square Downspout W/O Kickout";
                    break;
                case "Jamb":
                    tTypeString = "Jamb Trim";
                    // With...
                    // sidewall 2
                    if ((b.RafterLength <= 244)) {
                        // add trim for single side
                        Trim = new clsTrim();
                        Trim.tLength = NearestTrimSize(b.RafterLength, 1, "Jamb", true);
                        Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                        Trim.Quantity = 2;
                        Trim.tType = tTypeString;
                        TrimCollection.Add;
                        Trim;
                        // decrease needed trim length
                        NetTrimLength = (NetTrimLength
                                    - (Trim.tLength * Trim.Quantity));
                    }

                    // sidewall 4
                    if ((b.rShape == "Gable")) {
                        if ((b.RafterLength <= 244)) {
                            // add trim for single side
                            Trim = new clsTrim();
                            Trim.tLength = NearestTrimSize(b.RafterLength, 1, "Jamb", true);
                            Trim.tMeasurement = ImperialMeasurementFormat(Trim.tLength);
                            Trim.Quantity = 2;
                            Trim.tType = tTypeString;
                            TrimCollection.Add;
                            Trim;
                            // decrease needed trim length
                            NetTrimLength = (NetTrimLength
                                        - (Trim.tLength * Trim.Quantity));
                        }

                    }

                    break;
                case "Head":
                    tTypeString = "Head Trim W/O Kickout";
                    break;
                case "Outside Angle":
                    tTypeString = "2x6 Outside Angle Trim";
                    break;
                case "Inside Angle":
                    tTypeString = "2x8 Inside Angle Trim";
                    break;
                case "Standard Wainscot":
                    tTypeString = "Standard Wainscot Trim";
                    break;
                case "Masonry Wainscot":
                    tTypeString = "Masonry Wainscot Trim";
                    break;
            }

            // exit sub if no remaining trim length
            if ((NetTrimLength <= 0)) {
                DuplicateMaterialRemoval(TrimCollection, "Trim");
                return;
            }

            // '' note: need to add a check to see if we should start with 20'4" trim or not (in cases where trim is shorter)
            Trim = new clsTrim();
            // '''''' Check for starting piece size
            switch (NetTrimLength) {
            }

            244;
            LargestTrimDivisor = 240;
            Trim.tMeasurement = "20'4""";
            Trim.tLength = 242;
            218;
            LargestTrimDivisor = 216;
            Trim.tMeasurement = "18'2""";
            Trim.tLength = 216;
            194;
            LargestTrimDivisor = 192;
            Trim.tMeasurement = "16'2""";
            Trim.tLength = 192;
            170;
            LargestTrimDivisor = 168;
            //
            Trim.tMeasurement = "14'2""";
            Trim.tLength = 168;
            146;
            LargestTrimDivisor = 144;
            Trim.tMeasurement = "12'2""";
            Trim.tLength = 144;
        }
        else {
            LargestTrimDivisor = 120;
            Trim.tMeasurement = "10'2""";
            Trim.tLength = 120;
        }

        // check for pieces of the largest size trim
        tQty = Application.WorksheetFunction.RoundDown((NetTrimLength / LargestTrimDivisor), 0);
        if ((TrimType != "Downspout")) {
            Trim.Quantity = tQty;
        }
        else if ((TrimType == "Downspout")) {
            Trim.Quantity = (tQty * DownspoutQty);
        }

        Trim.tType = tTypeString;
        Trim.clsType = "Trim";
        TrimCollection.Add;
        Trim;
        // ''find other trim size
        Trim = new clsTrim();
        Trim.tType = tTypeString;
        // find remaining length
        RemainingLength = (NetTrimLength
                    - (LargestTrimDivisor * tQty));
        if ((RemainingLength != 0)) {
            // find size, write length and qty to class
            if (((TrimType != "Jamb")
                        && (TrimType != "Head"))) {
                Trim.tMeasurement = NearestTrimSize(RemainingLength, 1);
                Trim.tLength = NearestTrimSize(RemainingLength, 1, ,, true);
            }
            else if ((TrimType == "Jamb")) {
                Trim.tMeasurement = NearestTrimSize(RemainingLength, 1, "Jamb");
                Trim.tLength = NearestTrimSize(RemainingLength, 1, "Jamb", true);
            }
            else if ((TrimType == "Head")) {
                Trim.tMeasurement = NearestTrimSize(RemainingLength, 1, "Head");
                Trim.tLength = NearestTrimSize(RemainingLength, 1, "Head", true);
            }

            // just increase quantity of the corresponding 20'4" (or largest starting size) if the remaining trim size rounds to it
            for (t = 1; (t <= TrimCollection.Count); t++) {
                if (((Trim.tMeasurement == TrimCollection(t).tMeasurement)
                            && (Trim.tType == TrimCollection(t).tType))) {
                    if ((TrimType != "Downspout")) {
                        TrimCollection(t).Quantity = (TrimCollection(t).Quantity + 1);
                    }
                    else if ((TrimType == "Downspout")) {
                        TrimCollection(t).Quantity = (TrimCollection(t).Quantity + DownspoutQty);
                    }

                    // mark duplicate as found, exit
                    DuplicateFound = true;
                    break;
                }

            }

            // if no duplicate found add to collection
            if ((DuplicateFound == false)) {
                if ((TrimType != "Downspout")) {
                    Trim.Quantity = 1;
                }
                else if ((TrimType == "Downspout")) {
                    Trim.Quantity = DownspoutQty;
                }

                Trim.Quantity = 1;
                Trim.clsType = "Trim";
                TrimCollection.Add;
                Trim;
            }

        }

        // ' function returns string of the nearest available rake trim size
        (<void>(NearestTrimSize((<void>(Length)), Variant, Optional, (<number>(Direction)), Optional, (<string>(UniqueTrimType)), Optional, (<boolean>(NumericOutput)))));
        // DESCRIPTION: Function returns the nearest value to a target
        // INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
        //  and an optional direction variable described below.
        // OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
        //  Set equal to -1 to find the closest value below your target
        //  set equal to 1 to find the closest value above your target
        // OUTPUT: The output is the number in the range closest to your target value.
        //  Because the output is a variant, the address of the closest number can also be returned when
        //  calling this function from another VBA macro.
        let t: Object;
        let u: Object;
        let Trims: Object;
        let Trim: Object;
        let tSize: Object;
        let NearestTrimSizeString: string;
        // check for trim type
        if ((UniqueTrimType == "")) {
            Trims = Array(122, 146, 170, 194, 218, 244);
        }
        else if ((UniqueTrimType == "Head")) {
            Trims = Array(42, 75, 87, 99, 122, 123, 147, 171, 195, 219, 244);
        }
        else if ((UniqueTrimType == "Jamb")) {
            Trims = Array(86, 122, 146, 170, 194, 218, 244);
        }

        t = 1.79769313486231E+308;
        // initialize
        for (Trim in Trims) {
            if (IsNumeric(Trim)) {
                u = Abs((Trim - Length));
                if (((Direction > 0)
                            && (Trim >= Length))) {
                    // only report if closer number is greater than the target
                    if ((u < t)) {
                        t = u;
                        tSize = Trim;
                    }

                }
                else if (((Direction < 0)
                            && (Trim <= Length))) {
                    // only report if closer number is less than the target
                    if ((u < t)) {
                        t = u;
                        tSize = Trim;
                    }

                }
                else if ((Direction == 0)) {
                    if ((u < t)) {
                        t = u;
                        tSize = Trim;
                    }

                }

            }

        }

        // return available trim name
        switch (tSize) {
            case 42:
                NearestTrimSizeString = "3'6""";
                break;
            case 75:
                NearestTrimSizeString = "6'3""";
                break;
            case 86:
                NearestTrimSizeString = "7'2""";
                break;
            case 87:
                NearestTrimSizeString = "7'3""";
                break;
            case 99:
                NearestTrimSizeString = "8'3""";
                break;
            case 122:
                NearestTrimSizeString = "10'2""";
                break;
            case 123:
                NearestTrimSizeString = "10'3""";
                break;
            case 146:
                NearestTrimSizeString = "12'2""";
                break;
            case 147:
                NearestTrimSizeString = "12'3""";
                break;
            case 170:
                NearestTrimSizeString = "14'2""";
                break;
            case 171:
                NearestTrimSizeString = "14'3""";
                break;
            case 194:
                NearestTrimSizeString = "16'2""";
                break;
            case 195:
                NearestTrimSizeString = "16'3""";
                break;
            case 218:
                NearestTrimSizeString = "18'2""";
                break;
            case 219:
                NearestTrimSizeString = "18'3""";
                break;
            case 244:
                NearestTrimSizeString = "20'4""";
                break;
        }

        // output
        if ((NumericOutput == false)) {
            NearestTrimSize = NearestTrimSizeString;
        }
        else if ((NumericOutput == true)) {
            NearestTrimSize = tSize;
        }

        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for generating framed opening materials
        FOMaterialGen((<Worksheet>(MatSht)), (<Collection>(TrimCollection)), (<Collection>(MiscCollection)));
        let FOCell: Range;
        let hTrim: clsTrim;
        let jTrim: clsTrim;
        let CanopyQty: number;
        let DoorSlabNoGlassQty: number;
        let DoorSlabGlassQty: number;
        let PDoorMaterials: Collection;
        let OHDoorMaterials: Collection;
        let WindowMaterials: Collection;
        let MiscFOMaterials: Collection;
        let PDoors: Collection;
        let OHDoors: Collection;
        let Windows: Collection;
        let MiscFOs: Collection;
        let FO: clsFO;
        let m: number;
        let m2: number;
        let WriteCell: Range;
        let FOMaterial: clsTrim;
        let FOTrimColor: string;
        let FOWidth: number;
        let FOHeight: number;
        let RemainingWidth: number;
        let RemainingHeight: number;
        let HeadTrimMeasurement: string;
        let HeadTrimLength: number;
        // Canopy
        let q: number;
        let CombinedLength: number;
        let tType: string;
        let NewQuantity: number;
        // new material collections
        PDoorMaterials = new Collection();
        OHDoorMaterials = new Collection();
        WindowMaterials = new Collection();
        MiscFOMaterials = new Collection();
        // new framed opening collections
        PDoors = new Collection();
        OHDoors = new Collection();
        Windows = new Collection();
        MiscFOs = new Collection();
        // Personnel Door vars
        let DoorSize: string;
        // ' Even though these aren't trim, trim class is used for convenience
        let JambKit: clsTrim;
        let DoorSlab: clsTrim;
        // With...
        // set trim color
        FOTrimColor = EstSht.Range;
        "FO_tColor".Value;
        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Personnel Doors
        for (FOCell in Range(EstSht.Range, "pDoorCell1", EstSht.Range, "pDoorCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                // new FO class
                FO = new clsFO();
                FO.FOType = "PDoor";
                FO.Height = (7 * 12);
                DoorSize = FOCell.offset(0, 1).Value;
                // based off of size, select trim measurement appropriately
                if ((DoorSize == "3070")) {
                    FO.Width = ((3 * 12)
                                + 3);
                }
                else if ((DoorSize == "4070")) {
                    FO.Width = ((4 * 12)
                                + 3);
                }

                // ''door slab
                // with glass
                if ((FOCell.offset(0, 3).Value == "Yes")) {
                    DoorSlab = new clsTrim();
                    // deadbolt
                    if ((FOCell.offset(0, 6).Value == "Yes")) {
                        DoorSlab.tType = ("Door Slab W/ Deadbolt, W/ Glass" + (" - " + DoorSize));
                    }
                    else {
                        DoorSlab.tType = ("Door Slab W/O Deadbolt, W/ Glass" + (" - " + DoorSize));
                    }

                    // without glass
                }
                else {
                    DoorSlab = new clsTrim();
                    // deadbolt
                    if ((FOCell.offset(0, 6).Value == "Yes")) {
                        DoorSlab.tType = ("Door Slab W/ Deadbolt, W/O Glass" + (" - " + DoorSize));
                    }
                    else {
                        DoorSlab.tType = ("Door Slab W/O Deadbolt, W/O Glass" + (" - " + DoorSize));
                    }

                }

                DoorSlab.tMeasurement = "N/A";
                DoorSlab.Quantity = 1;
                DoorSlab.Color = "N/A";
                PDoorMaterials.Add;
                DoorSlab;
                // ''canopy
                // jamb kit
                JambKit = new clsTrim();
                JambKit.Quantity = 1;
                JambKit.tMeasurement = FOCell.offset(0, 5).Value;
                JambKit.Color = "N/A";
                if ((FOCell.offset(0, 6).Value == "Yes")) {
                    JambKit.tType = ("Jamb W/ Deadbolt" + (" - " + DoorSize));
                }
                else {
                    JambKit.tType = ("Jamb W/O Deadbolt" + (" - " + DoorSize));
                }

                // add materials to collection
                PDoorMaterials.Add;
                JambKit;
                // add FO to collection
                PDoors.Add;
                FO;
            }

        }

        // ' generate Personnel Door Trim
        if ((PDoors.Count != 0)) {
            OptimalFOTrimGen(PDoorMaterials, PDoors, "PDoors");
        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhead
        for (FOCell in Range(EstSht.Range, "OHDoorCell1", EstSht.Range, "OHDoorCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                FO = new clsFO();
                FO.FOType = "OHDoor";
                FOWidth = (FOCell.offset(0, 1).Value * 12);
                FOHeight = (FOCell.offset(0, 2).Value * 12);
                // fo class info
                FO.Height = FOHeight;
                FO.Width = (FOWidth + 3);
                OHDoors.Add;
                FO;
            }

        }

        // ' generate Overhead Door Trim
        if ((OHDoors.Count != 0)) {
            OptimalFOTrimGen(OHDoorMaterials, OHDoors, "OHDoors");
        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Windows
        for (FOCell in Range(EstSht.Range, "WindowCell1", EstSht.Range, "WindowCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                FO = new clsFO();
                FO.FOType = "Window";
                FOWidth = (FOCell.offset(0, 1).Value + 3);
                FOHeight = FOCell.offset(0, 2).Value;
                FO.Width = FOWidth;
                FO.Height = FOHeight;
                // add window to FO collection
                Windows.Add;
                FO;
            }

        }

        // ' generate Window Trim
        if ((Windows.Count != 0)) {
            OptimalFOTrimGen(WindowMaterials, Windows, "Windows");
        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Misc FO
        for (FOCell in Range(EstSht.Range, "MiscFOCell1", EstSht.Range, "MiscFOCell12")) {
            // if cell isn't hidden, door size is entered
            if (((FOCell.EntireRow.Hidden == false)
                        && (FOCell.offset(0, 1).Value != ""))) {
                FO = new clsFO();
                FO.FOType = "MiscFO";
                FOWidth = ((FOCell.offset(0, 1).Value * 12)
                            + 3);
                FOHeight = (FOCell.offset(0, 2).Value * 12);
                // add lengths to FO class
                FO.Width = FOWidth;
                FO.Height = FOHeight;
                // add FO to MiscFO collection
                MiscFOs.Add;
                FO;
            }

        }

        // ' generate Misc FO Trim
        if ((MiscFOs.Count != 0)) {
            OptimalFOTrimGen(MiscFOMaterials, MiscFOs, "MiscFOs");
        }

        // remove duplicate materials, combine
        DuplicateMaterialRemoval(PDoorMaterials, "Trim");
        DuplicateMaterialRemoval(OHDoorMaterials, "Trim");
        DuplicateMaterialRemoval(WindowMaterials, "Trim");
        DuplicateMaterialRemoval(MiscFOMaterials, "Trim");
        TrimCombine(OHDoorMaterials);
        TrimCombine(WindowMaterials);
        TrimCombine(MiscFOMaterials);
        // write
        // With...
        // if no trim, then delete all headings and exit sub
        if (((PDoorMaterials.Count == 0)
                    && ((OHDoorMaterials.Count == 0)
                    && ((WindowMaterials.Count == 0)
                    && (MiscFOMaterials.Count == 0))))) {
            Range(MatSht.Range, "PDoorMatQtyCell1".offset(-4, 0), MatSht.Range, "MiscFOMatQtyCell1".offset(2, 0)).EntireRow.Delete;
            return;
        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Personnel Doors
        if ((PDoorMaterials.Count == 0)) {
            MatSht.Range;
            "PDoorMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
        }
        else {
            WriteCell = MatSht.Range;
            "PDoorMatQtyCell1";
            for (FOMaterial in PDoorMaterials) {
                // insert new row if not the first write cell in the section
                if ((WriteCell != MatSht.Range)) {
                    "PDoorMatQtyCell1";
                    MatSht.Rows;
                    (WriteCell.Row + 1).Insert;
                    Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge;
                    // add materials
                    WriteCell.Value = FOMaterial.Quantity;
                    WriteCell.offset(0, 1).Value = FOMaterial.tType;
                    WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement;
                    WriteCell.offset(0, 4).Value = FOMaterial.Color;
                    // update write cell
                    WriteCell = WriteCell.offset(1, 0);
                    FOMaterial;
                    // ''Canopys
                    if ((CanopyQty != 0)) {
                        MatSht.Rows;
                        (WriteCell.Row + 1).Insert;
                        Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge;
                        WriteCell.Value = CanopyQty;
                        WriteCell.offset(0, 1).Value = "Canopy";
                        WriteCell.offset(0, 3).Value = "N/A";
                        WriteCell.offset(0, 4).Value = "N/A";
                        WriteCell = WriteCell.offset(1, 0);
                    }

                }

                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Overhead Doors
                if ((OHDoorMaterials.Count == 0)) {
                    MatSht.Range;
                    "OHDoorMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                }
                else {
                    WriteCell = MatSht.Range;
                    "OHDoorMatQtyCell1";
                    for (FOMaterial in OHDoorMaterials) {
                        // insert new row if not the first write cell in the section
                        if ((WriteCell.Address != MatSht.Range)) {
                            "OHDoorMatQtyCell1".Address;
                            MatSht.Rows;
                            (WriteCell.Row + 1).Insert;
                            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge;
                            // add materials
                            WriteCell.Value = FOMaterial.Quantity;
                            WriteCell.offset(0, 1).Value = FOMaterial.tType;
                            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement;
                            WriteCell.offset(0, 4).Value = FOMaterial.Color;
                            // update write cell
                            WriteCell = WriteCell.offset(1, 0);
                            FOMaterial;
                        }

                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Windows
                        if ((WindowMaterials.Count == 0)) {
                            MatSht.Range;
                            "WindowMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                        }
                        else {
                            WriteCell = MatSht.Range;
                            "WindowMatQtyCell1";
                            for (FOMaterial in WindowMaterials) {
                                // insert new row if not the first write cell in the section
                                if ((WriteCell.Address != MatSht.Range)) {
                                    "WindowMatQtyCell1".Address;
                                    MatSht.Rows;
                                    (WriteCell.Row + 1).Insert;
                                    Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge;
                                    // add materials
                                    WriteCell.Value = FOMaterial.Quantity;
                                    WriteCell.offset(0, 1).Value = FOMaterial.tType;
                                    WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement;
                                    WriteCell.offset(0, 4).Value = FOMaterial.Color;
                                    // update write cell
                                    WriteCell = WriteCell.offset(1, 0);
                                    FOMaterial;
                                }

                                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Misc FOs
                                if ((MiscFOMaterials.Count == 0)) {
                                    MatSht.Range;
                                    "MiscFOMatQtyCell1".offset(-2, 0).Resize(4, 1).EntireRow.Delete;
                                }
                                else {
                                    WriteCell = MatSht.Range;
                                    "MiscFOMatQtyCell1";
                                    for (FOMaterial in MiscFOMaterials) {
                                        // insert new row if not the first write cell in the section
                                        if ((WriteCell.Address != MatSht.Range)) {
                                            "MiscFOMatQtyCell1".Address;
                                            MatSht.Rows;
                                            (WriteCell.Row + 1).Insert;
                                            Range(WriteCell.offset(0, 1), WriteCell.offset(0, 2)).Merge;
                                            // add materials
                                            WriteCell.Value = FOMaterial.Quantity;
                                            WriteCell.offset(0, 1).Value = FOMaterial.tType;
                                            WriteCell.offset(0, 3).Value = FOMaterial.tMeasurement;
                                            WriteCell.offset(0, 4).Value = FOMaterial.Color;
                                            // update write cell
                                            WriteCell = WriteCell.offset(1, 0);
                                            FOMaterial;
                                        }

                                        // With...
                                        for (FOMaterial in PDoorMaterials) {
                                            if ((((FOMaterial.tType.IndexOf("Head Trim", 0) + 1)
                                                        != 0)
                                                        || (FOMaterial.tType == "Jamb Trim"))) {
                                                FOMaterial.tShape = "R-Loc";
                                            }
                                            else {
                                                // door slaps, windows, canpoies, etc.
                                                FOMaterial.tShape = "N/A";
                                            }

                                            TrimCollection.Add;
                                            FOMaterial;
                                        }

                                        for (FOMaterial in OHDoorMaterials) {
                                            if ((((FOMaterial.tType.IndexOf("Head Trim", 0) + 1)
                                                        != 0)
                                                        || (FOMaterial.tType == "Jamb Trim"))) {
                                                FOMaterial.tShape = "R-Loc";
                                            }

                                            TrimCollection.Add;
                                            FOMaterial;
                                        }

                                        for (FOMaterial in WindowMaterials) {
                                            if ((((FOMaterial.tType.IndexOf("Head Trim", 0) + 1)
                                                        != 0)
                                                        || (FOMaterial.tType == "Jamb Trim"))) {
                                                FOMaterial.tShape = "R-Loc";
                                            }

                                            TrimCollection.Add;
                                            FOMaterial;
                                        }

                                        for (FOMaterial in MiscFOMaterials) {
                                            if ((((FOMaterial.tType.IndexOf("Head Trim", 0) + 1)
                                                        != 0)
                                                        || (FOMaterial.tType == "Jamb Trim"))) {
                                                FOMaterial.tShape = "R-Loc";
                                            }

                                            TrimCollection.Add;
                                            FOMaterial;
                                        }

                                        DuplicateMaterialRemoval(/* ref */(<Collection>(MaterialCollection)), Optional, (<string>(CollectionType)));
                                        let m: number;
                                        let m2: number;
                                        if ((CollectionType == "Trim")) {
                                            for (m = 1; (m <= MaterialCollection.Count); m++) {
                                                // check that not already flagged
                                                if (((MaterialCollection(m).DeleteFlag == false)
                                                            && (MaterialCollection(m).clsType == "Trim"))) {
                                                    // check for duplicate measurements
                                                    for (m2 = 1; (m2 <= MaterialCollection.Count); m2++) {
                                                        if ((MaterialCollection(m).clsType == MaterialCollection(m2).clsType)) {
                                                            // check that not the same material, not flagged for deletion, and duplicate measurement
                                                            if (((m2 != m)
                                                                        && ((MaterialCollection(m).tType == MaterialCollection(m2).tType)
                                                                        && ((MaterialCollection(m).Color == MaterialCollection(m2).Color)
                                                                        && ((MaterialCollection(m2).DeleteFlag == false)
                                                                        && (MaterialCollection(m2).tMeasurement == MaterialCollection(m).tMeasurement)))))) {
                                                                // add quantity to existing class
                                                                MaterialCollection(m).Quantity = (MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity);
                                                                // flag duplicate for deletion
                                                                MaterialCollection(m2).DeleteFlag = true;
                                                            }

                                                        }

                                                    }

                                                }

                                            }

                                        }
                                        else if ((CollectionType == "Panel")) {
                                            for (m = 1; (m <= MaterialCollection.Count); m++) {
                                                // check that not already flagged
                                                if ((MaterialCollection(m).DeleteFlag == false)) {
                                                    // check for duplicate measurements
                                                    for (m2 = 1; (m2 <= MaterialCollection.Count); m2++) {
                                                        // check that not the same material, not flagged for deletion, and duplicate measurement
                                                        if (((m2 != m)
                                                                    && ((MaterialCollection(m).PanelType == MaterialCollection(m2).PanelType)
                                                                    && ((MaterialCollection(m).PanelColor == MaterialCollection(m2).PanelColor)
                                                                    && ((MaterialCollection(m2).DeleteFlag == false)
                                                                    && (MaterialCollection(m2).PanelMeasurement == MaterialCollection(m).PanelMeasurement)))))) {
                                                            // add quantity to existing class
                                                            MaterialCollection(m).Quantity = (MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity);
                                                            // flag duplicate for deletion
                                                            MaterialCollection(m2).DeleteFlag = true;
                                                        }

                                                    }

                                                }

                                            }

                                        }
                                        else if ((CollectionType == "Misc")) {
                                            for (m = 1; (m <= MaterialCollection.Count); m++) {
                                                // check that not already flagged
                                                if ((MaterialCollection(m).DeleteFlag == false)) {
                                                    // check for duplicate measurements
                                                    for (m2 = 1; (m2 <= MaterialCollection.Count); m2++) {
                                                        // check that not the same material, not flagged for deletion, and duplicate measurement
                                                        if (((m2 != m)
                                                                    && ((MaterialCollection(m).Name == MaterialCollection(m2).Name)
                                                                    && ((MaterialCollection(m).Color == MaterialCollection(m2).Color)
                                                                    && ((MaterialCollection(m2).DeleteFlag == false)
                                                                    && (MaterialCollection(m2).Measurement == MaterialCollection(m).Measurement)))))) {
                                                            // add quantity to existing class
                                                            MaterialCollection(m).Quantity = (MaterialCollection(m).Quantity + MaterialCollection(m2).Quantity);
                                                            // flag duplicate for deletion
                                                            MaterialCollection(m2).DeleteFlag = true;
                                                        }

                                                    }

                                                }

                                            }

                                        }
                                        else if ((CollectionType == "Steel")) {
                                            for (m = 1; (m <= MaterialCollection.Count); m++) {
                                                // Check that not already flagged
                                                if ((MaterialCollection(m).DeleteFlag == false)) {
                                                    // check for duplicate measurements
                                                    for (m2 = 1; (m2 <= MaterialCollection.Count); m2++) {
                                                        if (((MaterialCollection(m).clsType == "Member")
                                                                    && (MaterialCollection(m2).clsType == "Member"))) {
                                                            // check that not the same material, not flagged for deletion, and duplicate measurement
                                                            if (((m2 != m)
                                                                        && ((MaterialCollection(m).Length == MaterialCollection(m2).Length)
                                                                        && ((MaterialCollection(m).Size == MaterialCollection(m2).Size)
                                                                        && ((MaterialCollection(m2).clsType == "Member")
                                                                        && (MaterialCollection(m2).DeleteFlag == false)))))) {
                                                                // Add quatities
                                                                // Debug.Print vbNewLine & "m - " & m & " - " & MaterialCollection(m).Length & " - " & MaterialCollection(m).DeleteFlag & " - " & MaterialCollection(m).Placement
                                                                // Debug.Print "m2 - " & m2 & " - " & MaterialCollection(m2).Length & " - " & MaterialCollection(m2).DeleteFlag & " - " & MaterialCollection(m2).Placement
                                                                MaterialCollection(m).Qty = (MaterialCollection(m).Qty + MaterialCollection(m2).Qty);
                                                                // flag duplicate for deletion
                                                                MaterialCollection(m2).DeleteFlag = true;
                                                            }

                                                        }

                                                    }

                                                }

                                            }

                                        }

                                        for (m = MaterialCollection.Count; (m <= 1); m = (m + -1)) {
                                            if ((MaterialCollection(m).DeleteFlag == true)) {
                                                MaterialCollection.Remove;
                                                m;
                                            }

                                        }

                                    }

                                }

                            }

                        }

                    }

                }

            }

        }

    }

}
private TrimCombine(/* ref */MaterialCollection: Collection) {
    let m: number;
    let tType: string;
    let CombinedLength: number;
    let FOMaterial: clsTrim;
    let q: number;
    for (m = 1; (m <= MaterialCollection.Count); m++) {
        // With...
        // skip materials other than trim
        if ((((MaterialCollection(m).tType.IndexOf("Door", 0) + 1)
                    != 0)
                    || (((MaterialCollection(m).tType.IndexOf("Deadbolt", 0) + 1)
                    != 0)
                    || ((MaterialCollection(m).tType.IndexOf("Canopy", 0) + 1)
                    != 0)))) {
            /* Warning! GOTO is not Implemented */}

        // '''''''''''''''''FOR NOW, Only combine 10'2 pieces of trim
        if ((MaterialCollection(m).tLength != 122)) {
            /* Warning! GOTO is not Implemented */}

        // '''''''''''''''''FOR NOW, Only combine 10'2 pieces of trim
        // check that trim can be combined
        if (((MaterialCollection(m).Quantity > 1)
                    && (MaterialCollection(m).tLength <= 122))) {
            // flag trim piece for deletion
            MaterialCollection(m).DeleteFlag = true;
            if (((MaterialCollection(m).tType.IndexOf("Head", 0) + 1)
                        != 0)) {
                tType = "Head";
            }
            else if (((MaterialCollection(m).tType.IndexOf("Jamb", 0) + 1)
                        != 0)) {
                tType = "Jamb";
            }
            else {
                tType = "";
            }

            // reset combined length
            CombinedLength = 0;
            for (q = 1; (q <= MaterialCollection(m).Quantity); q++) {
                // check if can add to previous piece and be less than 20'4"
                if (((CombinedLength + MaterialCollection(m).tLength)
                            <= 244)) {
                    // combine with previous trim piece
                    CombinedLength = (CombinedLength + MaterialCollection(m).tLength);
                }
                else {
                    // add a new piece with the new trim length
                    FOMaterial = new clsTrim();
                    FOMaterial.Color = MaterialCollection(m).Color;
                    FOMaterial.tType = MaterialCollection(m).tType;
                    FOMaterial.tMeasurement = NearestTrimSize(CombinedLength, 1, tType);
                    FOMaterial.tLength = NearestTrimSize(CombinedLength, 1, tType, true);
                    FOMaterial.Quantity = 1;
                    MaterialCollection.Add;
                    FOMaterial;
                    // reset combined to current piece
                    CombinedLength = (0 + MaterialCollection(m).tLength);
                }

            }

            // If left over material, add to collection
            if ((CombinedLength != 0)) {
                FOMaterial = new clsTrim();
                FOMaterial.Color = MaterialCollection(m).Color;
                FOMaterial.tType = MaterialCollection(m).tType;
                FOMaterial.tMeasurement = NearestTrimSize(CombinedLength, 1, tType);
                FOMaterial.tLength = NearestTrimSize(CombinedLength, 1, tType, true);
                FOMaterial.Quantity = 1;
                MaterialCollection.Add;
                FOMaterial;
            }

        }

        /* Warning! Labeled Statements are not Implemented */}

    // remove duplicate material
    DuplicateMaterialRemoval(MaterialCollection, "Trim");
}

private OptimalFOTrimGen(/* ref */MaterialCollection: Collection, /* ref */FOs: Collection, FOType: string) {
    let FO: clsFO;
    let SplitFO: clsFO;
    let CombinedWidth: number;
    let CombinedHeight: number;
    let FOMaterial: clsTrim;
    let TrimPiece: clsTrim;
    let item: Object;
    let tPiece1Length: number;
    let tPiece2Length: number;
    let tPiece3Length: number;
    let tColor: string;
    let m: number;
    let jTrimTotalLength: number;
    let jTrimRemainder: number;
    let jTrim20FtPieces: number;
    let Trim20FtPieceCount: number;
    let AltGrouping: boolean;
    let jTrimRemaining: number;
    let SplitJambTrim: boolean;
    // BPP Solver Adaptation Vars
    let BPP_TrimCollection: Collection;
    let NumTrimPieces: number;
    // '' Debug flag for skipping section
    let DebugFlag: boolean;
    // '' debug mode
    DebugFlag = false;
    if ((DebugFlag == true)) {
        return;
    }

    // find FO Trim color
    tColor = EstSht.Range("FO_tColor").Value;
    // '' 'generate jamb trim collection
    BPP_TrimCollection = new Collection();
    for (FO in FOs) {
        // determine the number of trim pieces
        switch (FO.Height) {
        }

        ((20 * 12)
                    + 4);
        FOMaterial = new clsTrim();
        NumTrimPieces = 1;
        FOMaterial.tLength = FO.Height;
        FOMaterial.tMeasurement = ImperialMeasurementFormat(FO.Height);
        FOMaterial.Quantity = 2;
        BPP_TrimCollection.Add;
        FOMaterial;
        // '' two pieces of trim
        ((20 * (2 * 12))
                    + ((4 * 2)
                    - 2));
        tPiece1Length = NearestTrimSize(((FO.Height / 2)
                        + 1), 0, "Jamb", true);
        tPiece2Length = NearestTrimSize(((FO.Height
                        - (tPiece1Length - 1))
                        + 1), 0, "Jamb", true);
        // add directly to material collection for now
        FOMaterial = new clsTrim();
        FOMaterial.tLength = tPiece1Length;
        FOMaterial.Quantity = 2;
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length);
        FOMaterial.tType = "Jamb Trim";
        BPP_TrimCollection.Add;
        FOMaterial;
        FOMaterial = new clsTrim();
        FOMaterial.tLength = tPiece2Length;
        FOMaterial.Quantity = 2;
        FOMaterial.tType = "Jamb Trim";
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length);
        BPP_TrimCollection.Add;
        FOMaterial;
        // '' three pieces of trim
        // find best overlapping trim sizes (accounting for 1" overlap)
        tPiece1Length = NearestTrimSize(((FO.Height / 3)
                        + 1), 0, "Jamb", true);
        tPiece2Length = NearestTrimSize(((FO.Height / 3)
                        + 1), 0, "Jamb", true);
        // add 2" overlap to the middle piece
        tPiece3Length = NearestTrimSize(((FO.Height
                        - ((tPiece1Length - 1)
                        - (tPiece2Length - 1)))
                        + 2), 0, "Jamb", true);
        // add directly to material collection for now
        FOMaterial = new clsTrim();
        FOMaterial.tLength = tPiece1Length;
        FOMaterial.Quantity = 2;
        FOMaterial.tType = "Jamb Trim";
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length);
        BPP_TrimCollection.Add;
        FOMaterial;
        FOMaterial = new clsTrim();
        FOMaterial.tLength = tPiece2Length;
        FOMaterial.tType = "Jamb Trim";
        FOMaterial.Quantity = 2;
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length);
        MaterialCollection.Add;
        FOMaterial;
        FOMaterial = new clsTrim();
        FOMaterial.tLength = tPiece3Length;
        FOMaterial.Quantity = 2;
        FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece3Length);
        BPP_TrimCollection.Add;
        FOMaterial;
        FO;
        // solve jamb trim collection
        DuplicateMaterialRemoval(BPP_TrimCollection, "Trim");
        JankyBPPSolver.BPP_Solver(MaterialCollection, BPP_TrimCollection, "Jamb", FOType);
        // '' 'generate Head trim collection
        BPP_TrimCollection = new Collection();
        for (FO in FOs) {
            // determine the number of trim pieces
            switch (FO.Width) {
            }

            ((20 * 12)
                        + 4);
            FOMaterial = new clsTrim();
            NumTrimPieces = 1;
            FOMaterial.tLength = FO.Width;
            FOMaterial.tMeasurement = ImperialMeasurementFormat(FO.Width);
            FOMaterial.Quantity = 1;
            BPP_TrimCollection.Add;
            FOMaterial;
            // '' two pieces of trim
            ((20 * (2 * 12))
                        + ((4 * 2)
                        - 2));
            tPiece1Length = NearestTrimSize(((FO.Width / 2)
                            + 1), 0, "Head", true);
            tPiece2Length = NearestTrimSize(((FO.Width
                            - (tPiece1Length - 1))
                            + 1), 0, "Head", true);
            FOMaterial = new clsTrim();
            FOMaterial.tLength = tPiece1Length;
            FOMaterial.tType = "Head Trim W/ Kickout";
            FOMaterial.Quantity = 1;
            FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length);
            BPP_TrimCollection.Add;
            FOMaterial;
            FOMaterial = new clsTrim();
            FOMaterial.tLength = tPiece2Length;
            FOMaterial.tType = "Head Trim W/ Kickout";
            FOMaterial.Quantity = 1;
            FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length);
            BPP_TrimCollection.Add;
            FOMaterial;
            // '' three pieces of trim
            // find best overlapping trim sizes (accounting for 1" overlap)
            tPiece1Length = NearestTrimSize(((FO.Width / 3)
                            + 1), 0, "Head", true);
            tPiece2Length = NearestTrimSize(((FO.Width / 3)
                            + 1), 0, "Head", true);
            // add 2" overlap to the middle piece
            tPiece3Length = NearestTrimSize(((FO.Width
                            - ((tPiece1Length - 1)
                            - (tPiece2Length - 1)))
                            + 2), 0, "Head", true);
            FOMaterial = new clsTrim();
            FOMaterial.tLength = tPiece1Length;
            FOMaterial.Quantity = 1;
            FOMaterial.tType = "Head Trim W/ Kickout";
            FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece1Length);
            BPP_TrimCollection.Add;
            FOMaterial;
            FOMaterial = new clsTrim();
            FOMaterial.tLength = tPiece2Length;
            FOMaterial.Quantity = 1;
            FOMaterial.tType = "Head Trim W/ Kickout";
            FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece2Length);
            BPP_TrimCollection.Add;
            FOMaterial;
            FOMaterial = new clsTrim();
            FOMaterial.tLength = tPiece3Length;
            FOMaterial.Quantity = 1;
            FOMaterial.tType = "Head Trim W/ Kickout";
            FOMaterial.tMeasurement = ImperialMeasurementFormat(tPiece3Length);
            BPP_TrimCollection.Add;
            FOMaterial;
            FO;
            // solve Head Trim collection
            DuplicateMaterialRemoval(BPP_TrimCollection, "Trim");
            JankyBPPSolver.BPP_Solver(MaterialCollection, BPP_TrimCollection, "Head", FOType);
            // add duplicate head trip without kickout for Windows and Misc Fos
            if (((FOType == "Windows")
                        || (FOType == "MiscFOs"))) {
                for (item in MaterialCollection) {
                    if ((item.clsType == "Trim")) {
                        FOMaterial = item;
                        // add duplicate head trim without kickout
                        if ((FOMaterial.tType == "Head Trim W/ Kickout")) {
                            TrimPiece = new clsTrim();
                            TrimPiece.Quantity = FOMaterial.Quantity;
                            TrimPiece.tMeasurement = FOMaterial.tMeasurement;
                            TrimPiece.tType = "Head Trim W/O Kickout";
                            MaterialCollection.Add;
                            TrimPiece;
                        }

                    }

                }

            }

            //     'add duplicate head trim without kickout if a Misc FO
            //     If FOType = "MiscFOs" Then
            //         For m = 1 To MaterialCollection.Count
            //             With MaterialCollection(m)
            //                 If .tType = "Head Trim W/ Kickout" Then
            //                     'add equivalent head trim without kickout
            //                      Set FOMaterial = New clsTrim
            //                     FOMaterial.Color = tColor
            //                     FOMaterial.tType = "Head Trim W/O Kickout"
            //                     FOMaterial.tMeasurement = .tMeasurement
            //                     FOMaterial.tLength = .tLength
            //                     FOMaterial.Quantity = 1
            //                     MaterialCollection.Add FOMaterial
            //                 End If
            //             End With
            //         Next m
            //     End If
            // add trim color
            for (item in MaterialCollection) {
                if ((item.clsType == "Trim")) {
                    FOMaterial = item;
                    switch (FOMaterial.tType) {
                        case "Head Trim W/ Kickout":
                        case "Head Trim W/O Kickout":
                        case "Jamb Trim":
                            FOMaterial.Color = tColor;
                            break;
                        default:
                            FOMaterial.Color = "N/A";
                            break;
                    }

                }

            }

            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for calculating panel collections for the roof
            // Note: Panel uses Rafter lengths with overhangs already factored in
            RoofPanelGen(/* ref */(<Collection>(PanelCollection)), (<number>(RafterSheetLength)), (<number>(EaveOverhang)), (<number>(RoofLength)), Optional, (<string>(rShape)), Optional, (<boolean>(EaveExtension)));
            let Overlap: number;
            let p1: clsPanel;
            let p2: clsPanel;
            let p3: clsPanel;
            let p4: clsPanel;
            let p5: clsPanel;
            let p1Length: number;
            let p1Measurement: string;
            let p2Length: number;
            let p2Measurement: string;
            let p3Length: number;
            let p3Measurement: string;
            let p4Length: number;
            let p4Measurement: string;
            let p5Length: number;
            let p5Measurement: string;
            let PanelQty: number;
            let IdealPLength: number;
            let RemainingLength: number;
            let LargeOverhang: boolean;
            let OnePanel: boolean;
            let TwoPanel: boolean;
            let ThreePanel: boolean;
            let FourPanel: boolean;
            let FivePanel: boolean;
            let SixPanel: boolean;
            // standard panel overlap of 6"
            Overlap = 6;
            //  Check for overhang greater than 1.5 ft
            if ((EaveOverhang > (1.5 * 12))) {
                LargeOverhang = true;
            }

            switch (true) {
                case (RafterSheetLength <= (42 * 12)):
                    OnePanel = true;
                    break;
                case (RafterSheetLength <= (83 * 12)):
                    if ((LargeOverhang == false)) {
                        p1Length = ((40 * 12)
                                    + EaveOverhang);
                    }
                    else {
                        p1Length = ((35 * 12)
                                    + EaveOverhang);
                    }

                    p2Length = (RafterSheetLength - p1Length);
                    if ((p2Length <= (42 * 12))) {
                        TwoPanel = true;
                    }
                    else {
                        ThreePanel = true;
                    }

                    // three panels
                    break;
                case (RafterSheetLength <= (124 * 12)):
                    if ((LargeOverhang == false)) {
                        p1Length = ((40 * 12)
                                    + EaveOverhang);
                    }
                    else {
                        p1Length = ((35 * 12)
                                    + EaveOverhang);
                    }

                    p2Length = (40 * 12);
                    p3Length = (RafterSheetLength
                                - (p1Length - p2Length));
                    if ((p3Length <= (42 * 12))) {
                        ThreePanel = true;
                    }
                    else {
                        FourPanel = true;
                    }

                    // four panels
                    break;
                case (RafterSheetLength <= (165 * 12)):
                    if ((LargeOverhang == false)) {
                        p1Length = ((40 * 12)
                                    + EaveOverhang);
                    }
                    else {
                        p1Length = ((35 * 12)
                                    + EaveOverhang);
                    }

                    p2Length = (40 * 12);
                    p3Length = p2Length;
                    p4Length = (RafterSheetLength
                                - (p1Length
                                - (p2Length - p3Length)));
                    if ((p4Length <= (42 * 12))) {
                        FourPanel = true;
                    }
                    else {
                        FivePanel = true;
                    }

                    // five panels
                    break;
                case (RafterSheetLength <= (206 * 12)):
                    if ((LargeOverhang == false)) {
                        p1Length = ((40 * 12)
                                    + EaveOverhang);
                    }
                    else {
                        p1Length = ((35 * 12)
                                    + EaveOverhang);
                    }

                    p2Length = (40 * 12);
                    p3Length = p2Length;
                    p4Length = p2Length;
                    p5Length = (RafterSheetLength
                                - (p1Length
                                - (p2Length
                                - (p3Length - p4Length))));
                    if ((p5Length <= (42 * 12))) {
                        FivePanel = true;
                    }
                    else {
                        SixPanel = true;
                    }

                    break;
                default:
                    SixPanel = true;
                    break;
            }

            //  panel quantity
            PanelQty = Application.WorksheetFunction.RoundUp((RoofLength / 3), 0);
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' roof panels for one sidewall'''''
            //  at least 1 panel
            p1 = new clsPanel();
            // Debug.Print "Rafter Measurement: " & ImperialMeasurementFormat(RafterLength)
            // Debug.Print "Rafter Length: " & RafterLength / 12
            // Debug.Print "Ideal Panel Length: " & (RafterLength / 12) / 2
            // Debug.Print "Ideal Panel Length (minus overhang): " & ((RafterLength - EaveOverhang) / 12) / 2
            // '' determine how many panels to make
            switch (true) {
                case OnePanel:
                    p1.PanelLength = RafterSheetLength;
                    // add underlap for eave extension panels
                    if ((EaveExtension == true)) {
                        p1.PanelLength = (p1.PanelLength + 6);
                    }

                    // '' convert to imperial
                    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength);
                    // quantity
                    p1.Quantity = PanelQty;
                    // add to collection
                    PanelCollection.Add;
                    p1;
                    // ' check for 2 divisions
                    break;
                case TwoPanel:
                    // new panel class
                    p2 = new clsPanel();
                    // include overhang on ideal panel length when comparing 2 panel options
                    IdealPLength = (RafterSheetLength / 2);
                    // determine panel 1 length
                    if ((LargeOverhang == false)) {
                        // round p1 length to panel length that doesn't exceed 40'
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang);
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                        }

                    }
                    else if ((LargeOverhang == true)) {
                        // round p1 length to panel length that doesn't exceed 35
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (35 * 12))) {
                            // manually check
                            if ((Abs((IdealPLength
                                            - (RafterSheetLength
                                            - (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang)))) < Abs((IdealPLength
                                            - (RafterSheetLength
                                            - (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang)))))) {
                                p1.PanelLength = (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang);
                            }
                            else {
                                p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            }

                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                        }

                    }

                    // add in overhang
                    // p1.PanelLength = p1.PanelLength + EaveOverhang
                    // find p2 length
                    p2.PanelLength = (RafterSheetLength - p1.PanelLength);
                    // add overlap
                    p2.PanelLength = (p2.PanelLength + Overlap);
                    p1.PanelLength = (p1.PanelLength + Overlap);
                    // convert lengths, calculate quantities
                    // '' convert to imperial
                    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength);
                    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength);
                    // quantity
                    p1.Quantity = PanelQty;
                    p2.Quantity = PanelQty;
                    // add to collection
                    PanelCollection.Add;
                    p1;
                    PanelCollection.Add;
                    p2;
                    // ' check for 3
                    break;
                case ThreePanel:
                    // new panel classes
                    p2 = new clsPanel();
                    p3 = new clsPanel();
                    IdealPLength = (RafterSheetLength / 3);
                    // determine panel 1 length (short side)
                    if ((LargeOverhang == false)) {
                        // round p1 length to panel length that doesn't exceed 40'
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                        }

                    }
                    else if ((LargeOverhang == true)) {
                        // round p1 length to panel length that doesn't exceed 35
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (35 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            // check about other panel lengths
                            if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                            }
                            else {
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            }

                        }

                    }

                    // add in overhang
                    // p1.PanelLength = p1.PanelLength + EaveOverhang
                    // find remaining rafter length
                    p3.PanelLength = (RafterSheetLength
                                - (p1.PanelLength - p2.PanelLength));
                    // add overlap, add undercut back in (because it isn't undercut)
                    p3.PanelLength = (p3.PanelLength + Overlap);
                    // add two overlaps, deduct overhang and add undercut back in
                    p2.PanelLength = (p2.PanelLength
                                + (Overlap * 2));
                    p1.PanelLength = (p1.PanelLength + Overlap);
                    // convert lengths, calculate quantities
                    // '' convert to imperial
                    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength);
                    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength);
                    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength);
                    // quantity
                    p1.Quantity = PanelQty;
                    p2.Quantity = PanelQty;
                    p3.Quantity = PanelQty;
                    // add to collection
                    PanelCollection.Add;
                    p1;
                    PanelCollection.Add;
                    p2;
                    PanelCollection.Add;
                    p3;
                    // remove duplicates
                    DuplicateMaterialRemoval(PanelCollection, "Panel");
                    // 'check for 4
                    break;
                case FourPanel:
                    // new panel classes
                    p2 = new clsPanel();
                    p3 = new clsPanel();
                    p4 = new clsPanel();
                    // ideal
                    IdealPLength = (RafterSheetLength / 4);
                    // determine panel 1 length (short side)
                    if ((LargeOverhang == false)) {
                        // round p1 length to panel length that doesn't exceed 40'
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                        }

                    }
                    else if ((LargeOverhang == true)) {
                        // round p1 length to panel length that doesn't exceed 35
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (35 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            // check about other panel lengths
                            if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                                p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                            }
                            else {
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                                p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            }

                        }

                    }

                    // add in overhang
                    // p1.PanelLength = p1.PanelLength + EaveOverhang
                    // determine remaining length
                    p4.PanelLength = (RafterSheetLength
                                - (p1.PanelLength
                                - (p2.PanelLength - p3.PanelLength)));
                    // add overlap, add undercut back in (bottom panel)
                    p4.PanelLength = (p4.PanelLength + Overlap);
                    // deduct overhang, add overlap for panel 1 (top panel)
                    p1.PanelLength = (p1.PanelLength + Overlap);
                    // add two overlaps for panel 2 deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
                    p2.PanelLength = (p2.PanelLength
                                + (Overlap * 2));
                    p3.PanelLength = (p3.PanelLength
                                + (Overlap * 2));
                    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength);
                    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength);
                    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength);
                    p4.PanelMeasurement = ImperialMeasurementFormat(p4.PanelLength);
                    // quantity
                    p1.Quantity = PanelQty;
                    p2.Quantity = PanelQty;
                    p3.Quantity = PanelQty;
                    p4.Quantity = PanelQty;
                    // add to collection
                    PanelCollection.Add;
                    p1;
                    PanelCollection.Add;
                    p2;
                    PanelCollection.Add;
                    p3;
                    PanelCollection.Add;
                    p4;
                    // ' check for 5
                    break;
                case FivePanel:
                    // new panel classes
                    p2 = new clsPanel();
                    p3 = new clsPanel();
                    p4 = new clsPanel();
                    p5 = new clsPanel();
                    // ideal
                    IdealPLength = (RafterSheetLength / 5);
                    // determine panel 1 length (short side)
                    if ((LargeOverhang == false)) {
                        // round p1 length to panel length that doesn't exceed 40'
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, 1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                            p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, (p1.PanelLength
                                                + (p2.PanelLength + p3.PanelLength))));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            p4.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                        }

                    }
                    else if ((LargeOverhang == true)) {
                        // round p1 length to panel length that doesn't exceed 35
                        if ((ClosestRoofPurlin(IdealPLength, 1) <= (35 * 12))) {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength) + EaveOverhang);
                            p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                            p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                            p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, (p1.PanelLength
                                                + (p2.PanelLength + p3.PanelLength))));
                        }
                        else {
                            p1.PanelLength = (ClosestRoofPurlin(IdealPLength, -1) + EaveOverhang);
                            // check about other panel lengths
                            if ((ClosestRoofPurlin(IdealPLength, 1) <= (40 * 12))) {
                                // check for other panel lengths
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 2, p1.PanelLength));
                                p3.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 3, (p1.PanelLength + p2.PanelLength)));
                                p4.PanelLength = ClosestRoofPurlin(IdealPLength, PanelOptionCompare(IdealPLength, 4, (p1.PanelLength
                                                    + (p2.PanelLength + p3.PanelLength))));
                            }
                            else {
                                p2.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                                p3.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                                p4.PanelLength = ClosestRoofPurlin(IdealPLength, -1);
                            }

                        }

                    }

                    // add in overhang
                    // p1.PanelLength = p1.PanelLength + EaveOverhang
                    // determine remaining length
                    p5.PanelLength = (RafterSheetLength
                                - (p1.PanelLength
                                - (p2.PanelLength
                                - (p3.PanelLength - p4.PanelLength))));
                    // add overlap, add undercut back in (bottom panel)
                    p5.PanelLength = (p5.PanelLength + Overlap);
                    // deduct overhang, add overlap for panel 1 (top panel)
                    p1.PanelLength = (p1.PanelLength + Overlap);
                    // add two overlaps for panel 2 deduct overhang and add undercut back (because it isn't overhanging or undercut) (middle panel)
                    p2.PanelLength = (p2.PanelLength
                                + (Overlap * 2));
                    p3.PanelLength = (p3.PanelLength
                                + (Overlap * 2));
                    p4.PanelLength = (p4.PanelLength
                                + (Overlap * 2));
                    p1.PanelMeasurement = ImperialMeasurementFormat(p1.PanelLength);
                    p2.PanelMeasurement = ImperialMeasurementFormat(p2.PanelLength);
                    p3.PanelMeasurement = ImperialMeasurementFormat(p3.PanelLength);
                    p4.PanelMeasurement = ImperialMeasurementFormat(p4.PanelLength);
                    p5.PanelMeasurement = ImperialMeasurementFormat(p5.PanelLength);
                    // quantity
                    p1.Quantity = PanelQty;
                    p2.Quantity = PanelQty;
                    p3.Quantity = PanelQty;
                    p4.Quantity = PanelQty;
                    p5.Quantity = PanelQty;
                    // add to collection
                    PanelCollection.Add;
                    p1;
                    PanelCollection.Add;
                    p2;
                    PanelCollection.Add;
                    p3;
                    PanelCollection.Add;
                    p4;
                    PanelCollection.Add;
                    p5;
                    break;
                default:
                    /* Warning! GOTO is not Implemented */break;
            }

            return;
            /* Warning! Labeled Statements are not Implemented */MsgBox;
            "It has been calculated that more than 5 seperate panels will be needed to cover the rafter length of the roof. Please perform this calculation manually.";
            vbExclamation;
            "Program Design Exceeded";
            return;
            SoffitGen((<Collection>(SoffitPanels)), (<Collection>(SoffitTrim)), (<string>(SoffitLocation)), (<clsBuilding>(b)), Optional, (<Collection>(s2RoofPanels)), Optional, (<Collection>(s4RoofPanels)));
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Sub for Generating Soffit Panels and Soffit Trim
            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            let RoofPanel: clsPanel;
            let SoffitPanel: clsPanel;
            let NamedRangeString: string;
            // Var for reading correct soffit panel/trim info cell
            let SoffitQty: number;
            let NetRafterLength: number;
            let NetOutsideAngleLength: number;
            let TrimPiece: clsTrim;
            let SoffitLength: number;
            let NetStandardEaveOverhang: number;
            // var for subtracting the 4.25" overhang as needed for a single slope
            let EaveExtBuildingLength: number;
            // eave extension length from endwall to endwall
            let EaveExtRafterLength: number;
            NamedRangeString = (SoffitLocation + "Soffit");
            // With...
            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Panels
            // determine Soffit Panel Quantity
            switch (SoffitLocation) {
                case "e1_GableOverhang":
                    SoffitQty = Application.WorksheetFunction.RoundUp(((b.e1Overhang / 12)
                                    / 5), 0);
                    SoffitLength = b.e1Overhang;
                    break;
                case "e3_GableOverhang":
                    SoffitQty = Application.WorksheetFunction.RoundUp(((b.e3Overhang / 12)
                                    / 5), 0);
                    SoffitLength = b.e3Overhang;
                    break;
                case "e1_GableExtension":
                    SoffitQty = Application.WorksheetFunction.RoundUp(((b.e1Extension / 12)
                                    / 5), 0);
                    SoffitLength = b.e1Extension;
                    break;
                case "e3_GableExtension":
                    SoffitQty = Application.WorksheetFunction.RoundUp(((b.e3Extension / 12)
                                    / 5), 0);
                    SoffitLength = b.e3Extension;
                    break;
                case "s2_EaveOverhang":
                case "s4_EaveOverhang":
                    SoffitQty = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
                    break;
                case "s2_EaveExtension":
                    EaveExtBuildingLength = (b.s2EaveExtensionBuildingLength / 12);
                    EaveExtRafterLength = b.s2ExtensionRafterLength;
                    break;
                case "s4_EaveExtension":
                    EaveExtBuildingLength = (b.s4EaveExtensionBuildingLength / 12);
                    EaveExtRafterLength = b.s4ExtensionRafterLength;
                    break;
            }

            // Generate of Gable Overhang/Extension Soffit Panels
            if (((SoffitLocation.IndexOf("Gable", 0) + 1)
                        != 0)) {
                if ((b.rShape == "Single Slope")) {
                    // subtract the standard eave overhang
                    if ((b.s4Overhang != 0)) {
                        NetStandardEaveOverhang = (4.25 * 2);
                    }
                    else {
                        NetStandardEaveOverhang = 4.25;
                    }

                    // add soffit corresponding to roof panels of sidewall 2, less the standard overhangs
                    RoofPanelGen(SoffitPanels, (b.s2RafterSheetLength - NetStandardEaveOverhang), 0, (SoffitLength / 12));
                    for (SoffitPanel in SoffitPanels) {
                        SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value;
                        SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value;
                        SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value;
                        SoffitPanel.clsType = "Panel";
                    }

                    // if a gable roof, don't subtract the eave overhang (due to the undercut) and just match each sidewall's rafter sheet length
                }
                else if ((b.rShape == "Gable")) {
                    RoofPanelGen(SoffitPanels, b.s2RafterSheetLength, b.s2Overhang, (SoffitLength / 12));
                    RoofPanelGen(SoffitPanels, b.s4RafterSheetLength, b.s4Overhang, (SoffitLength / 12));
                    for (SoffitPanel in SoffitPanels) {
                        SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value;
                        SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value;
                        SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value;
                        SoffitPanel.clsType = "Panel";
                    }

                }

                // Remove Duplicate Soffit Panels
                DuplicateMaterialRemoval(SoffitPanels, "Panel");
                // Generation of Eave Overhang Soffit Panels
            }
            else if (((SoffitLocation.IndexOf("EaveOverhang", 0) + 1)
                        != 0)) {
                SoffitPanel = new clsPanel();
                // make soffit panel collection
                if ((SoffitLocation == "s2_EaveOverhang")) {
                    SoffitPanel.PanelMeasurement = ImperialMeasurementFormat(b.s2Overhang);
                }
                else if ((SoffitLocation == "s4_EaveOverhang")) {
                    SoffitPanel.PanelMeasurement = ImperialMeasurementFormat(b.s4Overhang);
                }

                SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value;
                SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value;
                SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value;
                SoffitPanel.Quantity = SoffitQty;
                SoffitPanel.clsType = "Panel";
                SoffitPanels.Add;
                SoffitPanel;
                // Generation of Eave Extension Soffit Panels
            }
            else if (((SoffitLocation.IndexOf("EaveExtension", 0) + 1)
                        != 0)) {
                RoofPanelGen(SoffitPanels, EaveExtRafterLength, 0, EaveExtBuildingLength);
                // update panel parameters
                for (SoffitPanel in SoffitPanels) {
                    SoffitPanel.PanelShape = EstSht.Range(NamedRangeString).offset(0, 1).Value;
                    SoffitPanel.PanelType = EstSht.Range(NamedRangeString).offset(0, 2).Value;
                    SoffitPanel.PanelColor = EstSht.Range(NamedRangeString).offset(0, 3).Value;
                    SoffitPanel.clsType = "Panel";
                }

            }

            // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Soffit Trim
            // '''''''''''''''''''''''''''''''' Gable Overhangs/Extensions
            // Jamb Trim and 2x6 Outside Angle
            if (((SoffitLocation.IndexOf("Gable", 0) + 1)
                        != 0)) {
                // calculate net rafter length for one endwall
                if ((b.rShape == "Gable")) {
                    NetRafterLength = (b.RafterLength * 2);
                }
                else if ((b.rShape == "Single Slope")) {
                    NetRafterLength = b.RafterLength;
                }

                // generate Jamb Trim
                TrimPieceCalc(SoffitTrim, NetRafterLength, "Jamb", ,, b);
                // Generate 2x6 outside angle trim
                if ((b.rShape == "Single Slope")) {
                    NetOutsideAngleLength = (b.s2RafterSheetLength
                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength));
                    TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", ,, b);
                }
                else if ((b.rShape == "Gable")) {
                    NetOutsideAngleLength = (b.s2RafterSheetLength
                                + (b.s4RafterSheetLength
                                + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength)));
                    TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", ,, b);
                }

                // Update Trim Information
                for (TrimPiece in SoffitTrim) {
                    // With...
                    if ((TrimPiece.tType == "Jamb Trim")) {
                        TrimPiece.tShape = b.wPanelShape;
                    }

                    if ((TrimPiece.tType == "2x6 Outside Angle")) {
                        TrimPiece.tShape = "N/A";
                    }

                    TrimPiece;
                    // '''''''''''''''''''''''''''''''' Eave Overhangs/Extensions
                    // Head Trim and 2x6 Outside Angle
                    ((SoffitLocation.IndexOf("Eave", 0) + 1)
                                != 0);
                    // Head trim for one sidewall
                    TrimPieceCalc(SoffitTrim, (TrimPiece.bLength * 12), "Head");
                    // ''''Generate 2x6 outside angle trim
                    // Eave Overhang (where 2x6 outside angle def covers the gable extensions/overhangs)
                    if (((SoffitLocation.IndexOf("Overhang", 0) + 1)
                                != 0)) {
                        NetOutsideAngleLength = ((TrimPiece.bLength * 12)
                                    + (TrimPiece.e1Overhang
                                    + (TrimPiece.e1Extension
                                    + (TrimPiece.e3Overhang + TrimPiece.e3Extension))));
                        // Eave Extension (where 2x6 outside angle may not cover the gable extensions/overhangs
                    }
                    else if (((SoffitLocation.IndexOf("Extension", 0) + 1)
                                != 0)) {
                        NetOutsideAngleLength = (EaveExtBuildingLength * 12);
                    }

                    TrimPieceCalc(SoffitTrim, NetOutsideAngleLength, "Outside Angle", ,, b);
                    // Update Trim Information
                    for (TrimPiece in SoffitTrim) {
                        // With...
                        if ((TrimPiece.tType == "Head Trim W/O Kickout")) {
                            TrimPiece.tShape = b.wPanelShape;
                        }

                        if ((TrimPiece.tType == "2x6 Outside Angle")) {
                            TrimPiece.tShape = "N/A";
                            if (((SoffitLocation.IndexOf("Extension", 0) + 1)
                                        != 0)) {
                                if ((SoffitLocation == "s2_EaveExtension")) {
                                    TrimPiece.tType = (TrimPiece.tType + (" "
                                                + (b.s2ExtensionPitch + ":12")));
                                }

                                if ((SoffitLocation == "s4_EaveExtension")) {
                                    TrimPiece.tType = (TrimPiece.tType + (" "
                                                + (b.s4ExtensionPitch + ":12")));
                                }
                                else if (((SoffitLocation.IndexOf("Overhang", 0) + 1)
                                            != 0)) {
                                    TrimPiece.tType = (TrimPiece.tType + (" "
                                                + (b.rPitch + ":12")));
                                }

                            }

                        }

                        TrimPiece;
                        // With...
                        ExtensionPanelGen((<Collection>(ExtensionPanels)), (<clsBuilding>(b)), (<string>(ExtensionLocation)), Optional, (<Collection>(s2RoofPanels)), Optional, (<Collection>(s4RoofPanels)));
                        let RoofPanel: clsPanel;
                        let ExtensionPanel: clsPanel;
                        let NamedRangeString: string;
                        // Var for reading correct extension panel info cell
                        let PanelQty: number;
                        let EaveExtBuildingLength: number;
                        // length measured from endwall to endwall
                        let EaveExtRafterLength: number;
                        let ExtensionLengthOverage: number;
                        NamedRangeString = ExtensionLocation;
                        // With...
                        switch (ExtensionLocation) {
                            case "e1_GableExtension":
                                ExtensionLengthOverage = (b.e1Extension % (3 * 12));
                                if (((ExtensionLengthOverage > 0)
                                            && (b.bLengthRoofPanelOverage >= ExtensionLengthOverage))) {
                                    // use roof panel overage
                                    PanelQty = Application.WorksheetFunction.RoundUp((((b.e1Extension - ExtensionLengthOverage)
                                                    / 12)
                                                    / 3), 0);
                                    // update roof panel overage remaining
                                    b.bLengthRoofPanelOverage = (b.bLengthRoofPanelOverage - ExtensionLengthOverage);
                                }
                                else {
                                    PanelQty = Application.WorksheetFunction.RoundUp(((b.e1Extension / 12)
                                                    / 3), 0);
                                }

                                // set extension panel quantity
                                b.e1ExtensionPanelQty = PanelQty;
                                break;
                            case "e3_GableExtension":
                                ExtensionLengthOverage = (b.e3Extension % (3 * 12));
                                if (((ExtensionLengthOverage > 0)
                                            && (b.bLengthRoofPanelOverage >= ExtensionLengthOverage))) {
                                    // use roof panel overage
                                    PanelQty = Application.WorksheetFunction.RoundUp((((b.e3Extension - ExtensionLengthOverage)
                                                    / 12)
                                                    / 3), 0);
                                    // update roof panel overage remaining
                                    b.bLengthRoofPanelOverage = (b.bLengthRoofPanelOverage - ExtensionLengthOverage);
                                }
                                else {
                                    PanelQty = Application.WorksheetFunction.RoundUp(((b.e3Extension / 12)
                                                    / 3), 0);
                                }

                                b.e3ExtensionPanelQty = PanelQty;
                                break;
                            case "s2_EaveExtension":
                                EaveExtBuildingLength = (b.s2EaveExtensionBuildingLength / 12);
                                // eave ext rafter length
                                EaveExtRafterLength = b.s2ExtensionRafterLength;
                                break;
                            case "s4_EaveExtension":
                                EaveExtBuildingLength = (b.s4EaveExtensionBuildingLength / 12);
                                // eave ext rafter length
                                EaveExtRafterLength = b.s4ExtensionRafterLength;
                                break;
                        }

                        // '''''''''''''''''''''''''''''''''''''''''''''' For Gable Extensions '''''''''''''''''''''''''''''''''''''''''
                        if (((ExtensionLocation.IndexOf("GableExtension", 0) + 1)
                                    != 0)) {
                            // '' Corresponding Extension Panels for Each Sidewall 2 Roof Panel type
                            for (RoofPanel in s2RoofPanels) {
                                ExtensionPanel = new clsPanel();
                                ExtensionPanel.PanelMeasurement = RoofPanel.PanelMeasurement;
                                ExtensionPanel.PanelShape = b.rPanelShape;
                                ExtensionPanel.PanelType = b.rPanelType;
                                ExtensionPanel.PanelColor = b.rPanelColor;
                                ExtensionPanel.Quantity = PanelQty;
                                ExtensionPanel.clsType = "Panel";
                                ExtensionPanels.Add;
                                ExtensionPanel;
                            }

                            // For a Gable Roof, Corresponding Extension Panels for Each Sidewall 4 Roof Panel type
                            if ((b.rShape == "Gable")) {
                                for (RoofPanel in s4RoofPanels) {
                                    ExtensionPanel = new clsPanel();
                                    ExtensionPanel.PanelMeasurement = RoofPanel.PanelMeasurement;
                                    ExtensionPanel.PanelShape = b.rPanelShape;
                                    ExtensionPanel.PanelType = b.rPanelType;
                                    ExtensionPanel.PanelColor = b.rPanelColor;
                                    ExtensionPanel.Quantity = PanelQty;
                                    ExtensionPanel.clsType = "Panel";
                                    ExtensionPanels.Add;
                                    ExtensionPanel;
                                }

                            }

                            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Eave Extension Panels
                        }
                        else if (((ExtensionLocation.IndexOf("EaveExtension", 0) + 1)
                                    != 0)) {
                            RoofPanelGen(ExtensionPanels, EaveExtRafterLength, 0, EaveExtBuildingLength, b.rShape, true);
                            // update panel parameters
                            for (ExtensionPanel in ExtensionPanels) {
                                ExtensionPanel.PanelShape = b.rPanelShape;
                                ExtensionPanel.PanelType = b.rPanelType;
                                ExtensionPanel.PanelColor = b.rPanelColor;
                                ExtensionPanel.clsType = "Panel";
                            }

                        }

                        // remove duplicate panels
                        DuplicateMaterialRemoval(ExtensionPanels, "Panel");
                        (<number>(PanelOptionCompare((<number>(IdealPLength)), (<number>(PanelCount)), (<number>(CurrentTotalLength)))));
                        // '' Function to determine whether or not to round the roof panel up or down to the nearest purlin to keep the total closest to the ideal
                        if ((Abs(((IdealPLength * PanelCount)
                                        - (CurrentTotalLength + ClosestRoofPurlin(IdealPLength, 1)))) < Abs(((IdealPLength * PanelCount)
                                        - (CurrentTotalLength + ClosestRoofPurlin(IdealPLength, -1)))))) {
                            PanelOptionCompare = 1;
                        }
                        else {
                            PanelOptionCompare = -1;
                        }

                    }

                }

            }

        }

    }

}
private RoofScrewGen(TekQty: number, LapQty: number, b: clsBuilding, rOverlaps: number) {
    let xLapSpaces: number;
    let yLapSpaces: number;
    let rPurlins: number;
    let s2ExtensionPurlins: number;
    let s4ExtensionPurlins: number;
    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Roof Screws ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    // With...
    rPurlins = Application.WorksheetFunction.RoundUp(((b.RafterLength / 12)
                    / 5), 0);
    // double purlins for gable roof
    if ((b.rShape == "Gable")) {
        rPurlins = (rPurlins * 2);
    }

    // ''add in overhang/extensions
    // sidewall 2 roof purlins
    if ((b.s2Extension == 0)) {
        // add in one purlin for an additional overhang
        if ((b.s2Overhang > 4.25)) {
            rPurlins = (rPurlins + 1);
        }
        else if ((b.s2Extension != 0)) {
            s2ExtensionPurlins = Application.WorksheetFunction.RoundUp(((b.s2ExtensionRafterLength / 12)
                            / 5), 0);
            rPurlins = (rPurlins + s2ExtensionPurlins);
        }

        // sidewall 4 roof purlins
        if ((b.s4Extension == 0)) {
            // add in one purlin for an eave extension
            if ((b.s4Overhang > 4.25)) {
                rPurlins = (rPurlins + 1);
            }
            else if ((b.s4Extension != 0)) {
                s4ExtensionPurlins = Application.WorksheetFunction.RoundUp(((b.s4ExtensionRafterLength / 12)
                                / 5), 0);
                rPurlins = (rPurlins + s4ExtensionPurlins);
            }

            // '''calculate tek screw quantity
            // one for every purlin per foot length, top and bottom putlin get 2 per ft length, and overlaps get two per ft length
            if ((b.rShape == "Single Slope")) {
                TekQty = ((rPurlins * b.RoofFtLength)
                            + ((2 * b.RoofFtLength)
                            + (rOverlaps * b.RoofFtLength)));
            }
            else if ((b.rShape == "Gable")) {
                // additional top and bottom purlin for s4
                TekQty = ((rPurlins * b.RoofFtLength)
                            + ((4 * b.RoofFtLength)
                            + (rOverlaps * b.RoofFtLength)));
            }

            // exclude intersections
            // for sidewall 2 intersections
            if ((b.s2Extension > 0)) {
                if (((b.e1Extension > 0)
                            && (b.s2e1ExtensionIntersection == false))) {
                    TekQty = (TekQty
                                - (s2ExtensionPurlins
                                * (b.e1Extension / 12)));
                }

                if (((b.e3Extension > 0)
                            && (b.s2e3ExtensionIntersection == false))) {
                    TekQty = (TekQty
                                - (s2ExtensionPurlins
                                * (b.e3Extension / 12)));
                }

                // for sidewall 4 intersections
                if ((b.s4Extension > 0)) {
                    if (((b.e1Extension > 0)
                                && (b.s4e1ExtensionIntersection == false))) {
                        TekQty = (TekQty
                                    - (s4ExtensionPurlins
                                    * (b.e1Extension / 12)));
                    }

                    if (((b.e3Extension > 0)
                                && (b.s4e3ExtensionIntersection == false))) {
                        TekQty = (TekQty
                                    - (s4ExtensionPurlins
                                    * (b.e3Extension / 12)));
                    }

                    // round up tek screw to the nearest 250
                    TekQty = (Application.WorksheetFunction.RoundUp((TekQty / 250), 0) * 250);
                    // ''''calculate lap screw quantity
                    // roof length spacses
                    yLapSpaces = (Application.WorksheetFunction.RoundUp(((b.RoofLength / 12)
                                    / 3), 0) + 1);
                    if ((b.rShape == "Single Slope")) {
                        // rafter length /3'
                        xLapSpaces = (Application.WorksheetFunction.RoundUp(((b.s2RafterSheetLength
                                        + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength))
                                        / 30), 0) + 1);
                        // calculate lap qty
                        LapQty = (xLapSpaces * yLapSpaces);
                    }
                    else if ((b.rShape == "Gable")) {
                        // sidewall 2
                        xLapSpaces = (Application.WorksheetFunction.RoundUp(((b.s2RafterSheetLength + b.s2ExtensionRafterLength)
                                        / 30), 0) + 1);
                        LapQty = (xLapSpaces * yLapSpaces);
                        xLapSpaces = (Application.WorksheetFunction.RoundUp(((b.s4RafterSheetLength + b.s4ExtensionRafterLength)
                                        / 30), 0) + 1);
                        LapQty = (LapQty
                                    + (xLapSpaces * yLapSpaces));
                    }

                    // '''''''''''''''''''''''''' Note: Lap screws NOT** Currently reduced for excluded extension intersections
                    // increase if gutters
                    if ((b.Gutters == true)) {
                        // add additional screws for gutters along s2
                        LapQty = (LapQty + (2 * b.RoofFtLength));
                        if ((b.rShape == "Gable")) {
                            LapQty = (LapQty + (2 * b.RoofFtLength));
                        }

                        // round up
                        LapQty = (Application.WorksheetFunction.RoundUp((LapQty / 250), 0) * 250);
                    }

                }

                WallScrewGen((<number>(TekQty)), (<number>(LapQty)), (<clsBuilding>(b)), (<number>(sOverlaps)), (<number>(eOverlaps)));
                let sPurlins: number;
                let ePurlins: number;
                let sTekScrews: number;
                let eTekScrews: number;
                let RemainingHeight: number;
                let xLapSpaces: number;
                let yLapSpaces: number;
                let e1HeightIncriment: number;
                let e3HeightIncriment: number;
                let s2RemainingHeight: number;
                let s4RemainingHeight: number;
                let e1Purlins: number;
                let e3Purlins: number;
                let s2Purlins: number;
                let s4Purlins: number;
                let e1MaxLength: number;
                let e3MaxLength: number;
                let s2MaxLength: number;
                let s4MaxLength: number;
                let e1GablePurlinTotalLength: number;
                let e3GablePurlinTotalLength: number;
                let e1WallPurlinTotalLength: number;
                let e3WallPurlinTotalLength: number;
                let TotalGableTekScrews: number;
                let InteriorRoofAngle: Object;
                let WallHeightCounter: number;
                // With...
                // calculate interior roof angle
                InteriorRoofAngle = WorksheetFunction.Asin((b.rPitch / Sqr((b.rPitch
                                    | ((2 + 12)
                                    | 2)))));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // ''''''''''''''''''''''''''''''' Find Purlin Count
                // ''calculate endwall purlins
                // peak height
                if ((b.rShape == "Single Slope")) {
                    switch (b.WallStatus("e1")) {
                        case "Exclude":
                            e1MaxLength = 0;
                            break;
                        case "Include":
                        case "Partial":
                            e1MaxLength = ((b.bHeight - b.LengthAboveFinishedFloor("e1"))
                                        + ((b.rPitch * b.bWidth)
                                        / 12));
                            break;
                        case "Gable Only":
                            e1MaxLength = ((b.rPitch * b.bWidth)
                                        / 12);
                            break;
                    }

                    switch (b.WallStatus("e3")) {
                        case "Exclude":
                            e3MaxLength = 0;
                            break;
                        case "Include":
                        case "Partial":
                            e3MaxLength = ((b.bHeight - b.LengthAboveFinishedFloor("e3"))
                                        + ((b.rPitch * b.bWidth)
                                        / 12));
                            break;
                        case "Gable Only":
                            e3MaxLength = ((b.rPitch * b.bWidth)
                                        / 12);
                            break;
                    }

                }
                else if ((b.rShape == "Gable")) {
                    switch (b.WallStatus("e1")) {
                        case "Exclude":
                            e1MaxLength = 0;
                            break;
                        case "Include":
                        case "Partial":
                            e1MaxLength = ((b.bHeight - b.LengthAboveFinishedFloor("e1"))
                                        + ((b.rPitch
                                        * (b.bWidth / 2))
                                        / 12));
                            break;
                        case "Gable Only":
                            e1MaxLength = ((b.rPitch
                                        * (b.bWidth / 2))
                                        / 12);
                            break;
                    }

                    switch (b.WallStatus("e3")) {
                        case "Exclude":
                            e3MaxLength = 0;
                            break;
                        case "Include":
                        case "Partial":
                            e3MaxLength = ((b.bHeight - b.LengthAboveFinishedFloor("e3"))
                                        + ((b.rPitch
                                        * (b.bWidth / 2))
                                        / 12));
                            break;
                        case "Gable Only":
                            e3MaxLength = ((b.rPitch
                                        * (b.bWidth / 2))
                                        / 12);
                            break;
                    }

                }

                // account for bottom purlins
                if ((b.WallStatus("e1") == "Include")) {
                    e1HeightIncriment = (7 + (2 / 12));
                }
                else if ((b.WallStatus("e1") != "Exclude")) {
                    e1HeightIncriment = 5;
                }

                if ((b.WallStatus("e3") == "Include")) {
                    e3HeightIncriment = (7 + (2 / 12));
                }
                else if ((b.WallStatus("e3") != "Exclude")) {
                    e3HeightIncriment = 5;
                }

                if (b.WallStatus) {
                    ("e1" != "Exclude");
                    for (
                    ; (e1HeightIncriment < e1MaxLength);
                    ) {
                        e1Purlins = (e1Purlins + 1);
                        switch (b.WallStatus("e1")) {
                            case "Include":
                            case "Partial":
                                if ((e1HeightIncriment
                                            > (b.bHeight - b.LengthAboveFinishedFloor("e1")))) {
                                    e1GablePurlinTotalLength = (e1GablePurlinTotalLength
                                                + ((e1MaxLength - e1HeightIncriment)
                                                / Tan(InteriorRoofAngle)));
                                }
                                else {
                                    e1WallPurlinTotalLength = (e1WallPurlinTotalLength + b.bWidth);
                                }

                                break;
                            case "Gable":
                                e1GablePurlinTotalLength = (e1GablePurlinTotalLength
                                            + ((e1MaxLength - e1HeightIncriment)
                                            / Tan(InteriorRoofAngle)));
                                break;
                        }

                        e1HeightIncriment = (e1HeightIncriment + 5);
                    }

                }

                if (b.WallStatus) {
                    ("e3" != "Exclude");
                    for (
                    ; (e3HeightIncriment < e3MaxLength);
                    ) {
                        e3Purlins = (e3Purlins + 1);
                        switch (b.WallStatus("e3")) {
                            case "Include":
                            case "Partial":
                                if ((e3HeightIncriment
                                            > (b.bHeight - b.LengthAboveFinishedFloor("e3")))) {
                                    e3GablePurlinTotalLength = (e3GablePurlinTotalLength
                                                + ((e3MaxLength - e3HeightIncriment)
                                                / Tan(InteriorRoofAngle)));
                                }
                                else {
                                    e3WallPurlinTotalLength = (e3WallPurlinTotalLength + b.bWidth);
                                }

                                break;
                            case "Gable":
                                e3GablePurlinTotalLength = (e3GablePurlinTotalLength
                                            + ((e3MaxLength - e3HeightIncriment)
                                            / Tan(InteriorRoofAngle)));
                                break;
                        }

                        e3HeightIncriment = (e3HeightIncriment + 5);
                    }

                }

                if ((b.rShape == "Gable")) {
                    e3GablePurlinTotalLength = (e3GablePurlinTotalLength * 2);
                    e1GablePurlinTotalLength = (e1GablePurlinTotalLength * 2);
                }

                if (b.WallStatus) {
                    "e1" = "Gable Only";
                    e1GablePurlinTotalLength = (e1GablePurlinTotalLength + b.bWidth);
                }
                else if (b.WallStatus) {
                    ("e1" != "Exclude");
                    e1WallPurlinTotalLength = (e1WallPurlinTotalLength + b.bWidth);
                }

                if (b.WallStatus) {
                    "e3" = "Gable Only";
                    e3GablePurlinTotalLength = (e3GablePurlinTotalLength + b.bWidth);
                }
                else if (b.WallStatus) {
                    ("e1" != "Exclude");
                    e1WallPurlinTotalLength = (e1WallPurlinTotalLength + b.bWidth);
                }

                // purlin length of top not accounted for?
                // ''''''''''''calculate sidewall purlins
                switch (b.WallStatus("s2")) {
                    case "Exclude":
                        s2MaxLength = 0;
                        break;
                    case "Include":
                    case "Partial":
                        s2MaxLength = (b.bHeight - b.LengthAboveFinishedFloor("s2"));
                        break;
                }

                switch (b.WallStatus("s4")) {
                    case "Exclude":
                        s4MaxLength = 0;
                        break;
                    case "Include":
                    case "Partial":
                        if ((b.rShape == "Gable")) {
                            s4MaxLength = (b.bHeight - b.LengthAboveFinishedFloor("s4"));
                        }
                        else if ((b.rShape == "Single Slope")) {
                            s4MaxLength = ((b.HighSideEaveHeight / 12)
                                        - b.LengthAboveFinishedFloor("s4"));
                        }

                        break;
                }

                // account for bottom purlins
                if ((b.WallStatus("s2") == "Include")) {
                    s2RemainingHeight = (s2MaxLength - (7 - (2 / 12)));
                }
                else if ((b.WallStatus("s2") != "Exclude")) {
                    s2RemainingHeight = (s2MaxLength - 5);
                }

                if ((b.WallStatus("s4") == "Include")) {
                    s4RemainingHeight = (s4MaxLength - (7 - (2 / 12)));
                }
                else if ((b.WallStatus("s4") != "Exclude")) {
                    s4RemainingHeight = (s4MaxLength - 5);
                }

                sPurlins = 1;
                // find purlins above
                while ((s2RemainingHeight >= 5)) {
                    s2Purlins = (s2Purlins + 1);
                    s2RemainingHeight = (s2RemainingHeight - 5);
                }

                while ((s4RemainingHeight >= 5)) {
                    s4Purlins = (s4Purlins + 1);
                    s4RemainingHeight = (s4RemainingHeight - 5);
                }

                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Tek screws
                //  one per ft per purlin, two for top and bottom purlin
                sTekScrews = ((s2Purlins * b.bLength)
                            + ((2 * b.bLength)
                            + ((s4Purlins * b.bLength)
                            + ((2 * b.bLength)
                            + (sOverlaps * b.bLength)))));
                eTekScrews = (e1WallPurlinTotalLength
                            + (e1GablePurlinTotalLength
                            + (e3WallPurlinTotalLength
                            + (e3GablePurlinTotalLength
                            + (eOverlaps * b.bWidth)))));
                TekQty = (Application.WorksheetFunction.RoundUp(((sTekScrews + eTekScrews)
                                / 250), 0) * 250);
                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Lap Screws
                // 'sidewall 2
                xLapSpaces = (Application.WorksheetFunction.RoundUp((b.bLength / 3), 0) + 1);
                yLapSpaces = (Application.WorksheetFunction.RoundUp(((s2MaxLength * 12)
                                / 30), 0) + 1);
                LapQty = (xLapSpaces * yLapSpaces);
                yLapSpaces = (Application.WorksheetFunction.RoundUp(((s4MaxLength * 12)
                                / 30), 0) + 1);
                LapQty = (LapQty
                            + (xLapSpaces * yLapSpaces));
                xLapSpaces = (Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0) + 1);
                yLapSpaces = (Application.WorksheetFunction.RoundUp(((e1MaxLength * 12)
                                / 30), 0) + 1);
                LapQty = (xLapSpaces * yLapSpaces);
                yLapSpaces = (Application.WorksheetFunction.RoundUp(((e3MaxLength * 12)
                                / 30), 0) + 1);
                LapQty = (LapQty
                            + (xLapSpaces + yLapSpaces));
                LapQty = (Application.WorksheetFunction.RoundUp((LapQty / 250), 0) * 250);
            }

            TrimScrewCalc((<Collection>(TrimScrews)), (<Collection>(RakeTrimPieces)), (<clsBuilding>(b)));
            let NetCornerLength: number;
            let NetRakeTrimLength: number;
            let Screw: clsFastener;
            let TrimPiece: clsTrim;
            // calculate net rake trim length
            for (TrimPiece in RakeTrimPieces) {
                NetRakeTrimLength = (NetRakeTrimLength
                            + (TrimPiece.tLength * TrimPiece.Quantity));
            }

            // calculate corner trim length
            if ((b.rShape == "Gable")) {
                // assume complete, exclude intersections if needed
                NetCornerLength = (b.bHeight * (4 * 12));
                if (((b.WallStatus("e1") != "Include")
                            && (b.WallStatus("s2") != "Include"))) {
                    NetCornerLength = (NetCornerLength
                                - (b.bHeight * 12));
                }

                if (((b.WallStatus("e1") != "Include")
                            && (b.WallStatus("s4") != "Include"))) {
                    NetCornerLength = (NetCornerLength
                                - (b.bHeight * 12));
                }

                if (((b.WallStatus("e3") != "Include")
                            && (b.WallStatus("s2") != "Include"))) {
                    NetCornerLength = (NetCornerLength
                                - (b.bHeight * 12));
                }

                if (((b.WallStatus("e3") != "Include")
                            && (b.WallStatus("s4") != "Include"))) {
                    NetCornerLength = (NetCornerLength
                                - (b.bHeight * 12));
                }
                else if ((b.rShape == "Single Slope")) {
                    // sidewall 2 corners + s4 corners
                    NetCornerLength = ((b.bHeight * (12 * 2))
                                + (b.HighSideEaveHeight * 2));
                    if (((b.WallStatus("s2") != "Include")
                                && (b.WallStatus("e1") != "Include"))) {
                        NetCornerLength = (NetCornerLength
                                    - (b.bHeight * 12));
                    }

                    if (((b.WallStatus("s2") != "Include")
                                && (b.WallStatus("e3") != "Include"))) {
                        NetCornerLength = (NetCornerLength
                                    - (b.bHeight * 12));
                    }

                    if (((b.WallStatus("s4") != "Include")
                                && (b.WallStatus("e1") != "Include"))) {
                        NetCornerLength = (NetCornerLength - b.HighSideEaveHeight);
                    }

                    if (((b.WallStatus("s4") != "Include")
                                && (b.WallStatus("e3") != "Include"))) {
                        NetCornerLength = (NetCornerLength - b.HighSideEaveHeight);
                    }

                    // '' Add to Screws Collection
                    //  If screws the same color, combine
                    if ((b.OutsideCornerTrimColor == b.RakeTrimColor)) {
                        Screw = new clsFastener();
                        Screw.Quantity = (Application.WorksheetFunction.RoundUp(((((NetCornerLength / 30)
                                        * 2)
                                        + ((NetRakeTrimLength / 12)
                                        + (NetRakeTrimLength / 30)))
                                        / 250), 0) * 250);
                        Screw.Color = b.OutsideCornerTrimColor;
                        TrimScrews.Add;
                        Screw;
                        // add seperately if different colors
                    }
                    else {
                        // rake trim screws
                        Screw = new clsFastener();
                        Screw.Quantity = (Application.WorksheetFunction.RoundUp((((NetRakeTrimLength / 12)
                                        + (NetRakeTrimLength / 30))
                                        / 250), 0) * 250);
                        Screw.Color = b.RakeTrimColor;
                        TrimScrews.Add;
                        Screw;
                        // outside corner trim screws
                        if ((NetCornerLength != 0)) {
                            Screw = new clsFastener();
                            Screw.Quantity = (Application.WorksheetFunction.RoundUp((((NetCornerLength / 30)
                                            * 2)
                                            / 250), 0) * 250);
                            Screw.Color = b.OutsideCornerTrimColor;
                            TrimScrews.Add;
                            Screw;
                        }

                    }

                }

                SoffitScrewCalc((<number>(ScrewQty)), (<string>(SoffitScrewColor)), (<string>(SoffitType)), (<clsBuilding>(b)));
                let Location: string;
                // Wall Location
                let pLines: number;
                // Purlin Lines
                // determine wall location
                switch (true) {
                    case ((SoffitType.IndexOf("e1", 0) + 1)
                                != 0):
                        Location = "e1";
                        break;
                    case ((SoffitType.IndexOf("s2", 0) + 1)
                                != 0):
                        Location = "s2";
                        break;
                    case ((SoffitType.IndexOf("e3", 0) + 1)
                                != 0):
                        Location = "e3";
                        break;
                    case ((SoffitType.IndexOf("s4", 0) + 1)
                                != 0):
                        Location = "s4";
                        break;
                }

                // determine soffit screw color (only possibility since soffits of different color wouldn't be on the same building)
                SoffitScrewColor = EstSht.Range(SoffitType).offset(0, 4).Value;
                // With...
                // find roof purlin lines if needed
                if (((Location == "e1")
                            || (Location == "e3"))) {
                    pLines = Application.WorksheetFunction.RoundUp(((b.RafterLength / 12)
                                    / 5), 0);
                    // double purlins for gable roof
                    if ((b.rShape == "Gable")) {
                        pLines = (pLines * 2);
                    }

                    // ''add in overhang/extensions
                    // sidewall 2 roof purlins
                    if ((b.s2Extension == 0)) {
                        // add in one purlin for an additional eave overhang
                        if ((b.s2Overhang > 4.25)) {
                            pLines = (pLines + 1);
                        }
                        else if ((b.s2Extension != 0)) {
                            pLines = (pLines + Application.WorksheetFunction.RoundUp(((b.s2ExtensionRafterLength / 12)
                                            / 5), 0));
                        }

                        // sidewall 4 roof purlins
                        if ((b.s4Extension == 0)) {
                            // add in one purlin for an eave extension
                            if ((b.s4Overhang > 4.25)) {
                                pLines = (pLines + 1);
                            }
                            else if ((b.s4Extension != 0)) {
                                pLines = (pLines + Application.WorksheetFunction.RoundUp(((b.s4ExtensionRafterLength / 12)
                                                / 5), 0));
                            }

                        }

                        // determine extension/overhang type, calculate screws
                        switch (true) {
                            case ((SoffitType.IndexOf("EaveOverhang", 0) + 1)
                                        != 0):
                                // just 2/ft along building length
                                ScrewQty = (ScrewQty
                                            + (b.bLength * 2));
                                break;
                            case ((SoffitType.IndexOf("EaveExtension", 0) + 1)
                                        != 0):
                                // calculate extension purlin lines
                                if ((Location == "s2")) {
                                    pLines = Application.WorksheetFunction.RoundUp(((b.s2ExtensionRafterLength / 12)
                                                    / 5), 0);
                                }
                                else if ((Location == "s4")) {
                                    pLines = Application.WorksheetFunction.RoundUp(((b.s4ExtensionRafterLength / 12)
                                                    / 5), 0);
                                }

                                // screw/ft along purlin lines
                                ScrewQty = (ScrewQty
                                            + (b.bLength * pLines));
                                break;
                            case ((SoffitType.IndexOf("GableOverhang", 0) + 1)
                                        != 0):
                                if ((Location == "e1")) {
                                    ScrewQty = (ScrewQty
                                                + ((b.e1Overhang / 12)
                                                * pLines));
                                }
                                else if ((Location == "e3")) {
                                    ScrewQty = (ScrewQty
                                                + ((b.e3Overhang / 12)
                                                * pLines));
                                }

                                break;
                            case ((SoffitType.IndexOf("GableExtension", 0) + 1)
                                        != 0):
                                if ((Location == "e1")) {
                                    ScrewQty = (ScrewQty
                                                + ((b.e1Extension / 12)
                                                * pLines));
                                }
                                else if ((Location == "e3")) {
                                    ScrewQty = (ScrewQty
                                                + ((b.e3Extension / 12)
                                                * pLines));
                                }

                                break;
                        }

                    }

                }

                MatListSectionWrite((<Worksheet>(OutputSht)), (<Range>(WriteCell)), (<Collection>(MatCollection)), (<string>(CollectionType)));
                let Panel: clsPanel;
                let TrimPiece: clsTrim;
                let item: clsMiscItem;
                let StartCell: Range;
                // save start cell
                StartCell = WriteCell;
                switch (CollectionType) {
                    case "Panel":
                        for (Panel in MatCollection) {
                            if ((WriteCell != StartCell)) {
                                OutputSht.Rows[(WriteCell.Row + 1)].Insert;
                            }

                            WriteCell.Value = Panel.Quantity;
                            WriteCell.offset(0, 1).Value = Panel.PanelShape;
                            WriteCell.offset(0, 2).Value = Panel.PanelType;
                            WriteCell.offset(0, 3).Value = Panel.PanelMeasurement;
                            WriteCell.offset(0, 4).Value = Panel.PanelColor;
                            WriteCell = WriteCell.offset(1, 0);
                        }

                        break;
                    case "Trim":
                        for (TrimPiece in MatCollection) {
                            if ((WriteCell != StartCell)) {
                                OutputSht.Rows[(WriteCell.Row + 1)].Insert;
                            }

                            WriteCell.Value = TrimPiece.Quantity;
                            WriteCell.offset(0, 1).Value = TrimPiece.tShape;
                            WriteCell.offset(0, 2).Value = TrimPiece.tType;
                            WriteCell.offset(0, 3).Value = TrimPiece.tMeasurement;
                            WriteCell.offset(0, 4).Value = TrimPiece.Color;
                            WriteCell = WriteCell.offset(1, 0);
                        }

                        break;
                }

            }

        }

    }

}
private MiscMaterialCalc(ButylTapeQty: number, InsideClosureQty: number, OutsideClosureQty: number, b: clsBuilding, rOverlaps: number) {
    // With...
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Butyl Tape
    ButylTapeQty = (Application.WorksheetFunction.RoundUp((b.RoofFtLength / 3), 0) + 1);
    if ((b.rShape == "Single Slope")) {
        // one total rafter length for s2
        ButylTapeQty = (ButylTapeQty
                    * ((b.s2RafterSheetLength
                    + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength))
                    / 12));
    }
    else if ((b.rShape == "Gable")) {
        // rafter length along s2 and s4
        ButylTapeQty = (ButylTapeQty
                    * (((b.s2RafterSheetLength + b.s2ExtensionRafterLength)
                    / 12)
                    + ((b.s4RafterSheetLength + b.s4ExtensionRafterLength)
                    / 12)));
        ButylTapeQty = (ButylTapeQty
                    + (b.bLength * 2));
    }

    // add tape for overlaps
    ButylTapeQty = (ButylTapeQty
                + (rOverlaps * b.bLength));
    ButylTapeQty = Application.WorksheetFunction.RoundUp(((ButylTapeQty * 1.05)
                    / 44), 0);
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Inside Closures & Outside Closures
    if ((b.rShape == "Single Slope")) {
        if (b.WallStatus) {
            "s2" = "Include";
            InsideClosureQty = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
            OutsideClosureQty = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
        }

        if (b.WallStatus) {
            "e1" = "Include";
            OutsideClosureQty = (OutsideClosureQty + Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0));
            if (b.WallStatus) {
                "e3" = "Include";
                OutsideClosureQty = (OutsideClosureQty + Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0));
            }
            else if ((b.rShape == "Gable")) {
                if (b.WallStatus) {
                    "s2" = "Include";
                    InsideClosureQty = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
                    if (b.WallStatus) {
                        "s4" = "Include";
                        InsideClosureQty = (InsideClosureQty + Application.WorksheetFunction.RoundUp((b.bLength / 3), 0));
                        if (b.WallStatus) {
                            "e1" = "Include";
                            OutsideClosureQty = (OutsideClosureQty + Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0));
                            if (b.WallStatus) {
                                "e3" = "Include";
                                OutsideClosureQty = (OutsideClosureQty + Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0));
                            }

                        }

                    }

                }

            }

        }

    }

}
// ''''''''''''''''''''' Sidewall Panel Generation
SidewallPanelGen(SidewallPanels: Collection, sWall: string, b: clsBuilding, FullHeightLinerPanels: boolean) {
    let Panel: clsPanel;
    // Warning!!! Optional parameters not supported
    let sP1: clsPanel;
    let sP2: clsPanel;
    let sP3: clsPanel;
    let WainscotPanel: clsPanel;
    let p1Length: number;
    let p2Length: number;
    let p3Length: number;
    let SpecialBottomPurlin: boolean;
    let WainscotFtLength: number;
    let FO: clsFO;
    let FOCutoutp1: clsPanel;
    let FOCollection: Collection;
    // With...
    // Check for Wainscot
    if (b.Wainscot) {
        (sWall != "None");
        WainscotPanel = new clsPanel();
        WainscotPanel.PanelLength = number.Parse(Left(b.Wainscot, sWall, 2));
        if ((FullHeightLinerPanels == false)) {
            WainscotFtLength = (WainscotPanel.PanelLength / 12);
        }

        // ''''''''''''''''''''''''''' Generate Sidewall Panels
        switch (b.WallStatus) {
            case sWall:
                break;
            case "Exclude":
                return;
                break;
            case "Include":
            case "Partial":
                if (b.WallStatus) {
                    sWall = "Partial";
                    SpecialBottomPurlin = true;
                    if (((b.rShape == "Single Slope")
                                && (sWall == "s4"))) {
                        // ''If high side eave is under 42 Feet
                        if ((((b.HighSideEaveHeight / 12)
                                    - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)) {
                            42;
                            sP1 = new clsPanel();
                            sP1.PanelLength = (b.HighSideEaveHeight
                                        - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12));
                            // FullHeightLinerPanels is only TRUE when this sub is being called specifically to gen full height liner panels
                            if ((FullHeightLinerPanels == true)) {
                                sP1.PanelLength = (sP1.PanelLength - 8);
                            }

                            // ''If high side eave is over 42 Feet and less than or equal to 84 feet
                            // '' Since the highest purlin under 42' is at 37' 3.5", max height for 2 is 37'3.5"+42'
                        }
                        else if ((((b.HighSideEaveHeight / 12)
                                    - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)) {
                            (42
                                        & (((b.HighSideEaveHeight / 12)
                                        - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength));
                            (79 + (3.5 / 12));
                            //  Panel #1
                            sP1 = new clsPanel();
                            sP1.PanelLength = (ClosestWallPurlin((b.HighSideEaveHeight
                                            - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12))) / 2);
                            0;
                            SpecialBottomPurlin;
                            // add overlap
                            sP1.PanelLength = (sP1.PanelLength + 1.5);
                            //  Panel #2
                            sP2 = new clsPanel();
                            sP2.PanelLength = (b.HighSideEaveHeight
                                        - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12));
                            (sP1.PanelLength * -1);
                            sP2.PanelLength = (sP2.PanelLength + 1.5);
                            if ((FullHeightLinerPanels == true)) {
                                sP2.PanelLength = (sP2.PanelLength - 8);
                            }

                            // '''''''''''''' if high side eave is over 79' 3.5"
                        }
                        else if ((((b.HighSideEaveHeight / 12)
                                    - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)) {
                            (79 + (3.5 / 12));
                            //  Panel #1
                            sP1 = new clsPanel();
                            sP1.PanelLength = (ClosestWallPurlin((b.HighSideEaveHeight
                                            - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12))) / 3);
                            0;
                            SpecialBottomPurlin;
                            // add overlap
                            sP1.PanelLength = (sP1.PanelLength + 1.5);
                            //  Panel #2
                            sP2 = new clsPanel();
                            sP2.PanelLength = ClosestWallPurlin((sP1.PanelLength
                                            + ((b.HighSideEaveHeight
                                            - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12))
                                            / 3)));
                            0;
                            SpecialBottomPurlin;
                            // two overlaps
                            sP2.PanelLength = (sP2.PanelLength + 3);
                            //  Panel #3
                            sP3 = new clsPanel();
                            sP3.PanelLength = (b.HighSideEaveHeight
                                        - ((WainscotFtLength + b.LengthAboveFinishedFloor)[sWall] * 12));
                            ((sP1.PanelLength - sP2.PanelLength)
                                        * -1);
                            // overlap
                            sP3.PanelLength = (sP3.PanelLength + 1.5);
                            if ((FullHeightLinerPanels == true)) {
                                sP3.PanelLength = (sP3.PanelLength - 8);
                            }

                            // '''''''''''''''''''normal handling for everything but s4 on a single slope
                        }
                        else {
                            // ''If building height is under 42 Feet
                            if (((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)) {
                                42;
                                sP1 = new clsPanel();
                                sP1.PanelLength = ((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength);
                                12;
                                if ((FullHeightLinerPanels == true)) {
                                    sP1.PanelLength = (sP1.PanelLength - 8);
                                }

                                // ''If building height is over 42 Feet
                            }
                            else if (((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)) {
                                42;
                                //  Panel #1
                                sP1 = new clsPanel();
                                if ((ClosestWallPurlin((((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)
                                                / 2)) * 12)) {
                                    0;
                                    SpecialBottomPurlin;
                                    (42 * 12);
                                    // if over 42, then find next closest below
                                    sP1.PanelLength = ClosestWallPurlin(((((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)
                                                    / 2)
                                                    * 12));
                                    -1;
                                    SpecialBottomPurlin;
                                }
                                else {
                                    // find closest sidewall purlin if divided in half
                                    sP1.PanelLength = (ClosestWallPurlin((((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)
                                                    / 2)) * 12);
                                    0;
                                    SpecialBottomPurlin;
                                }

                                // add overlap
                                sP1.PanelLength = (sP1.PanelLength + 1.5);
                                // Panel #2
                                sP2 = new clsPanel();
                                sP2.PanelLength = (((b.bHeight - b.LengthAboveFinishedFloor)[sWall] - WainscotFtLength)
                                            * 12);
                                (sP1.PanelLength * -1);
                                // overlap
                                sP2.PanelLength = (sP2.PanelLength + 1.5);
                                if ((FullHeightLinerPanels == true)) {
                                    sP2.PanelLength = (sP2.PanelLength - 8);
                                }

                            }

                        }

                    }

                    // ''''''' Add Quantities
                    if (!(sP1 == null)) {
                        sP1.Quantity = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
                    }

                    if (!(sP2 == null)) {
                        sP2.Quantity = sP1.Quantity;
                    }

                    if (!(sP3 == null)) {
                        sP3.Quantity = sP1.Quantity;
                    }

                    // '''Modify the sidewall panel collection (which is currently a solid wall) by removing panels that are covering a qualifying framed opening.
                    // ' Qualifying framed openings are: Overhead Doors greater than or equal to 7' in width
                    if (!(sP1 == null)) {
                        // set correct wall fo collection
                        if ((sWall == "s2")) {

                        }

                        FOCollection = b.s2FOs;
                    }
                    else {

                    }

                    FOCollection = b.s4FOs;
                    for (FO in FOCollection) {
                        if ((((FO.FOType == "OHDoor")
                                    || (FO.FOType == "MiscFO"))
                                    && ((FO.Width >= (7 * 12))
                                    && (FO.bEdgeHeight == 0)))) {
                            FOCutoutp1 = new clsPanel();
                            FOCutoutp1.Quantity = (Application.WorksheetFunction.RoundUp((FO.Width / (3 * 12)), 0) - 2);
                            // Calculate number of panels to cut short
                            FOCutoutp1.PanelLength = (sP1.PanelLength
                                        - (FO.Height
                                        - (WainscotFtLength * 12)));
                            sP1.Quantity = (sP1.Quantity - FOCutoutp1.Quantity);
                            // If there is more than 1 sidewall panel required for a well, check that only an overlap section isn't being added
                            if (!(sP2 == null)) {
                                if ((FOCutoutp1.PanelLength > 1.5)) {
                                    SidewallPanels.Add;
                                }

                                FOCutoutp1;
                            }
                            else {
                                SidewallPanels.Add;
                                FOCutoutp1;
                            }

                        }

                    }

                    // add modified sidewall panel 1 to sidewall panel collection
                    SidewallPanels.Add;
                    sP1;
                }

                if (!(sP2 == null)) {
                    SidewallPanels.Add;
                }

                sP2;
                if (!(sP3 == null)) {
                    SidewallPanels.Add;
                }

                sP3;
                // add parameters
                for (Panel in SidewallPanels) {
                    Panel.PanelMeasurement = ImperialMeasurementFormat(Panel.PanelLength);
                    Panel.PanelShape = b.wPanelShape;
                    Panel.PanelType = b.wPanelType;
                    Panel.PanelColor = b.wPanelColor;
                }

                // wainscot
                if ((!(WainscotPanel == null)
                            && (FullHeightLinerPanels == false))) {
                    WainscotPanel.PanelMeasurement = ImperialMeasurementFormat(WainscotPanel.PanelLength);
                    WainscotPanel.Quantity = sP1.Quantity;
                    // only add quantity of full length sidewall panels
                    WainscotPanel.PanelColor = EstSht.Range((sWall + "_Wainscot")).offset(0, 2).Value;
                    WainscotPanel.PanelType = EstSht.Range((sWall + "_Wainscot")).offset(0, 1).Value;
                    WainscotPanel.PanelShape = b.wPanelShape;
                    SidewallPanels.Add;
                    WainscotPanel;
                }

                break;
        }

    }

}
private LinerPanelGen(LinerPanels: Collection, b: clsBuilding, Location: string) {
    let LinerPanel: clsPanel;
    // ''' Full Height Panels
    if ((b.LinerPanels(Location) == "Full Height")) {
        // roof liner panels
        if ((Location == "Roof")) {
            RoofPanelGen(LinerPanels, (b.RafterLength - 8), 0, b.bLength, b.rShape);
            if ((b.rShape == "Gable")) {
                for (LinerPanel in LinerPanels) {
                    LinerPanel.Quantity = (LinerPanel.Quantity * 2);
                }

            }

            // wall liner panels
        }
        else {
            switch (Location) {
                case "e1":
                case "e3":
                    EndwallPanelGen(LinerPanels, Location, b, true);
                    break;
                case "s2":
                case "s4":
                    SidewallPanelGen(LinerPanels, Location, b, true);
                    break;
            }

        }

        // ''' 8' Liner Panels
    }
    else if ((b.LinerPanels(Location) == "8'")) {
        switch (Location) {
            case "e1":
            case "e3":
                LinerPanel = new clsPanel();
                LinerPanel.Quantity = Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0);
                LinerPanel.PanelLength = (8 * 12);
                LinerPanels.Add;
                LinerPanel;
                break;
            case "s2":
            case "s4":
                LinerPanel = new clsPanel();
                LinerPanel.Quantity = Application.WorksheetFunction.RoundUp((b.bLength / 3), 0);
                LinerPanel.PanelLength = (8 * 12);
                LinerPanels.Add;
                LinerPanel;
                break;
        }

    }
    else {
        return;
    }

    // add liner panel parameters
    for (LinerPanel in LinerPanels) {
        LinerPanel.PanelMeasurement = ImperialMeasurementFormat(LinerPanel.PanelLength);
        LinerPanel.PanelShape = EstSht.Range((Location + "_LinerPanels")).offset(0, 1).Value;
        LinerPanel.PanelType = EstSht.Range((Location + "_LinerPanels")).offset(0, 2).Value;
        LinerPanel.PanelColor = EstSht.Range((Location + "_LinerPanels")).offset(0, 3).Value;
    }

}
