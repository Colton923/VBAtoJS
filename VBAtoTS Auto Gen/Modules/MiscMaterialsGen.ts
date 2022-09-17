// TODO: Option Explicit ... Warning!!! not translated


MiscMaterialCalc(MiscMaterials: Collection, WriteCell: Range, b: clsBuilding) {
    let FOCell: Range;
    let NewMiscMaterials: Collection;
    let MiscMaterial: clsMiscItem;
    let FOArea: number;
    let WallArea: number;
    // sf
    let RoofArea: number;
    // sf
    // initalize new materials collection
    NewMiscMaterials = new Collection();
    // With...
    // additional PDoor Misc Materials
    for (FOCell in Range(EstSht.Range, "pDoorCell1", EstSht.Range, "pDoorCell12")) {
        // check that cell isn't hidden, door size is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // check for a canopy
            if ((FOCell.offset(0, 4).Value == "4' x 4'6""")) {
                MiscMaterial = new clsMiscItem();
                MiscMaterial.Measurement = "4' x 4'6""";
            }
            else if ((FOCell.offset(0, 4).Value == "4' x 7'6""")) {
                MiscMaterial = new clsMiscItem();
                MiscMaterial.Measurement = "4' x 7'6""";
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Quantity = 1;
                MiscMaterial.Name = "Door Canopy";
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

        }

    }

    // additional OHDoor materials
    for (FOCell in Range(EstSht.Range, "OHDoorCell1", EstSht.Range, "OHDoorCell12")) {
        // check that cell isn't hidden, door size is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // calculate FO area
            FOArea = (FOCell.offset(0, 1).Value * FOCell.offset(0, 2).Value);
            // add OH door
            if ((FOCell.offset(0, 4).Value == "Sectional")) {
                MiscMaterial = new clsMiscItem();
                MiscMaterial.Name = "Sectional OH Door";
            }
            else if ((FOCell.offset(0, 4).Value == "RUD")) {
                MiscMaterial = new clsMiscItem();
                MiscMaterial.Name = "Roll Up OH Door";
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Quantity = 1;
                MiscMaterial.Width = FOCell.offset(0, 1).Value;
                MiscMaterial.Height = FOCell.offset(0, 2).Value;
                MiscMaterial.Area = (MiscMaterial.Width * MiscMaterial.Height);
                MiscMaterial.Measurement = (FOCell.offset(0, 1).Text + (" x " + FOCell.offset(0, 2).Text));
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

            // '''check for insulation
            if ((FOCell.offset(0, 5).Value == "Vinyl Backed")) {
                MiscMaterial = new clsMiscItem();
                // 1 SF pieces
                MiscMaterial.Quantity = FOArea;
                MiscMaterial.Name = "Vinyl Backed Insulation";
            }
            else if ((FOCell.offset(0, 5).Value == "Steel Backed")) {
                MiscMaterial = new clsMiscItem();
                // 1 SF pieces
                MiscMaterial.Quantity = FOArea;
                MiscMaterial.Name = "Steel Backed Insulation";
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Measurement = ("1 ft" + (<string>(178)));
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

            // '''Door Operation
            switch (FOCell.offset(0, 6).Value) {
                case "Manual":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Manual Opener";
                    break;
                case "Chain Hoist":
                    if ((FOCell.offset(0, 4).Value != "RUD")) {
                        MiscMaterial = new clsMiscItem();
                        MiscMaterial.Name = "Chain Hoist Opener";
                    }

                    break;
                case "Electric Opener":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = ("Electric Opener - OH Door #" + FOCell.Value);
                    break;
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Quantity = 1;
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

            // '''High Lift
            if ((FOCell.offset(0, 7).Value == "Yes")) {
                MiscMaterial = new clsMiscItem();
                MiscMaterial.Quantity = 1;
                MiscMaterial.Name = "High Lift";
                MiscMaterial.Measurement = HighLiftSize((b.bHeight - FOCell.offset(0, 2).Value), -1);
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

            // '''OH Door Windows
            switch (FOCell.offset(0, 8).Value) {
                case "Non-Insulated":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Non-Insulated Window";
                    MiscMaterial.Measurement = "4'";
                    MiscMaterial.Quantity = Application.WorksheetFunction.RoundDown((FOCell.offset(0, 1).Value / 4), 0);
                    break;
                case "Insulated":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Insulated Window";
                    MiscMaterial.Measurement = "4'";
                    MiscMaterial.Quantity = Application.WorksheetFunction.RoundDown((FOCell.offset(0, 1).Value / 4), 0);
                    break;
                case "Full Glass Panel":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Full Glass Panel Window";
                    MiscMaterial.Measurement = "1'";
                    MiscMaterial.Quantity = FOCell.offset(0, 1).Value;
                    break;
            }

            if (!(MiscMaterial == null)) {
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

        }

    }

    // additional Window materials
    for (FOCell in Range(EstSht.Range, "WindowCell1", EstSht.Range, "WindowCell12")) {
        // check that cell isn't hidden, door size is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // add window
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Quantity = 1;
            MiscMaterial.Name = "Standard Window";
            MiscMaterial.Measurement = (FOCell.offset(0, 1).Text + (" x " + FOCell.offset(0, 2).Text));
            // area in square feet
            MiscMaterial.Area = ((FOCell.offset(0, 1).Value * FOCell.offset(0, 2).Value)
                        / 144);
            NewMiscMaterials.Add;
            MiscMaterial;
            MiscMaterial = null;
        }

    }

    // additional Misc FO materials
    for (FOCell in Range(EstSht.Range, "MiscFOCell1", EstSht.Range, "MiscFOCell12")) {
        // check that cell isn't hidden, door size is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // ''' Exhaust Fans/Louvers
            switch (FOCell.offset(0, 4).Value) {
                case "24"" Exhaust Fan":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Exhaust Fan";
                    MiscMaterial.Measurement = "24""";
                    break;
                case "30"" Exhaust Fan":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Exhaust Fan";
                    MiscMaterial.Measurement = "30""";
                    break;
                case "36"" Exhaust Fan":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Exhaust Fan";
                    MiscMaterial.Measurement = "36""";
                    break;
                case "24"" Louver":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Louver";
                    MiscMaterial.Measurement = "24""";
                    break;
                case "30"" Louver":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Louver";
                    MiscMaterial.Measurement = "30""";
                    break;
                case "36"" Louver":
                    MiscMaterial = new clsMiscItem();
                    MiscMaterial.Name = "Louver";
                    MiscMaterial.Measurement = "36""";
                    break;
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Quantity = 1;
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

            // ''' Weather Hoods
            switch (FOCell.offset(0, 5).Value) {
                case "24""":
                case "30""":
                case "36""":
                    MiscMaterial = new clsMiscItem();
                    break;
            }

            if (!(MiscMaterial == null)) {
                MiscMaterial.Quantity = 1;
                MiscMaterial.Name = "Weather Hood";
                MiscMaterial.Measurement = FOCell.offset(0, 5).Text;
                NewMiscMaterials.Add;
                MiscMaterial;
                MiscMaterial = null;
            }

        }

    }

    // '''''Insulation
    // Building Areas
    // With...
    if ((b.rShape == "Gable")) {
        WallArea = Application.WorksheetFunction.RoundUp(((2
                        * (b.bHeight * b.bLength))
                        + ((2
                        * (b.bHeight * b.bWidth))
                        + (b.bWidth
                        * ((b.bWidth * b.rPitch)
                        / 12)))), 0);
        RoofArea = Application.WorksheetFunction.RoundUp((b.bLength
                        * ((b.RafterLength / 12)
                        * 2)), 0);
    }
    else if ((b.rShape == "Single Slope")) {
        WallArea = Application.WorksheetFunction.RoundUp(((b.bHeight * b.bLength)
                        + ((b.bHeight
                        * (b.HighSideEaveHeight / 12))
                        + ((2
                        * (b.bHeight * b.bWidth))
                        + (b.bWidth
                        * ((b.bWidth * b.rPitch)
                        / 12))))), 0);
        RoofArea = Application.WorksheetFunction.RoundUp((b.bLength
                        * (b.RafterLength / 12)), 0);
    }

    // subtract OH door area from wall area
    for (MiscMaterial in NewMiscMaterials) {
        if (((MiscMaterial.Name.IndexOf("OH Door", 0) + 1)
                    != 0)) {
            // subtract area
            WallArea = (WallArea
                        - (MiscMaterial.Width * MiscMaterial.Height));
        }

    }

    // reset object
    // '' Wall Insulation '''
    MiscMaterial = null;
    switch (EstSht.Range) {
        case "WallInsulation".Value:
            break;
        case "3"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "3"" VRR Wall Insulation";
            break;
        case "4"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "4"" VRR Wall Insulation";
            break;
        case "6"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "6"" VRR Wall Insulation";
            break;
        case "1"" Spray Foam":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "1"" Spray Foam Wall Insulation";
            break;
        case "2"" Spray Foam":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "2"" Spray Foam Wall Insulation";
            break;
    }

    if (!(MiscMaterial == null)) {
        MiscMaterial.Measurement = ("1 ft" + (<string>(178)));
        MiscMaterial.Quantity = WallArea;
        NewMiscMaterials.Add;
        MiscMaterial;
        MiscMaterial = null;
    }

    // '' Roof Insulation '''
    switch (EstSht.Range) {
        case "RoofInsulation".Value:
            break;
        case "3"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "3"" VRR Roof Insulation";
            break;
        case "4"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "4"" VRR Roof Insulation";
            break;
        case "6"" VRR":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "6"" VRR Roof Insulation";
            break;
        case "1"" Spray Foam":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "1"" Spray Foam Roof Insulation";
            break;
        case "2"" Spray Foam":
            MiscMaterial = new clsMiscItem();
            MiscMaterial.Name = "2"" Spray Foam Roof Insulation";
            break;
    }

    if (!(MiscMaterial == null)) {
        MiscMaterial.Measurement = ("1 ft" + (<string>(178)));
        MiscMaterial.Quantity = RoofArea;
        NewMiscMaterials.Add;
        MiscMaterial;
        MiscMaterial = null;
    }

    // ''' Ridge Vents
    if (EstSht.Range) {
        ("RidgeVentQty".Value != 0);
        switch (EstSht.Range) {
            case "RidgeVentType".Value:
                break;
            case "Standard":
            case "Low Profile":
                MiscMaterial = new clsMiscItem();
                break;
        }

        if (!(MiscMaterial == null)) {
            MiscMaterial.Quantity = EstSht.Range;
            "RidgeVentQty".Value;
            MiscMaterial.Measurement = "10'";
            MiscMaterial.Name = EstSht.Range;
            ("RidgeVentType".Value + " Ridge Vent");
            NewMiscMaterials.Add;
            MiscMaterial;
            MiscMaterial = null;
        }

    }

    // remove duplicates
    MaterialsListGen.DuplicateMaterialRemoval(NewMiscMaterials, "Misc");
    // ''' Output to employee materials list, add to master misc materials collection
    for (MiscMaterial in NewMiscMaterials) {
        // output to employee materials list
        WriteCell.Value = MiscMaterial.Quantity;
        WriteCell.offset(0, 1).Value = MiscMaterial.Name;
        WriteCell.offset(0, 3).Value = MiscMaterial.Measurement;
        WriteCell.offset(0, 4).Value = MiscMaterial.Color;
        WriteCell = WriteCell.offset(1, 0);
        // add to master collection
        MiscMaterials.Add;
        MiscMaterial;
    }

}

// ' function returns string of the nearest available rake HighLift size
HighLiftSize() {
    (<number>(Direction));
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
    let HighLifts: Object;
    let HighLift: Object;
    let hSize: Object;
    let NearestHighLiftSizeString: string;
    let Length: Object;
    HighLifts = Array(36, 54, 72, 96, 120);
    // Convert Ft Length to inches
    Length = (FtLength * 12);
    t = 1.79769313486231E+308;
    // initialize
    for (HighLift in HighLifts) {
        if (IsNumeric(HighLift)) {
            u = Abs((HighLift - Length));
            if (((Direction > 0)
                        && (HighLift >= Length))) {
                // only report if closer number is greater than the target
                if ((u < t)) {
                    t = u;
                    hSize = HighLift;
                }

            }
            else if (((Direction < 0)
                        && (HighLift <= Length))) {
                // only report if closer number is less than the target
                if ((u < t)) {
                    t = u;
                    hSize = HighLift;
                }

            }
            else if ((Direction == 0)) {
                if ((u < t)) {
                    t = u;
                    hSize = HighLift;
                }

            }

        }

    }

    // return available High Lift size
    switch (hSize) {
        case 36:
            NearestHighLiftSizeString = "36""";
            break;
        case 54:
            NearestHighLiftSizeString = "54""";
            break;
        case 72:
            NearestHighLiftSizeString = "72""";
            break;
        case 96:
            NearestHighLiftSizeString = "96""";
            break;
        case 120:
            NearestHighLiftSizeString = "120""";
            break;
    }

    return NearestHighLiftSizeString;
}
