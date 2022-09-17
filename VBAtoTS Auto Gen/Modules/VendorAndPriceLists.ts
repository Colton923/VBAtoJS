// TODO: Option Explicit ... Warning!!! not translated


VendorMaterialListsGen(PanelCollection: Collection, TrimCollection: Collection, MiscCollection: Collection) {
    let VendorSht: Worksheet;
    let WriteCell: Range;
    let item: clsMiscItem;
    let Panel: clsPanel;
    let Trim: clsTrim;
    let N: number;
    // delete old output sheets
    Application.DisplayAlerts = false;
    for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
        if ((ThisWorkbook.Sheets(N).Name == "Vendor Sheet Metal Materials")) {
            ThisWorkbook.Sheets(N).Delete;
            break;
        }

    }

    for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
        if ((ThisWorkbook.Sheets(N).Name == "Vendor Misc. Materials")) {
            ThisWorkbook.Sheets(N).Delete;
            break;
        }

    }

    Application.DisplayAlerts = true;
    VendorSheetMetalShtTmp.Copy;
    /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Optimized Cut List");
    VendorSht = ThisWorkbook.Sheets("VendorSheetMetalShtTmp (2)");
    // rename
    VendorSht.Name = "Vendor Sheet Metal Materials";
    VendorSht.Visible = xlSheetVisible;
    // populate
    WriteCell = VendorSht.Range("MatListQtyCell1");
    for (Panel in PanelCollection) {
        WriteCell.Value = Panel.Quantity;
        WriteCell.offset(0, 1).Value = Panel.PanelShape;
        WriteCell.offset(0, 2).Value = Panel.PanelType;
        WriteCell.offset(0, 3).Value = Panel.PanelMeasurement;
        WriteCell.offset(0, 4).Value = Panel.PanelColor;
        WriteCell = WriteCell.offset(1, 0);
    }

    for (Trim in TrimCollection) {
        WriteCell.Value = Trim.Quantity;
        WriteCell.offset(0, 1).Value = Trim.tShape;
        WriteCell.offset(0, 2).Value = Trim.tType;
        WriteCell.offset(0, 3).Value = Trim.tMeasurement;
        WriteCell.offset(0, 4).Value = Trim.Color;
        WriteCell = WriteCell.offset(1, 0);
    }

    // write misc items that need to be sent to the sheet metal vendor list
    for (item in MiscCollection) {
        // With...
        if ((((item.Name.IndexOf("Formed Ridge Cap", 0) + 1)
                    != 0)
                    || ((item.Name == "Sculptured Gutter End Cap")
                    || ((item.Name == "Gutter Strap")
                    || ((item.Name == "Downspout Strap")
                    || ((item.Name == "Pop Rivets")
                    || ((item.Name == "Tek Screws")
                    || ((item.Name == "Lap Screws")
                    || ((item.Name == "Butyl Tape")
                    || ((item.Name == "Inside Closures")
                    || (item.Name == "Outside Closures"))))))))))) {
            WriteCell.Value = item.Quantity;
            WriteCell.offset(0, 1).Value = item.Shape;
            WriteCell.offset(0, 2).Value = item.Name;
            WriteCell.offset(0, 3).Value = item.Measurement;
            WriteCell.offset(0, 4).Value = item.Color;
            WriteCell = WriteCell.offset(1, 0);
        }

    }

    // format
    VendorSht.Columns.AutoFit;
    // '' Vendor Misc Materials List
    // set new output sheet
    VendorMiscMaterialsShtTmp.Copy;
    /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Optimized Cut List");
    VendorSht = ThisWorkbook.Sheets("VendorMiscMaterialsShtTmp (2)");
    // rename sheet
    VendorSht.Name = "Vendor Misc. Materials";
    VendorSht.Visible = xlSheetVisible;
    WriteCell = VendorSht.Range("MatListQtyCell1");
    // write misc items that need to be sent to the misc materials vendor list
    for (item in MiscCollection) {
        // With...
        if ((((item.Name.IndexOf("Formed Ridge Cap", 0) + 1)
                    == 0)
                    && ((item.Name != "Sculptured Gutter End Cap")
                    && ((item.Name != "Gutter Strap")
                    && ((item.Name != "Downspout Strap")
                    && ((item.Name != "Pop Rivets")
                    && ((item.Name != "Tek Screws")
                    && ((item.Name != "Lap Screws")
                    && ((item.Name != "Butyl Tape")
                    && ((item.Name != "Inside Closures")
                    && (item.Name != "Outside Closures"))))))))))) {
            WriteCell.Value = item.Quantity;
            WriteCell.offset(0, 1).Value = item.Name;
            WriteCell.offset(0, 3).Value = item.Measurement;
            WriteCell.offset(0, 4).Value = item.Color;
            WriteCell = WriteCell.offset(1, 0);
            // merge cells
            WriteCell.offset(0, 1).Resize(1, 2).Merge;
        }

    }

    // format
    VendorSht.Columns.AutoFit;
}

PriceListGen(PanelCollection: Collection, TrimCollection: Collection, MiscCollection: Collection) {
    let item: clsMiscItem;
    let Panel: clsPanel;
    let Trim: clsTrim;
    let LookupCol: number;
    let PriceSht: Worksheet;
    let WriteCell: Range;
    let N: number;
    let PriceTbl: ListObject;
    let Row: number;
    let LookupName: string;
    let SectionalOHDoorPriceTbl: ListObject;
    let PricingQty: number;
    let PanelType: string;
    // set master price table
    PriceTbl = MasterPriceSht.ListObjects("MasterPriceTbl");
    SectionalOHDoorPriceTbl = MasterPriceSht.ListObjects("SectionalOHDoorPriceTbl");
    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Lookup Item Prices '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    // '''' Panels
    for (Panel in PanelCollection) {
        // With...
        if ((((Panel.PanelType.IndexOf("Skylights", 0) + 1)
                    == 0)
                    && ((Panel.PanelType.IndexOf("Reverse", 0) + 1)
                    == 0))) {
            // ''footage cost for normal panels
            // check for errors
            if ((IsError(Application.VLookup(Panel.PanelType, PriceTbl.Range, 4, false)) == true)) {
                "Unknown".TotalCost = "Item Not Found";
                "Unknown".UnitCost = "Item Not Found";
                Panel.FootageCost = "Item Not Found";
            }
            else {
                // successful lookup
                (Panel.FootageCost * (Panel.PanelLength / 12).TotalCost) = (Panel.UnitCost * Panel.Quantity);
                Application.WorksheetFunction.VLookup(Panel.PanelType, PriceTbl.Range, 4, false).UnitCost = (Panel.UnitCost * Panel.Quantity);
                Panel.FootageCost = (Panel.UnitCost * Panel.Quantity);
            }

        }
        else if (((Panel.PanelType.IndexOf("Skylights", 0) + 1)
                    != 0)) {
            // unit cost for skylight panels
            if ((IsError(Application.WorksheetFunction.VLookup((Panel.PanelType + ", 12'"), PriceTbl.Range, 3, false)) == true)) {
                "Unknown".TotalCost = "Item Not Found";
                Panel.UnitCost = "Item Not Found";
            }
            else {
                // successful lookup
                Application.WorksheetFunction.VLookup((Panel.PanelType + ", 12'"), PriceTbl.Range, 3, false).TotalCost = (Panel.UnitCost * Panel.Quantity);
                Panel.UnitCost = (Panel.UnitCost * Panel.Quantity);
            }

        }
        else if (((Panel.PanelType.IndexOf("Reverse", 0) + 1)
                    != 0)) {
            // unit cost for reverse panels
            // check for errors
            PanelType = Panel.PanelType.Replace("Reverse ", "");
            if ((IsError(Application.VLookup(PanelType, PriceTbl.Range, 4, false)) == true)) {
                "Unknown".TotalCost = "Item Not Found";
                "Unknown".UnitCost = "Item Not Found";
                Panel.FootageCost = "Item Not Found";
            }
            else {
                // successful lookup
                (Panel.FootageCost * (Panel.PanelLength / 12).TotalCost) = (Panel.UnitCost * Panel.Quantity);
                Application.WorksheetFunction.VLookup(PanelType, PriceTbl.Range, 4, false).UnitCost = (Panel.UnitCost * Panel.Quantity);
                Panel.FootageCost = (Panel.UnitCost * Panel.Quantity);
            }

        }

    }

    // '''' Trim
    for (Trim in TrimCollection) {
        // With...
        // determine color
        switch (Trim.Color) {
            case "Galvalume":
                LookupCol = 7;
                break;
            case "Copper Metallic":
                LookupCol = 5;
                break;
            default:
                LookupCol = 6;
                break;
        }

        // trim name
        switch (true) {
            case ((Trim.tType.IndexOf("Jamb W/ Deadbolt", 0) + 1)
                        != 0):
            case ((Trim.tType.IndexOf("Jamb W/O Deadbolt", 0) + 1)
                        != 0):
                // lookup name
                Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, false).TotalCost = (Trim.UnitCost * Trim.Quantity);
                (Trim.tType.Substring((Trim.tType.Length - 4)) + (" "
                            + (Trim.tMeasurement + ("""" + (" Jamb Kit " + Trim.tType.Substring(0, (Trim.tType.Length - 7)).Substring((Trim.tType.Substring(0, (Trim.tType.Length - 7)).Length - (Trim.tType.Substring(0, (Trim.tType.Length - 7)).Length - 5))).UnitCost))))) = (Trim.UnitCost * Trim.Quantity);
                LookupName = (Trim.UnitCost * Trim.Quantity);
                // '' Door Slab are in collection with trim
                break;
            case ((Trim.tType.IndexOf("Door Slab", 0) + 1)
                        != 0):
                // lookup name
                (Trim.tType.Substring((Trim.tType.Length - 4)) + (" " + Trim.tType.Substring(0, (Trim.tType.Length - 7)).UnitCost)) = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, false);
                LookupName = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, false);
                // TBD Placeholder
                if ((Trim.UnitCost != "TBD")) {
                    Trim.TotalCost = (Trim.UnitCost * Trim.Quantity);
                }
                else {
                    Trim.TotalCost = "TBD";
                }

                // '' Pitch String Items
                break;
        }

        Application.WorksheetFunction.VLookup("High-Side Eave", PriceTbl.Range, LookupCol, false).UnitCost = (Trim.UnitCost * Trim.Quantity);
        ((Trim.tType.IndexOf("High-Side Eave", 0) + 1)
                    != 0.FootageCost) = (Trim.UnitCost * Trim.Quantity);
        (Trim.FootageCost * (Trim.tLength / 12).TotalCost) = (Trim.UnitCost * Trim.Quantity);
        Application.WorksheetFunction.VLookup("Short Eave", PriceTbl.Range, LookupCol, false).UnitCost = (Trim.UnitCost * Trim.Quantity);
        ((Trim.tType.IndexOf("Short Eave", 0) + 1)
                    != 0.FootageCost) = (Trim.UnitCost * Trim.Quantity);
        (Trim.FootageCost * (Trim.tLength / 12).TotalCost) = (Trim.UnitCost * Trim.Quantity);
        Application.WorksheetFunction.VLookup("Sculptured Gutter Hang-On", PriceTbl.Range, LookupCol, false).UnitCost = (Trim.UnitCost * Trim.Quantity);
        ((Trim.tType.IndexOf("Sculptured Gutter Hang-On", 0) + 1)
                    != 0.FootageCost) = (Trim.UnitCost * Trim.Quantity);
        // '' Head Trim (assumed to be the same cost with or without kickout)
        (Trim.FootageCost * (Trim.tLength / 12).TotalCost) = (Trim.UnitCost * Trim.Quantity);
        Application.WorksheetFunction.VLookup("Head Trim", PriceTbl.Range, LookupCol, false).UnitCost = (Trim.UnitCost * Trim.Quantity);
        ((Trim.tType.IndexOf("Head Trim", 0) + 1)
                    != 0.FootageCost) = (Trim.UnitCost * Trim.Quantity);
        // '' Flat Rate Items
        ((Trim.tType.IndexOf("Formed Ridge Cap", 0) + 1)
                    != 0);
        ((Trim.tType.IndexOf("Sculptured Gutter End Cap", 0) + 1)
                    != 0);
        ((Trim.tType.IndexOf("Gutter Strap", 0) + 1)
                    != 0);
        // only a flat unit cost
        Application.WorksheetFunction.VLookup(Trim.tType, PriceTbl.Range, (LookupCol + 3), false).TotalCost = (Trim.UnitCost * Trim.Quantity);
        Trim.UnitCost = (Trim.UnitCost * Trim.Quantity);
        // '' Normal Items
        // '' Perform vlookup as normal
        // Check for errors
        if ((IsError(Application.VLookup(Trim.tType, PriceTbl.Range, LookupCol, false)) == true)) {
            "Unknown".TotalCost = "Item Not Found";
            "Unknown".UnitCost = "Item Not Found";
            Trim.FootageCost = "Item Not Found";
        }
        else {
            //  No lookup error
            (Trim.FootageCost * (Trim.tLength / 12).TotalCost) = (Trim.UnitCost * Trim.Quantity);
            Application.WorksheetFunction.VLookup(Trim.tType, PriceTbl.Range, LookupCol, false).UnitCost = (Trim.UnitCost * Trim.Quantity);
            Trim.FootageCost = (Trim.UnitCost * Trim.Quantity);
        }

        // With...
        Trim;
        // ''''' Misc Items
        for (item in MiscCollection) {
            // With...
            // determine color
            switch (item.Color) {
                case "Galvalume":
                    LookupCol = 7;
                    break;
                case "Copper Metallic":
                    LookupCol = 5;
                    break;
                default:
                    LookupCol = 6;
                    break;
            }

            switch (true) {
                case ((item.Name.IndexOf("Pop Rivets", 0) + 1)
                            != 0):
                case ((item.Name.IndexOf("Tek Screws", 0) + 1)
                            != 0):
                case ((item.Name.IndexOf("Lap Screws", 0) + 1)
                            != 0):
                    if ((item.Name == "Pop Rivets")) {
                        PricingQty = Application.WorksheetFunction.RoundUp((item.Quantity / 100), 0);
                    }
                    else if (((item.Name == "Tek Screws")
                                || (item.Name == "Lap Screws"))) {
                        PricingQty = Application.WorksheetFunction.RoundUp((item.Quantity / 250), 0);
                    }

                    Application.WorksheetFunction.VLookup(item.Name, PriceTbl.Range, 3, false).TotalCost = (item.UnitCost * PricingQty);
                    item.UnitCost = (item.UnitCost * PricingQty);
                    // Sectional OH Doors
                    break;
                case ((item.Name.IndexOf("Sectional OH Door", 0) + 1)
                            != 0):
                    // '''find width, height
                    // TODO: On Error Resume Next Warning!!!: The statement is not translatable
                    "Size Not Found".TotalCost = "Size Not Found";
                    item.UnitCost = "Size Not Found";
                    for (Row = 1; (Row <= SectionalOHDoorPriceTbl.ListRows.Count); Row++) {
                        if ((SectionalOHDoorPriceTbl.DataBodyRange(Row, 1) == item.Width)) {
                            SectionalOHDoorPriceTbl.DataBodyRange(Row, SectionalOHDoorPriceTbl.ListColumns(item.Height.ToString()).Index).TotalCost = (item.UnitCost * item.Quantity);
                            item.UnitCost = (item.UnitCost * item.Quantity);
                            break;
                        }

                    }

                    // resume error handling
                    // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
                    // masterpricesht.ListObjects("SectionalOHDoorPriceTbl").DataBodyRange
                    // Flat Rate Colored Items
                    break;
                case ((item.Name.IndexOf("Formed Ridge Cap", 0) + 1)
                            != 0):
                case ((item.Name.IndexOf("Sculptured Gutter End Cap", 0) + 1)
                            != 0):
                case ((item.Name.IndexOf("Gutter Strap", 0) + 1)
                            != 0):
                    // ''only a flat unit cost
                    // '' Pitch String
                    if (((item.Name.IndexOf("Formed Ridge Cap", 0) + 1)
                                != 0)) {
                        Application.WorksheetFunction.VLookup("Formed Ridge Cap", PriceTbl.Range, (LookupCol + 3), false).TotalCost = (item.UnitCost * item.Quantity);
                        item.UnitCost = (item.UnitCost * item.Quantity);
                    }
                    else {
                        // '' Normal Items
                        Application.WorksheetFunction.VLookup(item.Name, PriceTbl.Range, (LookupCol + 3), false).TotalCost = (item.UnitCost * item.Quantity);
                        item.UnitCost = (item.UnitCost * item.Quantity);
                    }

                    // '' other items
                    break;
                default:
                    if (((item.Name.IndexOf("Wall Insulation", 0) + 1)
                                != 0)) {
                        // remove "Wall" from name, lookup
                        Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 4, false).TotalCost = (item.UnitCost * item.Quantity);
                        (item.Name.Substring(0, (item.Name.IndexOf(" Wall", 0) + 1)) + "Insulation".UnitCost) = (item.UnitCost * item.Quantity);
                        LookupName = (item.UnitCost * item.Quantity);
                    }
                    else if (((item.Name.IndexOf("Roof Insulation", 0) + 1)
                                != 0)) {
                        // remove "Wall" from name, lookup
                        Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 4, false).TotalCost = (item.UnitCost * item.Quantity);
                        (item.Name.Substring(0, (item.Name.IndexOf(" Roof", 0) + 1)) + "Insulation".UnitCost) = (item.UnitCost * item.Quantity);
                        LookupName = (item.UnitCost * item.Quantity);
                    }
                    else if ((((item.Name.IndexOf("High Lift", 0) + 1)
                                != 0)
                                || (((item.Name.IndexOf("Door Canopy", 0) + 1)
                                != 0)
                                || (((item.Name.IndexOf("Exhaust Fan", 0) + 1)
                                != 0)
                                || (((item.Name.IndexOf("Louver", 0) + 1)
                                != 0)
                                || ((item.Name.IndexOf("Weather Hood", 0) + 1)
                                != 0)))))) {
                        // put measurement first
                        Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, false).TotalCost = (item.UnitCost * item.Quantity);
                        (item.Measurement + (" " + item.Name.UnitCost)) = (item.UnitCost * item.Quantity);
                        LookupName = (item.UnitCost * item.Quantity);
                    }

                    // ''' Normal Items
                    for (Row = 1; (Row <= PriceTbl.ListRows.Count); Row++) {
                        if ((PriceTbl.DataBodyRange(Row, 1) == item.Name)) {
                            // check for price per measurement data
                            if (((PriceTbl.DataBodyRange(Row, 4) != "-")
                                        && (PriceTbl.DataBodyRange(Row, 4) != ""))) {
                                // items needed to be priced by area that doesn't match to the stored quantity
                                if ((((item.Name.IndexOf("Roll Up OH Door", 0) + 1)
                                            != 0)
                                            || (((item.Name.IndexOf("Standard Window", 0) + 1)
                                            != 0)
                                            || ((item.Name.IndexOf("Full Glass Panel Window", 0) + 1)
                                            != 0)))) {
                                    // footage cost is actually cost per SF but keeping var name for now
                                    (item.Area * PriceTbl.DataBodyRange(Row, 4).TotalCost) = (item.UnitCost * item.Quantity);
                                    PriceTbl.DataBodyRange(Row, 4).UnitCost = (item.UnitCost * item.Quantity);
                                    item.FootageCost = (item.UnitCost * item.Quantity);
                                    // Items with Quantity matching measurement
                                }
                                else {
                                    PriceTbl.DataBodyRange(Row, 4).TotalCost = (item.UnitCost * item.Quantity);
                                    item.UnitCost = (item.UnitCost * item.Quantity);
                                }

                                // check for flat rate data
                            }
                            else if (((PriceTbl.DataBodyRange(Row, 3) != "-")
                                        && (PriceTbl.DataBodyRange(Row, 3) != ""))) {
                                item.UnitCost = PriceTbl.DataBodyRange(Row, 3);
                                // TBD Placeholder
                                if ((item.UnitCost != "TBD")) {
                                    item.TotalCost = (item.UnitCost * item.Quantity);
                                }
                                else {
                                    item.TotalCost = "TBD";
                                }

                            }

                            break;
                        }

                    }

                    // '' Electric Opener and Unknown Items
                    // Electric Opener
                    if (((item.Name.IndexOf("Electric Opener", 0) + 1)
                                != 0)) {
                        "Input Required".TotalCost = "Input Required";
                        item.UnitCost = "Input Required";
                    }
                    else if ((item.TotalCost == "")) {
                        "Unknown".TotalCost = "Item Not Found";
                        "Unknown".UnitCost = "Item Not Found";
                        item.FootageCost = "Item Not Found";
                    }

                    break;
            }

        }

        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Generate Price Sheet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        // delete old output sheet
        Application.DisplayAlerts = false;
        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
            if ((ThisWorkbook.Sheets(N).Name == "Materials Price List")) {
                ThisWorkbook.Sheets(N).Delete;
                break;
            }

        }

        Application.DisplayAlerts = true;
        MaterialsPriceShtTmp.Copy;
        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Project Details");
        PriceSht = ThisWorkbook.Sheets("MaterialsPriceShtTmp (2)");
        // rename
        PriceSht.Name = "Materials Price List";
        PriceSht.Visible = xlSheetVisible;
        //  Begin Output
        WriteCell = PriceSht.Range("PriceListQtyCell1");
        for (Panel in PanelCollection) {
            WriteCell.Value = Panel.Quantity;
            WriteCell.offset(0, 1).Value = Panel.PanelShape;
            WriteCell.offset(0, 2).Value = Panel.PanelType;
            WriteCell.offset(0, 3).Value = Panel.PanelMeasurement;
            WriteCell.offset(0, 4).Value = Panel.PanelColor;
            WriteCell.offset(0, 5).Value = Panel.FootageCost;
            WriteCell.offset(0, 6).Value = Panel.UnitCost;
            WriteCell.offset(0, 7).Value = Panel.TotalCost;
            WriteCell = WriteCell.offset(1, 0);
        }

        for (Trim in TrimCollection) {
            WriteCell.Value = Trim.Quantity;
            WriteCell.offset(0, 1).Value = Trim.tShape;
            WriteCell.offset(0, 2).Value = Trim.tType;
            WriteCell.offset(0, 3).Value = Trim.tMeasurement;
            WriteCell.offset(0, 4).Value = Trim.Color;
            WriteCell.offset(0, 5).Value = Trim.FootageCost;
            WriteCell.offset(0, 6).Value = Trim.UnitCost;
            WriteCell.offset(0, 7).Value = Trim.TotalCost;
            WriteCell = WriteCell.offset(1, 0);
        }

        for (item in MiscCollection) {
            WriteCell.Value = item.Quantity;
            WriteCell.offset(0, 1).Value = item.Shape;
            WriteCell.offset(0, 2).Value = item.Name;
            WriteCell.offset(0, 3).Value = item.Measurement;
            WriteCell.offset(0, 4).Value = item.Color;
            WriteCell.offset(0, 5).Value = item.FootageCost;
            WriteCell.offset(0, 6).Value = item.UnitCost;
            WriteCell.offset(0, 7).Value = item.TotalCost;
            WriteCell = WriteCell.offset(1, 0);
        }

        // format
        PriceSht.Columns.AutoFit;
        CostEstimateGen((<Collection>(PanelCollection)), (<Collection>(TrimCollection)), (<Collection>(MiscCollection)), (<clsBuilding>(b)));
        let StructuralSteel: Collection = new Collection();
        let SheetMetal: Collection = new Collection();
        let OHDoors: Collection = new Collection();
        let Insulation: Collection = new Collection();
        let ElectricOpeners: Collection = new Collection();
        let Windows: Collection = new Collection();
        let RidgeVents: Collection = new Collection();
        let DoorCanopies: Collection = new Collection();
        let ExhaustFansLouversWeatherhoods: Collection = new Collection();
        let clsItem: clsMiscItem;
        let clsPanel: clsPanel;
        let clsTrim: clsTrim;
        let CollectionItem: Object;
        // Material Name Sorting Arrays
        let OHDoorNames: Object;
        let InsulationNames: Object;
        let CollectionTotal: Currency;
        let N: number;
        let CostEstimateSht: Worksheet;
        // Oh Door Material Collections
        OHDoorNames = Array("Roll Up OH Door", "Sectional OH Door", "Chain Hoist Opener", "High Lift", "Non-Insulated Window", "Insulated Window", "Full Glass Panel Window", "Vinyl Backed Insulation", "Steel Backed Insulation");
        InsulationNames = Array("3"" VRR", "4"" VRR", "6"" VRR", "1"" Spray Foam", "2"" Spray Foam");
        // delete old output sheets
        Application.DisplayAlerts = false;
        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
            if ((ThisWorkbook.Sheets(N).Name == "Cost Estimate")) {
                ThisWorkbook.Sheets(N).Delete;
                break;
            }

        }

        Application.DisplayAlerts = true;
        CostEstimateShtTmp.Copy;
        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Structural Steel Price List");
        CostEstimateSht = ThisWorkbook.Sheets("CostEstimateShtTmp (2)");
        // rename
        CostEstimateSht.Name = "Cost Estimate";
        CostEstimateSht.Columns.AutoFit;
        CostEstimateSht.Visible = xlSheetVisible;
        // Set header info
        // With...
        EstSht.Range("CustomerName").Value.Range("B4").Value = Now();
        ("Job Information - "
                    + (EstSht.Range("CustomerName").Value
                    + (b.bWidth + ("' x "
                    + (b.bLength + ("' x "
                    + (b.bHeight + "'".Range("B3").Value))))))) = Now();
        CostEstimateSht.Range("A1").Value = Now();
        // debug
        Debug.Print;
        "------------------------------------- Cost Estimate Prep ------------------------------";
        Debug.Print;
        "---------------------- Panel Collection ---------------";
        for (clsPanel in PanelCollection) {
            SheetMetal.Add;
            clsPanel;
            Debug.Print;
            (clsPanel.PanelType + " added to SheetMetal");
        }

        Debug.Print;
        "---------------------- Trim Collection ---------------";
        for (clsTrim in TrimCollection) {
            // check for OH doors
            for (N = LBound(OHDoorNames); (N <= UBound(OHDoorNames)); N++) {
                if (((clsTrim.tType.IndexOf(OHDoorNames[N], 0) + 1)
                            != 0)) {
                    OHDoors.Add;
                    clsTrim;
                    Debug.Print;
                    (clsTrim.tType + " added to OHDoors");
                    /* Warning! GOTO is not Implemented */}

            }

            // Check for Insulation
            for (N = LBound(InsulationNames); (N <= UBound(InsulationNames)); N++) {
                if (((clsTrim.tType.IndexOf(InsulationNames[N], 0) + 1)
                            != 0)) {
                    Insulation.Add;
                    clsTrim;
                    Debug.Print;
                    (clsTrim.tType + " added to Insulation");
                    /* Warning! GOTO is not Implemented */}

            }

            // Check for Windows
            if ((clsTrim.tType == "Standard Window")) {
                Windows.Add;
                clsTrim;
                Debug.Print;
                (clsTrim.tType + " added to Windows");
                /* Warning! GOTO is not Implemented */// check for electric openers
            }
            else if (((clsTrim.tType.IndexOf("Electric Opener", 0) + 1)
                        != 0)) {
                ElectricOpeners.Add;
                clsTrim;
                Debug.Print;
                (clsTrim.tType + " added to Electric Openers");
                /* Warning! GOTO is not Implemented */// Check for Ridge Vents
            }
            else if (((clsTrim.tType.IndexOf("Ridge Vent", 0) + 1)
                        != 0)) {
                RidgeVents.Add;
                clsTrim;
                Debug.Print;
                (clsTrim.tType + " added to Ridge Vents");
                /* Warning! GOTO is not Implemented */// check for door canopies
            }
            else if (((clsTrim.tType.IndexOf("Door Canopy", 0) + 1)
                        != 0)) {
                DoorCanopies.Add;
                clsTrim;
                Debug.Print;
                (clsTrim.tType + " added to Door Canopies");
                /* Warning! GOTO is not Implemented */// check for exhaust fans, louvers, or weather hoods
            }
            else if ((((clsTrim.tType.IndexOf("Exhaust Fan", 0) + 1)
                        != 0)
                        || (((clsTrim.tType.IndexOf("Louver", 0) + 1)
                        != 0)
                        || ((clsTrim.tType.IndexOf("Weather Hood", 0) + 1)
                        != 0)))) {
                ExhaustFansLouversWeatherhoods.Add;
                clsTrim;
                Debug.Print;
                (clsTrim.tType + " added to ExhaustFans,Louvers,Weatherhoods");
                /* Warning! GOTO is not Implemented */}

            // otherwise, add to sheet metal collection
            SheetMetal.Add;
            clsTrim;
            Debug.Print;
            (clsTrim.tType + " added to Sheet Metal");
            /* Warning! Labeled Statements are not Implemented */}

        Debug.Print;
        "---------------------- MiscItem Collection ---------------";
        for (clsItem in MiscCollection) {
            // Check for electric Opener
            if (((clsItem.Name.IndexOf("Electric Opener", 0) + 1)
                        != 0)) {
                // With...
                /* Warning! Labeled Statements are not Implemented */xlDown.Range("Insulation_TotalCost").offset(-1, -1).Value = clsItem.Name;
                // cost, markup %
                CostEstimateSht.Range("Insulation_TotalCost").EntireRow.Insert.Range;
                "<Enter Cost>".Range("Insulation_TotalCost").offset(-1, 2).Value = CostEstimateSht.Range("Insulation_TotalCost").EntireRow.Insert.Range;
                "Insulation_TotalCost".offset(-1, 0).Value = CostEstimateSht.Range("Insulation_TotalCost").EntireRow.Insert.Range;
                "OHDoors_TotalCost".offset(0, 2).Value;
                // add formulas
                CostEstimateSht.Range("Insulation_TotalCost").EntireRow.Insert.Range;
                "Insulation_TotalCost".offset(-1, 1).Resize(2, 1).FillUp.Range("Insulation_TotalCost").offset(-1, 3).Resize(2, 1).FillUp.Range("Insulation_TotalCost").offset(-1, 4).Resize(2, 1).FillUp;
            }

            // check for OH doors
            for (N = LBound(OHDoorNames); (N <= UBound(OHDoorNames)); N++) {
                if (((clsItem.Name.IndexOf(OHDoorNames[N], 0) + 1)
                            != 0)) {
                    OHDoors.Add;
                    clsItem;
                    Debug.Print;
                    (clsItem.Name + " added to OHDoors");
                    /* Warning! GOTO is not Implemented */}

            }

            // Check for Insulation
            for (N = LBound(InsulationNames); (N <= UBound(InsulationNames)); N++) {
                if (((clsItem.Name.IndexOf(InsulationNames[N], 0) + 1)
                            != 0)) {
                    Insulation.Add;
                    clsItem;
                    Debug.Print;
                    (clsItem.Name + " added to Insulation");
                    /* Warning! GOTO is not Implemented */}

            }

            // Check for Windows
            if ((clsItem.Name == "Standard Window")) {
                Windows.Add;
                clsItem;
                Debug.Print;
                (clsItem.Name + " added to Windows");
                /* Warning! GOTO is not Implemented */// check for electric openers
            }
            else if (((clsItem.Name.IndexOf("Electric Opener", 0) + 1)
                        != 0)) {
                ElectricOpeners.Add;
                clsItem;
                Debug.Print;
                (clsItem.Name + " added to Electric Openers");
                /* Warning! GOTO is not Implemented */// Check for Ridge Vents
            }
            else if (((clsItem.Name.IndexOf("Ridge Vent", 0) + 1)
                        != 0)) {
                RidgeVents.Add;
                clsItem;
                Debug.Print;
                (clsItem.Name + " added to Ridge Vents");
                /* Warning! GOTO is not Implemented */// check for door canopies
            }
            else if (((clsItem.Name.IndexOf("Door Canopy", 0) + 1)
                        != 0)) {
                DoorCanopies.Add;
                clsItem;
                Debug.Print;
                (clsItem.Name + " added to Door Canopies");
                /* Warning! GOTO is not Implemented */// check for exhaust fans, louvers, or weather hoods
            }
            else if ((((clsItem.Name.IndexOf("Exhaust Fan", 0) + 1)
                        != 0)
                        || (((clsItem.Name.IndexOf("Louver", 0) + 1)
                        != 0)
                        || ((clsItem.Name.IndexOf("Weather Hood", 0) + 1)
                        != 0)))) {
                ExhaustFansLouversWeatherhoods.Add;
                clsItem;
                Debug.Print;
                (clsItem.Name + " added to ExhaustFans,Louvers,Weatherhoods");
                /* Warning! GOTO is not Implemented */}

            // otherwise, add to sheet metal collection
            SheetMetal.Add;
            clsItem;
            Debug.Print;
            (clsItem.Name + " added to Sheet Metal");
            /* Warning! Labeled Statements are not Implemented */}

        // '''''''''''''''''''''''''' Total Collections and Output
        // With...
        // Structural Steel
        CostEstimateSht.Range;
        "StructuralSteel_TotalCost".Value = b.SSTotalCost;
        // Sheet Metal
        for (CollectionItem in SheetMetal) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("SheetMetal_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // OH Doors
        for (CollectionItem in OHDoors) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("OHDoors_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // Insulation
        for (CollectionItem in Insulation) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("Insulation_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // Windows
        for (CollectionItem in Windows) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("Windows_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // Ridge Vents
        for (CollectionItem in RidgeVents) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("RidgeVents_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // Door Canopies
        for (CollectionItem in DoorCanopies) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("DoorCanopies_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // Exhaust Fans, Louvers, Weatherhoods
        for (CollectionItem in ExhaustFansLouversWeatherhoods) {
            if ((IsNumeric(CollectionItem.TotalCost) == true)) {
                CollectionTotal = (CollectionTotal + CollectionItem.TotalCost);
            }

        }

        CollectionItem.Range("ExhaustFansLouversWeatherhoods_TotalCost").Value = CollectionTotal;
        CollectionTotal = 0;
        // delete empty line items (other than structural steel, misc. items and electric openers)
        if (CostEstimateSht.Range) {
            "OHDoors_TotalCost".Value = 0;
            CostEstimateSht.Range;
            "OHDoors_TotalCost".EntireRow.Delete;
            /* Warning! Labeled Statements are not Implemented */xlUp;
            if (CostEstimateSht.Range) {
                "Insulation_TotalCost".Value = 0;
                CostEstimateSht.Range;
                "Insulation_TotalCost".EntireRow.Delete;
                /* Warning! Labeled Statements are not Implemented */xlUp;
                if (CostEstimateSht.Range) {
                    "Windows_TotalCost".Value = 0;
                    CostEstimateSht.Range;
                    "Windows_TotalCost".EntireRow.Delete;
                    /* Warning! Labeled Statements are not Implemented */xlUp;
                    if (CostEstimateSht.Range) {
                        "RidgeVents_TotalCost".Value = 0;
                        CostEstimateSht.Range;
                        "RidgeVents_TotalCost".EntireRow.Delete;
                        /* Warning! Labeled Statements are not Implemented */xlUp;
                        if (CostEstimateSht.Range) {
                            "DoorCanopies_TotalCost".Value = 0;
                            CostEstimateSht.Range;
                            "DoorCanopies_TotalCost".EntireRow.Delete;
                            /* Warning! Labeled Statements are not Implemented */xlUp;
                            if (CostEstimateSht.Range) {
                                "ExhaustFansLouversWeatherhoods_TotalCost".Value = 0;
                                CostEstimateSht.Range;
                                "ExhaustFansLouversWeatherhoods_TotalCost".EntireRow.Delete;
                                /* Warning! Labeled Statements are not Implemented */xlUp;
                            }

                            // Populate Labor Section
                            LaborGen(CostEstimateSht, b, TrimCollection, MiscCollection);
                        }

                        LaborGen((<Worksheet>(CostEstSht)), (<clsBuilding>(b)), (<Collection>(TrimCollection)), (<Collection>(MiscCollection)));
                        let FOCell: Range;
                        let ItemCount: number;
                        let ItemLF: number;
                        let ItemSF: number;
                        let clsTrim: clsTrim;
                        let clsItem: clsMiscItem;
                        let Row: number;
                        // With...
                        // ''Erection
                        // Building Width * Building Length
                        CostEstSht.Range;
                        "Erection".Value = (b.bLength * b.bWidth);
                        // ''Height Premium
                        // Wall Square Footage over 17'
                        if ((b.bHeight > 17)) {
                            // calculate SF for overage
                            ItemSF = ((b.bHeight - 17)
                                        * ((b.bLength * 2)
                                        + (b.bWidth * 2)));
                        }

                        CostEstSht.Range;
                        "HeightPremium".Value = ItemSF;
                        ItemSF = 0;
                        // ''Pitch Premium
                        // Roof Area and Endwall Area due to Pitch
                        if ((b.rShape == "Single Slope")) {
                            // additional roof area
                            ItemSF = (((b.RafterLength / 12)
                                        * b.bLength)
                                        - (b.bLength * b.bWidth));
                            ItemSF = (ItemSF
                                        + (b.bWidth
                                        * ((b.HighSideEaveHeight / 12)
                                        - b.bHeight)));
                        }
                        else if ((b.rShape == "Gable")) {
                            // additional roof area
                            ItemSF = (((b.RafterLength / 12) * (2 * b.bLength))
                                        - (b.bLength * b.bWidth));
                            ItemSF = (ItemSF
                                        + ((b.bWidth / 2)
                                        * (((b.bWidth / 2)
                                        * b.rPitch)
                                        / 12)));
                        }

                        CostEstSht.Range;
                        "PitchPremium".Value = ItemSF;
                        ItemSF = 0;
                        // ''PDoors
                        // PDoor Count
                        for (FOCell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                            if ((FOCell.offset(0, 1).Value != "")) {
                                ItemCount = (ItemCount + 1);
                            }

                        }

                        FOCell.Range("PDoors").Value = ItemCount;
                        ItemCount = 0;
                        // ''Door Canopies
                        // Canopy Count
                        for (clsItem in MiscCollection) {
                            if (((clsItem.Name.IndexOf("Door Canopy", 0) + 1)
                                        != 0)) {
                                ItemCount = (ItemCount + 1);
                            }

                        }

                        clsItem.Range("DoorCanopies").Value = ItemCount;
                        ItemCount = 0;
                        // ''OH Doors
                        // OHdoor Count
                        for (FOCell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                            if ((FOCell.offset(0, 1).Value != "")) {
                                ItemCount = (ItemCount + 1);
                            }

                        }

                        FOCell.Range("OHDoors").Value = ItemCount;
                        ItemCount = 0;
                        // ''Windows
                        // window count
                        for (FOCell in Range(EstSht.Range("WindowCell1"), EstSht.Range("WindowCell12"))) {
                            if ((FOCell.offset(0, 1).Value != "")) {
                                ItemCount = (ItemCount + 1);
                            }

                        }

                        FOCell.Range("Windows").Value = ItemCount;
                        ItemCount = 0;
                        // ''Misc FOs
                        // misc FO count
                        for (FOCell in Range(EstSht.Range("MiscFOCell1"), EstSht.Range("MiscFOCell12"))) {
                            if ((FOCell.offset(0, 1).Value != "")) {
                                ItemCount = (ItemCount + 1);
                            }

                        }

                        FOCell.Range("MiscFOs").Value = ItemCount;
                        ItemCount = 0;
                        // ''VRR Wall & Roof Insulation
                        // Wall/Roof Insulation Square Footage
                        for (clsItem in MiscCollection) {
                            switch (true) {
                            }

                            clsItem.Name = clsItem.Quantity.Range("VRRWallInsulation4Inch").EntireRow.Delete;
                            "4"" VRR Wall Insulation".Range("VRRWallInsulation4Inch").Value = clsItem.Quantity.Range("VRRWallInsulation3Inch").EntireRow.Delete;
                            clsItem.Name = clsItem.Quantity.Range("VRRWallInsulation3Inch").EntireRow.Delete;
                            "3"" VRR Roof Insulation".Range("VRRRoofInsulation3Inch").Value = clsItem.Quantity.Range("VRRRoofInsulation4Inch").EntireRow.Delete.Range("VRRRoofInsulation6Inch").EntireRow.Delete;
                            clsItem.Name = clsItem.Quantity.Range("VRRRoofInsulation4Inch").EntireRow.Delete.Range("VRRRoofInsulation6Inch").EntireRow.Delete;
                            "4"" VRR Roof Insulation".Range("VRRRoofInsulation4Inch").Value = clsItem.Quantity.Range("VRRRoofInsulation3Inch").EntireRow.Delete.Range("VRRRoofInsulation6Inch").EntireRow.Delete;
                            clsItem.Name = clsItem.Quantity.Range("VRRRoofInsulation3Inch").EntireRow.Delete.Range("VRRRoofInsulation6Inch").EntireRow.Delete;
                            "6"" VRR Roof Insulation".Range("VRRRoofInsulation6Inch").Value = clsItem.Quantity.Range("VRRRoofInsulation3Inch").EntireRow.Delete.Range("VRRRoofInsulation4Inch").EntireRow.Delete;
                            clsItem.Name = clsItem.Quantity.Range("VRRRoofInsulation3Inch").EntireRow.Delete.Range("VRRRoofInsulation4Inch").EntireRow.Delete;
                            clsItem;
                            // ''Ridge Vents
                            // Ridge Vent Count
                            for (clsItem in MiscCollection) {
                                if (((clsItem.Name.IndexOf("Ridge Vent", 0) + 1)
                                            != 0)) {
                                    CostEstSht.Range;
                                }

                                "RidgeVents".Value = clsItem.Quantity;
                            }

                            // ''Gutters
                            // LF of gutter hang-ons
                            for (clsTrim in TrimCollection) {
                                if (((clsTrim.tType.IndexOf("Sculptured Gutter Hang-On", 0) + 1)
                                            != 0)) {
                                    ItemCount = (ItemCount
                                                + (clsTrim.tLength * clsTrim.Quantity));
                                }

                            }

                            clsTrim.Range("Gutters").Value = (ItemCount / 12);
                            ItemCount = 0;
                            // '''' SF for the 8 below items
                            // With...
                            // ''Gable Overhangs
                            if ((b.rShape == "Single Slope")) {
                                ItemSF = (((b.s2RafterSheetLength
                                            + (b.s2ExtensionRafterLength + b.s4ExtensionRafterLength))
                                            * (b.e1Overhang + b.e3Overhang))
                                            / 144);
                            }
                            else if ((b.rShape == "Gable")) {
                                ItemSF = (((b.s2RafterSheetLength + b.s2ExtensionRafterLength)
                                            * (b.e1Overhang + b.e3Overhang))
                                            / 144);
                                // s2
                                ItemSF = (ItemSF
                                            + (((b.s4RafterSheetLength + b.s4ExtensionRafterLength)
                                            * (b.e1Overhang + b.e3Overhang))
                                            / 144));
                            }

                            CostEstSht.Range("GableOverhangs").Value = ItemSF;
                            ItemSF = 0;
                            // ''Eave Overhangs
                            ItemSF = (((b.s2Overhang - 4.25)
                                        * b.RoofLength)
                                        / 144);
                            // s2
                            if (((b.s4Overhang != 4.25)
                                        && (b.s4Overhang != 0))) {
                                ItemSF = (ItemSF
                                            + (((b.s4Overhang - 4.25)
                                            * b.RoofLength)
                                            / 144));
                            }

                            CostEstSht.Range("EaveOverhangs").Value = ItemSF;
                            ItemSF = 0;
                            // ''Gable Extensions
                            if ((b.rShape == "Single Slope")) {
                                ItemSF = ((b.s2RafterSheetLength
                                            * (b.e1Extension + b.e3Extension))
                                            / 144);
                            }
                            else if ((b.rShape == "Gable")) {
                                ItemSF = ((b.s2RafterSheetLength
                                            * (b.e1Extension + b.e3Extension))
                                            / 144);
                                // s2
                                ItemSF = (ItemSF
                                            + ((b.s4RafterSheetLength
                                            * (b.e1Extension + b.e3Extension))
                                            / 144));
                            }

                            CostEstSht.Range("GableExtensions").Value = ItemSF;
                            ItemSF = 0;
                            // ''Eave Extensions
                            ItemSF = ((b.s2EaveExtensionBuildingLength * b.s2ExtensionRafterLength)
                                        / 144);
                            // s2
                            ItemSF = (ItemSF
                                        + ((b.s4EaveExtensionBuildingLength * b.s4ExtensionRafterLength)
                                        / 144));
                            CostEstSht.Range("EaveExtensions").Value = ItemSF;
                            ItemSF = 0;
                            // ''Gable Overhang Soffit
                            if ((b.rShape == "Single Slope")) {
                                if ((b.e1GableOverhangSoffit == true)) {
                                    ItemSF = ((b.s2RafterSheetLength * b.e1Overhang)
                                                / 144);
                                }

                                if ((b.e3GableOverhangSoffit == true)) {
                                    ItemSF = (ItemSF
                                                + ((b.s2RafterSheetLength * b.e3Overhang)
                                                / 144));
                                }
                                else if ((b.rShape == "Gable")) {
                                    if ((b.e1GableOverhangSoffit == true)) {
                                        ItemSF = ((b.s2RafterSheetLength * b.e1Overhang)
                                                    / 144);
                                        ItemSF = (ItemSF
                                                    + ((b.s4RafterSheetLength * b.e1Overhang)
                                                    / 144));
                                    }

                                    if ((b.e3GableOverhangSoffit == true)) {
                                        ItemSF = (ItemSF
                                                    + ((b.s2RafterSheetLength * b.e3Overhang)
                                                    / 144));
                                        ItemSF = (ItemSF
                                                    + ((b.s2RafterSheetLength * b.e3Overhang)
                                                    / 144));
                                        ItemSF = (ItemSF
                                                    + ((b.s4RafterSheetLength * b.e3Overhang)
                                                    / 144));
                                    }

                                }

                                CostEstSht.Range("GableOverhangSoffit").Value = ItemSF;
                                ItemSF = 0;
                                // ''Eave Overhang Soffit
                                if ((b.s2EaveOverhangSoffit == true)) {
                                    ItemSF = ((b.RoofLength
                                                * (b.s2Overhang - 4.25))
                                                / 144);
                                }

                                if ((b.s4EaveOverhangSoffit == true)) {
                                    ItemSF = (ItemSF
                                                + (((b.s4Overhang - 4.25)
                                                * b.RoofLength)
                                                / 144));
                                }

                                CostEstSht.Range("EaveOverhangSoffit").Value = ItemSF;
                                ItemSF = 0;
                                // ''Gable Extension Soffit
                                if ((b.rShape == "Single Slope")) {
                                    if ((b.e1GableExtensionSoffit == true)) {
                                        ItemSF = ((b.s2RafterSheetLength * b.e1Extension)
                                                    / 144);
                                    }

                                    if ((b.e3GableExtensionSoffit == true)) {
                                        ItemSF = (ItemSF
                                                    + ((b.s2RafterSheetLength * b.e3Extension)
                                                    / 144));
                                    }
                                    else if ((b.rShape == "Gable")) {
                                        if ((b.e1GableExtensionSoffit == true)) {
                                            ItemSF = ((b.s2RafterSheetLength * b.e1Extension)
                                                        / 144);
                                            ItemSF = (ItemSF
                                                        + ((b.s4RafterSheetLength * b.e1Extension)
                                                        / 144));
                                        }

                                        if ((b.e3GableExtensionSoffit == true)) {
                                            ItemSF = (ItemSF
                                                        + ((b.s2RafterSheetLength * b.e3Extension)
                                                        / 144));
                                            ItemSF = (ItemSF
                                                        + ((b.s2RafterSheetLength * b.e3Extension)
                                                        / 144));
                                            ItemSF = (ItemSF
                                                        + ((b.s4RafterSheetLength * b.e3Extension)
                                                        / 144));
                                        }

                                    }

                                    CostEstSht.Range("GableExtensionSoffit").Value = ItemSF;
                                    ItemSF = 0;
                                    // ''Eave Extension Soffit
                                    if ((b.s2EaveOverhangSoffit == true)) {
                                        ItemSF = ((b.s2EaveExtensionBuildingLength * b.s2ExtensionRafterLength)
                                                    / 144);
                                    }

                                    // s2
                                    if ((b.s4EaveOverhangSoffit == true)) {
                                        ItemSF = (ItemSF
                                                    + ((b.s4EaveExtensionBuildingLength * b.s4ExtensionRafterLength)
                                                    / 144));
                                    }

                                    CostEstSht.Range("EaveExtensionSoffit").Value = ItemSF;
                                    ItemSF = 0;
                                    // ''Wainscot
                                    if (b.Wainscot) {
                                        ("e1" != "None");
                                        ItemLF = (number.Parse(Left(b.Wainscot, "e1", 2)) * Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0));
                                        if (b.Wainscot) {
                                            ("e3" != "None");
                                            ItemLF = (ItemLF
                                                        + (number.Parse(Left(b.Wainscot, "e3", 2)) * Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0)));
                                            if (b.Wainscot) {
                                                ("s2" != "None");
                                                ItemLF = (ItemLF
                                                            + (number.Parse(Left(b.Wainscot, "s2", 2)) * Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0)));
                                                if (b.Wainscot) {
                                                    ("s4" != "None");
                                                    ItemLF = (ItemLF
                                                                + (number.Parse(Left(b.Wainscot, "s4", 2)) * Application.WorksheetFunction.RoundUp((b.bWidth / 3), 0)));
                                                    ItemLF = (ItemLF / 12);
                                                    //  convert to FT
                                                    CostEstSht.Range("Wainscot").Value = ItemLF;
                                                    ItemLF = 0;
                                                    // ''Liner Panels
                                                    // endwall 1
                                                    if (b.LinerPanels) {
                                                        "e1" = "8'";
                                                        ItemSF = (ItemSF + (8 * b.bWidth));
                                                    }
                                                    else if (b.LinerPanels) {
                                                        "e1" = "Full Height";
                                                        if ((b.rShape == "Gable")) {
                                                            ItemSF = (ItemSF
                                                                        + ((b.bWidth * b.bHeight)
                                                                        + ((b.bWidth / 2)
                                                                        * (((b.bWidth / 2)
                                                                        * b.rPitch)
                                                                        / 12))));
                                                        }
                                                        else if ((b.rShape == "Single Slope")) {
                                                            ItemSF = (ItemSF
                                                                        + ((b.bWidth * b.bHeight)
                                                                        + ((b.bWidth / 2)
                                                                        * ((b.bWidth * b.rPitch)
                                                                        / 12))));
                                                        }

                                                        ItemSF = (ItemSF
                                                                    - ((8 / 12)
                                                                    * b.bWidth));
                                                    }

                                                    // sidewall 2
                                                    if (b.LinerPanels) {
                                                        "s2" = "8'";
                                                        ItemSF = (ItemSF + (8 * b.bLength));
                                                    }
                                                    else if (b.LinerPanels) {
                                                        "s2" = "Full Height";
                                                        ItemSF = (ItemSF
                                                                    + (b.bLength
                                                                    * (b.bHeight - (8 / 12))));
                                                    }

                                                    // endwall 3
                                                    if (b.LinerPanels) {
                                                        "e3" = "8'";
                                                        ItemSF = (ItemSF + (8 * b.bWidth));
                                                    }
                                                    else if (b.LinerPanels) {
                                                        "e3" = "Full Height";
                                                        if ((b.rShape == "Gable")) {
                                                            ItemSF = (ItemSF
                                                                        + ((b.bWidth * b.bHeight)
                                                                        + ((b.bWidth / 2)
                                                                        * (((b.bWidth / 2)
                                                                        * b.rPitch)
                                                                        / 12))));
                                                        }
                                                        else if ((b.rShape == "Single Slope")) {
                                                            ItemSF = (ItemSF
                                                                        + ((b.bWidth * b.bHeight)
                                                                        + ((b.bWidth / 2)
                                                                        * ((b.bWidth * b.rPitch)
                                                                        / 12))));
                                                        }

                                                        ItemSF = (ItemSF
                                                                    - ((8 / 12)
                                                                    * b.bWidth));
                                                    }

                                                    // sidewall 4
                                                    if (b.LinerPanels) {
                                                        "s4" = "8'";
                                                        ItemSF = (ItemSF + (8 * b.bLength));
                                                    }
                                                    else if (b.LinerPanels) {
                                                        "s4" = "Full Height";
                                                        if ((b.rShape == "Gable")) {
                                                            ItemSF = (ItemSF
                                                                        + (b.bLength
                                                                        * (b.bHeight - (8 / 12))));
                                                        }
                                                        else if ((b.rShape == "Single Slope")) {
                                                            ItemSF = (ItemSF
                                                                        + (b.bLength
                                                                        * ((b.HighSideEaveHeight - 8)
                                                                        / 12)));
                                                        }

                                                    }

                                                    // Roof
                                                    if (b.LinerPanels) {
                                                        "Roof" = "Full Height";
                                                        if ((b.rShape == "Single Slope")) {
                                                            ItemSF = (ItemSF
                                                                        + (((b.RafterLength - 8)
                                                                        / 12)
                                                                        * b.bLength));
                                                        }
                                                        else if ((b.rShape == "Gable")) {
                                                            ItemSF = (ItemSF
                                                                        + ((((b.RafterLength - 8)
                                                                        / 12)
                                                                        * b.bLength)
                                                                        * 2));
                                                        }

                                                    }

                                                    CostEstSht.Range("LinerPanels").Value = ItemSF;
                                                }

                                                // '' Delete Blank Line Items
                                                for (Row = b.Range; ; Row++) {
                                                    "LinerPanels".Row;
                                                    b.Range;
                                                    "Erection".Row;
                                                    -1;
                                                    if (b.Cells) {
                                                        Row;
                                                        2;
                                                        b.Value = 0;
                                                        b.Cells;
                                                        Row;
                                                        2;
                                                        b.EntireRow.Delete;
                                                        Row;
                                                    }

                                                    DescriptionGen((<clsBuilding>(b)));
                                                    let dStr: string;
                                                    let BayStr: string;
                                                    let cell: Range;
                                                    let DescriptionSht: Worksheet;
                                                    let N: number;
                                                    let i: number;
                                                    let PDoorTypes: Object;
                                                    let OHDoorTypes: Object;
                                                    let WindowTypes: Object;
                                                    let FOTypes: Object;
                                                    let TempDesc: string;
                                                    let dCell: Range;
                                                    let RowCount: number;
                                                    // delete old output sheets
                                                    Application.DisplayAlerts = false;
                                                    for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
                                                        if ((ThisWorkbook.Sheets(N).Name == "Project Description")) {
                                                            ThisWorkbook.Sheets(N).Delete;
                                                            break;
                                                        }

                                                    }

                                                    // set new output sheet
                                                    DescriptionShtTmp.Copy;
                                                    /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Project Details");
                                                    DescriptionSht = ThisWorkbook.Sheets("DescriptionShtTmp (2)");
                                                    // rename
                                                    DescriptionSht.Name = "Project Description";
                                                    DescriptionSht.Visible = xlSheetVisible;
                                                    dCell = DescriptionSht.Range("DescriptionCell");
                                                    RowCount = 0;
                                                    // With...
                                                    // main building description
                                                    dStr = (EstSht.Range("BusinessName").Value + " agrees to provide the material, labor, and equipment to erect the following metal building: ");
                                                    dCell.Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Width: " + b.bWidth);
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Length: " + b.bLength);
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Eave height: " + b.bHeight);
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Pitch: " + b.rPitch);
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Roof Shape: " + b.rShape);
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = (EstSht.Range("BayNum").Value + " bays: ");
                                                    for (i = 1; (i <= EstSht.Range("BayNum").Value); i++) {
                                                        if ((i != 1)) {
                                                            dStr = (dStr + ", ");
                                                        }

                                                        dStr = (dStr + ("bay #"
                                                                    + (i + (": "
                                                                    + (EstSht.Range("Bay1_Length").offset((i - 1), 0).Value + "'")))));
                                                    }

                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    // Walls and Wall Status
                                                    for (i = 1; (i <= 4); i++) {
                                                        if (((i == 1)
                                                                    || (i == 3))) {
                                                            dStr = ("Endwall "
                                                                        + (i + ": "));
                                                        }
                                                        else {
                                                            dStr = ("Sidewall "
                                                                        + (i + ": "));
                                                        }

                                                        // status
                                                        if ((EstSht.Range("e1_WallStatus").offset((i - 1), 0).Value != "Include")) {
                                                            dStr = (dStr + EstSht.Range("e1_WallStatus").offset((i - 1), 0).Value);
                                                        }
                                                        else {
                                                            dStr = (dStr + "included");
                                                        }

                                                        // expandable
                                                        if ((EstSht.Range("e1_WallStatus").offset((i - 1), 1).Value == "Yes")) {
                                                            dStr = (dStr + ", expandable");
                                                        }

                                                        // Ft above finished floor IF PARTIAL
                                                        if ((EstSht.Range("e1_WallStatus").offset((i - 1), 0).Value == "Partial")) {
                                                            dStr = (dStr + (", "
                                                                        + (EstSht.Range("e1_WallStatus").offset((i - 1), 2).Value + "ft above finished floor")));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    // Wall and Roof Panels
                                                    dStr = ("Wall panels: "
                                                                + (b.wPanelColor.ToLower() + (" " + b.wPanelShape.ToLower())));
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Roof panels: "
                                                                + (b.rPanelColor.ToLower() + (" " + b.rPanelShape.ToLower())));
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    // liner panels
                                                    if (b.LinerPanels) {
                                                        ("e1" != "None");
                                                        dStr = b.LinerPanels;
                                                        ("e1" + (" "
                                                                    + (EstSht.Range("e1_LinerPanels").offset(0, 3).Value.ToLower() + (" "
                                                                    + (EstSht.Range("e1_LinerPanels").offset(0, 2).Value.ToLower() + " endwall #1 liner panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.LinerPanels) {
                                                        ("s2" != "None");
                                                        dStr = b.LinerPanels;
                                                        ("s2" + (" "
                                                                    + (EstSht.Range("s2_LinerPanels").offset(0, 3).Value.ToLower() + (" "
                                                                    + (EstSht.Range("s2_LinerPanels").offset(0, 2).Value.ToLower() + " sidewall #2 liner panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.LinerPanels) {
                                                        ("e3" != "None");
                                                        dStr = b.LinerPanels;
                                                        ("e3" + (" "
                                                                    + (EstSht.Range("e3_LinerPanels").offset(0, 3).Value.ToLower() + (" "
                                                                    + (EstSht.Range("e3_LinerPanels").offset(0, 2).Value.ToLower() + " endwall #3 liner panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.LinerPanels) {
                                                        ("s4" != "None");
                                                        dStr = b.LinerPanels;
                                                        ("s4" + (" "
                                                                    + (EstSht.Range("s4_LinerPanels").offset(0, 3).Value.ToLower() + (" "
                                                                    + (EstSht.Range("s4_LinerPanels").offset(0, 2).Value.ToLower() + " sidewall #4 liner panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.LinerPanels) {
                                                        ("Roof" != "None");
                                                        dStr = b.LinerPanels;
                                                        ("Roof" + (" "
                                                                    + (EstSht.Range("Roof_LinerPanels").offset(0, 3).Value.ToLower() + (" "
                                                                    + (EstSht.Range("Roof_LinerPanels").offset(0, 2).Value.ToLower() + " roof liner panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    // '''trim
                                                    // FO trim
                                                    dStr = ("Framed opening trim: " + EstSht.Range("FO_tColor").Value.ToLower());
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    // base trim
                                                    if ((b.BaseTrim == true)) {
                                                        dStr = ("Base opening trim: " + EstSht.Range("Base_tColor").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    // rake, eave, corner trim
                                                    dStr = ("Rake trim: " + EstSht.Range("Rake_tColor").Value.ToLower());
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Eave trim: " + EstSht.Range("Eave_tColor").Value.ToLower());
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    dStr = ("Corner trim: " + EstSht.Range("OutsideCorner_tColor").Value.ToLower());
                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                    RowCount = (RowCount + 1);
                                                    // downspouts/gutters
                                                    if ((EstSht.Range("GutterAndDownspouts").Value == "Yes")) {
                                                        dStr = ("Downspouts: " + EstSht.Range("DownspoutColor").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                        dStr = ("Gutters: " + EstSht.Range("GutterColor").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    // wainscot
                                                    if (b.Wainscot) {
                                                        ("e1" != "None");
                                                        dStr = b.Wainscot;
                                                        ("e1" + (" endwall #1 wainscot, "
                                                                    + (EstSht.Range("e1_Wainscot").offset(0, 1).Value.ToLower() + (" "
                                                                    + (EstSht.Range("e1_Wainscot").offset(0, 2).Value.ToLower() + " panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.Wainscot) {
                                                        ("s2" != "None");
                                                        dStr = b.Wainscot;
                                                        ("s2" + (" sidewall #2 wainscot, "
                                                                    + (EstSht.Range("s2_Wainscot").offset(0, 1).Value.ToLower() + (" "
                                                                    + (EstSht.Range("s2_Wainscot").offset(0, 2).Value.ToLower() + " panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.Wainscot) {
                                                        ("e3" != "None");
                                                        dStr = b.Wainscot;
                                                        ("e3" + (" endwall #3 wainscot, "
                                                                    + (EstSht.Range("e3_Wainscot").offset(0, 1).Value.ToLower() + (" "
                                                                    + (EstSht.Range("e3_Wainscot").offset(0, 2).Value.ToLower() + " panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (b.Wainscot) {
                                                        ("s4" != "None");
                                                        dStr = b.Wainscot;
                                                        ("s4" + (" sidewall #4 wainscot, "
                                                                    + (EstSht.Range("s4_Wainscot").offset(0, 1).Value.ToLower() + (" "
                                                                    + (EstSht.Range("s4_Wainscot").offset(0, 2).Value.ToLower() + " panels")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if ((EstSht.Range("Wainscot_tColor").Value != "")) {
                                                        dStr = ("Wainscot trim: " + EstSht.Range("Wainscot_tColor").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    for (cell in Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))) {
                                                        if ((cell.offset(0, 1).Value != "")) {
                                                            TempDesc = (TempDesc
                                                                        + (cell.offset(0, 1).Value + " size personnel door"));
                                                            if ((cell.offset(0, 3).Value == "Yes")) {
                                                                TempDesc = (TempDesc + " with half glass");
                                                            }

                                                            if ((cell.offset(0, 4).Value != "No")) {
                                                                TempDesc = (TempDesc + (" with "
                                                                            + (cell.offset(0, 4).Value + " canopy")));
                                                            }

                                                            if ((cell.offset(0, 6).Value == "Yes")) {
                                                                TempDesc = (TempDesc + " with dead bolt");
                                                            }

                                                            TempDesc = TempDesc;
                                                            // loop through array for similar PO Doors
                                                            for (i = 0; (i <= 11); i++) {
                                                                if (!IsEmpty(PDoorTypes[i, 0])) {
                                                                    if ((TempDesc == PDoorTypes[i, 1])) {
                                                                        PDoorTypes[i, 0] = (PDoorTypes[i, 0] + 1);
                                                                        TempDesc = "";
                                                                    }

                                                                }
                                                                else if ((TempDesc != "")) {
                                                                    PDoorTypes[i, 0] = 1;
                                                                    PDoorTypes[i, 1] = TempDesc;
                                                                    TempDesc = "";
                                                                }

                                                            }

                                                        }

                                                    }

                                                    // add array values to dStr
                                                    for (i = 0; (i <= 11); i++) {
                                                        if (!IsEmpty(PDoorTypes[i, 0])) {
                                                            dStr = (dStr + ("("
                                                                        + (PDoorTypes[i, 0] + (") "
                                                                        + (PDoorTypes[i, 1] + ", ")))));
                                                        }

                                                    }

                                                    if ((dStr != "")) {
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    TempDesc = "";
                                                    dStr = "";
                                                    for (cell in Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))) {
                                                        if ((cell.offset(0, 1).Value != "")) {
                                                            TempDesc = (TempDesc
                                                                        + (cell.offset(0, 1).Value + ("' x "
                                                                        + (cell.offset(0, 2).Value + ("' "
                                                                        + (cell.offset(0, 4).Value.ToLower() + " overhead door"))))));
                                                            if ((cell.offset(0, 5).Value != "None")) {
                                                                TempDesc = (TempDesc + (" with "
                                                                            + (cell.offset(0, 5).Value.ToLower() + " insulation")));
                                                            }
                                                            else {
                                                                TempDesc = (TempDesc + " non-insulated");
                                                            }

                                                            switch (cell.offset(0, 6).Value) {
                                                                case "Manual":
                                                                    TempDesc = (TempDesc + " with manual operation");
                                                                    break;
                                                                case "Chain Hoisr":
                                                                    TempDesc = (TempDesc + " with chain hoist");
                                                                    break;
                                                                case "Electric Opener":
                                                                    TempDesc = (TempDesc + " with electric opener");
                                                                    break;
                                                            }

                                                            if ((cell.offset(0, 7).Value != "No")) {
                                                                TempDesc = (TempDesc + " with high lift");
                                                            }

                                                            if ((cell.offset(0, 8).Value != "None")) {
                                                                TempDesc = (TempDesc + (" with "
                                                                            + (cell.offset(0, 8).Value.ToLower() + " windows")));
                                                            }

                                                            TempDesc = TempDesc;
                                                        }

                                                        for (i = 0; (i <= 11); i++) {
                                                            if (!IsEmpty(OHDoorTypes[i, 0])) {
                                                                if ((TempDesc == OHDoorTypes[i, 1])) {
                                                                    OHDoorTypes[i, 0] = (OHDoorTypes[i, 0] + 1);
                                                                    TempDesc = "";
                                                                }

                                                            }
                                                            else if ((TempDesc != "")) {
                                                                OHDoorTypes[i, 0] = 1;
                                                                OHDoorTypes[i, 1] = TempDesc;
                                                                TempDesc = "";
                                                            }

                                                        }

                                                    }

                                                    // add array values to dStr
                                                    for (i = 0; (i <= 11); i++) {
                                                        if (!IsEmpty(OHDoorTypes[i, 0])) {
                                                            dStr = (dStr + ("("
                                                                        + (OHDoorTypes[i, 0] + (") " + OHDoorTypes[i, 1]))));
                                                        }

                                                    }

                                                    TempDesc = "";
                                                    if ((dStr != "")) {
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    for (cell in Range(EstSht.Range("WindowCell1"), EstSht.Range("WindowCell12"))) {
                                                        if ((cell.offset(0, 1).Value != "")) {
                                                            TempDesc = (TempDesc
                                                                        + (cell.offset(0, 1).Value + ("' x "
                                                                        + (cell.offset(0, 2).Value + """ window, "))));
                                                            TempDesc = TempDesc;
                                                        }

                                                        for (i = 0; (i <= 23); i++) {
                                                            if (!IsEmpty(WindowTypes[i, 0])) {
                                                                if ((TempDesc == WindowTypes[i, 1])) {
                                                                    WindowTypes[i, 0] = (WindowTypes[i, 0] + 1);
                                                                    TempDesc = "";
                                                                }

                                                            }
                                                            else if ((TempDesc != "")) {
                                                                WindowTypes[i, 0] = 1;
                                                                WindowTypes[i, 1] = TempDesc;
                                                                TempDesc = "";
                                                            }

                                                        }

                                                    }

                                                    // add array values to dStr
                                                    for (i = 0; (i <= 23); i++) {
                                                        if (!IsEmpty(WindowTypes[i, 0])) {
                                                            dStr = (dStr + ("("
                                                                        + (WindowTypes[i, 0] + (") " + WindowTypes[i, 1]))));
                                                        }

                                                    }

                                                    TempDesc = "";
                                                    if ((dStr != "")) {
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    for (cell in Range(EstSht.Range("MiscFOCell1"), EstSht.Range("MiscFOCell12"))) {
                                                        if ((cell.offset(0, 1).Value != "")) {
                                                            TempDesc = (TempDesc
                                                                        + (cell.offset(0, 1).Value + ("' x "
                                                                        + (cell.offset(0, 2).Value + """ Misc FO, "))));
                                                            if ((cell.offset(0, 4).Value != "None")) {
                                                                TempDesc = (TempDesc + (" with " + cell.offset(0, 4).Value.ToLower()));
                                                            }

                                                            if ((cell.offset(0, 5).Value != "None")) {
                                                                TempDesc = (TempDesc + (" with "
                                                                            + (cell.offset(0, 5).Value.ToLower() + " weather hood")));
                                                            }

                                                            TempDesc = TempDesc;
                                                        }

                                                        for (i = 0; (i <= 11); i++) {
                                                            if (!IsEmpty(FOTypes[i, 0])) {
                                                                if ((TempDesc == FOTypes[i, 1])) {
                                                                    FOTypes[i, 0] = (FOTypes[i, 0] + 1);
                                                                    TempDesc = "";
                                                                }

                                                            }
                                                            else if ((TempDesc != "")) {
                                                                FOTypes[i, 0] = 1;
                                                                FOTypes[i, 1] = TempDesc;
                                                                TempDesc = "";
                                                            }

                                                        }

                                                    }

                                                    // add array values to dStr
                                                    for (i = 0; (i <= 11); i++) {
                                                        if (!IsEmpty(FOTypes[i, 0])) {
                                                            dStr = (dStr + ("("
                                                                        + (FOTypes[i, 0] + (") " + FOTypes[i, 1]))));
                                                        }

                                                    }

                                                    TempDesc = "";
                                                    if ((dStr != "")) {
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (((EstSht.Range("WallInsulation").Value != "None")
                                                                && (EstSht.Range("WallInsulation").Value != ""))) {
                                                        dStr = ("Wall insulation: " + EstSht.Range("WallInsulation").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (((EstSht.Range("RoofInsulation").Value != "None")
                                                                && (EstSht.Range("RoofInsulation").Value != ""))) {
                                                        dStr = ("Roof insulation: " + EstSht.Range("RoofInsulation").Value.ToLower());
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    // ridge vents, wall panels, skilights
                                                    if (((EstSht.Range("RidgeVentQty").Value != 0)
                                                                && (EstSht.Range("RidgeVentQty").Value != ""))) {
                                                        dStr = ("("
                                                                    + (EstSht.Range("RidgeVentQty").Value + (") "
                                                                    + (EstSht.Range("RidgeVentType").Value.ToLower() + " ridge vent(s)"))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (((EstSht.Range("TranslucentWallPanelQty").Value != 0)
                                                                && (EstSht.Range("TranslucentWallPanelQty").Value != ""))) {
                                                        dStr = ("("
                                                                    + (EstSht.Range("TranslucentWallPanelQty").Value + (") "
                                                                    + (EstSht.Range("TranslucentWallPanelLength").Value + ("' translucent wall panel(s)" + "
")))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    if (((EstSht.Range("SkylightQty").Value != 0)
                                                                && (EstSht.Range("SkylightQty").Value != ""))) {
                                                        dStr = ("("
                                                                    + (EstSht.Range("SkylightQty").Value + (") "
                                                                    + (EstSht.Range("SkylightLength").Value + "' skylight(s)"))));
                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    // With...
                                                    // e1 overhang
                                                    if (EstSht.Range) {
                                                        ("e1_GableOverhang".Value != "");
                                                        dStr = (dStr + EstSht.Range);
                                                        ("Building_Width".Value + ("' wide x " + EstSht.Range));
                                                        ("e1_GableOverhang".Value + "' long gable overhang on endwall #1, ");
                                                        if (EstSht.Range) {
                                                            "e1_GableOverhangSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "e1_GableOverhangSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "e1_GableOverhangSoffit".offset(0, 2).Value) + (" endwall #1 soffit panels with "
                                                                        + (LCase(EstSht.Range, "e1_GableOverhangSoffit".offset(0, 4).Value) + " trim"))))));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (EstSht.Range) {
                                                        ("s2_EaveOverhang".Value != "");
                                                        dStr = (dStr + EstSht.Range);
                                                        ("s2_EaveOverhang".Value + ("' wide x " + EstSht.Range));
                                                        ("Building_Length".Value + "' long eave overhang on sidewall #2, ");
                                                        if (EstSht.Range) {
                                                            "s2_EaveOverhangSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "s2_EaveOverhangSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "s2_EaveOverhangSoffit".offset(0, 2).Value) + (" sidewall #2 soffit panels with "
                                                                        + (LCase(EstSht.Range, "s2_EaveOverhangSoffit".offset(0, 4).Value) + " trim, "))))));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (EstSht.Range) {
                                                        ("e3_GableOverhang".Value != "");
                                                        dStr = (dStr + EstSht.Range);
                                                        ("Building_Width".Value + ("' wide x " + EstSht.Range));
                                                        ("e3_GableOverhang".Value + "' long gable overhang on endwall #3, ");
                                                        if (EstSht.Range) {
                                                            "e3_GableOverhangSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "e3_GableOverhangSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "e3_GableOverhangSoffit".offset(0, 2).Value) + (" endwall #3 soffit panels with "
                                                                        + (LCase(EstSht.Range, "e3_GableOverhangSoffit".offset(0, 4).Value) + " trim, "))))));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (EstSht.Range) {
                                                        ("s4_EaveOverhang".Value != "");
                                                        dStr = (dStr + EstSht.Range);
                                                        ("s4_EaveOverhang".Value + ("' wide x " + EstSht.Range));
                                                        ("Building_Length".Value + "' long eave overhang on sidewall #4, ");
                                                        if (EstSht.Range) {
                                                            "s4_EaveOverhangSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "s4_EaveOverhangSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "s4_EaveOverhangSoffit".offset(0, 2).Value) + (" sidewall #4 soffit panels with "
                                                                        + (LCase(EstSht.Range, "s4_EaveOverhangSoffit".offset(0, 4).Value) + " trim, "))))));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (EstSht.Range) {
                                                        ("e1_GableExtension".Value != "");
                                                        dStr = (dStr + EstSht.Range);
                                                        ("Building_Width".Value + ("' wide x " + EstSht.Range));
                                                        ("e1_GableExtension".Value + "' long gable extension roof structure on endwall #1, ");
                                                        if (EstSht.Range) {
                                                            "e1_GableExtensionSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "e1_GableExtensionSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "e1_GableExtensionSoffit".offset(0, 2).Value) + (" endwall #1 soffit panels with "
                                                                        + (LCase(EstSht.Range, "e1_GableExtensionSoffit".offset(0, 4).Value) + " trim, "))))));
                                                        }

                                                        dCell.offset(RowCount, 0).Value = dStr;
                                                        RowCount = (RowCount + 1);
                                                    }

                                                    dStr = "";
                                                    if (EstSht.Range) {
                                                        ("s2_EaveExtension".Value != "");
                                                        if (EstSht.Range) {
                                                            "s2_EaveExtensionPitch".Value = "Match Roof";
                                                            dStr = (dStr + EstSht.Range);
                                                            ("s2_EaveExtension".Value + ("' wide x " + EstSht.Range));
                                                            ("Building_Length".Value + ("' long "
                                                                        + (b.rPitch + "/12 pitch eave extension roof structure on sidewall #2, ")));
                                                        }
                                                        else {
                                                            dStr = (dStr + EstSht.Range);
                                                            ("s2_EaveExtension".Value + ("' wide x " + EstSht.Range));
                                                            ("Building_Length".Value + ("' long " + EstSht.Range));
                                                            ("s2_EaveExtensionPitch".Value + "/12 pitch eave extension roof structure on sidewall #2, ");
                                                        }

                                                        if (EstSht.Range) {
                                                            "s2_EaveExtensionSoffit".Value = "Yes";
                                                            dStr = (dStr
                                                                        + (LCase(EstSht.Range, "s2_EaveExtensionSoffit".offset(0, 3).Value) + (" "
                                                                        + (LCase(EstSht.Range, "s2_EaveExtensionSoffit".offset(0, 2).Value) + (" sidewall #2 soffit panels with "
                                                                        + (LCase(EstSht.Range, "s2_EaveExtensionSoffit".offset(0, 4).Value) + " trim, "))))));
                                                            dCell.offset(RowCount, 0).Value = dStr;
                                                            RowCount = (RowCount + 1);
                                                        }

                                                        dStr = "";
                                                        if (EstSht.Range) {
                                                            ("e3_GableExtension".Value != "");
                                                            dStr = (dStr + EstSht.Range);
                                                            ("Building_Width".Value + ("' wide x " + EstSht.Range));
                                                            ("e3_GableExtension".Value + "' long Gable extension roof structure on endwall #3, ");
                                                            if (EstSht.Range) {
                                                                "e3_GableExtensionSoffit".Value = "Yes";
                                                                dStr = (dStr
                                                                            + (LCase(EstSht.Range, "e3_GableExtensionSoffit".offset(0, 3).Value) + (" "
                                                                            + (LCase(EstSht.Range, "e3_GableExtensionSoffit".offset(0, 2).Value) + (" endwall #3 soffit panels with "
                                                                            + (LCase(EstSht.Range, "e3_GableExtensionSoffit".offset(0, 4).Value) + " trim, "))))));
                                                                dCell.offset(RowCount, 0).Value = dStr;
                                                                RowCount = (RowCount + 1);
                                                            }

                                                            dStr = "";
                                                            if (EstSht.Range) {
                                                                ("s4_EaveExtension".Value != "");
                                                                if (EstSht.Range) {
                                                                    "s4_EaveExtensionPitch".Value = "Match Roof";
                                                                    dStr = (dStr + EstSht.Range);
                                                                    ("s4_EaveExtension".Value + ("' wide x " + EstSht.Range));
                                                                    ("Building_Length".Value + ("' long "
                                                                                + (b.rPitch + "/12 pitch eave extension roof structure on sidewall #4, ")));
                                                                }
                                                                else {
                                                                    dStr = (dStr + EstSht.Range);
                                                                    ("s4_EaveExtension".Value + ("' wide x " + EstSht.Range));
                                                                    ("Building_Length".Value + ("' long " + EstSht.Range));
                                                                    ("s4_EaveExtensionPitch".Value + "/12 pitch eave extension roof structure on sidewall #4, ");
                                                                }

                                                                if (EstSht.Range) {
                                                                    "s4_EaveExtensionSoffit".Value = "Yes";
                                                                    dStr = (dStr
                                                                                + (LCase(EstSht.Range, "s4_EaveExtensionSoffit".offset(0, 3).Value) + (" "
                                                                                + (LCase(EstSht.Range, "s4_EaveExtensionSoffit".offset(0, 2).Value) + (" sidewall #4 soffit panels with "
                                                                                + (LCase(EstSht.Range, "s4_EaveExtensionSoffit".offset(0, 4).Value) + " trim, "))))));
                                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                                    RowCount = (RowCount + 1);
                                                                }

                                                                dStr = "";
                                                                if ((b.s2e1ExtensionIntersection == true)) {
                                                                    dStr = "Sidewall #2 and endwall #1 extension intersections";
                                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                                    RowCount = (RowCount + 1);
                                                                }

                                                                if ((b.s2e3ExtensionIntersection == true)) {
                                                                    dStr = "Sidewall #2 and endwall #3 extension intersections";
                                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                                    RowCount = (RowCount + 1);
                                                                }

                                                                if ((b.s4e1ExtensionIntersection == true)) {
                                                                    dStr = "Sidewall #4 and endwall #1 extension intersections";
                                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                                    RowCount = (RowCount + 1);
                                                                }

                                                                if ((b.s4e3ExtensionIntersection == true)) {
                                                                    dStr = "Sidewall #4 and endwall #3 extension intersections";
                                                                    dCell.offset(RowCount, 0).Value = dStr;
                                                                    RowCount = (RowCount + 1);
                                                                }

                                                                dStr = Trim[dStr];
                                                                // remove last comma
                                                                // dStr = Left(dStr, Len(dStr) - 1)
                                                            }

                                                            DescriptionSht.Columns[1].AutoFit;
                                                            // DescriptionSht.Range("DescriptionCell").Value = dStr
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
