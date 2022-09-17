// TODO: Option Explicit ... Warning!!! not translated


private Enable() {
    Application.EnableEvents = true;
}

private CommandButton1_Click() {
    GeneralMod.MaterialsListCaller;
}

public NewBusinessName() {
    let CurName: string;
    let NewName: string;
    CurName = this.Range("BusinessName").Value;
    NewName = InputBox(("Enter the business name you'd like to use for this Estimate. (old name: "
                    + (CurName + ")")), "Name Change");
    if ((NewName == "")) {
        // do nothing
    }
    else {
        this.Unprotect;
        "WhiteTruckMafia";
        this.Range("BusinessName").Value = NewName;
        this.Protect;
        "WhiteTruckMafia";
    }

}

private Worksheet_Change(Target: Range) {
    let BayStart: Range;
    let BayEnd: Range;
    let Downspouts: Range;
    let PDoorStart: Range;
    let PDoorEnd: Range;
    let OHDoorStart: Range;
    let OHDoorEnd: Range;
    let WindowStart: Range;
    let WindowEnd: Range;
    let FOStart: Range;
    let FOEnd: Range;
    let OverhangTbl: Range;
    let ExtensionTbl: Range;
    let Overhangs: Range;
    let Extensions: Range;
    let cell: Range;
    let WallAvailability: boolean[,];
    let i: number;
    if ((Target.Count > 1)) {
        FullSheetCheck();
        return;
    }

    // ranges
    BayStart = EstSht.Range("Building_Height").offset(2, -1);
    BayEnd = BayStart.offset(12, 0);
    PDoorStart = EstSht.Range("pDoorCell1").offset(-1, 0);
    PDoorEnd = EstSht.Range("pDoorCell12");
    OHDoorStart = EstSht.Range("OHDoorCell1").offset(-1, 0);
    OHDoorEnd = EstSht.Range("OHDoorCell12");
    WindowStart = EstSht.Range("WindowCell1").offset(-1, 0);
    WindowEnd = EstSht.Range("WindowCell12");
    FOStart = EstSht.Range("MiscFOCell1").offset(-1, 0);
    FOEnd = EstSht.Range("MiscFOCell12");
    OverhangTbl = EstSht.Range("e1_GableOverhang").offset(-1, -1).Resize(5, 7);
    ExtensionTbl = EstSht.Range("e1_GableExtension").offset(-1, -1).Resize(5, 7);
    Overhangs = Range(EstSht.Range("e1_GableOverhang"), EstSht.Range("s4_EaveOverhang"));
    Extensions = Range(EstSht.Range("e1_GableExtension"), EstSht.Range("s4_EaveExtension"));
    // 'Wall Availability for Liners, Wainscot, FOs
    // Assume all available, change if not
    WallAvailability[1] = true;
    WallAvailability[2] = true;
    WallAvailability[3] = true;
    WallAvailability[4] = true;
    if ((this.Range("e1_WallStatus").Value != "Include")) {
        WallAvailability[1] = false;
    }

    if ((this.Range("s2_WallStatus").Value != "Include")) {
        WallAvailability[2] = false;
    }

    if ((this.Range("e3_WallStatus").Value != "Include")) {
        WallAvailability[3] = false;
    }

    if ((this.Range("s4_WallStatus").Value != "Include")) {
        WallAvailability[4] = false;
    }

    if (!(Intersect(Target, EstSht.Range("BayNum")) == null)) {
        // unprotect
        EstSht.Unprotect;
        "WhiteTruckMafia";
        if ((Target.Value == 0)) {
            if ((BayStart.offset(-1, 0).EntireRow.Hidden == false)) {
                BayStart.offset(-1, 0).EntireRow.Hidden = true;
            }
            else {
                if ((BayStart.offset(-1, 0).EntireRow.Hidden == true)) {
                    BayStart.offset(-1, 0).EntireRow.Hidden = false;
                }

                // change all bay lengths to 0
                Application.EnableEvents = false;
                EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0;
                Application.EnableEvents = true;
                switch (Target.Value) {
                    case "":
                        Target.Value = "0";
                        break;
                    case "0":
                        if ((EstSht.Range(BayStart, BayEnd).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = true;
                        }

                        break;
                    case "1":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(2, 1).EntireRow.Hidden = false;
                        break;
                    case "2":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(3, 1).EntireRow.Hidden = false;
                        break;
                    case "3":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(4, 1).EntireRow.Hidden = false;
                        break;
                    case "4":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(5, 1).EntireRow.Hidden = false;
                        break;
                    case "5":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(6, 1).EntireRow.Hidden = false;
                        break;
                    case "6":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(7, 1).EntireRow.Hidden = false;
                        break;
                    case "7":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(8, 1).EntireRow.Hidden = false;
                        break;
                    case "8":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(9, 1).EntireRow.Hidden = false;
                        break;
                    case "9":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(10, 1).EntireRow.Hidden = false;
                        break;
                    case "10":
                        if ((EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                            EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                        }

                        BayStart.Resize(11, 1).EntireRow.Hidden = false;
                        break;
                    case "11":
                        if ((BayEnd.EntireRow.Hidden == false)) {
                            BayEnd.EntireRow.Hidden = true;
                        }

                        BayStart.Resize(12, 1).EntireRow.Hidden = false;
                        break;
                    case "12":
                        EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = false;
                        break;
                }

                // protect
                EstSht.Protect;
                "WhiteTruckMafia";
            }

            // '''' Alter Walls
            if (!(Intersect(Target, EstSht.Range("AlterWalls")) == null)) {
                Application.ScreenUpdating = false;
                EstSht.Unprotect;
                "WhiteTruckMafia";
                switch (Target.Value) {
                    case "":
                        Target.Value = "No";
                        break;
                    case "No":
                        if ((EstSht.Range("Wainscot").Value != "Yes")) {
                            // do nothing
                            // remove row seperating alter walls and wainscot table
                        }
                        else if ((EstSht.Range("Wainscot").Value == "Yes")) {
                            if ((EstSht.Range("LinerPanels").Value == "No")) {
                                // resize section heading row
                                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                            }

                        }

                        // Lock Cells
                        EstSht.Range("e1_WallStatus").Resize(4, 3).Locked = true;
                        if (((EstSht.Range("Wainscot").Value == "Yes")
                                    && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                            EstSht.Range("Wainscot").offset(-3, 0).EntireRow.Hidden = false;
                        }

                        UpdatesEventsProtection(false);
                        Range(this.Range("e1_WallStatus"), this.Range("s4_WallStatus")).Value = "Include";
                        Range(this.Range("e1_WallStatus"), this.Range("s4_WallStatus")).offset(0, 2).Value = 0;
                        this.Range("e1_Expandable").Value = "No";
                        this.Range("e3_Expandable").Value = "No";
                        WallAvailability[1] = true;
                        WallAvailability[2] = true;
                        WallAvailability[3] = true;
                        WallAvailability[4] = true;
                        AlterAvailableWalls(WallAvailability);
                        UpdatesEventsProtection(true);
                        break;
                    case "Yes":
                        EstSht.Range("e1_WallStatus").Resize(4, 3).Locked = false;
                        if ((EstSht.Range("Wainscot").Value == "Yes")) {
                            if ((EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden == true)) {
                                EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = false;
                            }

                            // '''''''''''''''''''''''''''''' Set Defaults
                            //         With EstSht
                            //             .Range("e1_WallStatus").Value = "Include"
                            //             .Range("s2_WallStatus").Value = "Include"
                            //             .Range("e3_WallStatus").Value = "Include"
                            //             .Range("s4_WallStatus").Value = "Include"
                            //             .Range("e1_Expandable").Value = "No"
                            //             .Range("e3_Expandable").Value = "No"
                            //         End With
                        }

                        // protect
                        EstSht.Protect;
                        "WhiteTruckMafia";
                        Application.ScreenUpdating = true;
                        break;
                }

            }

            // '''' Wall Status Changes
            // With...
            if (!(Intersect(Target, Range(this.Range, "e1_WallStatus", this.Range, "s4_WallStatus")) == null)) {
                UpdatesEventsProtection(false);
                switch (Target.Address) {
                    case this.Range:
                        "e1_WallStatus".Address;
                        if (this.Range) {
                            "e1_WallStatus".Value = "Partial";
                            this.Range;
                            false.Range("e1_WallStatus").offset(0, 2).Value = 0;
                            "e1_WallStatus".offset(0, 2).Locked = 0;
                        }
                        else {
                            if (this.Range) {
                                "e1_WallStatus".Value = "";
                                this.Range;
                                "e1_WallStatus".Value = "Include";
                                WallAvailability[1] = true;
                            }

                            this.Range;
                            true.Range("e1_WallStatus").offset(0, 2).Value = "N/A";
                            "e1_WallStatus".offset(0, 2).Locked = "N/A";
                        }

                        break;
                    case this.Range:
                        "s2_WallStatus".Address;
                        if (this.Range) {
                            "s2_WallStatus".Value = "Partial";
                            this.Range;
                            false.Range("s2_WallStatus").offset(0, 2).Value = 0;
                            "s2_WallStatus".offset(0, 2).Locked = 0;
                        }
                        else {
                            if (this.Range) {
                                "s2_WallStatus".Value = "";
                                this.Range;
                                "s2_WallStatus".Value = "Include";
                                WallAvailability[2] = true;
                            }

                            this.Range;
                            true.Range("s2_WallStatus").offset(0, 2).Value = "N/A";
                            "s2_WallStatus".offset(0, 2).Locked = "N/A";
                        }

                        break;
                    case this.Range:
                        "e3_WallStatus".Address;
                        if (this.Range) {
                            "e3_WallStatus".Value = "Partial";
                            this.Range;
                            false.Range("e3_WallStatus").offset(0, 2).Value = 0;
                            "e3_WallStatus".offset(0, 2).Locked = 0;
                        }
                        else {
                            if (this.Range) {
                                "e3_WallStatus".Value = "";
                                this.Range;
                                "e3_WallStatus".Value = "Include";
                                WallAvailability[3] = true;
                            }

                            this.Range;
                            true.Range("e3_WallStatus").offset(0, 2).Value = "N/A";
                            "e3_WallStatus".offset(0, 2).Locked = "N/A";
                        }

                        break;
                    case this.Range:
                        "s4_WallStatus".Address;
                        if (this.Range) {
                            "s4_WallStatus".Value = "Partial";
                            this.Range;
                            false.Range("s4_WallStatus").offset(0, 2).Value = 0;
                            "s4_WallStatus".offset(0, 2).Locked = 0;
                        }
                        else {
                            if (this.Range) {
                                "s4_WallStatus".Value = "";
                                this.Range;
                                "s4_WallStatus".Value = "Include";
                                WallAvailability[4] = true;
                            }

                            this.Range;
                            true.Range("s4_WallStatus").offset(0, 2).Value = "N/A";
                            "s4_WallStatus".offset(0, 2).Locked = "N/A";
                        }

                        break;
                }

                // update wall availability
                AlterAvailableWalls(WallAvailability);
                UpdatesEventsProtection(true);
            }

            // '''' Liner Panels Section
            if (!(Intersect(Target, EstSht.Range("LinerPanels")) == null)) {
                // unprotect
                EstSht.Unprotect;
                "WhiteTruckMafia";
                switch (Target.Value) {
                    case "":
                        Target.Value = "No";
                        break;
                    case "No":
                        EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                        // hide liner panels section
                        if ((Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == false)) {
                            Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = true;
                        }

                        // unhide row above wainscot table if needed
                        if (((EstSht.Range("Wainscot").Value == "Yes")
                                    && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                            EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = false;
                        }

                        UpdatesEventsProtection(false);
                        Range(this.Range("e1_LinerPanels"), this.Range("Roof_LinerPanels")).Value = "None";
                        Range(this.Range("e1_LinerPanels"), this.Range("Roof_LinerPanels")).offset(0, 1).Resize(5, 4).Value = "";
                        UpdatesEventsProtection(true);
                        break;
                    case "Yes":
                        if ((Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == true)) {
                            Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = false;
                        }

                        if ((Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == true)) {
                            Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = false;
                        }

                        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.AutoFit;
                        break;
                }

                // protect
                EstSht.Protect;
                "WhiteTruckMafia";
            }

            // '''' Liner Panels Options Change
            if (!(Intersect(Target, Range(EstSht.Range("e1_LinerPanels"), EstSht.Range("Roof_LinerPanels"))) == null)) {
                UpdatesEventsProtection(false);
                if ((Target.Value == "")) {
                    Target.Value = "None";
                }

                if ((Target.Value == "None")) {
                    Target.offset(0, 1).Value = "";
                    Target.offset(0, 2).Value = "";
                    Target.offset(0, 3).Value = "";
                }

                UpdatesEventsProtection(true);
            }

            // '''' Wainscot Section
            if (!(Intersect(Target, EstSht.Range("Wainscot")) == null)) {
                // unprotect
                EstSht.Unprotect;
                "WhiteTruckMafia";
                switch (Target.Value) {
                    case "":
                        Target.Value = "No";
                        break;
                    case "No":
                        EstSht.Range("e1_Wainscot").Resize(4, 3).Locked = true;
                        EstSht.Range("Wainscot_tColor").Locked = true;
                        if ((EstSht.Range("AlterWalls").Value != "Yes")) {
                            // do nothing
                            if ((EstSht.Range("LinerPanels").Value == "No")) {
                                // resize section heading row
                                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                            }

                        }
                        else if ((EstSht.Range("AlterWalls").Value == "Yes")) {
                            // Do nothing
                            // hide row seperating walls and wainscot table
                            if ((EstSht.Range("LinerPanels").Value == "No")) {
                                // resize section heading row
                                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                            }

                        }

                        // '''''''''''''''''''''''''''''' reset rows
                        UpdatesEventsProtection(false);
                        // With...
                        "None".Range("Wainscot_tColor").Value = "";
                        "".Range("e3_Wainscot").offset(0, 1).Resize(1, 3).Value = "";
                        "".Range("s2_Wainscot").offset(0, 1).Resize(1, 3).Value = "";
                        EstSht.Range("e1_Wainscot").offset(0, 1).Resize(1, 3).Value = "";
                        Range(., Range("e1_Wainscot"), ., Range("s4_Wainscot")).Value = "";
                        UpdatesEventsProtection(true);
                        break;
                    case "Yes":
                        EstSht.Range("e1_Wainscot").Resize(4, 3).Locked = false;
                        EstSht.Range("Wainscot_tColor").Locked = false;
                        if (((EstSht.Range("LinerPanels").Value != "Yes")
                                    && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                            EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = false;
                        }

                        UpdatesEventsProtection(false);
                        // With...
                        "".Range("s4_Wainscot").offset(0, 1).Resize(1, 3).Value = "None";
                        "".Range("e3_Wainscot").offset(0, 1).Resize(1, 3).Value = "None";
                        "".Range("s2_Wainscot").offset(0, 1).Resize(1, 3).Value = "None";
                        "None".Range("e1_Wainscot").offset(0, 1).Resize(1, 3).Value = "None";
                        "None".Range("s4_Wainscot").Value = "None";
                        "None".Range("e3_Wainscot").Value = "None";
                        "None".Range("s2_Wainscot").Value = "None";
                        EstSht.Range("e1_Wainscot").Value = "None";
                        UpdatesEventsProtection(true);
                        break;
                }

                // protect
                EstSht.Protect;
                "WhiteTruckMafia";
                Application.ScreenUpdating = true;
                Application.EnableEvents = true;
            }

            // '''' Wainscot Options Change
            if (!(Intersect(Target, Range(EstSht.Range("e1_Wainscot"), EstSht.Range("s4_Wainscot"))) == null)) {
                UpdatesEventsProtection(false);
                if ((Target.Value == "")) {
                    Target.Value = "None";
                }

                if ((Target.Value == "None")) {
                    Target.offset(0, 1).Value = "";
                    Target.offset(0, 2).Value = "";
                }

                UpdatesEventsProtection(true);
            }

            // ''' Gutter & Downspouts
            if (!(Intersect(Target, EstSht.Range("GutterAndDownspouts")) == null)) {
                // check if yes or no
                if ((Target.Value == "")) {
                    Target.Value = "No";
                }

                // ''' personnel door number
                if (!(Intersect(Target, EstSht.Range("PDoorNum")) == null)) {
                    UpdatesEventsProtection(false);
                    // hide row under quantity box when quantity is 0
                    if (((Target.Value == 0)
                                || (Target.Value == ""))) {
                        Target.Value = 0;
                        // reset table to blank
                        this.Range("pDoorCell1").offset(0, 1).Resize(12, 7).Value = "";
                        if ((PDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden == false)) {
                            PDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = true;
                        }
                        else {
                            if ((PDoorStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                PDoorStart.offset(-1, 0).EntireRow.Hidden = false;
                            }

                            for (i = 0; (i <= 12); i++) {
                                if ((i <= Target.Value)) {
                                    if ((PDoorStart.offset(i, 0).EntireRow.Hidden == true)) {
                                        PDoorStart.offset(i, 0).EntireRow.Hidden = false;
                                    }
                                    else {
                                        if ((PDoorStart.offset(i, 0).EntireRow.Hidden == false)) {
                                            PDoorStart.offset(i, 0).EntireRow.Hidden = true;
                                        }

                                        i;
                                    }

                                    UpdatesEventsProtection(true);
                                }

                                // ''' OH door
                                if (!(Intersect(Target, EstSht.Range("OHDoorNum")) == null)) {
                                    UpdatesEventsProtection(false);
                                    // hide row under quantity box when quantity is 0
                                    if (((Target.Value == 0)
                                                || (Target.Value == ""))) {
                                        Target.Value = 0;
                                        // reset table to blank
                                        this.Range("OHDoorCell1").offset(0, 1).Resize(12, 9).Value = "";
                                        if ((OHDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden == false)) {
                                            OHDoorStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = true;
                                        }
                                        else {
                                            if ((OHDoorStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                OHDoorStart.offset(-1, 0).EntireRow.Hidden = false;
                                            }

                                            for (i = 0; (i <= 12); i++) {
                                                if ((i <= Target.Value)) {
                                                    if ((OHDoorStart.offset(i, 0).EntireRow.Hidden == true)) {
                                                        OHDoorStart.offset(i, 0).EntireRow.Hidden = false;
                                                    }
                                                    else {
                                                        if ((OHDoorStart.offset(i, 0).EntireRow.Hidden == false)) {
                                                            OHDoorStart.offset(i, 0).EntireRow.Hidden = true;
                                                        }

                                                        i;
                                                    }

                                                    UpdatesEventsProtection(true);
                                                }

                                                // ''' Windows
                                                if (!(Intersect(Target, EstSht.Range("WindowNum")) == null)) {
                                                    UpdatesEventsProtection(false);
                                                    // hide row under quantity box when quantity is 0
                                                    if (((Target.Value == 0)
                                                                || (Target.Value == ""))) {
                                                        Target.Value = 0;
                                                        // reset table to blank
                                                        this.Range("WindowCell1").offset(0, 1).Resize(24, 3).Value = "";
                                                        this.Range("WindowCell1").offset(0, 6).Resize(24, 1).Value = "";
                                                        if ((WindowStart.offset(-1, 0).Resize(26, 1).EntireRow.Hidden == false)) {
                                                            WindowStart.offset(-1, 0).Resize(26, 1).EntireRow.Hidden = true;
                                                        }
                                                        else {
                                                            if ((WindowStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                WindowStart.offset(-1, 0).EntireRow.Hidden = false;
                                                            }

                                                            for (i = 0; (i <= 24); i++) {
                                                                if ((i <= Target.Value)) {
                                                                    if ((WindowStart.offset(i, 0).EntireRow.Hidden == true)) {
                                                                        WindowStart.offset(i, 0).EntireRow.Hidden = false;
                                                                    }
                                                                    else {
                                                                        if ((WindowStart.offset(i, 0).EntireRow.Hidden == false)) {
                                                                            WindowStart.offset(i, 0).EntireRow.Hidden = true;
                                                                        }

                                                                        i;
                                                                    }

                                                                    UpdatesEventsProtection(true);
                                                                }

                                                                // '''Window Default Values
                                                                // '''increase top edge height so that windows can't be lower than building
                                                                if (!(Intersect(Target, this.Range("WindowCell1").offset(0, 2).Resize(24, 1)) == null)) {
                                                                    if ((Target.Value > 86)) {
                                                                        Target.offset(0, 3).Value = (Target.Value / 12);
                                                                    }

                                                                }

                                                                // '''increase top edge height so that MiscFOs can't be lower than building
                                                                if (!(Intersect(Target, this.Range("MiscFOCell1").offset(0, 2).Resize(12, 1)) == null)) {
                                                                    if ((Target.Value > (86 / 12))) {
                                                                        Target.offset(0, 5).Value = Target.Value;
                                                                    }

                                                                }

                                                                // if MiscFO is 'field located', only allow 7'2" jambs w/ stool only
                                                                if (!(Intersect(Target, this.Range("MiscFOCell1").offset(0, 3).Resize(12, 1)) == null)) {
                                                                    if ((Target.Value == "Field Locate")) {
                                                                        Target.offset(0, 7).Value = "7'2"" Jambs w/ Stool";
                                                                    }

                                                                }

                                                                // ''' Misc Framed Openings
                                                                if (!(Intersect(Target, EstSht.Range("MiscFONum")) == null)) {
                                                                    UpdatesEventsProtection(false);
                                                                    // hide row under quantity box when quantity is 0
                                                                    if (((Target.Value == 0)
                                                                                || (Target.Value == ""))) {
                                                                        Target.Value = 0;
                                                                        // reset table to blank
                                                                        this.Range("MiscFOCell1").offset(0, 1).Resize(12, 5).Value = "";
                                                                        this.Range("MiscFOCell1").offset(0, 8).Resize(12, 1).Value = "";
                                                                        this.Range("MiscFOCell1").offset(0, 10).Resize(12, 1).Value = "";
                                                                        if ((FOStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden == false)) {
                                                                            FOStart.offset(-1, 0).Resize(14, 1).EntireRow.Hidden = true;
                                                                        }
                                                                        else {
                                                                            if ((FOStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                FOStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                            }

                                                                            for (i = 0; (i <= 12); i++) {
                                                                                if ((i <= Target.Value)) {
                                                                                    if ((FOStart.offset(i, 0).EntireRow.Hidden == true)) {
                                                                                        FOStart.offset(i, 0).EntireRow.Hidden = false;
                                                                                    }
                                                                                    else {
                                                                                        if ((FOStart.offset(i, 0).EntireRow.Hidden == false)) {
                                                                                            FOStart.offset(i, 0).EntireRow.Hidden = true;
                                                                                        }

                                                                                        i;
                                                                                    }

                                                                                    UpdatesEventsProtection(true);
                                                                                }

                                                                                // ''' building length change
                                                                                if (!(Intersect(Target, EstSht.Range("Building_Length")) == null)) {
                                                                                    // unprotect
                                                                                    EstSht.Unprotect;
                                                                                    "WhiteTruckMafia";
                                                                                    if ((EstSht.Range("Building_Length").Value == "")) {
                                                                                        EstSht.Range("Building_Length").Value = 0;
                                                                                    }

                                                                                    // change all bay lengths to 0
                                                                                    Application.EnableEvents = false;
                                                                                    EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0;
                                                                                    Application.EnableEvents = true;
                                                                                    EstSht.Protect;
                                                                                    "WhiteTruckMafia";
                                                                                }

                                                                                // ''' bay length change
                                                                                if (!(Intersect(Target, EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1)) == null)) {
                                                                                    // unprotect
                                                                                    EstSht.Unprotect;
                                                                                    "WhiteTruckMafia";
                                                                                    Application.EnableEvents = false;
                                                                                    BayUpdate(EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1), EstSht.Range("Building_Length"), Intersect(Target, EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1)));
                                                                                    Application.EnableEvents = true;
                                                                                    EstSht.Protect;
                                                                                    "WhiteTruckMafia";
                                                                                }

                                                                                // ''' Trim Color Bulk Change
                                                                                if (!(Intersect(Target, EstSht.Range("All_tColors")) == null)) {
                                                                                    Application.EnableEvents = false;
                                                                                    if ((Target.Value == "")) {
                                                                                        Target.Value = "N/A";
                                                                                    }
                                                                                    else {
                                                                                        // change trim colors to match
                                                                                        // With...
                                                                                        "None".Range("DownspoutColor").Value = Target.Value;
                                                                                        Target.Value.Range("Base_tColor").Value = Target.Value;
                                                                                        Target.Value.Range("FO_tColor").Value = Target.Value;
                                                                                        Target.Value.Range("OutsideCorner_tColor").Value = Target.Value;
                                                                                        Target.Value.Range("Eave_tColor").Value = Target.Value;
                                                                                        EstSht.Range("Rake_tColor").Value = Target.Value;
                                                                                    }

                                                                                    Application.EnableEvents = true;
                                                                                }

                                                                                // ''' overhang table clear
                                                                                if (!(Intersect(Target, OverhangTbl) == null)) {
                                                                                    Application.EnableEvents = false;
                                                                                    for (cell in Overhangs) {
                                                                                        if ((cell.Row == Target.Row)) {
                                                                                            if (((cell.Value == "")
                                                                                                        || (cell.Value == 0))) {
                                                                                                // clear soffits
                                                                                                cell.offset(0, 1).Value = "";
                                                                                                cell.offset(0, 2).Value = "";
                                                                                                cell.offset(0, 3).Value = "";
                                                                                                cell.offset(0, 4).Value = "";
                                                                                                cell.offset(0, 5).Value = "";
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    Application.EnableEvents = true;
                                                                                }

                                                                                // extension table clear
                                                                                if (!(Intersect(Target, ExtensionTbl) == null)) {
                                                                                    Application.EnableEvents = false;
                                                                                    for (cell in Extensions) {
                                                                                        if ((cell.Row == Target.Row)) {
                                                                                            if (((cell.Value == "")
                                                                                                        || (cell.Value == 0))) {
                                                                                                // clear soffits
                                                                                                cell.offset(0, 1).Value = "";
                                                                                                cell.offset(0, 2).Value = "";
                                                                                                cell.offset(0, 3).Value = "";
                                                                                                cell.offset(0, 4).Value = "";
                                                                                                cell.offset(0, 5).Value = "";
                                                                                            }

                                                                                        }

                                                                                    }

                                                                                    Application.EnableEvents = true;
                                                                                }

                                                                                // Show/Hide Eave Extension Pitch and Set Intersection default values
                                                                                // With...
                                                                                // s2 eave extension
                                                                                if (!(Intersect(Target, EstSht.Range, "s2_EaveExtension") == null)) {
                                                                                    Application.ScreenUpdating = false;
                                                                                    EstSht.Unprotect;
                                                                                    "WhiteTruckMafia";
                                                                                    if (EstSht.Range) {
                                                                                        "s2_EaveExtension".Value = "";
                                                                                        if (EstSht.Range) {
                                                                                            "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = false;
                                                                                            EstSht.Range;
                                                                                            "N/A".Range("s2e3_Intersection").Value = "N/A";
                                                                                            true.Range("s2e1_Intersection").Value = "N/A";
                                                                                            "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "N/A";
                                                                                        }
                                                                                        else {
                                                                                            // If previously hidden, unhide and set default option to include intersection
                                                                                            if (EstSht.Range) {
                                                                                                "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = true;
                                                                                                EstSht.Range;
                                                                                                false.Range("s2_EaveExtensionPitch").Value = "Match Roof";
                                                                                                "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "Match Roof";
                                                                                                if (EstSht.Range) {
                                                                                                    ("e1_GableExtension".Value != "");
                                                                                                    EstSht.Range;
                                                                                                    "s2e1_Intersection".Value = "Include";
                                                                                                    if (EstSht.Range) {
                                                                                                        ("e3_GableExtension".Value != "");
                                                                                                        EstSht.Range;
                                                                                                        "s2e3_Intersection".Value = "Include";
                                                                                                    }

                                                                                                }

                                                                                                EstSht.Protect;
                                                                                                "WhiteTruckMafia";
                                                                                                Application.ScreenUpdating = true;
                                                                                            }

                                                                                            // s4 eave extension
                                                                                            if (!(Intersect(Target, EstSht.Range, "s4_EaveExtension") == null)) {
                                                                                                Application.ScreenUpdating = false;
                                                                                                EstSht.Unprotect;
                                                                                                "WhiteTruckMafia";
                                                                                                if (EstSht.Range) {
                                                                                                    "s4_EaveExtension".Value = "";
                                                                                                    if (EstSht.Range) {
                                                                                                        "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = false;
                                                                                                        EstSht.Range;
                                                                                                        "N/A".Range("s4e3_Intersection").Value = "N/A";
                                                                                                        true.Range("s4e1_Intersection").Value = "N/A";
                                                                                                        "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "N/A";
                                                                                                    }
                                                                                                    else {
                                                                                                        // If previously hidden, unhide and set default option to include intersection
                                                                                                        if (EstSht.Range) {
                                                                                                            "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = true;
                                                                                                            EstSht.Range;
                                                                                                            false.Range("s4_EaveExtensionPitch").Value = "Match Roof";
                                                                                                            "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "Match Roof";
                                                                                                            if (EstSht.Range) {
                                                                                                                ("e1_GableExtension".Value != "");
                                                                                                                EstSht.Range;
                                                                                                                "s4e1_Intersection".Value = "Include";
                                                                                                                if (EstSht.Range) {
                                                                                                                    ("e3_GableExtension".Value != "");
                                                                                                                    EstSht.Range;
                                                                                                                    "s4e3_Intersection".Value = "Include";
                                                                                                                }

                                                                                                            }

                                                                                                            EstSht.Protect;
                                                                                                            "WhiteTruckMafia";
                                                                                                            Application.ScreenUpdating = true;
                                                                                                        }

                                                                                                        // e1 gable extension Intersection Option
                                                                                                        if (!(Intersect(Target, EstSht.Range, "e1_GableExtension") == null)) {
                                                                                                            UpdatesEventsProtection(false);
                                                                                                            if (EstSht.Range) {
                                                                                                                "e1_GableExtension".Value = "";
                                                                                                                EstSht.Range;
                                                                                                                "N/A".Range("s4e1_Intersection").MergeArea.Locked = true;
                                                                                                                true.Range("s4e1_Intersection").Value = true;
                                                                                                                "N/A".Range("s2e1_Intersection").MergeArea.Locked = true;
                                                                                                                "s2e1_Intersection".Value = true;
                                                                                                            }
                                                                                                            else {
                                                                                                                // intersection for s2 e1
                                                                                                                if (EstSht.Range) {
                                                                                                                    "s2e1_Intersection".Value = "N/A";
                                                                                                                    EstSht.Range;
                                                                                                                    "s2e1_Intersection".MergeArea.Locked = false;
                                                                                                                    if (EstSht.Range) {
                                                                                                                        ("s2_EaveExtension".Value != "");
                                                                                                                        EstSht.Range;
                                                                                                                        "s2e1_Intersection".Value = "Include";
                                                                                                                    }

                                                                                                                    // intersection for s4 e1
                                                                                                                    if (EstSht.Range) {
                                                                                                                        "s4e1_Intersection".Value = "N/A";
                                                                                                                        EstSht.Range;
                                                                                                                        "s4e1_Intersection".MergeArea.Locked = false;
                                                                                                                        if (EstSht.Range) {
                                                                                                                            ("s4_EaveExtension".Value != "");
                                                                                                                            EstSht.Range;
                                                                                                                            "s4e1_Intersection".Value = "Include";
                                                                                                                        }

                                                                                                                    }

                                                                                                                    UpdatesEventsProtection(true);
                                                                                                                }

                                                                                                                // e3 gable extension Intersection Option
                                                                                                                if (!(Intersect(Target, EstSht.Range, "e3_GableExtension") == null)) {
                                                                                                                    UpdatesEventsProtection(false);
                                                                                                                    if (EstSht.Range) {
                                                                                                                        "e3_GableExtension".Value = "";
                                                                                                                        EstSht.Range;
                                                                                                                        "N/A".Range("s4e3_Intersection").MergeArea.Locked = true;
                                                                                                                        true.Range("s4e3_Intersection").Value = true;
                                                                                                                        "N/A".Range("s2e3_Intersection").MergeArea.Locked = true;
                                                                                                                        "s2e3_Intersection".Value = true;
                                                                                                                    }
                                                                                                                    else {
                                                                                                                        // intersection for s2 e3
                                                                                                                        if (EstSht.Range) {
                                                                                                                            "s2e3_Intersection".Value = "N/A";
                                                                                                                            EstSht.Range;
                                                                                                                            false.Range("s2e3_Intersection").Value = "Include";
                                                                                                                            "s2e3_Intersection".MergeArea.Locked = "Include";
                                                                                                                        }

                                                                                                                        // intersection for s4 e3
                                                                                                                        if (EstSht.Range) {
                                                                                                                            "s4e3_Intersection".Value = "N/A";
                                                                                                                            EstSht.Range;
                                                                                                                            false.Range("s4e3_Intersection").Value = "Include";
                                                                                                                            "s4e3_Intersection".MergeArea.Locked = "Include";
                                                                                                                        }

                                                                                                                    }

                                                                                                                    UpdatesEventsProtection(true);
                                                                                                                }

                                                                                                            }

                                                                                                            // ''' Roof/Wall Panel Shape Change - Disable Translucent Wall Panels and Skylights
                                                                                                            // With...
                                                                                                            if (!(Intersect(Target, Range(EstSht.Range, "Wall_pShape", EstSht.Range, "Roof_pShape")) == null)) {
                                                                                                                EstSht.Unprotect;
                                                                                                                "WhiteTruckMafia";
                                                                                                                Application.EnableEvents = false;
                                                                                                                if (EstSht.Range) {
                                                                                                                    "Wall_pShape".Value = ("M-Loc" | EstSht.Range);
                                                                                                                    "Roof_pShape".Value = "M-Loc";
                                                                                                                    EstSht.Range;
                                                                                                                    "".Range("SkylightLength").Value = "";
                                                                                                                    "".Range("TranslucentWallPanelLength").Value = "";
                                                                                                                    "".Range("SkylightQty").Value = "";
                                                                                                                    "TranslucentWallPanelQty".Value = "";
                                                                                                                    if (EstSht.Range) {
                                                                                                                        EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                        EstSht.Range;
                                                                                                                        "SkylightQty";
                                                                                                                        EstSht.EntireRow.Hidden = false;
                                                                                                                        EstSht.Range;
                                                                                                                        EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                        EstSht.Range;
                                                                                                                        "SkylightQty";
                                                                                                                        EstSht.EntireRow.Hidden = true;
                                                                                                                    }
                                                                                                                    else if (EstSht.Range) {
                                                                                                                        (("Wall_pShape".Value != "M-Loc")
                                                                                                                                    & EstSht.Range);
                                                                                                                        ("Roof_pShape".Value != "M-Loc");
                                                                                                                        // unhide rows if needed
                                                                                                                        if (EstSht.Range) {
                                                                                                                            EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                            EstSht.Range;
                                                                                                                            "SkylightQty";
                                                                                                                            EstSht.EntireRow.Hidden = true;
                                                                                                                            EstSht.Range;
                                                                                                                            EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                            EstSht.Range;
                                                                                                                            "SkylightQty";
                                                                                                                            EstSht.EntireRow.Hidden = false;
                                                                                                                        }

                                                                                                                        EstSht.Protect;
                                                                                                                        "WhiteTruckMafia";
                                                                                                                        Application.EnableEvents = true;
                                                                                                                    }

                                                                                                                }

                                                                                                                // '''''''''' only allow galvalume for panel color selection when prime acrylic galvalume panel is select
                                                                                                                // With...
                                                                                                                // wall, roof panels
                                                                                                                if (!(Intersect(Target, EstSht.Range, "Wall_pType") == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "Wall_pType", EstSht.Range, "Wall_Color");
                                                                                                                }
                                                                                                                else if (!(Intersect(Target, EstSht.Range, "Roof_pType") == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "Roof_pType", EstSht.Range, "Roof_Color");
                                                                                                                }

                                                                                                                // liner panels
                                                                                                                if (!(Intersect(Target, EstSht.Range, "e1_LinerPanels".offset(0, 2)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e1_LinerPanels".offset(0, 2), EstSht.Range, "e1_LinerPanels".offset(0, 3));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s2_LinerPanels".offset(0, 2)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s2_LinerPanels".offset(0, 2), EstSht.Range, "s2_LinerPanels".offset(0, 3));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "e3_LinerPanels".offset(0, 2)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e3_LinerPanels".offset(0, 2), EstSht.Range, "e3_LinerPanels".offset(0, 3));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s4_LinerPanels".offset(0, 2)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s4_LinerPanels".offset(0, 2), EstSht.Range, "s4_LinerPanels".offset(0, 3));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "Roof_LinerPanels".offset(0, 2)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "Roof_LinerPanels".offset(0, 2), EstSht.Range, "Roof_LinerPanels".offset(0, 3));
                                                                                                                }

                                                                                                                // wainscot
                                                                                                                if (!(Intersect(Target, EstSht.Range, "e1_Wainscot".offset(0, 1)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e1_Wainscot".offset(0, 1), EstSht.Range, "e1_Wainscot".offset(0, 2));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s2_Wainscot".offset(0, 1)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s2_Wainscot".offset(0, 1), EstSht.Range, "s2_Wainscot".offset(0, 2));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "e3_Wainscot".offset(0, 1)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e3_Wainscot".offset(0, 1), EstSht.Range, "e3_Wainscot".offset(0, 2));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s4_Wainscot".offset(0, 1)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s4_Wainscot".offset(0, 1), EstSht.Range, "s4_Wainscot".offset(0, 2));
                                                                                                                }

                                                                                                                // overhangs
                                                                                                                if (!(Intersect(Target, EstSht.Range, "e1_GableOverhang".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e1_GableOverhang".offset(0, 3), EstSht.Range, "e1_GableOverhang".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s2_EaveOverhang".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s2_EaveOverhang".offset(0, 3), EstSht.Range, "s2_EaveOverhang".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "e3_GableOverhang".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e3_GableOverhang".offset(0, 3), EstSht.Range, "e3_GableOverhang".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s4_EaveOverhang".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s4_EaveOverhang".offset(0, 3), EstSht.Range, "s4_EaveOverhang".offset(0, 4));
                                                                                                                }

                                                                                                                // extensions
                                                                                                                if (!(Intersect(Target, EstSht.Range, "e1_GableExtension".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e1_GableExtension".offset(0, 3), EstSht.Range, "e1_GableExtension".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s2_EaveExtension".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s2_EaveExtension".offset(0, 3), EstSht.Range, "s2_EaveExtension".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "e3_GableExtension".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "e3_GableExtension".offset(0, 3), EstSht.Range, "e3_GableExtension".offset(0, 4));
                                                                                                                }

                                                                                                                if (!(Intersect(Target, EstSht.Range, "s4_EaveExtension".offset(0, 3)) == null)) {
                                                                                                                    PanelColorOptionCheck(EstSht.Range, "s4_EaveExtension".offset(0, 3), EstSht.Range, "s4_EaveExtension".offset(0, 4));
                                                                                                                }

                                                                                                                // ''''''''''''''''''''''''''''''''''''''''''''''''' Framed Opening Option Changes '''''''''''''''''''''
                                                                                                                // With...
                                                                                                                // ''''''''''''''''''''''''''''''''''''''''' Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                if (!(Intersect(Target, Range(EstSht.Range, "pDoorCell1".offset(0, 1), EstSht.Range, "pDoorCell12".offset(0, 1))) == null)) {
                                                                                                                    EstSht.Unprotect;
                                                                                                                    "WhiteTruckMafia";
                                                                                                                    Application.EnableEvents = false;
                                                                                                                    if ((Target.Value == "4070")) {
                                                                                                                        // Remove half glass option
                                                                                                                        // With...
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                        true.ShowInput = true;
                                                                                                                        true.InCellDropdown = true;
                                                                                                                        "No".IgnoreBlank = true;
                                                                                                                        // Remove dead bolt option
                                                                                                                        // With...
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                        true.ShowInput = true;
                                                                                                                        true.InCellDropdown = true;
                                                                                                                        "No".IgnoreBlank = true;
                                                                                                                        // Set Default Values for 4070
                                                                                                                        Target.offset(0, 2).Value = "No";
                                                                                                                        Target.offset(0, 5).Value = "No";
                                                                                                                        if ((Target.offset(0, 3).Value == "")) {
                                                                                                                            Target.offset(0, 3).Value = "No";
                                                                                                                        }

                                                                                                                        if ((Target.offset(0, 4).Value == "")) {
                                                                                                                            Target.offset(0, 4).Value = "8.25";
                                                                                                                        }
                                                                                                                        else if (((Target.Value == "3070")
                                                                                                                                    || (Target.Value == ""))) {
                                                                                                                            // restore half glass, deadbolt
                                                                                                                            // With...
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                            true.ShowInput = true;
                                                                                                                            true.InCellDropdown = true;
                                                                                                                            "Yes,No".IgnoreBlank = true;
                                                                                                                            if ((Target.offset(0, 2).Validation.Value == false)) {
                                                                                                                                Target.offset(0, 2).Value = "";
                                                                                                                            }

                                                                                                                            // With...
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                            true.ShowInput = true;
                                                                                                                            true.InCellDropdown = true;
                                                                                                                            "Yes,No".IgnoreBlank = true;
                                                                                                                            // Set Default Values for 3070 and reset for blank
                                                                                                                            if ((Target.Value == "3070")) {
                                                                                                                                if ((Target.offset(0, 2).Value == "")) {
                                                                                                                                    Target.offset(0, 2).Value = "No";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                    Target.offset(0, 3).Value = "No";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                    Target.offset(0, 4).Value = 8.25;
                                                                                                                                }

                                                                                                                                // only if blank, keep selected values on change
                                                                                                                                if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                    Target.offset(0, 5).Value = "No";
                                                                                                                                }
                                                                                                                                else if ((Target.Value == "")) {
                                                                                                                                    Target.offset(0, 2).Value = "";
                                                                                                                                    Target.offset(0, 3).Value = "";
                                                                                                                                    Target.offset(0, 4).Value = "";
                                                                                                                                    Target.offset(0, 5).Value = "";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 5).Validation.Value == false)) {
                                                                                                                                    Target.offset(0, 5).Value = "";
                                                                                                                                }

                                                                                                                                EstSht.Protect;
                                                                                                                                "WhiteTruckMafia";
                                                                                                                                Application.EnableEvents = true;
                                                                                                                            }

                                                                                                                            // ''''''''''''''''''''''''''''''''''''''''' Overhead Doors '''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                            // '''''''''''''''Set Default Values''''''''''''''''
                                                                                                                            // If height or width is entered, set default values; "Type" change (RUD/Sectional) will fix validation as defined below
                                                                                                                            // Width
                                                                                                                            if (!(Intersect(Target, Range(EstSht.Range, "OHDoorCell1".offset(0, 1), EstSht.Range, "OHDoorCell12".offset(0, 1))) == null)) {
                                                                                                                                if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                    Target.offset(0, 3).Value = "Sectional";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                    Target.offset(0, 4).Value = "None";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                    Target.offset(0, 5).Value = "Manual";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 6).Value == "")) {
                                                                                                                                    Target.offset(0, 6).Value = "No";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 7).Value == "")) {
                                                                                                                                    Target.offset(0, 7).Value = "None";
                                                                                                                                }

                                                                                                                                if ((Target.offset(0, 8).Value == "")) {
                                                                                                                                    Target.offset(0, 8).Value = 0;
                                                                                                                                }

                                                                                                                                // Height
                                                                                                                                if (!(Intersect(Target, Range(EstSht.Range, "OHDoorCell1".offset(0, 2), EstSht.Range, "OHDoorCell12".offset(0, 2))) == null)) {
                                                                                                                                    if ((Target.offset(0, 2).Value == "")) {
                                                                                                                                        Target.offset(0, 2).Value = "Sectional";
                                                                                                                                    }

                                                                                                                                    if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                        Target.offset(0, 3).Value = "None";
                                                                                                                                    }

                                                                                                                                    if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                        Target.offset(0, 4).Value = "Manual";
                                                                                                                                    }

                                                                                                                                    if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                        Target.offset(0, 5).Value = "No";
                                                                                                                                    }

                                                                                                                                    if ((Target.offset(0, 6).Value == "")) {
                                                                                                                                        Target.offset(0, 6).Value = "None";
                                                                                                                                    }

                                                                                                                                    if ((Target.offset(0, 7).Value == "")) {
                                                                                                                                        Target.offset(0, 7).Value = 0;
                                                                                                                                    }

                                                                                                                                    if (!(Intersect(Target, Range(EstSht.Range, "OHDoorCell1".offset(0, 4), EstSht.Range, "OHDoorCell12".offset(0, 4))) == null)) {
                                                                                                                                        EstSht.Unprotect;
                                                                                                                                        "WhiteTruckMafia";
                                                                                                                                        Application.EnableEvents = false;
                                                                                                                                        if ((Target.Value == "RUD")) {
                                                                                                                                            // ''sizing options
                                                                                                                                            // width
                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("RUDWidth").Address.IgnoreBlank) = true;
                                                                                                                                            // clear if invalid
                                                                                                                                            if ((Target.offset(0, -3).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, -3).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("RUDHeight").Address.IgnoreBlank) = true;
                                                                                                                                            // clear if invalid
                                                                                                                                            if ((Target.offset(0, -2).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, -2).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            "None".IgnoreBlank = true;
                                                                                                                                            Target.offset(0, 1).Value = "None";
                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            "Chain Hoist".IgnoreBlank = true;
                                                                                                                                            Target.offset(0, 2).Value = "Chain Hoist";
                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            "No".IgnoreBlank = true;
                                                                                                                                            Target.offset(0, 3).Value = "No";
                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            "None".IgnoreBlank = true;
                                                                                                                                            Target.offset(0, 4).Value = "None";
                                                                                                                                        }
                                                                                                                                        else if (((Target.Value == "Sectional")
                                                                                                                                                    || (Target.Value == ""))) {
                                                                                                                                            // ''sizing options
                                                                                                                                            // width
                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("SectionalOHDoorWidth").Address.IgnoreBlank) = true;
                                                                                                                                            if ((Target.offset(0, -3).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, -3).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("SectionalOHDoorHeight").Address.IgnoreBlank) = true;
                                                                                                                                            if ((Target.offset(0, -2).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, -2).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("OHDoorInsulationOptions").Address.IgnoreBlank) = true;
                                                                                                                                            if ((Target.offset(0, 1).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, 1).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("OHDoorOperationOptions").Address.IgnoreBlank) = true;
                                                                                                                                            if ((Target.offset(0, 2).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, 2).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            "Yes,No".IgnoreBlank = true;
                                                                                                                                            if ((Target.offset(0, 3).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, 3).Value = "";
                                                                                                                                            }

                                                                                                                                            // With...
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                            true.ShowInput = true;
                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                            ("=Lists!" + ListSht.Range("OHDoorWindowOptions").Address.IgnoreBlank) = true;
                                                                                                                                            if ((Target.offset(0, 4).Validation.Value == false)) {
                                                                                                                                                Target.offset(0, 4).Value = "";
                                                                                                                                            }

                                                                                                                                            EstSht.Protect;
                                                                                                                                            "WhiteTruckMafia";
                                                                                                                                            Application.EnableEvents = true;
                                                                                                                                        }

                                                                                                                                        // ''''''''''''''Misc FOs '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                        // Exhaust Fans/Louvers default to "None"'
                                                                                                                                        // Width
                                                                                                                                        if (!(Intersect(Target, Range(EstSht.Range, "MiscFOCell1".offset(0, 1), EstSht.Range, "MiscFOCell12".offset(0, 1))) == null)) {
                                                                                                                                            if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                                Target.offset(0, 3).Value = "None";
                                                                                                                                            }

                                                                                                                                            if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                                Target.offset(0, 4).Value = "None";
                                                                                                                                            }

                                                                                                                                            if ((Target.offset(0, 6).Value == "")) {
                                                                                                                                                Target.offset(0, 6).Formula = "=86/12";
                                                                                                                                            }

                                                                                                                                            if ((Target.offset(0, 7).Value == "")) {
                                                                                                                                                Target.offset(0, 7).Value = 0;
                                                                                                                                            }

                                                                                                                                            // Height
                                                                                                                                            if (!(Intersect(Target, Range(EstSht.Range, "MiscFOCell1".offset(0, 2), EstSht.Range, "MiscFOCell12".offset(0, 2))) == null)) {
                                                                                                                                                if ((Target.offset(0, 2).Value == "")) {
                                                                                                                                                    Target.offset(0, 2).Value = "None";
                                                                                                                                                }

                                                                                                                                                if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                                    Target.offset(0, 3).Value = "None";
                                                                                                                                                }

                                                                                                                                                if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                                    Target.offset(0, 5).Formula = "=86/12";
                                                                                                                                                }

                                                                                                                                                if ((Target.offset(0, 6).Value == "")) {
                                                                                                                                                    Target.offset(0, 6).Value = 0;
                                                                                                                                                }

                                                                                                                                                // ''''''''''''''Windows '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                // Width
                                                                                                                                                if (!(Intersect(Target, Range(EstSht.Range, "WindowCell1".offset(0, 1), EstSht.Range, "WindowCell12".offset(0, 1))) == null)) {
                                                                                                                                                    if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                                        Target.offset(0, 4).Formula = "=86/12";
                                                                                                                                                    }

                                                                                                                                                    if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                                        Target.offset(0, 5).Value = 0;
                                                                                                                                                    }

                                                                                                                                                    // Height
                                                                                                                                                    if (!(Intersect(Target, Range(EstSht.Range, "MiscFOCell1".offset(0, 2), EstSht.Range, "MiscFOCell12".offset(0, 2))) == null)) {
                                                                                                                                                        if ((Target.offset(0, 3).Value == "")) {
                                                                                                                                                            Target.offset(0, 3).Formula = "=86/12";
                                                                                                                                                        }

                                                                                                                                                        if ((Target.offset(0, 4).Value == "")) {
                                                                                                                                                            Target.offset(0, 4).Value = 0;
                                                                                                                                                        }

                                                                                                                                                        // ''''''''''''''Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set Default Values '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                        // Width
                                                                                                                                                        if (!(Intersect(Target, Range(EstSht.Range, "pDoorCell1".offset(0, 1), EstSht.Range, "pDoorCell12".offset(0, 1))) == null)) {
                                                                                                                                                            if ((Target.offset(0, 6).Value == "")) {
                                                                                                                                                                Target.offset(0, 6).Value = 0;
                                                                                                                                                            }

                                                                                                                                                            // Height
                                                                                                                                                            if (!(Intersect(Target, Range(EstSht.Range, "MiscFOCell1".offset(0, 2), EstSht.Range, "MiscFOCell12".offset(0, 2))) == null)) {
                                                                                                                                                                if ((Target.offset(0, 5).Value == "")) {
                                                                                                                                                                    Target.offset(0, 5).Value = 0;
                                                                                                                                                                }

                                                                                                                                                                // ''''''''''''''Overhangs and Extensions''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                                // Overhangs Set Default Values
                                                                                                                                                                if (!(Intersect(Target, Range(EstSht.Range, "e1_GableOverhang", EstSht.Range, "s4_EaveOverhang")) == null)) {
                                                                                                                                                                    if ((Target.offset(0, 1).Value == "")) {
                                                                                                                                                                        Target.offset(0, 1).Value = "No";
                                                                                                                                                                    }

                                                                                                                                                                    // Extensions Set Default Values
                                                                                                                                                                    if (!(Intersect(Target, Range(EstSht.Range, "e1_GableExtension", EstSht.Range, "s4_EaveExtension")) == null)) {
                                                                                                                                                                        if ((Target.offset(0, 1).Value == "")) {
                                                                                                                                                                            Target.offset(0, 1).Value = "No";
                                                                                                                                                                        }

                                                                                                                                                                    }

                                                                                                                                                                }

                                                                                                                                                                Worksheet_SelectionChange((<Range>(Target)));
                                                                                                                                                                let BayStart: Range;
                                                                                                                                                                let BayEnd: Range;
                                                                                                                                                                // ranges
                                                                                                                                                                BayStart = EstSht.Range("A19");
                                                                                                                                                                BayEnd = EstSht.Range("A31");
                                                                                                                                                                // '' Check if bay length <> building length when moving to a new section
                                                                                                                                                                if ((Target.Row >= BayEnd.offset(1, 0).Row)) {
                                                                                                                                                                    // check if building length and # of bays <> 0
                                                                                                                                                                    if (((EstSht.Range("Building_Length").Value != 0)
                                                                                                                                                                                && (EstSht.Range("BayNum").Value != 0))) {
                                                                                                                                                                        // check that total bay length = building length
                                                                                                                                                                        if ((HiddenSht.Range("TotalBayLength").Value != EstSht.Range("Building_Length").Value)) {
                                                                                                                                                                            // alert that length totals not equal, move to first cell of bay length table
                                                                                                                                                                            MsgBox;
                                                                                                                                                                            "The total bay length does not match the building length! Please correct the data before proceeding.";
                                                                                                                                                                            vbExclamation;
                                                                                                                                                                            "Invalid Bay Lengths";
                                                                                                                                                                            BayStart.offset(1, 1).Select;
                                                                                                                                                                        }

                                                                                                                                                                    }

                                                                                                                                                                }

                                                                                                                                                            }

                                                                                                                                                            AlterAvailableWalls((<boolean>(WallAvailability[])));
                                                                                                                                                            let WallSelection: string;
                                                                                                                                                            let N: number;
                                                                                                                                                            UpdatesEventsProtection(false);
                                                                                                                                                            for (N = 1; (N <= 4); N++) {
                                                                                                                                                                if ((WallAvailability[N] == true)) {
                                                                                                                                                                    switch (N) {
                                                                                                                                                                        case 1:
                                                                                                                                                                            WallSelection = "Endwall 1";
                                                                                                                                                                            break;
                                                                                                                                                                        case 2:
                                                                                                                                                                            if ((WallSelection != "")) {
                                                                                                                                                                                WallSelection = (WallSelection + ("," + "Sidewall 2"));
                                                                                                                                                                            }
                                                                                                                                                                            else {
                                                                                                                                                                                WallSelection = "Sidewall 2";
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                        case 3:
                                                                                                                                                                            if ((WallSelection != "")) {
                                                                                                                                                                                WallSelection = (WallSelection + ("," + "Endwall 3"));
                                                                                                                                                                            }
                                                                                                                                                                            else {
                                                                                                                                                                                WallSelection = "Endwall 3";
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                        case 4:
                                                                                                                                                                            if ((WallSelection != "")) {
                                                                                                                                                                                WallSelection = (WallSelection + ("," + "Sidewall 4"));
                                                                                                                                                                            }
                                                                                                                                                                            else {
                                                                                                                                                                                WallSelection = "Sidewall 4";
                                                                                                                                                                            }

                                                                                                                                                                            break;
                                                                                                                                                                    }

                                                                                                                                                                }

                                                                                                                                                            }

                                                                                                                                                            // With...
                                                                                                                                                            if ((WallAvailability[1] == false)) {
                                                                                                                                                                if (this.Range) {
                                                                                                                                                                    "e1_WallStatus".Value = "Exclude";
                                                                                                                                                                    this.Range;
                                                                                                                                                                    true.Range("e1_LinerPanels").Value = "None";
                                                                                                                                                                    "".Range("e1_LinerPanels").Resize(1, 4).Locked = "None";
                                                                                                                                                                    "e1_LinerPanels".Resize(1, 4).Value = "None";
                                                                                                                                                                }
                                                                                                                                                                else {
                                                                                                                                                                    this.Range;
                                                                                                                                                                    "e1_LinerPanels".Resize(1, 4).Locked = false;
                                                                                                                                                                }

                                                                                                                                                                this.Range;
                                                                                                                                                                true.Range("e1_Wainscot").Value = "None";
                                                                                                                                                                "".Range("e1_Wainscot").Resize(1, 3).Locked = "None";
                                                                                                                                                                "e1_Wainscot".Resize(1, 3).Value = "None";
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                this.Range;
                                                                                                                                                                false.Range("e1_Wainscot").Resize(1, 3).Locked = false;
                                                                                                                                                                "e1_LinerPanels".Resize(1, 4).Locked = false;
                                                                                                                                                            }

                                                                                                                                                            if ((WallAvailability[2] == false)) {
                                                                                                                                                                if (this.Range) {
                                                                                                                                                                    "s2_WallStatus".Value = "Exclude";
                                                                                                                                                                    this.Range;
                                                                                                                                                                    true.Range("s2_LinerPanels").Value = "None";
                                                                                                                                                                    "".Range("s2_LinerPanels").Resize(1, 4).Locked = "None";
                                                                                                                                                                    "s2_LinerPanels".Resize(1, 4).Value = "None";
                                                                                                                                                                }

                                                                                                                                                                this.Range;
                                                                                                                                                                true.Range("s2_Wainscot").Value = "None";
                                                                                                                                                                "".Range("s2_Wainscot").Resize(1, 3).Locked = "None";
                                                                                                                                                                "s2_Wainscot".Resize(1, 3).Value = "None";
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                this.Range;
                                                                                                                                                                false.Range("s2_Wainscot").Resize(1, 3).Locked = false;
                                                                                                                                                                "s2_LinerPanels".Resize(1, 4).Locked = false;
                                                                                                                                                            }

                                                                                                                                                            if ((WallAvailability[3] == false)) {
                                                                                                                                                                if (this.Range) {
                                                                                                                                                                    "e3_WallStatus".Value = "Exclude";
                                                                                                                                                                    this.Range;
                                                                                                                                                                    true.Range("e3_LinerPanels").Value = "None";
                                                                                                                                                                    "".Range("e3_LinerPanels").Resize(1, 4).Locked = "None";
                                                                                                                                                                    "e3_LinerPanels".Resize(1, 4).Value = "None";
                                                                                                                                                                }

                                                                                                                                                                this.Range;
                                                                                                                                                                true.Range("e3_Wainscot").Value = "None";
                                                                                                                                                                "".Range("e3_Wainscot").Resize(1, 3).Locked = "None";
                                                                                                                                                                "e3_Wainscot".Resize(1, 3).Value = "None";
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                this.Range;
                                                                                                                                                                false.Range("e3_Wainscot").Resize(1, 3).Locked = false;
                                                                                                                                                                "e3_LinerPanels".Resize(1, 4).Locked = false;
                                                                                                                                                            }

                                                                                                                                                            if ((WallAvailability[4] == false)) {
                                                                                                                                                                if (this.Range) {
                                                                                                                                                                    "s4_WallStatus".Value = "Exclude";
                                                                                                                                                                    this.Range;
                                                                                                                                                                    true.Range("s4_LinerPanels").Value = "None";
                                                                                                                                                                    "".Range("s4_LinerPanels").Resize(1, 4).Locked = "None";
                                                                                                                                                                    "s4_LinerPanels".Resize(1, 4).Value = "None";
                                                                                                                                                                }

                                                                                                                                                                this.Range;
                                                                                                                                                                true.Range("s4_Wainscot").Value = "None";
                                                                                                                                                                "".Range("s4_Wainscot").Resize(1, 3).Locked = "None";
                                                                                                                                                                "s4_Wainscot".Resize(1, 3).Value = "None";
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                this.Range;
                                                                                                                                                                false.Range("s4_Wainscot").Resize(1, 3).Locked = false;
                                                                                                                                                                "s4_LinerPanels".Resize(1, 4).Locked = false;
                                                                                                                                                            }

                                                                                                                                                            // '' Wainscot
                                                                                                                                                            // no included walls, remove PDoors and OHDoors
                                                                                                                                                            if ((WallSelection == "")) {
                                                                                                                                                                0.Range("OHDoorNum").Value = 0;
                                                                                                                                                                UpdatesEventsProtection(true).Range("PDoorNum").Value = 0;
                                                                                                                                                                UpdatesEventsProtection(false);
                                                                                                                                                            }

                                                                                                                                                            // if all walls excluded, remove Windows and MiscFOs
                                                                                                                                                            if (this.Range) {
                                                                                                                                                                "e1_WallStatus" = ("Exclude" & this.Range);
                                                                                                                                                                "s2_WallStatus" = ("Exclude" & this.Range);
                                                                                                                                                                "e3_WallStatus" = ("Exclude" & this.Range);
                                                                                                                                                                "s4_WallStatus" = "Exclude";
                                                                                                                                                                0.Range("MiscFONum").Value = 0;
                                                                                                                                                                UpdatesEventsProtection(true).Range("WindowNum").Value = 0;
                                                                                                                                                                UpdatesEventsProtection(false);
                                                                                                                                                                // at least one wall is included, allow PDoors and OHDoors
                                                                                                                                                            }
                                                                                                                                                            else if ((WallSelection != "")) {
                                                                                                                                                                let FieldLocateWallSelection: string;
                                                                                                                                                                FieldLocateWallSelection = (WallSelection + ("," + "Field Locate"));
                                                                                                                                                                // With...
                                                                                                                                                                "pDoorCell1".offset(0, 2).Resize(12, 1).Validation.Delete.Add;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                FieldLocateWallSelection.IgnoreBlank = true;
                                                                                                                                                                this.Range;
                                                                                                                                                                "pDoorCell1".offset(0, 2).Resize(12, 1).Value = "";
                                                                                                                                                                // With...
                                                                                                                                                                "OHDoorCell1".offset(0, 3).Resize(12, 1).Validation.Delete.Add;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                WallSelection.IgnoreBlank = true;
                                                                                                                                                                this.Range;
                                                                                                                                                                "OHDoorCell1".offset(0, 3).Resize(12, 1).Value = "";
                                                                                                                                                            }

                                                                                                                                                            // Reset Wall Selection, allow for Partial or Included for Windows/MisFOs
                                                                                                                                                            WallSelection = "";
                                                                                                                                                            if (this.Range) {
                                                                                                                                                                ("e1_WallStatus".Value != "Exclude");
                                                                                                                                                                WallSelection = (WallSelection + ("," + "Endwall 1"));
                                                                                                                                                            }

                                                                                                                                                            if (this.Range) {
                                                                                                                                                                ("s2_WallStatus".Value != "Exclude");
                                                                                                                                                                WallSelection = (WallSelection + ("," + "Sidewall 2"));
                                                                                                                                                            }

                                                                                                                                                            if (this.Range) {
                                                                                                                                                                ("e3_WallStatus".Value != "Exclude");
                                                                                                                                                                WallSelection = (WallSelection + ("," + "Endwall 3"));
                                                                                                                                                            }

                                                                                                                                                            if (this.Range) {
                                                                                                                                                                ("s4_WallStatus".Value != "Exclude");
                                                                                                                                                                WallSelection = (WallSelection + ("," + "Sidewall 4"));
                                                                                                                                                            }

                                                                                                                                                            if ((WallSelection == "")) {
                                                                                                                                                                // Do nothing
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                // '''Windows
                                                                                                                                                                FieldLocateWallSelection = (WallSelection + ("," + "Field Locate"));
                                                                                                                                                                // With...
                                                                                                                                                                "WindowCell1".offset(0, 3).Resize(24, 1).Validation.Delete.Add;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                FieldLocateWallSelection.IgnoreBlank = true;
                                                                                                                                                                this.Range;
                                                                                                                                                                "WindowCell1".offset(0, 3).Resize(24, 1).Value = "";
                                                                                                                                                                // With...
                                                                                                                                                                "MiscFOCell1".offset(0, 3).Resize(12, 1).Validation.Delete.Add;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                FieldLocateWallSelection.IgnoreBlank = true;
                                                                                                                                                                this.Range;
                                                                                                                                                                "MiscFOCell1".offset(0, 3).Resize(12, 1).Value = "";
                                                                                                                                                            }

                                                                                                                                                            UpdatesEventsProtection(true);
                                                                                                                                                        }

                                                                                                                                                        // ''''''''''''''''''''''''''''''''''''''''''''' Sub for checking everything if a paste has happened
                                                                                                                                                        FullSheetCheck();
                                                                                                                                                        let BayStart: Range;
                                                                                                                                                        let BayEnd: Range;
                                                                                                                                                        let Downspouts: Range;
                                                                                                                                                        let PDoorStart: Range;
                                                                                                                                                        let PDoorEnd: Range;
                                                                                                                                                        let OHDoorStart: Range;
                                                                                                                                                        let OHDoorEnd: Range;
                                                                                                                                                        let WindowStart: Range;
                                                                                                                                                        let WindowEnd: Range;
                                                                                                                                                        let FOStart: Range;
                                                                                                                                                        let FOEnd: Range;
                                                                                                                                                        let OverhangTbl: Range;
                                                                                                                                                        let ExtensionTbl: Range;
                                                                                                                                                        let Overhangs: Range;
                                                                                                                                                        let Extensions: Range;
                                                                                                                                                        let cell: Range;
                                                                                                                                                        let WallAvailability: boolean[,];
                                                                                                                                                        // ranges
                                                                                                                                                        BayStart = EstSht.Range("Building_Height").offset(2, -1);
                                                                                                                                                        BayEnd = BayStart.offset(12, 0);
                                                                                                                                                        PDoorStart = EstSht.Range("pDoorCell1").offset(-1, 0);
                                                                                                                                                        PDoorEnd = EstSht.Range("pDoorCell12");
                                                                                                                                                        OHDoorStart = EstSht.Range("OHDoorCell1").offset(-1, 0);
                                                                                                                                                        OHDoorEnd = EstSht.Range("OHDoorCell12");
                                                                                                                                                        WindowStart = EstSht.Range("WindowCell1").offset(-1, 0);
                                                                                                                                                        WindowEnd = EstSht.Range("WindowCell12");
                                                                                                                                                        FOStart = EstSht.Range("MiscFOCell1").offset(-1, 0);
                                                                                                                                                        FOEnd = EstSht.Range("MiscFOCell12");
                                                                                                                                                        OverhangTbl = EstSht.Range("e1_GableOverhang").offset(-1, -1).Resize(5, 7);
                                                                                                                                                        ExtensionTbl = EstSht.Range("e1_GableExtension").offset(-1, -1).Resize(5, 7);
                                                                                                                                                        Overhangs = Range(EstSht.Range("e1_GableOverhang"), EstSht.Range("s4_EaveOverhang"));
                                                                                                                                                        Extensions = Range(EstSht.Range("e1_GableExtension"), EstSht.Range("s4_EaveExtension"));
                                                                                                                                                        // 'Wall Availability for Liners, Wainscot, FOs
                                                                                                                                                        // Assume all available, change if not
                                                                                                                                                        WallAvailability[1] = true;
                                                                                                                                                        WallAvailability[2] = true;
                                                                                                                                                        WallAvailability[3] = true;
                                                                                                                                                        WallAvailability[4] = true;
                                                                                                                                                        if ((this.Range("e1_WallStatus").Value != "Include")) {
                                                                                                                                                            WallAvailability[1] = false;
                                                                                                                                                        }

                                                                                                                                                        if ((this.Range("s2_WallStatus").Value != "Include")) {
                                                                                                                                                            WallAvailability[2] = false;
                                                                                                                                                        }

                                                                                                                                                        if ((this.Range("e3_WallStatus").Value != "Include")) {
                                                                                                                                                            WallAvailability[3] = false;
                                                                                                                                                        }

                                                                                                                                                        if ((this.Range("s4_WallStatus").Value != "Include")) {
                                                                                                                                                            WallAvailability[4] = false;
                                                                                                                                                        }

                                                                                                                                                        UpdatesEventsProtection(false);
                                                                                                                                                        // ''' bay number change
                                                                                                                                                        // unprotect
                                                                                                                                                        // hide row under bay number box when bay number is 0
                                                                                                                                                        if ((this.Range("BayNum").Value == 0)) {
                                                                                                                                                            if ((BayStart.offset(-1, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                BayStart.offset(-1, 0).EntireRow.Hidden = true;
                                                                                                                                                            }
                                                                                                                                                            else {
                                                                                                                                                                if ((BayStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                    BayStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                                                                                                                }

                                                                                                                                                                // change all bay lengths to 0
                                                                                                                                                                EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0;
                                                                                                                                                                // check  value
                                                                                                                                                                switch (this.Range("BayNum").Value) {
                                                                                                                                                                    case "":
                                                                                                                                                                        this.Range("BayNum").Value = "0";
                                                                                                                                                                        break;
                                                                                                                                                                    case "0":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        break;
                                                                                                                                                                    case "1":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(2, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "2":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(3, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "3":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "4":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(5, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "5":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(6, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "6":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(7, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "7":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(8, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "8":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(9, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "9":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(10, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "10":
                                                                                                                                                                        if ((EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(11, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "11":
                                                                                                                                                                        if ((BayEnd.EntireRow.Hidden == false)) {
                                                                                                                                                                            BayEnd.EntireRow.Hidden = true;
                                                                                                                                                                        }

                                                                                                                                                                        BayStart.Resize(12, 1).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                    case "12":
                                                                                                                                                                        EstSht.Range(BayStart, BayEnd).EntireRow.Hidden = false;
                                                                                                                                                                        break;
                                                                                                                                                                }

                                                                                                                                                                // '''' Alter Walls
                                                                                                                                                                // check if yes or no
                                                                                                                                                                switch (EstSht.Range("AlterWalls").Value) {
                                                                                                                                                                    case "":
                                                                                                                                                                        EstSht.Range("AlterWalls").Value = "No";
                                                                                                                                                                        break;
                                                                                                                                                                    case "No":
                                                                                                                                                                        if ((EstSht.Range("Wainscot").Value != "Yes")) {
                                                                                                                                                                            // hide column k, format J
                                                                                                                                                                            if ((EstSht.Columns["K:K"].Hidden == false)) {
                                                                                                                                                                                EstSht.Columns["K:K"].Hidden = true;
                                                                                                                                                                            }

                                                                                                                                                                            EstSht.Columns["J:J"].ColumnWidth = 5;
                                                                                                                                                                            // remove row seperating alter walls and wainscot table
                                                                                                                                                                        }
                                                                                                                                                                        else if ((EstSht.Range("Wainscot").Value == "Yes")) {
                                                                                                                                                                            if ((EstSht.Range("LinerPanels").Value == "No")) {
                                                                                                                                                                                // resize section heading row
                                                                                                                                                                                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                                                                                                                                                                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                            }

                                                                                                                                                                        }

                                                                                                                                                                        // unhide row above wainscot table if needed
                                                                                                                                                                        if (((EstSht.Range("Wainscot").Value == "Yes")
                                                                                                                                                                                    && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                                                                                                                                                                            EstSht.Range("Wainscot").offset(-3, 0).EntireRow.Hidden = false;
                                                                                                                                                                        }

                                                                                                                                                                        Range(this.Range("e1_WallStatus"), this.Range("s4_WallStatus")).Value = "Include";
                                                                                                                                                                        Range(this.Range("e1_WallStatus"), this.Range("s4_WallStatus")).offset(0, 2).Value = 0;
                                                                                                                                                                        this.Range("e1_Expandable").Value = "No";
                                                                                                                                                                        this.Range("e3_Expandable").Value = "No";
                                                                                                                                                                        WallAvailability[1] = true;
                                                                                                                                                                        WallAvailability[2] = true;
                                                                                                                                                                        WallAvailability[3] = true;
                                                                                                                                                                        WallAvailability[4] = true;
                                                                                                                                                                        break;
                                                                                                                                                                    case "Yes":
                                                                                                                                                                        if ((EstSht.Columns["K:K"].Hidden == true)) {
                                                                                                                                                                            EstSht.Columns["K:K"].Hidden = false;
                                                                                                                                                                        }

                                                                                                                                                                        EstSht.Columns["J:J"].ColumnWidth = 30;
                                                                                                                                                                        // unhide last table row
                                                                                                                                                                        if ((EstSht.Range("Wainscot").Value == "Yes")) {
                                                                                                                                                                            if ((EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                                EstSht.Range("AlterWalls").offset(2, 0).EntireRow.Hidden = false;
                                                                                                                                                                            }

                                                                                                                                                                            // '''''''''''''''''''''''''''''' Set Defaults
                                                                                                                                                                            // With...
                                                                                                                                                                            "Include".Range("e1_Expandable").Value = "No";
                                                                                                                                                                            "Include".Range("s4_WallStatus").Value = "No";
                                                                                                                                                                            "Include".Range("e3_WallStatus").Value = "No";
                                                                                                                                                                            "Include".Range("s2_WallStatus").Value = "No";
                                                                                                                                                                            EstSht.Range("e1_WallStatus").Value = "No";
                                                                                                                                                                        }

                                                                                                                                                                        // '''' Wall Status Changes
                                                                                                                                                                        // With...
                                                                                                                                                                        if (this.Range) {
                                                                                                                                                                            "e1_WallStatus".Value = "Partial";
                                                                                                                                                                            this.Range;
                                                                                                                                                                            false.Range("e1_WallStatus").offset(0, 2).Value = 0;
                                                                                                                                                                            "e1_WallStatus".offset(0, 2).Locked = 0;
                                                                                                                                                                        }
                                                                                                                                                                        else {
                                                                                                                                                                            if (this.Range) {
                                                                                                                                                                                "e1_WallStatus".Value = "";
                                                                                                                                                                                this.Range;
                                                                                                                                                                                "e1_WallStatus".Value = "Include";
                                                                                                                                                                                WallAvailability[1] = true;
                                                                                                                                                                            }

                                                                                                                                                                            this.Range;
                                                                                                                                                                            true.Range("e1_WallStatus").offset(0, 2).Value = "N/A";
                                                                                                                                                                            "e1_WallStatus".offset(0, 2).Locked = "N/A";
                                                                                                                                                                        }

                                                                                                                                                                        if (this.Range) {
                                                                                                                                                                            "s2_WallStatus".Value = "Partial";
                                                                                                                                                                            this.Range;
                                                                                                                                                                            false.Range("s2_WallStatus").offset(0, 2).Value = 0;
                                                                                                                                                                            "s2_WallStatus".offset(0, 2).Locked = 0;
                                                                                                                                                                        }
                                                                                                                                                                        else {
                                                                                                                                                                            if (this.Range) {
                                                                                                                                                                                "s2_WallStatus".Value = "";
                                                                                                                                                                                this.Range;
                                                                                                                                                                                "s2_WallStatus".Value = "Include";
                                                                                                                                                                                WallAvailability[2] = true;
                                                                                                                                                                            }

                                                                                                                                                                            this.Range;
                                                                                                                                                                            true.Range("s2_WallStatus").offset(0, 2).Value = "N/A";
                                                                                                                                                                            "s2_WallStatus".offset(0, 2).Locked = "N/A";
                                                                                                                                                                        }

                                                                                                                                                                        if (this.Range) {
                                                                                                                                                                            "e3_WallStatus".Value = "Partial";
                                                                                                                                                                            this.Range;
                                                                                                                                                                            false.Range("e3_WallStatus").offset(0, 2).Value = 0;
                                                                                                                                                                            "e3_WallStatus".offset(0, 2).Locked = 0;
                                                                                                                                                                        }
                                                                                                                                                                        else {
                                                                                                                                                                            if (this.Range) {
                                                                                                                                                                                "e3_WallStatus".Value = "";
                                                                                                                                                                                this.Range;
                                                                                                                                                                                "e3_WallStatus".Value = "Include";
                                                                                                                                                                                WallAvailability[3] = true;
                                                                                                                                                                            }

                                                                                                                                                                            this.Range;
                                                                                                                                                                            true.Range("e3_WallStatus").offset(0, 2).Value = "N/A";
                                                                                                                                                                            "e3_WallStatus".offset(0, 2).Locked = "N/A";
                                                                                                                                                                        }

                                                                                                                                                                        if (this.Range) {
                                                                                                                                                                            "s4_WallStatus".Value = "Partial";
                                                                                                                                                                            this.Range;
                                                                                                                                                                            false.Range("s4_WallStatus").offset(0, 2).Value = 0;
                                                                                                                                                                            "s4_WallStatus".offset(0, 2).Locked = 0;
                                                                                                                                                                        }
                                                                                                                                                                        else {
                                                                                                                                                                            if (this.Range) {
                                                                                                                                                                                "s4_WallStatus".Value = "";
                                                                                                                                                                                this.Range;
                                                                                                                                                                                "s4_WallStatus".Value = "Include";
                                                                                                                                                                                WallAvailability[4] = true;
                                                                                                                                                                            }

                                                                                                                                                                            this.Range;
                                                                                                                                                                            true.Range("s4_WallStatus").offset(0, 2).Value = "N/A";
                                                                                                                                                                            "s4_WallStatus".offset(0, 2).Locked = "N/A";
                                                                                                                                                                        }

                                                                                                                                                                        // '''' Liner Panels Section
                                                                                                                                                                        // check if yes or no
                                                                                                                                                                        switch (this.Range("LinerPanels").Value) {
                                                                                                                                                                            case "":
                                                                                                                                                                                EstSht.Range("LinerPanels").Value = "No";
                                                                                                                                                                                break;
                                                                                                                                                                            case "No":
                                                                                                                                                                                EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                                                                                                                                                                // hide liner panels section
                                                                                                                                                                                if ((Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == false)) {
                                                                                                                                                                                    Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = true;
                                                                                                                                                                                }

                                                                                                                                                                                if (((EstSht.Range("Wainscot").Value == "Yes")
                                                                                                                                                                                            && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                                                                                                                                                                                    EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                Range(this.Range("e1_LinerPanels"), this.Range("Roof_LinerPanels")).Value = "None";
                                                                                                                                                                                Range(this.Range("e1_LinerPanels"), this.Range("Roof_LinerPanels")).offset(0, 1).Resize(4, 4).Value = "";
                                                                                                                                                                                break;
                                                                                                                                                                            case "Yes":
                                                                                                                                                                                if ((Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == true)) {
                                                                                                                                                                                    Range(EstSht.Range("LinerPanels").offset(2, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                if ((Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden == true)) {
                                                                                                                                                                                    Range(EstSht.Range("LinerPanels").offset(3, 0), EstSht.Range("LinerPanels").offset(10, 0)).EntireRow.Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                EstSht.Range("LinerPanels").offset(2, 0).EntireRow.AutoFit;
                                                                                                                                                                                break;
                                                                                                                                                                        }

                                                                                                                                                                        // '''' Liner Panels Options Change
                                                                                                                                                                        for (cell in Range(EstSht.Range("e1_LinerPanels"), EstSht.Range("Roof_LinerPanels"))) {
                                                                                                                                                                            if ((cell.Value == "")) {
                                                                                                                                                                                cell.Value = "None";
                                                                                                                                                                            }

                                                                                                                                                                            if ((cell.Value == "None")) {
                                                                                                                                                                                cell.offset(0, 1).Value = "";
                                                                                                                                                                                cell.offset(0, 2).Value = "";
                                                                                                                                                                                cell.offset(0, 3).Value = "";
                                                                                                                                                                            }

                                                                                                                                                                        }

                                                                                                                                                                        // '''' Wainscot Section
                                                                                                                                                                        // check if yes or no
                                                                                                                                                                        switch (EstSht.Range("Wainscot").Value) {
                                                                                                                                                                            case "":
                                                                                                                                                                                EstSht.Range("Wainscot").Value = "No";
                                                                                                                                                                                break;
                                                                                                                                                                            case "No":
                                                                                                                                                                                if ((EstSht.Range("AlterWalls").Value != "Yes")) {
                                                                                                                                                                                    // hide column k, format J
                                                                                                                                                                                    if ((EstSht.Columns["K:K"].Hidden == false)) {
                                                                                                                                                                                        EstSht.Columns["K:K"].Hidden = true;
                                                                                                                                                                                    }

                                                                                                                                                                                    EstSht.Columns["J:J"].ColumnWidth = 5;
                                                                                                                                                                                    if ((EstSht.Range("LinerPanels").Value == "No")) {
                                                                                                                                                                                        // resize section heading row
                                                                                                                                                                                        EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                                                                                                                                                                        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                    }

                                                                                                                                                                                }
                                                                                                                                                                                else if ((EstSht.Range("AlterWalls").Value == "Yes")) {
                                                                                                                                                                                    // hide row seperating walls and wainscot table
                                                                                                                                                                                    if ((EstSht.Range("LinerPanels").Value == "No")) {
                                                                                                                                                                                        // resize section heading row
                                                                                                                                                                                        EstSht.Range("LinerPanels").offset(2, 0).RowHeight = 15;
                                                                                                                                                                                        EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                    }

                                                                                                                                                                                }

                                                                                                                                                                                // '''''''''''''''''''''''''''''' reset rows
                                                                                                                                                                                // With...
                                                                                                                                                                                Range(., Range("e1_Wainscot"), ., Range("s4_Wainscot")).Value = "None";
                                                                                                                                                                                "".Range("e3_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                "".Range("s2_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                EstSht.Range("e1_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                break;
                                                                                                                                                                            case "Yes":
                                                                                                                                                                                if ((EstSht.Columns["K:K"].Hidden == true)) {
                                                                                                                                                                                    EstSht.Columns["K:K"].Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                EstSht.Columns["J:J"].ColumnWidth = 30;
                                                                                                                                                                                //         'unhide row above table if needed
                                                                                                                                                                                if (((EstSht.Range("LinerPanels").Value != "Yes")
                                                                                                                                                                                            && (EstSht.Range("AlterWalls").Value == "Yes"))) {
                                                                                                                                                                                    EstSht.Range("LinerPanels").offset(2, 0).EntireRow.Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                // With...
                                                                                                                                                                                "".Range("e3_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                "".Range("s2_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                "None".Range("e1_Wainscot").offset(0, 1).Resize(1, 2).Value = "";
                                                                                                                                                                                "None".Range("s4_Wainscot").Value = "";
                                                                                                                                                                                "None".Range("e3_Wainscot").Value = "";
                                                                                                                                                                                "None".Range("s2_Wainscot").Value = "";
                                                                                                                                                                                EstSht.Range("e1_Wainscot").Value = "";
                                                                                                                                                                                break;
                                                                                                                                                                        }

                                                                                                                                                                        // '''' Wainscot Options Change
                                                                                                                                                                        for (cell in Range(EstSht.Range("e1_Wainscot"), EstSht.Range("s4_Wainscot"))) {
                                                                                                                                                                            if ((cell.Value == "")) {
                                                                                                                                                                                cell.Value = "None";
                                                                                                                                                                            }

                                                                                                                                                                            if ((cell.Value == "None")) {
                                                                                                                                                                                cell.offset(0, 1).Value = "";
                                                                                                                                                                                cell.offset(0, 2).Value = "";
                                                                                                                                                                            }

                                                                                                                                                                        }

                                                                                                                                                                        // ''' Gutter & Downspouts
                                                                                                                                                                        // check if yes or no
                                                                                                                                                                        if ((EstSht.Range("GutterAndDownspouts").Value == "")) {
                                                                                                                                                                            EstSht.Range("GutterAndDownspouts").Value = "No";
                                                                                                                                                                        }

                                                                                                                                                                        if (((EstSht.Range("PDoorNum").Value == 0)
                                                                                                                                                                                    || (EstSht.Range("PDoorNum").Value == ""))) {
                                                                                                                                                                            EstSht.Range("PDoorNum").Value = 0;
                                                                                                                                                                            // reset table to blank
                                                                                                                                                                            this.Range("pDoorCell1").offset(0, 1).Resize(12, 6).Value = "";
                                                                                                                                                                            if ((PDoorStart.offset(-1, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                PDoorStart.offset(-1, 0).EntireRow.Hidden = true;
                                                                                                                                                                            }
                                                                                                                                                                            else {
                                                                                                                                                                                if ((PDoorStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                                    PDoorStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                                                                                                                                }

                                                                                                                                                                                // check target value
                                                                                                                                                                                switch (EstSht.Range("PDoorNum").Value) {
                                                                                                                                                                                    case "0":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        break;
                                                                                                                                                                                    case "1":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(2, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "2":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(3, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "3":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "4":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(5, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "5":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(6, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "6":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(7, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "7":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(8, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "8":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(9, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "9":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(10, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "10":
                                                                                                                                                                                        if ((EstSht.Range(PDoorStart, PDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                            EstSht.Range(PDoorStart, PDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(11, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "11":
                                                                                                                                                                                        if ((PDoorEnd.EntireRow.Hidden == false)) {
                                                                                                                                                                                            PDoorEnd.EntireRow.Hidden = true;
                                                                                                                                                                                        }

                                                                                                                                                                                        PDoorStart.Resize(12, 1).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                    case "12":
                                                                                                                                                                                        EstSht.Range(PDoorStart, PDoorEnd).EntireRow.Hidden = false;
                                                                                                                                                                                        break;
                                                                                                                                                                                }

                                                                                                                                                                                // ''' OH door
                                                                                                                                                                                // hide row under quantity box when quantity is 0
                                                                                                                                                                                if (((EstSht.Range("OHDoorNum").Value == 0)
                                                                                                                                                                                            || (EstSht.Range("OHDoorNum").Value == ""))) {
                                                                                                                                                                                    EstSht.Range("OHDoorNum").Value = 0;
                                                                                                                                                                                    // reset table to blank
                                                                                                                                                                                    this.Range("OHDoorCell1").offset(0, 1).Resize(12, 8).Value = "";
                                                                                                                                                                                    if ((OHDoorStart.offset(-1, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                        OHDoorStart.offset(-1, 0).EntireRow.Hidden = true;
                                                                                                                                                                                    }
                                                                                                                                                                                    else {
                                                                                                                                                                                        if ((OHDoorStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                                            OHDoorStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                                                                                                                                        }

                                                                                                                                                                                        // check target value
                                                                                                                                                                                        switch (EstSht.Range("OHDoorNum").Value) {
                                                                                                                                                                                            case "0":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                break;
                                                                                                                                                                                            case "1":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(2, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "2":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(3, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "3":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "4":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(5, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "5":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(6, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "6":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(7, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "7":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(8, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "8":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(9, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "9":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(10, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "10":
                                                                                                                                                                                                if ((EstSht.Range(OHDoorStart, OHDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                    EstSht.Range(OHDoorStart, OHDoorEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(11, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "11":
                                                                                                                                                                                                if ((OHDoorEnd.EntireRow.Hidden == false)) {
                                                                                                                                                                                                    OHDoorEnd.EntireRow.Hidden = true;
                                                                                                                                                                                                }

                                                                                                                                                                                                OHDoorStart.Resize(12, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                            case "12":
                                                                                                                                                                                                EstSht.Range(OHDoorStart, OHDoorEnd).EntireRow.Hidden = false;
                                                                                                                                                                                                break;
                                                                                                                                                                                        }

                                                                                                                                                                                        // ''' Windows
                                                                                                                                                                                        // hide row under quantity box when quantity is 0
                                                                                                                                                                                        if (((EstSht.Range("WindowNum").Value == 0)
                                                                                                                                                                                                    || (EstSht.Range("WindowNum").Value == ""))) {
                                                                                                                                                                                            EstSht.Range("WindowNum").Value = 0;
                                                                                                                                                                                            // reset table to blank
                                                                                                                                                                                            this.Range("WindowCell1").offset(0, 1).Resize(24, 3).Value = "";
                                                                                                                                                                                            if ((WindowStart.offset(-1, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                WindowStart.offset(-1, 0).EntireRow.Hidden = true;
                                                                                                                                                                                            }
                                                                                                                                                                                            else {
                                                                                                                                                                                                if ((WindowStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                                                    WindowStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                                                                                                                                                }

                                                                                                                                                                                                // check target value
                                                                                                                                                                                                switch (EstSht.Range("WindowNum").Value) {
                                                                                                                                                                                                    case "0":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "1":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(2, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "2":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(3, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "3":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "4":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(5, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "5":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(6, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "6":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(7, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "7":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(8, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "8":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(9, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "9":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(10, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "10":
                                                                                                                                                                                                        if ((EstSht.Range(WindowStart, WindowEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                            EstSht.Range(WindowStart, WindowEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(11, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "11":
                                                                                                                                                                                                        if ((WindowEnd.EntireRow.Hidden == false)) {
                                                                                                                                                                                                            WindowEnd.EntireRow.Hidden = true;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        WindowStart.Resize(12, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                    case "12":
                                                                                                                                                                                                        EstSht.Range(WindowStart, WindowEnd).EntireRow.Hidden = false;
                                                                                                                                                                                                        break;
                                                                                                                                                                                                }

                                                                                                                                                                                                // ''' Misc Framed Openings
                                                                                                                                                                                                // hide row under quantity box when quantity is 0
                                                                                                                                                                                                if (((EstSht.Range("MiscFONum").Value == 0)
                                                                                                                                                                                                            || (EstSht.Range("MiscFONum").Value == ""))) {
                                                                                                                                                                                                    EstSht.Range("MiscFONum").Value = 0;
                                                                                                                                                                                                    // reset table to blank
                                                                                                                                                                                                    this.Range("MiscFOCell1").offset(0, 1).Resize(12, 5).Value = "";
                                                                                                                                                                                                    if ((FOStart.offset(-1, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                        FOStart.offset(-1, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                    }
                                                                                                                                                                                                    else {
                                                                                                                                                                                                        if ((FOStart.offset(-1, 0).EntireRow.Hidden == true)) {
                                                                                                                                                                                                            FOStart.offset(-1, 0).EntireRow.Hidden = false;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        // check target value
                                                                                                                                                                                                        switch (EstSht.Range("MiscFONum").Value) {
                                                                                                                                                                                                            case "0":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "1":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(11, 1).offset(2, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(2, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "2":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(10, 1).offset(3, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(3, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "3":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(9, 1).offset(4, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "4":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(8, 1).offset(5, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(5, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "5":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(7, 1).offset(6, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(6, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "6":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(6, 1).offset(7, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(7, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "7":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(5, 1).offset(8, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(8, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "8":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(4, 1).offset(9, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(9, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "9":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(3, 1).offset(10, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(10, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "10":
                                                                                                                                                                                                                if ((EstSht.Range(FOStart, FOEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    EstSht.Range(FOStart, FOEnd).Resize(2, 1).offset(11, 0).EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(11, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "11":
                                                                                                                                                                                                                if ((FOEnd.EntireRow.Hidden == false)) {
                                                                                                                                                                                                                    FOEnd.EntireRow.Hidden = true;
                                                                                                                                                                                                                }

                                                                                                                                                                                                                FOStart.Resize(12, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                            case "12":
                                                                                                                                                                                                                EstSht.Range(FOStart, FOEnd).EntireRow.Hidden = false;
                                                                                                                                                                                                                break;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        // ''' building length change
                                                                                                                                                                                                        // if blank building length, reset to 0
                                                                                                                                                                                                        if ((EstSht.Range("Building_Length").Value == "")) {
                                                                                                                                                                                                            EstSht.Range("Building_Length").Value = 0;
                                                                                                                                                                                                            EstSht.Range(BayStart, BayEnd).Resize(12, 1).offset(1, 1).Value = 0;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        // ''' Trim Color Bulk Change
                                                                                                                                                                                                        if ((EstSht.Range("All_tColors").Value == "")) {
                                                                                                                                                                                                            EstSht.Range("All_tColors").Value = "N/A";
                                                                                                                                                                                                        }
                                                                                                                                                                                                        else {
                                                                                                                                                                                                            // change trim colors to match
                                                                                                                                                                                                            // With...
                                                                                                                                                                                                            // change gutter/downspout colors as well
                                                                                                                                                                                                            EstSht.Range("All_tColors").Value.Range("FO_tColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                            EstSht.Range("All_tColors").Value.Range("OutsideCorner_tColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                            EstSht.Range("All_tColors").Value.Range("Eave_tColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                            EstSht.Range("Rake_tColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                            EstSht.Range("All_tColors").Value.Range("GutterColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                            Range("DownspoutColor").Value = EstSht.Range("All_tColors").Value;
                                                                                                                                                                                                        }

                                                                                                                                                                                                        // ''' overhang table clear
                                                                                                                                                                                                        // clear soffits
                                                                                                                                                                                                        for (cell in Overhangs) {
                                                                                                                                                                                                            if (((cell.Value == "")
                                                                                                                                                                                                                        || (cell.Value == 0))) {
                                                                                                                                                                                                                // clear soffits
                                                                                                                                                                                                                cell.offset(0, 1).Value = "";
                                                                                                                                                                                                                cell.offset(0, 2).Value = "";
                                                                                                                                                                                                                cell.offset(0, 3).Value = "";
                                                                                                                                                                                                                cell.offset(0, 4).Value = "";
                                                                                                                                                                                                                cell.offset(0, 5).Value = "";
                                                                                                                                                                                                            }

                                                                                                                                                                                                        }

                                                                                                                                                                                                        // extension table clear
                                                                                                                                                                                                        // clear soffits
                                                                                                                                                                                                        for (cell in Extensions) {
                                                                                                                                                                                                            if (((cell.Value == "")
                                                                                                                                                                                                                        || (cell.Value == 0))) {
                                                                                                                                                                                                                // clear soffits
                                                                                                                                                                                                                cell.offset(0, 1).Value = "";
                                                                                                                                                                                                                cell.offset(0, 2).Value = "";
                                                                                                                                                                                                                cell.offset(0, 3).Value = "";
                                                                                                                                                                                                                cell.offset(0, 4).Value = "";
                                                                                                                                                                                                                cell.offset(0, 5).Value = "";
                                                                                                                                                                                                            }

                                                                                                                                                                                                        }

                                                                                                                                                                                                        // Show/Hide Eave Extension Pitch and Set Intersection default values
                                                                                                                                                                                                        // With...
                                                                                                                                                                                                        // s2 eave extension
                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                            "s2_EaveExtension".Value = "";
                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                "N/A".Range("s2e3_Intersection").Value = "N/A";
                                                                                                                                                                                                                true.Range("s2e1_Intersection").Value = "N/A";
                                                                                                                                                                                                                "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "N/A";
                                                                                                                                                                                                            }
                                                                                                                                                                                                            else {
                                                                                                                                                                                                                // If previously hidden, unhide and set default option to include intersection
                                                                                                                                                                                                                if (EstSht.Range) {
                                                                                                                                                                                                                    "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = true;
                                                                                                                                                                                                                    EstSht.Range;
                                                                                                                                                                                                                    false.Range("s2_EaveExtensionPitch").Value = "Match Roof";
                                                                                                                                                                                                                    "s2_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "Match Roof";
                                                                                                                                                                                                                    if (EstSht.Range) {
                                                                                                                                                                                                                        ("e1_GableExtension".Value != "");
                                                                                                                                                                                                                        EstSht.Range;
                                                                                                                                                                                                                        "s2e1_Intersection".Value = "Include";
                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                            ("e3_GableExtension".Value != "");
                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                            "s2e3_Intersection".Value = "Include";
                                                                                                                                                                                                                        }

                                                                                                                                                                                                                    }

                                                                                                                                                                                                                    // s4 eave extension
                                                                                                                                                                                                                    if (EstSht.Range) {
                                                                                                                                                                                                                        "s4_EaveExtension".Value = "";
                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                            "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = false;
                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                            "N/A".Range("s4e3_Intersection").Value = "N/A";
                                                                                                                                                                                                                            true.Range("s4e1_Intersection").Value = "N/A";
                                                                                                                                                                                                                            "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "N/A";
                                                                                                                                                                                                                        }
                                                                                                                                                                                                                        else {
                                                                                                                                                                                                                            // If previously hidden, unhide and set default option to include intersection
                                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                                "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = true;
                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                false.Range("s4_EaveExtensionPitch").Value = "Match Roof";
                                                                                                                                                                                                                                "s4_EaveExtensionPitch".offset(-1, 0).Resize(4, 1).EntireRow.Hidden = "Match Roof";
                                                                                                                                                                                                                                if (EstSht.Range) {
                                                                                                                                                                                                                                    ("e1_GableExtension".Value != "");
                                                                                                                                                                                                                                    EstSht.Range;
                                                                                                                                                                                                                                    "s4e1_Intersection".Value = "Include";
                                                                                                                                                                                                                                    if (EstSht.Range) {
                                                                                                                                                                                                                                        ("e3_GableExtension".Value != "");
                                                                                                                                                                                                                                        EstSht.Range;
                                                                                                                                                                                                                                        "s4e3_Intersection".Value = "Include";
                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                // e1 gable extension Intersection Option
                                                                                                                                                                                                                                if (EstSht.Range) {
                                                                                                                                                                                                                                    "e1_GableExtension".Value = "";
                                                                                                                                                                                                                                    EstSht.Range;
                                                                                                                                                                                                                                    "N/A".Range("s4e1_Intersection").MergeArea.Locked = true;
                                                                                                                                                                                                                                    true.Range("s4e1_Intersection").Value = true;
                                                                                                                                                                                                                                    "N/A".Range("s2e1_Intersection").MergeArea.Locked = true;
                                                                                                                                                                                                                                    "s2e1_Intersection".Value = true;
                                                                                                                                                                                                                                }
                                                                                                                                                                                                                                else {
                                                                                                                                                                                                                                    // intersection for s2 e1
                                                                                                                                                                                                                                    if (EstSht.Range) {
                                                                                                                                                                                                                                        "s2e1_Intersection".Value = "N/A";
                                                                                                                                                                                                                                        EstSht.Range;
                                                                                                                                                                                                                                        "s2e1_Intersection".MergeArea.Locked = false;
                                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                                            ("s2_EaveExtension".Value != "");
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            "s2e1_Intersection".Value = "Include";
                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                        // intersection for s4 e1
                                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                                            "s4e1_Intersection".Value = "N/A";
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            "s4e1_Intersection".MergeArea.Locked = false;
                                                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                                                ("s4_EaveExtension".Value != "");
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                "s4e1_Intersection".Value = "Include";
                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                        // e3 gable extension Intersection Option
                                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                                            "e3_GableExtension".Value = "";
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            "N/A".Range("s4e3_Intersection").MergeArea.Locked = true;
                                                                                                                                                                                                                                            true.Range("s4e3_Intersection").Value = true;
                                                                                                                                                                                                                                            "N/A".Range("s2e3_Intersection").MergeArea.Locked = true;
                                                                                                                                                                                                                                            "s2e3_Intersection".Value = true;
                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                        else {
                                                                                                                                                                                                                                            // intersection for s2 e3
                                                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                                                "s2e3_Intersection".Value = "N/A";
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                false.Range("s2e3_Intersection").Value = "Include";
                                                                                                                                                                                                                                                "s2e3_Intersection".MergeArea.Locked = "Include";
                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                            // intersection for s4 e3
                                                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                                                "s4e3_Intersection".Value = "N/A";
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                false.Range("s4e3_Intersection").Value = "Include";
                                                                                                                                                                                                                                                "s4e3_Intersection".MergeArea.Locked = "Include";
                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                    // ''' Roof/Wall Panel Shape Change - Disable Translucent Wall Panels and Skylights
                                                                                                                                                                                                                                    // With...
                                                                                                                                                                                                                                    // ''' Panel Shape
                                                                                                                                                                                                                                    // Hide skylights/translucent wall panels for m-loc
                                                                                                                                                                                                                                    if (EstSht.Range) {
                                                                                                                                                                                                                                        "Wall_pShape".Value = ("M-Loc" | EstSht.Range);
                                                                                                                                                                                                                                        "Roof_pShape".Value = "M-Loc";
                                                                                                                                                                                                                                        EstSht.Range;
                                                                                                                                                                                                                                        "".Range("SkylightLength").Value = "";
                                                                                                                                                                                                                                        "".Range("TranslucentWallPanelLength").Value = "";
                                                                                                                                                                                                                                        "".Range("SkylightQty").Value = "";
                                                                                                                                                                                                                                        "TranslucentWallPanelQty".Value = "";
                                                                                                                                                                                                                                        if (EstSht.Range) {
                                                                                                                                                                                                                                            EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            "SkylightQty";
                                                                                                                                                                                                                                            EstSht.EntireRow.Hidden = false;
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                                                                                                                                            EstSht.Range;
                                                                                                                                                                                                                                            "SkylightQty";
                                                                                                                                                                                                                                            EstSht.EntireRow.Hidden = true;
                                                                                                                                                                                                                                        }
                                                                                                                                                                                                                                        else if (EstSht.Range) {
                                                                                                                                                                                                                                            (("Wall_pShape".Value != "M-Loc")
                                                                                                                                                                                                                                                        & EstSht.Range);
                                                                                                                                                                                                                                            ("Roof_pShape".Value != "M-Loc");
                                                                                                                                                                                                                                            // unhide rows if needed
                                                                                                                                                                                                                                            if (EstSht.Range) {
                                                                                                                                                                                                                                                EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                "SkylightQty";
                                                                                                                                                                                                                                                EstSht.EntireRow.Hidden = true;
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                EstSht.Range["TranslucentWallPanelQty"];
                                                                                                                                                                                                                                                EstSht.Range;
                                                                                                                                                                                                                                                "SkylightQty";
                                                                                                                                                                                                                                                EstSht.EntireRow.Hidden = false;
                                                                                                                                                                                                                                            }

                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                        // ''''''''''''''''''''''''''''''''''''''''''''''''' Framed Opening Option Changes '''''''''''''''''''''
                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                        // ''''''''''''''''''''''''''''''''''''''''' Personnel Doors '''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                                                                                                        for (cell in Range(EstSht.Range, "pDoorCell1".offset(0, 1), EstSht.Range, "pDoorCell12".offset(0, 1))) {
                                                                                                                                                                                                                                            if ((cell.Value == "4070")) {
                                                                                                                                                                                                                                                // Remove half glass option
                                                                                                                                                                                                                                                // With...
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                                                                                                "No".IgnoreBlank = true;
                                                                                                                                                                                                                                                cell.offset(0, 2).Value = "No";
                                                                                                                                                                                                                                                // With...
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                                                                                                "No".IgnoreBlank = true;
                                                                                                                                                                                                                                                cell.offset(0, 5).Value = "No";
                                                                                                                                                                                                                                            }
                                                                                                                                                                                                                                            else if (((cell.Value == "3070")
                                                                                                                                                                                                                                                        || (cell.Value == ""))) {
                                                                                                                                                                                                                                                // restore half glass, deadbolt
                                                                                                                                                                                                                                                // With...
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                                                                                                "Yes,No".IgnoreBlank = true;
                                                                                                                                                                                                                                                if ((cell.offset(0, 2).Validation.Value == false)) {
                                                                                                                                                                                                                                                    cell.offset(0, 2).Value = "";
                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                                // With...
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                true.ShowInput = true;
                                                                                                                                                                                                                                                true.InCellDropdown = true;
                                                                                                                                                                                                                                                "Yes,No".IgnoreBlank = true;
                                                                                                                                                                                                                                                if ((cell.offset(0, 5).Validation.Value == false)) {
                                                                                                                                                                                                                                                    cell.offset(0, 5).Value = "";
                                                                                                                                                                                                                                                }

                                                                                                                                                                                                                                                cell;
                                                                                                                                                                                                                                                // ''''''''''''''''''''''''''''''''''''''''' Overhead Doors '''''''''''''''''''''''''''''''''''''''''''''''
                                                                                                                                                                                                                                                for (cell in Range(EstSht.Range, "OHDoorCell1".offset(0, 4), EstSht.Range, "OHDoorCell12".offset(0, 4))) {
                                                                                                                                                                                                                                                    // '''''''''''''''''''''''' Roll Up Doors''''''''''''''''''''''
                                                                                                                                                                                                                                                    if ((cell.Value == "RUD")) {
                                                                                                                                                                                                                                                        // ''sizing options
                                                                                                                                                                                                                                                        // width
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("RUDWidth").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        // clear if invalid
                                                                                                                                                                                                                                                        if ((cell.offset(0, -3).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, -3).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("RUDHeight").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        // clear if invalid
                                                                                                                                                                                                                                                        if ((cell.offset(0, -2).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, -2).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "None".IgnoreBlank = true;
                                                                                                                                                                                                                                                        cell.offset(0, 1).Value = "None";
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "Chain Hoist".IgnoreBlank = true;
                                                                                                                                                                                                                                                        cell.offset(0, 2).Value = "Chain Hoist";
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "No".IgnoreBlank = true;
                                                                                                                                                                                                                                                        cell.offset(0, 3).Value = "No";
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "None".IgnoreBlank = true;
                                                                                                                                                                                                                                                        cell.offset(0, 4).Value = "None";
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    else if (((cell.Value == "Sectional")
                                                                                                                                                                                                                                                                || (cell.Value == ""))) {
                                                                                                                                                                                                                                                        // ''sizing options
                                                                                                                                                                                                                                                        // width
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("SectionalOHDoorWidth").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, -3).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, -3).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("SectionalOHDoorHeight").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, -2).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, -2).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("OHDoorInsulationOptions").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, 1).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, 1).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("OHDoorOperationOptions").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, 2).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, 2).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "Yes,No".IgnoreBlank = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, 3).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, 3).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        ("=Lists!" + ListSht.Range("OHDoorWindowOptions").Address.IgnoreBlank) = true;
                                                                                                                                                                                                                                                        if ((cell.offset(0, 4).Validation.Value == false)) {
                                                                                                                                                                                                                                                            cell.offset(0, 4).Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                        cell;
                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                    AlterAvailableWalls(WallAvailability);
                                                                                                                                                                                                                                                    UpdatesEventsProtection(true);
                                                                                                                                                                                                                                                    UpdatesEventsProtection((<boolean>(Setting)));
                                                                                                                                                                                                                                                    if ((Setting == true)) {
                                                                                                                                                                                                                                                        Application.ScreenUpdating = true;
                                                                                                                                                                                                                                                        Application.EnableEvents = true;
                                                                                                                                                                                                                                                        this.Protect;
                                                                                                                                                                                                                                                        "WhiteTruckMafia";
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    else {
                                                                                                                                                                                                                                                        Application.ScreenUpdating = false;
                                                                                                                                                                                                                                                        Application.EnableEvents = false;
                                                                                                                                                                                                                                                        this.Unprotect;
                                                                                                                                                                                                                                                        "WhiteTruckMafia";
                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                    PanelColorOptionCheck((<Range>(PanelTypeCell)), (<Range>(DropDownCell)));
                                                                                                                                                                                                                                                    UpdatesEventsProtection(false);
                                                                                                                                                                                                                                                    if (PanelTypeCell.Value) {
                                                                                                                                                                                                                                                        ("*" + ("Copper Metallic" + "*"));
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "Copper Metallic".IgnoreBlank = true;
                                                                                                                                                                                                                                                        DropDownCell.Value = "Copper Metallic";
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    else if (PanelTypeCell.Value) {
                                                                                                                                                                                                                                                        ("*" + ("Galvalume" + "*"));
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "Galvalume".IgnoreBlank = true;
                                                                                                                                                                                                                                                        DropDownCell.Value = "Galvalume";
                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    else if (((PanelTypeCell.Value.IndexOf("Thrifty", 0) + 1)
                                                                                                                                                                                                                                                                != 0)) {
                                                                                                                                                                                                                                                        if ((DropDownCell.Validation.Formula1 != "=ThriftyDropDownColors")) {
                                                                                                                                                                                                                                                            // With...
                                                                                                                                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                            /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                            /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                            true.ShowInput = true;
                                                                                                                                                                                                                                                            true.InCellDropdown = true;
                                                                                                                                                                                                                                                            "=ThriftyDropDownColors".IgnoreBlank = true;
                                                                                                                                                                                                                                                            DropDownCell.Value = "";
                                                                                                                                                                                                                                                        }

                                                                                                                                                                                                                                                    }
                                                                                                                                                                                                                                                    else if ((DropDownCell.Validation.Formula1 != "=DropDownColors")) {
                                                                                                                                                                                                                                                        // With...
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidateList;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlValidAlertStop;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */xlBetween;
                                                                                                                                                                                                                                                        /* Warning! Labeled Statements are not Implemented */true.ShowError = true;
                                                                                                                                                                                                                                                        true.ShowInput = true;
                                                                                                                                                                                                                                                        true.InCellDropdown = true;
                                                                                                                                                                                                                                                        "=DropDownColors".IgnoreBlank = true;
                                                                                                                                                                                                                                                        DropDownCell.Value = "";
                                                                                                                                                                                                                                                    }

                                                                                                                                                                                                                                                    UpdatesEventsProtection(true);
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

                                                                                                                                                                        break;
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
