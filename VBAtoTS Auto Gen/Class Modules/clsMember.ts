// TODO: Option Explicit ... Warning!!! not translated
let mType: string;
let Location: string;
let Length: number;
let Depth: number;
let Width: number;
let tType: string;
let Measurement: string;
let Qty: number;
let CL: number;
let rEdgePosition: number;
let DeleteFlag: boolean;
let bEdgeHeight: number;
let clsType: string;
let tEdgeHeight: number;
let Placement: string;
let ComponentMembers: Collection;
let LoadBearing: boolean;
let RafterLeftEdge: number;
let Size: string;


    // ''' Size string is the specific dimensions of the mType i.e. - "W8x12" or "8" C Purlin"
    public lEdgePosition(): number {
        // for receiver cee's, 0 width for the purpose of positioning since purlins will essentually fit flush into it
        if (((mType.IndexOf("Receiver Cee", 0) + 1)
                    == true)) {
            // receiver cee should never have a l/r edge position even if we're tracking other column's edges because of their orintation. *This is at least true when they're functioning as jambs*
            lEdgePosition = rEdgePosition;
        }
        else {
            lEdgePosition = (rEdgePosition + Width);
        }

    }

    // '''''''''''''''''''''''' Sub for finding the member's size string (and width) using the structural steel lookup table
    public SetSize(b: clsBuilding, ColumnOrRafter: string, Location: string, HorizontalReferenceDistance: number, CustomNonExpandable: string) {
        // ''' Valid Location Options: "Interior", "e1","s2","e3",and "s4"
        // Warning!!! Optional parameters not supported
        let LookupTbl: ListObject;
        let LookupHeight: number;
        let LookupHorizontalIndex: number;
        let LookupSizeString: string;
        let NearestHorizontalValue: number;
        if ((ColumnOrRafter == "Rafter")) {
            LookupTbl = LookupTblMatch(b, ColumnOrRafter, Location);
            LookupHeight = (Application.WorksheetFunction.RoundUp(((tEdgeHeight / 12)
                            / 10), 0) * 10);
            if ((HorizontalReferenceDistance <= (25 * 12))) {
                LookupHorizontalIndex = 1;
                //  default to 30' minimum for a given horizontal distance of less than 30'
            }
            else {
                LookupHorizontalIndex = (Application.WorksheetFunction.RoundUp((((HorizontalReferenceDistance / 12)
                                / 5)
                                - 5), 0) + 1);
                // LookupHorizontalIndex = NearestHorizontalValue
                // LookupHorizontalIndex = Application.WorksheetFunction.RoundDown((((Application.WorksheetFunction.RoundUp((HorizontalReferenceDistance / 12) / 10, 0) * 10) - 25) / 10) + 1, 0)
            }

            if ((LookupHeight > 80)) {
                /* Warning! GOTO is not Implemented */}

            if ((LookupHeight < 20)) {
                LookupHeight = 20;
            }

            if ((LookupHorizontalIndex > 12)) {
                if (((Location == "e1")
                            || (Location == "e3"))) {
                    LookupHorizontalIndex = 12;
                }
                else {
                    /* Warning! GOTO is not Implemented */}

            }

            // With...
            Size = LookupTbl.DataBodyRange;
            LookupTbl.ListRows[LookupHorizontalIndex.ToString()].Index;
            LookupTbl.ListColumns;
            LookupHeight.ToString().Index;
            if (((Size.IndexOf("TS", 0) + 1)
                        != 0)) {
                Width = 4;
            }
            else if (((Size.IndexOf("W", 0) + 1)
                        != 0)) {
                Width = Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                - 1)).Substring((Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                    - 1)).Length - (Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                    - 1)).Length - 1)));
            }

        }
        else if ((ColumnOrRafter == "Column")) {
            if ((CustomNonExpandable == "NonExpandable")) {
                LookupTbl = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl");
            }
            else {
                LookupTbl = LookupTblMatch(b, ColumnOrRafter, Location);
            }

            LookupHeight = (Application.WorksheetFunction.RoundUp(((tEdgeHeight / 12)
                            / 10), 0) * 10);
            if ((HorizontalReferenceDistance < (30 * 12))) {
                LookupHorizontalIndex = 1;
                //  default to 30' minimum for a given horizontal distance of less than 30'
            }
            else {
                LookupHorizontalIndex = ((((Application.WorksheetFunction.RoundUp(((HorizontalReferenceDistance / 12)
                                / 10), 0) * 10)
                            - 30)
                            / 10)
                            + 1);
            }

            if ((LookupHeight > 80)) {
                /* Warning! GOTO is not Implemented */}

            if ((LookupHeight < 20)) {
                LookupHeight = 20;
            }

            if ((LookupHorizontalIndex > 6)) {
                if (((Location == "e1")
                            || (Location == "e3"))) {
                    LookupHorizontalIndex = 6;
                }
                else {
                    // GoTo BadLookupData WHY IS S2 and S4 sending this to BADDATA??????????
                    LookupHorizontalIndex = 6;
                }

            }

            // With...
            Size = LookupTbl.DataBodyRange;
            LookupTbl.ListRows[LookupHorizontalIndex.ToString()].Index;
            LookupTbl.ListColumns;
            LookupHeight.ToString().Index;
            if (((Size.IndexOf("TS", 0) + 1)
                        != 0)) {
                Width = 4;
            }
            else if (((Size.IndexOf("W", 0) + 1)
                        != 0)) {
                Width = Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                - 1)).Substring((Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                    - 1)).Length - (Size.Substring(0, ((Size.IndexOf("x", 0) + 1)
                                    - 1)).Length - 1)));
            }

        }

        return;
        LookupTbl = null;
        /* Warning! Labeled Statements are not Implemented */if ((LookupHorizontalIndex > 80)) {
            MsgBox;
            "A horizontal lookup distance of greater than 80' has been calculated!";
            vbCritical;
            "Member Lookup Error";
        }
        else if ((LookupHorizontalIndex > 80)) {
            MsgBox;
            "A lookup height of greater than 80' has been calculated!";
            vbCritical;
            "Member Lookup Error";
        }

        return;
        /* Warning! Labeled Statements are not Implemented */MsgBox;
        "Member size lookup failed! Bad lookup string returned.";
        vbCritical;
        "Member Lookup Error";
    }

    // '''''''''''''''''''''''' Function that sets the correct steel lookup table
    private LookupTblMatch(b: clsBuilding, ColumnsOrRafters: string, Wall: string): ListObject {
        // Function Note: This does not properly handle non expandable endwall rafter lines, which should be set to either 8" receiver cee or 10" receiver cee depending on the length of the adjacent bay.
        // Warning!!! Optional parameters not supported
        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        // ''''''Columns or Rafters
        if ((ColumnsOrRafters == "Rafter")) {
            LookupTblMatch = SteelLookupSht.ListObjects("MainRafterAndExpandableEndwallRafterTbl");
        }
        else if ((ColumnsOrRafters == "Column")) {
            // ''''''''''''''''''' For columns, select table based off of walls ''''''''''''''''''''''''''
            switch (Wall) {
                case "s2":
                case "s4":
                    LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl");
                    break;
                case "e1":
                    if ((b.ExpandableEndwall("e1") == true)) {
                        LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl");
                    }
                    else if ((b.ExpandableEndwall("e1") == false)) {
                        LookupTblMatch = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl");
                    }

                    break;
                case "e3":
                    if ((b.ExpandableEndwall("e3") == true)) {
                        LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl");
                    }
                    else if ((b.ExpandableEndwall("e3") == false)) {
                        LookupTblMatch = SteelLookupSht.ListObjects("NonExpandableEndwallColumnTbl");
                    }

                    break;
                case "Interior":
                    LookupTblMatch = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl");
                    break;
            }

        }

    }

    public SetType(mType: void, mName: string) {
        // ''''''''mName is nonfunctional for now- this wil be taken from the Structural Steel Lookup tables. Unsure the best way to do this yet.
        // Warning!!! Optional parameters not supported
        switch (mType) {
            case "TS":
                Depth = 4;
                // this information is true, but it should be currently unused
                Width = 4;
                break;
            case "W-Beam":
                Depth = mName.Substring(0, ((mName.IndexOf("x", 0) + 1)
                                - 1)).Substring((mName.Substring(0, ((mName.IndexOf("x", 0) + 1)
                                    - 1)).Length - (mName.Substring(0, ((mName.IndexOf("x", 0) + 1)
                                    - 1)).Length - 1)));
                // mid(left(activecell.Value,instr(1,activecell.Value,"x")-1),len(left(activecell.Value,instr(1,activecell.Value,"x")-1))-1)
                break;
            case "8"" Receiver Cee":
                Width = 8;
                break;
            case "10"" Receiver Cee":
                Width = 10;
                break;
            case "C Purlin":
                break;
        }

    }

    private Class_Initialize() {
        // default qty to 1
        Qty = 1;
        clsType = "Member";
        ComponentMembers = new Collection();
        LoadBearing = false;
    }
