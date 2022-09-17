// TODO: Option Explicit ... Warning!!! not translated
let EaveStrutCount: number;


    ColTest() {
        let col: Collection = new Collection();
        let Panel: clsPanel = new clsPanel();
        Panel.Quantity = 1;
        Panel.PanelLength = 45;
        Panel.rEdgePosition = 3;
        col.Add;
        Panel;
        Panel = new clsPanel();
        Panel = col[1];
        Panel.rEdgePosition = 2;
        col.Add;
        Panel;
        Debug.Print;
        ".";
    }

    Test32() {
        Application.EnableEvents = true;
    }

    MoveExtensionOverhangMembers(b: clsBuilding) {
        let Member: clsMember;
        for (Member in b.e1Rafters) {
            if (Member.Placement) {
                "*Extension*";
                b.e1ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e3Rafters) {
            if (Member.Placement) {
                "*Extension*";
                b.e3ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e1Columns) {
            if (Member.Placement) {
                "*Extension*";
                b.e1ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e3Columns) {
            if (Member.Placement) {
                "*Extension*";
                b.e3ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.s2Columns) {
            if (Member.Placement) {
                "*Extension*";
                b.s2ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.s4Columns) {
            if (Member.Placement) {
                "*Extension*";
                b.s4ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e1Rafters) {
            if (Member.Placement) {
                "*Overhang*";
                b.e1ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e3Rafters) {
            if (Member.Placement) {
                "*Overhang*";
                b.e3ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e1Columns) {
            if (Member.Placement) {
                "*Overhang*";
                b.e1ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.e3Columns) {
            if (Member.Placement) {
                "*Overhang*";
                b.e3ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.s2Columns) {
            if (Member.Placement) {
                "*Overhang*";
                b.s2ExtensionMembers.Add;
                Member;
            }

        }

        for (Member in b.s4Columns) {
            if (Member.Placement) {
                "*Overhang*";
                b.s4ExtensionMembers.Add;
                Member;
            }

        }

    }

    WeldPlateGen(RafterLine: string, b: clsBuilding) {
        let Columns: Collection;
        let Column: clsMember;
        let WeldPlate: clsMiscItem;
        let WeldPlateRng: Range;
        let mCell: Range;
        WeldPlateRng = SteelLookupSht.Range("WeldPlateTblStart", "WeldPlateTblEnd");
        switch (RafterLine) {
            case "e1":
                Columns = b.e1Columns;
                break;
            case "e3":
                Columns = b.e3Columns;
                break;
            case "int":
                Columns = b.InteriorColumns;
                break;
        }

        for (Column in Columns) {
            WeldPlate = new clsMiscItem();
            WeldPlate.clsType = "Weld Plate";
            WeldPlate.Quantity = Column.Qty;
            for (mCell in WeldPlateRng) {
                if ((Column.Size == mCell.Value)) {
                    WeldPlate.Name = mCell.offset(0, 1).Value;
                    WeldPlate.Measurement = WeldPlate.Name;
                    WeldPlate.FootageCost = mCell.offset(0, 2).Value;
                    WeldPlate.Height = WeldPlate.Name.Substring((WeldPlate.Name.Length - (WeldPlate.Name.Length
                                    - (WeldPlate.Name.IndexOf("x", 0) + 1))));
                    WeldPlate.Width = WeldPlate.Name.Substring(0, ((WeldPlate.Name.IndexOf("x", 0) + 1)
                                    - 1));
                    Column.ComponentMembers.Add;
                    WeldPlate;
                }

            }

        }

    }

    CombineWeldPlates(b: clsBuilding) {
        let Column: clsMember;
        let WeldPlate: clsMiscItem;
        // With...
        for (Column in b.e1Columns) {
            for (WeldPlate in Column.ComponentMembers) {
                if ((WeldPlate.clsType == "Weld Plate")) {
                    b.WeldPlates.Add;
                    WeldPlate;
                }

            }

        }

        for (Column in b.e3Columns) {
            for (WeldPlate in Column.ComponentMembers) {
                if ((WeldPlate.clsType == "Weld Plate")) {
                    b.WeldPlates.Add;
                    WeldPlate;
                }

            }

        }

        for (Column in b.InteriorColumns) {
            for (WeldPlate in Column.ComponentMembers) {
                if ((WeldPlate.clsType == "Weld Plate")) {
                    b.WeldPlates.Add;
                    WeldPlate;
                }

            }

        }

        DuplicateMaterialRemoval(b.WeldPlates, "Misc");
    }

    //  ------------------- Sub used for testing ----------------------
    // currently being called from materialslistgen sub
    TestingSub2(b: clsBuilding) {
        Application.ScreenUpdating = false;
        EaveStrutCount = 0;
        let N: number;
        let SteelSht: Worksheet;
        let tempGirtsCollection: Collection;
        let manualGirtOptimization: Collection;
        let NewOptimizedCol: Collection;
        let GirtsCollection: Collection;
        let PurlinsCollection: Collection;
        let ReceiverCeeCollection: Collection;
        let EaveStrutCollection: Collection;
        let Member: clsMember;
        let Span: clsMember;
        let NewMember: clsMember;
        let RoofPurlins: Collection;
        let temp8RoofPurlins: Collection;
        let temp10RoofPurlins: Collection;
        let RoofPurlins8: Collection;
        let manualRoofPurlinOptimization: Collection;
        let NewRoofOptimizedCol: Collection;
        let RoofPurlins10: Collection;
        let temp8Receivers: Collection;
        let temp10Receivers: Collection;
        let Receivers8: Collection;
        let Receivers10: Collection;
        let EndwallColumns: Collection;
        let EndwallRafters: Collection;
        let SteelCollectionCls: clsSteelCollection;
        let NewSteelCollectionCls: clsSteelCollection;
        let IBeamClsCollection: Collection;
        let TSClsCollection: Collection;
        let IBeams: Collection;
        let TS: Collection;
        let Size: string;
        let Created: boolean;
        let FO: clsFO;
        let item: Object;
        let i: number;
        let NextSpanNum: number;
        if ((EstSht.Range("BayNum").Value > 1)) {
            IntColumnsGen;
            b;
            AdjustSidewallColumns(b, "s2");
            AdjustSidewallColumns(b, "s4");
        }

        EndwallColumnCLCalc(b, "e1");
        EndwallColumnCLCalc(b, "e3");
        FOJambsCalc(b, "e1");
        FOJambsCalc(b, "e3");
        FOJambsCalc(b, "s2");
        FOJambsCalc(b, "s4");
        if (b.ExpandableEndwall("e1")) {
            RemoveEndwallColumns(b, "e1");
        }

        if (b.ExpandableEndwall("e3")) {
            RemoveEndwallColumns(b, "e3");
        }

        RafterGen(b, "e1");
        RafterGen(b, "e3");
        RafterGen(b, "int");
        EndwallGirtLengthCalc(b, "e1");
        EndwallGirtLengthCalc(b, "e3");
        EndwallGirtLengthCalc(b, "s2");
        EndwallGirtLengthCalc(b, "s4");
        EaveStrutTypes(b, "s2");
        EaveStrutTypes(b, "s4");
        AdjustEndwallColumns(b, "e1");
        AdjustEndwallColumns(b, "e3");
        AdjustEndwallColumns(b, "Int");
        AdjustFOMembers(b, "e1");
        AdjustFOMembers(b, "e3");
        AdjustFOMembers(b, "s2");
        AdjustFOMembers(b, "s4");
        OverhangExtensionMembersGen(b);
        RoofPurlinGen(b);
        WeldPlateGen("e1", b);
        WeldPlateGen("e3", b);
        WeldPlateGen("int", b);
        CombineWeldPlates(b);
        BaseAngleTrimGen(b);
        AdditionalWeldClips(b, "e1");
        AdditionalWeldClips(b, "e3");
        AdditionalWeldClips(b, "s2");
        AdditionalWeldClips(b, "s4");
        FieldLocateFOCalc(b);
        // delete old output sheets
        Application.DisplayAlerts = false;
        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
            if ((ThisWorkbook.Sheets(N).Name == "Structural Steel Price List")) {
                ThisWorkbook.Sheets(N).Delete;
                break;
            }

        }

        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
            if ((ThisWorkbook.Sheets(N).Name == "Optimized Cut List")) {
                ThisWorkbook.Sheets(N).Delete;
                break;
            }

        }

        for (N = ThisWorkbook.Sheets.Count; (N <= 1); N = (N + -1)) {
            if ((ThisWorkbook.Sheets(N).Name == "Structural Steel Materials List")) {
                ThisWorkbook.Sheets(N).Delete;
                break;
            }

        }

        // '' Structural Steel Lists
        SteelCompleteMemberShtTmp.Copy;
        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Project Details");
        SteelSht = ThisWorkbook.Sheets("SteelCompleteMemberShtTmp (2)");
        // rename
        SteelSht.Name = "Optimized Cut List";
        SteelSht.Visible = xlSheetVisible;
        SteelMaterialsListTmp.Copy;
        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Project Details");
        SteelSht = ThisWorkbook.Sheets("SteelMaterialsListTmp (2)");
        // rename
        SteelSht.Name = "Structural Steel Materials List";
        SteelSht.Visible = xlSheetVisible;
        // set new output sheets
        SteelOutputShtTmp.Copy;
        /* Warning! Labeled Statements are not Implemented */ThisWorkbook.Sheets("Project Details");
        SteelSht = ThisWorkbook.Sheets("SteelOutputShtTmp (2)");
        // rename
        SteelSht.Name = "Structural Steel Price List";
        SteelSht.Visible = xlSheetVisible;
        Application.DisplayAlerts = true;
        SteelMaterialOutput(b);
        DrawItems;
        b;
        EstSht.Activate;
        GirtsCollection = new Collection();
        tempGirtsCollection = new Collection();
        manualGirtOptimization = new Collection();
        EaveStrutCollection = new Collection();
        temp8RoofPurlins = new Collection();
        manualRoofPurlinOptimization = new Collection();
        NewOptimizedCol = new Collection();
        temp10RoofPurlins = new Collection();
        RoofPurlins8 = new Collection();
        RoofPurlins10 = new Collection();
        NewRoofOptimizedCol = new Collection();
        RoofPurlins = new Collection();
        temp8Receivers = new Collection();
        temp10Receivers = new Collection();
        Receivers8 = new Collection();
        Receivers10 = new Collection();
        EndwallRafters = new Collection();
        EndwallColumns = new Collection();
        IBeamClsCollection = new Collection();
        TSClsCollection = new Collection();
        IBeams = new Collection();
        TS = new Collection();
        SteelCollectionCls = new clsSteelCollection();
        // Girts
        ParseGirts;
        b;
        tempGirtsCollection;
        b.e1Girts;
        EaveStrutCollection;
        manualGirtOptimization;
        ParseGirts;
        b;
        tempGirtsCollection;
        b.s2Girts;
        EaveStrutCollection;
        manualGirtOptimization;
        ParseGirts;
        b;
        tempGirtsCollection;
        b.e3Girts;
        EaveStrutCollection;
        manualGirtOptimization;
        ParseGirts;
        b;
        tempGirtsCollection;
        b.s4Girts;
        EaveStrutCollection;
        manualGirtOptimization;
        // FO Purlins
        ParseFOPurlins;
        tempGirtsCollection;
        temp8Receivers;
        b;
        b.e1FOs;
        manualGirtOptimization;
        ParseFOPurlins;
        tempGirtsCollection;
        temp8Receivers;
        b;
        b.s2FOs;
        manualGirtOptimization;
        ParseFOPurlins;
        tempGirtsCollection;
        temp8Receivers;
        b;
        b.e3FOs;
        manualGirtOptimization;
        ParseFOPurlins;
        tempGirtsCollection;
        temp8Receivers;
        b;
        b.s4FOs;
        manualGirtOptimization;
        ParseFOPurlins;
        tempGirtsCollection;
        temp8Receivers;
        b;
        b.fieldlocateFOs;
        manualGirtOptimization;
        // roof purlins - 8/10 C Purlins + Eave struts
        for (i = b.RoofPurlins.Count; (i <= 1); i = (i + -1)) {
            Member = b.RoofPurlins(i);
            if ((Member.Size == "8"" C Purlin")) {
                if ((Member.Length == (15 * 12))) {
                    manualRoofPurlinOptimization.Add;
                    Member;
                }
                else {
                    temp8RoofPurlins.Add;
                    Member;
                }

            }
            else if ((Member.Size == "10"" C Purlin")) {
                temp10RoofPurlins.Add;
                Member;
            }
            else if ((Member.mType == "Eave Strut")) {
                EaveStrutCollection.Add;
                Member;
                b.RoofPurlins.Remove(i);
            }

        }

        // Rafters and Columns
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.e1Columns;
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.e3Columns;
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.e1Rafters;
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.e3Rafters;
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.InteriorColumns;
        ParseColumns;
        tempGirtsCollection;
        manualGirtOptimization;
        temp8Receivers;
        temp10Receivers;
        EndwallColumns;
        IBeams;
        TS;
        b.intRafters;
        CombinePurlins;
        b;
        manualGirtOptimization;
        tempGirtsCollection;
        NewOptimizedCol;
        CombinePurlins;
        b;
        manualRoofPurlinOptimization;
        temp8RoofPurlins;
        NewRoofOptimizedCol;
        // all Ibeams are in IBeams
        // all Tube Steel are in TS
        // create class for each size Ibeam
        for (Member in IBeams) {
            Created = false;
            Size = Member.Size;
            for (SteelCollectionCls in IBeamClsCollection) {
                if ((SteelCollectionCls.Size == Size)) {
                    // already created
                    Created = true;
                    SteelCollectionCls.Members.Add;
                    Member;
                }

            }

            if ((Created == false)) {
                NewSteelCollectionCls = new clsSteelCollection();
                NewSteelCollectionCls.Size = Member.Size;
                NewSteelCollectionCls.Members.Add;
                Member;
                IBeamClsCollection.Add;
                NewSteelCollectionCls;
            }

        }

        // create class for each size TS
        for (Member in TS) {
            Created = false;
            Size = Member.Size;
            for (SteelCollectionCls in TSClsCollection) {
                if ((SteelCollectionCls.Size == Size)) {
                    // already created
                    Created = true;
                    SteelCollectionCls.Members.Add;
                    Member;
                }

            }

            if ((Created == false)) {
                NewSteelCollectionCls = new clsSteelCollection();
                NewSteelCollectionCls.Size = Member.Size;
                NewSteelCollectionCls.Members.Add;
                Member;
                TSClsCollection.Add;
                NewSteelCollectionCls;
            }

        }

        DuplicateMaterialRemoval(tempGirtsCollection, "Steel");
        DuplicateMaterialRemoval(temp8RoofPurlins, "Steel");
        DuplicateMaterialRemoval(temp10RoofPurlins, "Steel");
        DuplicateMaterialRemoval(temp8Receivers, "Steel");
        DuplicateMaterialRemoval(temp10Receivers, "Steel");
        if ((tempGirtsCollection.Count > 0)) {
            JankyBPPSolver.BPP_Solver(GirtsCollection, tempGirtsCollection, "Girt", "Steel", "e1");
        }

        if ((temp8RoofPurlins.Count > 0)) {
            JankyBPPSolver.BPP_Solver(RoofPurlins8, temp8RoofPurlins, "Roof Purlin", "Steel", "e1");
        }

        if ((temp10RoofPurlins.Count > 0)) {
            JankyBPPSolver.BPP_Solver(RoofPurlins10, temp10RoofPurlins, "Roof Purlin", "Steel", "e1");
        }

        if ((temp8Receivers.Count > 0)) {
            JankyBPPSolver.BPP_Solver(Receivers8, temp8Receivers, "Girt", "Steel", "e1");
        }

        if ((temp10Receivers.Count > 0)) {
            JankyBPPSolver.BPP_Solver(Receivers10, temp10Receivers, "Girt", "Steel", "e1");
        }

        TS = new Collection();
        IBeams = new Collection();
        for (SteelCollectionCls in TSClsCollection) {
            JankyBPPSolver.BPP_Solver(TS, SteelCollectionCls.Members, "TS", "Steel", "e1");
        }

        for (SteelCollectionCls in IBeamClsCollection) {
            JankyBPPSolver.BPP_Solver(IBeams, SteelCollectionCls.Members, "IBeam", "Steel", "e1");
        }

        DuplicateMaterialRemoval(EaveStrutCollection, "Steel");
        // rename all members passed through BPPSolver
        for (Member in GirtsCollection) {
            Member.Size = "8"" C Purlin";
            Member.mType = "C Purlin";
            for (Span in Member.ComponentMembers) {
                Span.Size = "8"" C Purlin";
            }

        }

        for (Member in RoofPurlins8) {
            Member.Size = "8"" C Purlin";
            Member.mType = "C Purlin";
            for (Span in Member.ComponentMembers) {
                Span.Size = "8"" C Purlin";
            }

        }

        for (Member in RoofPurlins10) {
            Member.Size = "10"" C Purlin";
            Member.mType = "C Purlin";
            for (Span in Member.ComponentMembers) {
                Span.Size = "10"" C Purlin";
            }

        }

        for (Member in Receivers8) {
            Member.Size = "8"" Receiver Cee";
            for (Span in Member.ComponentMembers) {
                Span.Size = "8"" Receiver Cee";
            }

        }

        for (Member in Receivers10) {
            Member.Size = "10"" Receiver Cee";
            for (Span in Member.ComponentMembers) {
                Span.Size = "10"" Receiver Cee";
            }

        }

        // Add combined purlins back to Girts Collection
        NextSpanNum = (GirtsCollection.Count + 1);
        for (Member in NewOptimizedCol) {
            Member.Placement = ("8"" C Purlin Span #" + NextSpanNum);
            Member.mType = "C Purlin";
            NextSpanNum = (NextSpanNum + 1);
            GirtsCollection.Add;
            Member;
        }

        NextSpanNum = (RoofPurlins8.Count + 1);
        for (Member in NewRoofOptimizedCol) {
            Member.Placement = ("Roof Purlin 8"" C Purlin Span #" + NextSpanNum);
            Member.mType = "C Purlin";
            NextSpanNum = (NextSpanNum + 1);
            RoofPurlins8.Add;
            Member;
        }

        // '''''''''Cut List Output
        CutListOutput(IBeams, "I Beams: Columns & Rafters");
        CutListOutput(TS, "Tube Steel: Columns & Rafters");
        CutListOutput(GirtsCollection, "Wall Girt");
        CutListOutput(RoofPurlins8, "8"" Roof Purlin");
        CutListOutput(RoofPurlins10, "10"" Roof Purlin");
        CutListOutput(Receivers8, "8"" Receiver Cee");
        CutListOutput(Receivers10, "10"" Receiver Cee");
        // '''''''''Price List Output
        // Call SteelPriceOutput(b.e1Columns, "endwall 1 column")
        // Call SteelPriceOutput(b.e3Columns, "endwall 3 column")
        // Call SteelPriceOutput(EndwallColumns, "Endwall Column")
        // Call SteelPriceOutput(b.InteriorColumns, "Main Rafter Line Column")
        // Call SteelPriceOutput(b.e1Rafters, "endwall 1 rafter")
        // Call SteelPriceOutput(b.e3Rafters, "endwall 3 rafter")
        // Call SteelPriceOutput(EndwallRafters, "Endwall Rafter")
        // Call SteelPriceOutput(b.intRafters, "Main Rafter")
        SteelPriceOutput(IBeams, "I Beams: Columns & Rafters");
        SteelPriceOutput(TS, "Tube Steel: Columns & Rafters");
        SteelPriceOutput(GirtsCollection, "Wall Girts, Non-Load Bearing Columns, Endwall Rafters, and FO Members");
        SteelPriceOutput(RoofPurlins8, "8"" Roof Purlins");
        SteelPriceOutput(RoofPurlins10, "10"" Roof Purlins");
        // Call SteelPriceOutput(b.e1Girts, "e1 Wall Girt")
        // Call SteelPriceOutput(b.s2Girts, "s2 Wall Girt")
        // Call SteelPriceOutput(b.e3Girts, "e3 Wall Girt")
        // Call SteelPriceOutput(b.s4Girts, "s4 Wall Girt")
        // Call SteelPriceOutput(b.RoofPurlins, "Roof Purlin")
        SteelPriceOutput(EaveStrutCollection, "Eave Strut");
        SteelPriceOutput(Receivers8, "8"" Endwall Rafters and FO Jambs");
        SteelPriceOutput(Receivers10, "10"" Endwall Rafters");
        // Call SteelPriceOutput(b.e1FOs, "FO", True)
        // Call SteelPriceOutput(b.s2FOs, "FO", True)
        // Call SteelPriceOutput(b.e3FOs, "FO", True)
        // Call SteelPriceOutput(b.s4FOs, "FO", True)
        SteelPriceOutput(b.e1OverhangMembers, "e1 gable overhang");
        SteelPriceOutput(b.e1ExtensionMembers, "e1 gable extension");
        SteelPriceOutput(b.s2OverhangMembers, "s2 eave overhang");
        SteelPriceOutput(b.s2ExtensionMembers, "s2 eave extension");
        SteelPriceOutput(b.e3OverhangMembers, "e3 gable overhang");
        SteelPriceOutput(b.e3ExtensionMembers, "e3 gable extension");
        SteelPriceOutput(b.s4OverhangMembers, "s4 eave overhang");
        SteelPriceOutput(b.s4ExtensionMembers, "s4 eave extension");
        SteelPriceOutput(b.BaseAngleTrim, "Base Angle");
        // ''''''''''''''''''Add Weld Clips to Output Sheet
        let FullMemberSht: Worksheet;
        let LastRow: number;
        FullMemberSht = ThisWorkbook.Sheets("Structural Steel Price List");
        if ((FullMemberSht.Range("A4").Value == "")) {
            LastRow = 4;
        }
        else {
            LastRow = FullMemberSht.Range("A3").End(xlDown).offset(1, 0).Row;
        }

        // With...
        // sum total structural steel costs
        "each".Range(("G" + LastRow)).Value = (1.57 * b.WeldClips);
        1.57.Range(("F" + LastRow)).Value = (1.57 * b.WeldClips);
        "n/a".Range(("E" + LastRow)).Value = (1.57 * b.WeldClips);
        "n/a".Range(("D" + LastRow)).Value = (1.57 * b.WeldClips);
        "Weld Clips".Range(("C" + LastRow)).Value = (1.57 * b.WeldClips);
        b.WeldClips.Range(("B" + LastRow)).Value = (1.57 * b.WeldClips);
        FullMemberSht.Range(("A" + LastRow)).Value = (1.57 * b.WeldClips);
        // report later in cost estimate sub
        for (i = 4; (i <= LastRow); i++) {
            if (IsNumeric(., Range(("H" + i)).Value)) {
                Range(("H" + i)).Value;
            }

        }

        // ''''''''''''''''''Add Weld Plates to Output Sheet
        let WeldPlate: clsMiscItem;
        let WeldPlateRng: Range;
        LastRow = (LastRow + 1);
        // With...
        for (WeldPlate in // TODO: Warning!!!! NULL EXPRESSION DETECTED...
        ) {
            LastRow = (LastRow + 1);
            "each".Range(("G" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
            WeldPlate.FootageCost.Range(("F" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
            WeldPlate.Name.Range(("E" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
            "n/a".Range(("D" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
            (WeldPlate.Name + " Weld Plate".Range(("C" + LastRow)).Value) = (WeldPlate.FootageCost * WeldPlate.Quantity);
            WeldPlate.Quantity.Range(("B" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
            b.WeldPlates.Range(("A" + LastRow)).Value = (WeldPlate.FootageCost * WeldPlate.Quantity);
        }

        Application.ScreenUpdating = true;
    }

    private ParseColumns(/* ref */tempGirtsCollection: Collection, /* ref */manualGirtOptimization: Collection, /* ref */temp8Receivers: Collection, /* ref */temp10Receivers: Collection, /* ref */EndwallColumns: Collection, IBeams: Collection, TS: Collection, /* ref */MemberCollection: Collection) {
        let Member: clsMember;
        // columns - receiver cees + C Purlins
        for (Member in MemberCollection) {
            if ((Member.Size == "8"" Receiver Cee")) {
                temp8Receivers.Add;
                Member;
            }
            else if ((Member.Size == "10"" Receiver Cee")) {
                temp10Receivers.Add;
                Member;
            }
            else if ((Member.Size == "8"" C Purlin")) {
                if ((Member.Length == (15 * 12))) {
                    manualGirtOptimization.Add;
                    Member;
                }
                else {
                    tempGirtsCollection.Add;
                    Member;
                }

            }
            else if (Member.Size) {
                "W*";
                IBeams.Add;
                Member;
            }
            else if (Member.Size) {
                "TS*";
                TS.Add;
                Member;
            }

        }

    }

    ParseFOPurlins(/* ref */tempGirtsCollection: Collection, /* ref */temp8Receivers: Collection, /* ref */b: clsBuilding, FOCollection: Collection, manualGirtOptimization: Collection) {
        // FO - Purlins and Receivers
        let FO: clsFO;
        let Member: clsMember;
        let i: number;
        for (FO in FOCollection) {
            for (i = FO.FOMaterials.Count; (i <= 1); i = (i + -1)) {
                Member = FO.FOMaterials(i);
                if ((Member.Size == "8"" C Purlin")) {
                    if ((Member.Length == (15 * 12))) {
                        manualGirtOptimization.Add;
                        Member;
                    }
                    else {
                        tempGirtsCollection.Add;
                        Member;
                    }

                }
                else if ((Member.Size == "8"" Receiver Cee")) {
                    temp8Receivers.Add;
                    Member;
                }

            }

        }

    }

    ParseGirts(b: clsBuilding, tempGirtsCollection: Collection, buildingGirts: Collection, EaveStrutCollection: Collection, manualGirtOptimization: Collection) {
        // Girts optimization collection
        let Member: clsMember;
        let i: number;
        for (i = buildingGirts.Count; (i <= 1); i = (i + -1)) {
            Member = buildingGirts(i);
            if ((Member.Size == "8"" C Purlin")) {
                if ((Member.Length == (15 * 12))) {
                    manualGirtOptimization.Add;
                    Member;
                }
                else {
                    tempGirtsCollection.Add;
                    Member;
                }

            }
            else {
                EaveStrutCollection.Add;
                Member;
                buildingGirts.Remove(i);
            }

        }

    }

    CombinePurlins(b: clsBuilding, manualGirtOptimization: Collection, tempGirtsCollection: Collection, NewOptimizedCol: Collection) {
        let FifteenFootPurlinCount: number;
        let FirstMember: clsMember;
        let SecondMember: clsMember;
        let FullSpan: clsMember;
        let i: number;
        let Member: clsMember;
        FifteenFootPurlinCount = manualGirtOptimization.Count;
        if (((FifteenFootPurlinCount % 2)
                    == 0)) {
            // even number, do nothing
        }
        else {
            // odd number, move 1 member to temp girt collection
            Member = manualGirtOptimization(1);
            tempGirtsCollection.Add;
            Member;
            manualGirtOptimization.Remove(1);
        }

        // combine members
        FifteenFootPurlinCount = manualGirtOptimization.Count;
        i = 1;
        while ((i < FifteenFootPurlinCount)) {
            FirstMember = manualGirtOptimization(i);
            SecondMember = manualGirtOptimization((i + 1));
            FullSpan = new clsMember();
            FullSpan.ComponentMembers.Add;
            FirstMember;
            FullSpan.ComponentMembers.Add;
            SecondMember;
            FullSpan.Size = "8"" C Purlin";
            FullSpan.Length = (30 * 12);
            NewOptimizedCol.Add;
            FullSpan;
            i = (i + 2);
        }

    }

    SteelMaterialOutput(b: clsBuilding) {
        let CurRow: number;
        let Member: clsMember;
        let SteelSht: Worksheet;
        let FullMemberSht: Worksheet;
        let FO: clsFO;
        let item: Object;
        let j: number;
        let i: number;
        let UnitPrice: number;
        let UnitMeasure: string;
        let UnitValue: number;
        let PriceTbl: ListObject;
        let ExtColRow: number;
        let ExtRafterRow: number;
        let SortRange: Range;
        let KeyRange: Range;
        let EaveStrutRow: number;
        // ''''''''''''''''''''''''''Steel Material OUtput Sheet
        SteelSht = ThisWorkbook.Sheets("Structural Steel Materials List");
        SteelSht.Activate;
        ExtColRow = 0;
        // e1 Columns
        // With...
        i = 0;
        for (Member in b.e1Columns) {
            if (!Member.Placement) {
                ("*Extension*"
                            & !Member.Placement);
                "*Overhang*";
                SteelSht.Range("e1_ColumnsStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("e1_ColumnsStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("e1_ColumnsStart").offset, (i + 1), 0].Row;
                SteelSht.Range("e1_ColumnsStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("Ext_ColumnsStart").offset((ExtColRow + 1), 0).Row].Insert;
                ExtColRow = (ExtColRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("e1_ColumnsStart").Resize(i, 5);
            KeyRange = SteelSht.Range("e1_ColumnsStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        // e3 Columns
        // With...
        i = 0;
        for (Member in b.e3Columns) {
            if (!Member.Placement) {
                ("*Extension*"
                            & !Member.Placement);
                "*Overhang*";
                SteelSht.Range("e3_ColumnsStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("e3_ColumnsStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("e3_ColumnsStart").offset, (i + 1), 0].Row;
                SteelSht.Range("e3_ColumnsStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("Ext_ColumnsStart").offset((ExtColRow + 1), 0).Row].Insert;
                ExtColRow = (ExtColRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("e3_ColumnsStart").Resize(i, 5);
            KeyRange = SteelSht.Range("e3_ColumnsStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        // Int Columns
        // With...
        i = 0;
        for (Member in b.InteriorColumns) {
            if (!Member.Placement) {
                ("*Extension*"
                            & !Member.Placement);
                "*Overhang*";
                SteelSht.Range("Int_ColumnsStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("Int_ColumnsStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("Int_ColumnsStart").offset, (i + 1), 0].Row;
                SteelSht.Range("Int_ColumnsStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement;
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("Ext_ColumnsStart").offset((ExtColRow + 1), 0).Row].Insert;
                ExtColRow = (ExtColRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("Int_ColumnsStart").Resize(i, 5);
            KeyRange = SteelSht.Range("Int_ColumnsStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        if ((ExtColRow > 0)) {
            SortRange = SteelSht.Range("Ext_ColumnsStart").Resize(ExtColRow, 5);
            KeyRange = SteelSht.Range("Ext_ColumnsStart").offset(0, 4).Resize(ExtColRow, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        ExtRafterRow = 0;
        // e1 Rafters
        // With...
        i = 0;
        for (Member in b.e1Rafters) {
            if (!Member.Placement) {
                ("*Extension*"
                            & !Member.Placement);
                ("*Overhang*"
                            & !Member.Placement);
                "*Stub*";
                SteelSht.Range("e1_RaftersStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("e1_RaftersStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("e1_RaftersStart").offset, (i + 1), 0].Row;
                SteelSht.Range("e1_RaftersStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("Ext_RaftersStart").offset((ExtRafterRow + 1), 0).Row].Insert;
                ExtRafterRow = (ExtRafterRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("e1_RaftersStart").Resize(i, 5);
            KeyRange = SteelSht.Range("e1_RaftersStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        // e3 Rafters
        // With...
        i = 0;
        for (Member in b.e3Rafters) {
            if (!Member.Placement) {
                ("*Extension*"
                            & !Member.Placement);
                ("*Overhang*"
                            & !Member.Placement);
                "*Stub*";
                SteelSht.Range("e3_RaftersStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("e3_RaftersStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("e3_RaftersStart").offset, (i + 1), 0].Row;
                SteelSht.Range("e3_RaftersStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement;
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("Ext_RaftersStart").offset((ExtRafterRow + 1), 0).Row].Insert;
                ExtRafterRow = (ExtRafterRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("e3_RaftersStart").Resize(i, 5);
            KeyRange = SteelSht.Range("e3_RaftersStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        if ((b.intRafters.Count > 0)) {
            // Int Rafters
            // With...
            i = 0;
            for (Member in b.intRafters) {
                if (!Member.Placement) {
                    ("*Extension*"
                                & !Member.Placement);
                    ("*Overhang*"
                                & !Member.Placement);
                    "*Stub*";
                    SteelSht.Range("Int_RaftersStart").offset;
                    i;
                    0;
                    ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                    Member.Placement.offset(i, 3).Value = Member.Length;
                    Member.Size.offset(i, 2).Value = Member.Length;
                    Member.Qty.offset(i, 1).Value = Member.Length;
                    SteelSht.Range("Int_RaftersStart").Value = Member.Length;
                    SteelSht.Rows[SteelSht.Range("Int_RaftersStart").offset, (i + 1), 0].Row;
                    SteelSht.Range("Int_RaftersStart").Insert;
                    i = (i + 1);
                }
                else {
                    SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty;
                    SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size;
                    SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement;
                    SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                    SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length;
                    // add next row
                    SteelSht.Rows[SteelSht.Range("Ext_RaftersStart").offset((ExtRafterRow + 1), 0).Row].Insert;
                    ExtRafterRow = (ExtRafterRow + 1);
                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("Int_RaftersStart").Resize(i, 5);
                KeyRange = SteelSht.Range("Int_RaftersStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((ExtRafterRow > 0)) {
            SortRange = SteelSht.Range("Ext_RaftersStart").Resize(ExtRafterRow, 5);
            KeyRange = SteelSht.Range("Ext_RaftersStart").offset(0, 4).Resize(ExtRafterRow, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        EaveStrutRow = 0;
        if ((b.e1Girts.Count > 0)) {
            // e1 Girts
            // With...
            i = 0;
            for (Member in b.e1Girts) {
                if (!Member.mType) {
                    "*Eave Strut*";
                    SteelSht.Range("e1_GirtsStart").offset;
                    i;
                    0;
                    ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                    Member.Placement.offset(i, 3).Value = Member.Length;
                    Member.Size.offset(i, 2).Value = Member.Length;
                    Member.Qty.offset(i, 1).Value = Member.Length;
                    SteelSht.Range("e1_GirtsStart").Value = Member.Length;
                    SteelSht.Rows[SteelSht.Range("e1_GirtsStart").offset, (i + 1), 0].Row;
                    SteelSht.Range("e1_GirtsStart").Insert;
                    i = (i + 1);
                }
                else {
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length;
                    // add next row
                    SteelSht.Rows[SteelSht.Range("EaveStrutsStart").offset((EaveStrutRow + 1), 0).Row].Insert;
                    EaveStrutRow = (EaveStrutRow + 1);
                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("e1_GirtsStart").Resize(i, 5);
                KeyRange = SteelSht.Range("e1_GirtsStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.s2Girts.Count > 0)) {
            // s2 Girts
            // With...
            i = 0;
            for (Member in b.s2Girts) {
                if (!Member.mType) {
                    "*Eave Strut*";
                    SteelSht.Range("s2_GirtsStart").offset;
                    i;
                    0;
                    ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                    Member.Placement.offset(i, 3).Value = Member.Length;
                    Member.Size.offset(i, 2).Value = Member.Length;
                    Member.Qty.offset(i, 1).Value = Member.Length;
                    SteelSht.Range("s2_GirtsStart").Value = Member.Length;
                    SteelSht.Rows[SteelSht.Range("s2_GirtsStart").offset, (i + 1), 0].Row;
                    SteelSht.Range("s2_GirtsStart").Insert;
                    i = (i + 1);
                }
                else {
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length;
                    // add next row
                    SteelSht.Rows[SteelSht.Range("EaveStrutsStart").offset((EaveStrutRow + 1), 0).Row].Insert;
                    EaveStrutRow = (EaveStrutRow + 1);
                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("s2_GirtsStart").Resize(i, 5);
                KeyRange = SteelSht.Range("s2_GirtsStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.e3Girts.Count > 0)) {
            // e3 Girts
            // With...
            i = 0;
            for (Member in b.e3Girts) {
                if (!Member.mType) {
                    "*Eave Strut*";
                    SteelSht.Range("e3_GirtsStart").offset;
                    i;
                    0;
                    ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                    Member.Placement.offset(i, 3).Value = Member.Length;
                    Member.Size.offset(i, 2).Value = Member.Length;
                    Member.Qty.offset(i, 1).Value = Member.Length;
                    SteelSht.Range("e3_GirtsStart").Value = Member.Length;
                    SteelSht.Rows[SteelSht.Range("e3_GirtsStart").offset, (i + 1), 0].Row;
                    SteelSht.Range("e3_GirtsStart").Insert;
                    i = (i + 1);
                }
                else {
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length;
                    // add next row
                    SteelSht.Rows[SteelSht.Range("EaveStrutsStart").offset((EaveStrutRow + 1), 0).Row].Insert;
                    EaveStrutRow = (EaveStrutRow + 1);
                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("e3_GirtsStart").Resize(i, 5);
                KeyRange = SteelSht.Range("e3_GirtsStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.s4Girts.Count > 0)) {
            // s4 Girts
            // With...
            i = 0;
            for (Member in b.s4Girts) {
                if (!Member.mType) {
                    "*Eave Strut*";
                    SteelSht.Range("s4_GirtsStart").offset;
                    i;
                    0;
                    ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                    Member.Placement.offset(i, 3).Value = Member.Length;
                    Member.Size.offset(i, 2).Value = Member.Length;
                    Member.Qty.offset(i, 1).Value = Member.Length;
                    SteelSht.Range("s4_GirtsStart").Value = Member.Length;
                    SteelSht.Rows[SteelSht.Range("s4_GirtsStart").offset, (i + 1), 0].Row;
                    SteelSht.Range("s4_GirtsStart").Insert;
                    i = (i + 1);
                }
                else {
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement;
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                    SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length;
                    // add next row
                    SteelSht.Rows[SteelSht.Range("EaveStrutsStart").offset((EaveStrutRow + 1), 0).Row].Insert;
                    EaveStrutRow = (EaveStrutRow + 1);
                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("s4_GirtsStart").Resize(i, 5);
                KeyRange = SteelSht.Range("s4_GirtsStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        // roof purlins / Eave Struts
        // With...
        i = 0;
        for (Member in b.RoofPurlins) {
            if (!Member.mType) {
                "*Eave Strut*";
                SteelSht.Range("RoofPurlinsStart").offset;
                i;
                0;
                ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                SteelSht.Range("RoofPurlinsStart").Value = Member.Length;
                SteelSht.Rows[SteelSht.Range("RoofPurlinsStart").offset, (i + 1), 0].Row;
                SteelSht.Range("RoofPurlinsStart").Insert;
                i = (i + 1);
            }
            else {
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty;
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size;
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement;
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length;
                // add next row
                SteelSht.Rows[SteelSht.Range("EaveStrutsStart").offset((EaveStrutRow + 1), 0).Row].Insert;
                EaveStrutRow = (EaveStrutRow + 1);
            }

        }

        if ((i > 0)) {
            SortRange = SteelSht.Range("RoofPurlinsStart").Resize(i, 5);
            KeyRange = SteelSht.Range("RoofPurlinsStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        if ((EaveStrutRow > 0)) {
            SortRange = SteelSht.Range("EaveStrutsStart").Resize(EaveStrutRow, 5);
            KeyRange = SteelSht.Range("EaveStrutsStart").offset(0, 4).Resize(EaveStrutRow, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        if ((b.e1FOs.Count > 0)) {
            // e1 FOs
            // With...
            i = 0;
            for (FO in b.e1FOs) {
                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                        Member.Placement.offset(i, 3).Value = Member.Length;
                        Member.Size.offset(i, 2).Value = Member.Length;
                        Member.Qty.offset(i, 1).Value = Member.Length;
                        item.offset(i, 0).Value = Member.Length;
                        Member = Member.Length;
                        SteelSht.Rows[SteelSht.Range("e1_FOStart").offset, (i + 1), 0].Row;
                        SteelSht.Range("e1_FOStart").Insert;
                        i = (i + 1);
                    }

                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("e1_FOStart").Resize(i, 5);
                KeyRange = SteelSht.Range("e1_FOStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.s2FOs.Count > 0)) {
            // s2 FOs
            // With...
            i = 0;
            for (FO in b.s2FOs) {
                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                        Member.Placement.offset(i, 3).Value = Member.Length;
                        Member.Size.offset(i, 2).Value = Member.Length;
                        Member.Qty.offset(i, 1).Value = Member.Length;
                        item.offset(i, 0).Value = Member.Length;
                        Member = Member.Length;
                        SteelSht.Rows[SteelSht.Range("s2_FOStart").offset, (i + 1), 0].Row;
                        SteelSht.Range("s2_FOStart").Insert;
                        i = (i + 1);
                    }

                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("s2_FOStart").Resize((i + 1), 5);
                KeyRange = SteelSht.Range("s2_FOStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.e3FOs.Count > 0)) {
            // e3 FOs
            // With...
            i = 0;
            for (FO in b.e3FOs) {
                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                        Member.Placement.offset(i, 3).Value = Member.Length;
                        Member.Size.offset(i, 2).Value = Member.Length;
                        Member.Qty.offset(i, 1).Value = Member.Length;
                        item.offset(i, 0).Value = Member.Length;
                        Member = Member.Length;
                        SteelSht.Rows[SteelSht.Range("e3_FOStart").offset, (i + 1), 0].Row;
                        SteelSht.Range("e3_FOStart").Insert;
                        i = (i + 1);
                    }

                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("e3_FOStart").Resize(i, 5);
                KeyRange = SteelSht.Range("e3_FOStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.s4FOs.Count > 0)) {
            // s4 FOs
            // With...
            i = 0;
            for (FO in b.s4FOs) {
                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                        Member.Placement.offset(i, 3).Value = Member.Length;
                        Member.Size.offset(i, 2).Value = Member.Length;
                        Member.Qty.offset(i, 1).Value = Member.Length;
                        item.offset(i, 0).Value = Member.Length;
                        Member = Member.Length;
                        SteelSht.Rows[SteelSht.Range("s4_FOStart").offset, (i + 1), 0].Row;
                        SteelSht.Range("s4_FOStart").Insert;
                        i = (i + 1);
                    }

                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("s4_FOStart").Resize(i, 5);
                KeyRange = SteelSht.Range("s4_FOStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.fieldlocateFOs.Count > 0)) {
            // s4 FOs
            // With...
            i = 0;
            for (FO in b.fieldlocateFOs) {
                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        ImperialMeasurementFormat(Member.Length).offset(i, 4).Value = Member.Length;
                        Member.Placement.offset(i, 3).Value = Member.Length;
                        Member.Size.offset(i, 2).Value = Member.Length;
                        Member.Qty.offset(i, 1).Value = Member.Length;
                        item.offset(i, 0).Value = Member.Length;
                        Member = Member.Length;
                        SteelSht.Rows[SteelSht.Range("FieldLocate_FOStart").offset, (i + 1), 0].Row;
                        SteelSht.Range("FieldLocate_FOStart").Insert;
                        i = (i + 1);
                    }

                }

            }

            if ((i > 0)) {
                SortRange = SteelSht.Range("FieldLocate_FOStart").Resize(i, 5);
                KeyRange = SteelSht.Range("FieldLocate_FOStart").offset(0, 4).Resize(i, 1);
                SortRange.Select;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
                ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
                /* Warning! Labeled Statements are not Implemented */KeyRange;
                /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
                /* Warning! Labeled Statements are not Implemented */xlAscending;
                /* Warning! Labeled Statements are not Implemented */xlSortNormal;
                // With...
                xlTopToBottom.SortMethod = xlPinYin.Apply;
                false.Orientation = xlPinYin.Apply;
                xlNo.MatchCase = xlPinYin.Apply;
                SortRange.Header = xlPinYin.Apply;
            }

        }

        if ((b.BaseAngleTrim.Count > 0)) {
            // Base Angle
            // With...
            i = 0;
            for (Member in // TODO: Warning!!!! NULL EXPRESSION DETECTED...
            ) {
                SteelSht.Rows[SteelSht.Range("BaseAngleStart").offset, (i + 1), 0].Row;
                Member.Placement.offset(i, 3).Value = Member.Length;
                Member.Size.offset(i, 2).Value = Member.Length;
                Member.Qty.offset(i, 1).Value = Member.Length;
                b.BaseAngleTrim.offset(i, 0).Value = Member.Length;
                SteelSht.Range("BaseAngleStart").Insert;
                i = (i + 1);
            }

            SortRange = SteelSht.Range("BaseAngleStart").Resize(i, 5);
            KeyRange = SteelSht.Range("BaseAngleStart").offset(0, 4).Resize(i, 1);
            SortRange.Select;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Clear;
            ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields.Add2;
            /* Warning! Labeled Statements are not Implemented */KeyRange;
            /* Warning! Labeled Statements are not Implemented */xlSortOnValues;
            /* Warning! Labeled Statements are not Implemented */xlAscending;
            /* Warning! Labeled Statements are not Implemented */xlSortNormal;
            // With...
            xlTopToBottom.SortMethod = xlPinYin.Apply;
            false.Orientation = xlPinYin.Apply;
            xlNo.MatchCase = xlPinYin.Apply;
            SortRange.Header = xlPinYin.Apply;
        }

        // Weld Clips
        SteelSht.Range("WeldClipsStart").Value = b.WeldClips;
        let WeldPlate: clsMiscItem;
        // Weld Plates
        if ((b.WeldPlates.Count > 0)) {
            // With...
            i = 0;
            for (WeldPlate in // TODO: Warning!!!! NULL EXPRESSION DETECTED...
            ) {
                // add next row
                b.WeldPlates.offset(i, 0).Value = WeldPlate.Name;
                SteelSht.Rows[SteelSht.Range("WeldPlateStart").offset, (i + 1), 0].Row;
                SteelSht.Range("WeldPlateStart").Insert;
                i = (i + 1);
            }

        }

    }
AdditionalWeldClips(b: clsBuilding, eWall: string) {
    let MemberCollection: Collection;
    let FOCollection: Collection;
    let Member: clsMember;
    let Jamb: clsMember;
    let item: Object;
    let FO: clsFO;
    let WeldClips: number;
    let RightClip: number;
    let LeftClip: number;
    WeldClips = b.WeldClips;
    switch (eWall) {
        case "e1":
            MemberCollection = b.e1Girts;
            FOCollection = b.e1FOs;
            if ((b.rShape == "Gable")) {
                WeldClips = (WeldClips + 1);
                if (!b.ExpandableEndwall("e1")) {
                    WeldClips = (WeldClips + 1);
                }

            }

            break;
        case "e3":
            MemberCollection = b.e3Girts;
            FOCollection = b.e3FOs;
            if (!b.ExpandableEndwall("e3")) {
                WeldClips = (WeldClips + 1);
            }

            break;
        case "s2":
            MemberCollection = b.s2Girts;
            FOCollection = b.s2FOs;
            break;
        case "s4":
            MemberCollection = b.s4Girts;
            FOCollection = b.s4FOs;
            break;
    }

    WeldClips = b.WeldClips;
    // ''''''''''''''''''''''''''''''''''roof purlin Weld Clips are added in Roof Purlin Gen'''''''''''''''''''''''''''''''''''
    // Add Weld Clips to Wall Girts
    for (Member in MemberCollection) {
        RightClip = 1;
        LeftClip = 1;
        if (((Member.Size == "8"" C Purlin")
                    || (Member.Size == "10"" C Purlin"))) {
            for (FO in FOCollection) {
                if (((Member.rEdgePosition >= FO.rEdgePosition)
                            && (((Member.lEdgePosition - Member.Length)
                            <= FO.rEdgePosition)
                            && ((Member.tEdgeHeight <= (30 * 12))
                            && (FO.FOType == "OHDoor"))))) {
                    RightClip = 0;
                    LeftClip = 0;
                }

                for (item in FO.FOMaterials) {
                    if ((item.clsType == "Member")) {
                        if (((item.CL == Member.rEdgePosition)
                                    && ((item.tEdgeHeight >= Member.tEdgeHeight)
                                    && (FO.rEdgePosition == Member.rEdgePosition)))) {
                            RightClip = 0;
                        }

                        if (((item.CL
                                    == (Member.lEdgePosition - Member.Length))
                                    && ((item.tEdgeHeight >= Member.tEdgeHeight)
                                    && (FO.lEdgePosition
                                    == (Member.lEdgePosition - Member.Length))))) {
                            LeftClip = 0;
                        }

                    }

                }

            }

        }
        else {
            RightClip = 0;
            LeftClip = 0;
        }

        WeldClips = (WeldClips
                    + (RightClip + LeftClip));
    }

    // FO Weld Clips
    for (FO in FOCollection) {
        for (item in FO.FOMaterials) {
            if ((item.clsType == "Member")) {
                if ((item.mType == "FO Receiver Jamb")) {
                    if ((item.bEdgeHeight > 0)) {
                        WeldClips = (WeldClips + item.Qty);
                    }
                    else if ((item.tEdgeHeight == b.DistanceToRoof(eWall, item.CL))) {
                        WeldClips = (WeldClips
                                    + (item.Qty * 2));
                    }

                }
                else if (((item.mType == "FO Header")
                            || (item.mType == "FO Stool"))) {
                    WeldClips = (WeldClips
                                + (item.Qty * 2));
                }

                WeldClips = (WeldClips + 2);
            }

        }

    }

    b.WeldClips = WeldClips;
}

// '''''''''' only used for non-expandable endwalls, returns FO edges that meet the following conditions:
// if OHDoor or MiscFO w/ Full Height Jambs option is within the max distance, returns edge that should be used as load bearing (if within 5' of ideal span), creates jambs as necessary
// if ideal column location lands on FO, returns edge that should be used as load bearing, creates jambs as necessary
// if no FOs qualify, returns IdealSpan
// Direction used to check for FOs going towards lower side of roof.
// e1 single slope: always positive direction (right to left)
// e3 single slope: always negative direction (left to right)
// gable roofs: both directions are used for both endwalls
// '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''StartPos, MaxDistance, IdealSpan all in inches
NonExpandableFOJambs(b: clsBuilding, eWall: string, StartPos: number, MaxDistance: number, IdealSpan: number, Direction: number): number {
    let WallColumns: Collection;
    let FOs: Collection;
    let FO: clsFO;
    let tempLocation: number;
    let Jamb: clsMember;
    let JambSupport: clsMember;
    let LoadBearingJamb: string;
    let lGtob: number;
    let rGtob: number;
    let FirstCorner: number;
    let LastCorner: number;
    if ((eWall == "e1")) {
        FOs = b.e1FOs;
        WallColumns = b.e1Columns;
    }
    else if ((eWall == "e3")) {
        FOs = b.e3FOs;
        WallColumns = b.e3Columns;
    }

    if ((Direction == 1)) {
        // positive direction
        if ((((b.bWidth * 12)
                    - StartPos)
                    < MaxDistance)) {
            return (b.bWidth * 12);

        }

        for (FO in FOs) {
            // if OHDoor or MiscFO w/ full height jambs is inside max distance and the furthest edge is at least within 5' of max distance,
            // then one of the jambs should be load bearing
            if (((FO.rEdgePosition > StartPos)
                        && ((FO.rEdgePosition
                        < (MaxDistance + StartPos))
                        && ((FO.lEdgePosition
                        > (StartPos
                        + (MaxDistance - 60)))
                        && ((FO.FOType == "OHDoor")
                        || (FO.StructuralSteelOption == "Full Height Jambs w/ Header & Stool")))))) {
                if ((FO.lEdgePosition
                            < (StartPos + MaxDistance))) {
                    // left edge should be used as load bearing
                    LoadBearingJamb = "Left";
                }
                else {
                    LoadBearingJamb = "Right";
                }

                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                // left jamb
                Jamb = new clsMember();
                Jamb.bEdgeHeight = 0;
                if ((lGtob
                            < ((30 * 12)
                            + 4))) {
                    // don't need jamb support
                    Jamb.tEdgeHeight = lGtob;
                    if ((LoadBearingJamb == "Left")) {
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.lEdgePosition;
                    }

                }
                else {
                    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = (30 * 12);
                    JambSupport = new clsMember();
                    JambSupport.bEdgeHeight = 0;
                    JambSupport.CL = FO.lEdgePosition;
                    if ((LoadBearingJamb == "Left")) {
                        JambSupport.LoadBearing = true;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    JambSupport.tEdgeHeight = lGtob;
                    JambSupport.Length = lGtob;
                    JambSupport.SetSize;
                    b;
                    "Column";
                    eWall;
                    30;
                    JambSupport.rEdgePosition = (JambSupport.CL
                                - (JambSupport.Width / 2));
                    WallColumns.Add;
                    JambSupport;
                }

                Jamb.Length = Jamb.tEdgeHeight;
                Jamb.Size = "8"" Receiver Cee";
                Jamb.Width = 2.5;
                Jamb.CL = FO.lEdgePosition;
                Jamb.rEdgePosition = (Jamb.CL
                            - (Jamb.Width / 2));
                if ((Jamb.LoadBearing == false)) {
                    FO.FOMaterials.Add;
                    Jamb;
                }
                else {
                    WallColumns.Add;
                    Jamb;
                }

                // right jamb
                Jamb = new clsMember();
                Jamb.bEdgeHeight = 0;
                if ((rGtob
                            < ((30 * 12)
                            + 4))) {
                    // don't need jamb support
                    Jamb.tEdgeHeight = rGtob;
                    if ((LoadBearingJamb == "Right")) {
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.rEdgePosition;
                    }

                }
                else {
                    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = (30 * 12);
                    JambSupport = new clsMember();
                    JambSupport.bEdgeHeight = 0;
                    JambSupport.CL = FO.rEdgePosition;
                    if ((LoadBearingJamb == "Right")) {
                        JambSupport.LoadBearing = true;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    JambSupport.tEdgeHeight = rGtob;
                    JambSupport.Length = rGtob;
                    JambSupport.SetSize;
                    b;
                    "Column";
                    eWall;
                    30;
                    JambSupport.rEdgePosition = (JambSupport.CL
                                - (JambSupport.Width / 2));
                    WallColumns.Add;
                    JambSupport;
                }

                Jamb.Length = Jamb.tEdgeHeight;
                Jamb.Size = "8"" Receiver Cee";
                Jamb.Width = 2.5;
                Jamb.CL = FO.rEdgePosition;
                Jamb.rEdgePosition = (Jamb.CL
                            - (Jamb.Width / 2));
                if ((Jamb.LoadBearing == false)) {
                    FO.FOMaterials.Add;
                    Jamb;
                }
                else {
                    WallColumns.Add;
                    Jamb;
                }

                // TODO: Exit Function: Warning!!! Need to return the value
                return;
            }
            else if (((FO.lEdgePosition
                        >= (StartPos + IdealSpan))
                        && ((FO.rEdgePosition
                        <= (StartPos + IdealSpan))
                        && (FO.FOType != "PDoor")))) {
                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                if ((FO.lEdgePosition
                            < (StartPos + MaxDistance))) {
                    LoadBearingJamb = "Left";
                    Jamb = new clsMember();
                    Jamb.bEdgeHeight = 0;
                    if ((lGtob
                                < ((30 * 12)
                                + 4))) {
                        // don't need jamb support
                        Jamb.tEdgeHeight = lGtob;
                        if ((LoadBearingJamb == "Left")) {
                            Jamb.LoadBearing = true;
                            NonExpandableFOJambs = FO.lEdgePosition;
                        }

                    }
                    else {
                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                        Jamb.tEdgeHeight = (30 * 12);
                        JambSupport = new clsMember();
                        JambSupport.bEdgeHeight = 0;
                        JambSupport.CL = FO.lEdgePosition;
                        if ((LoadBearingJamb == "Left")) {
                            JambSupport.LoadBearing = true;
                            NonExpandableFOJambs = JambSupport.CL;
                        }

                        JambSupport.tEdgeHeight = lGtob;
                        JambSupport.Length = lGtob;
                        JambSupport.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        JambSupport.rEdgePosition = (JambSupport.CL
                                    - (JambSupport.Width / 2));
                        WallColumns.Add;
                        JambSupport;
                    }

                    Jamb.Length = Jamb.tEdgeHeight;
                    Jamb.Size = "8"" Receiver Cee";
                    Jamb.Width = 2.5;
                    Jamb.CL = FO.lEdgePosition;
                    Jamb.rEdgePosition = (Jamb.CL
                                - (Jamb.Width / 2));
                    if ((Jamb.LoadBearing == false)) {
                        FO.FOMaterials.Add;
                        Jamb;
                    }
                    else {
                        WallColumns.Add;
                        Jamb;
                    }

                }
                else {
                    LoadBearingJamb = "Right";
                    Jamb = new clsMember();
                    Jamb.bEdgeHeight = 0;
                    if ((rGtob
                                < ((30 * 12)
                                + 4))) {
                        // don't need jamb support
                        Jamb.tEdgeHeight = rGtob;
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.rEdgePosition;
                    }
                    else {
                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                        Jamb.tEdgeHeight = (30 * 12);
                        JambSupport = new clsMember();
                        JambSupport.bEdgeHeight = 0;
                        JambSupport.LoadBearing = true;
                        JambSupport.tEdgeHeight = rGtob;
                        JambSupport.Length = rGtob;
                        JambSupport.CL = FO.rEdgePosition;
                        JambSupport.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        JambSupport.rEdgePosition = (JambSupport.CL
                                    - (JambSupport.Width / 2));
                        WallColumns.Add;
                        JambSupport;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    Jamb.Length = Jamb.tEdgeHeight;
                    Jamb.Size = "8"" Receiver Cee";
                    Jamb.Width = 2.5;
                    Jamb.CL = FO.rEdgePosition;
                    Jamb.rEdgePosition = (Jamb.CL
                                - (Jamb.Width / 2));
                    if ((Jamb.LoadBearing == false)) {
                        FO.FOMaterials.Add;
                        Jamb;
                    }
                    else {
                        WallColumns.Add;
                        Jamb;
                    }

                }

                // TODO: Exit Function: Warning!!! Need to return the value
                return;
            }
            else {
                NonExpandableFOJambs = StartPos;
            }

        }

    }
    else {
        if ((StartPos < MaxDistance)) {
            return FirstCorner;

        }

        for (FO in FOs) {
            // if OHDoor or MiscFO w/ full height jambs is inside max distance and the furthest edge is at least within 5' of max distance,
            // then one of the jambs should be load bearing
            if (((FO.lEdgePosition < StartPos)
                        && ((FO.lEdgePosition
                        > (StartPos - MaxDistance))
                        && ((FO.rEdgePosition
                        < ((StartPos - MaxDistance)
                        + 60))
                        && ((FO.FOType == "OHDoor")
                        || (FO.StructuralSteelOption == "Full Height Jambs w/ Header & Stool")))))) {
                if ((FO.rEdgePosition
                            > (StartPos - MaxDistance))) {
                    // right edge should be used as load bearing
                    LoadBearingJamb = "Right";
                }
                else {
                    LoadBearingJamb = "Left";
                }

                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                // left jamb
                Jamb = new clsMember();
                Jamb.bEdgeHeight = 0;
                if ((lGtob
                            < ((30 * 12)
                            + 4))) {
                    // don't need jamb support
                    Jamb.tEdgeHeight = lGtob;
                    if ((LoadBearingJamb == "Left")) {
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.lEdgePosition;
                    }

                }
                else {
                    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = (30 * 12);
                    JambSupport = new clsMember();
                    JambSupport.bEdgeHeight = 0;
                    JambSupport.CL = FO.lEdgePosition;
                    if ((LoadBearingJamb == "Left")) {
                        JambSupport.LoadBearing = true;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    JambSupport.tEdgeHeight = lGtob;
                    JambSupport.Length = lGtob;
                    JambSupport.SetSize;
                    b;
                    "Column";
                    eWall;
                    30;
                    JambSupport.rEdgePosition = (JambSupport.CL
                                - (JambSupport.Width / 2));
                    WallColumns.Add;
                    JambSupport;
                }

                Jamb.Length = Jamb.tEdgeHeight;
                Jamb.Size = "8"" Receiver Cee";
                Jamb.Width = 2.5;
                Jamb.CL = FO.lEdgePosition;
                Jamb.rEdgePosition = (Jamb.CL
                            - (Jamb.Width / 2));
                if ((Jamb.LoadBearing == false)) {
                    FO.FOMaterials.Add;
                    Jamb;
                }
                else {
                    WallColumns.Add;
                    Jamb;
                }

                // right jamb
                Jamb = new clsMember();
                Jamb.bEdgeHeight = 0;
                if ((rGtob
                            < ((30 * 12)
                            + 4))) {
                    // don't need jamb support
                    Jamb.tEdgeHeight = rGtob;
                    if ((LoadBearingJamb == "Right")) {
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.rEdgePosition;
                    }

                }
                else {
                    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = (30 * 12);
                    JambSupport = new clsMember();
                    JambSupport.bEdgeHeight = 0;
                    JambSupport.CL = FO.rEdgePosition;
                    if ((LoadBearingJamb == "Right")) {
                        JambSupport.LoadBearing = true;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    JambSupport.tEdgeHeight = rGtob;
                    JambSupport.Length = rGtob;
                    JambSupport.SetSize;
                    b;
                    "Column";
                    eWall;
                    30;
                    JambSupport.rEdgePosition = (JambSupport.CL
                                - (JambSupport.Width / 2));
                    WallColumns.Add;
                    JambSupport;
                }

                Jamb.Length = Jamb.tEdgeHeight;
                Jamb.Size = "8"" Receiver Cee";
                Jamb.Width = 2.5;
                Jamb.CL = FO.rEdgePosition;
                Jamb.rEdgePosition = (Jamb.CL
                            - (Jamb.Width / 2));
                if ((Jamb.LoadBearing == false)) {
                    FO.FOMaterials.Add;
                    Jamb;
                }
                else {
                    WallColumns.Add;
                    Jamb;
                }

                // TODO: Exit Function: Warning!!! Need to return the value
                return;
            }
            else if (((FO.lEdgePosition
                        >= (StartPos - IdealSpan))
                        && ((FO.rEdgePosition
                        <= (StartPos - IdealSpan))
                        && (FO.FOType != "PDoor")))) {
                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                if ((FO.rEdgePosition
                            < (StartPos - MaxDistance))) {
                    LoadBearingJamb = "Left";
                    Jamb = new clsMember();
                    Jamb.bEdgeHeight = 0;
                    if ((lGtob
                                < ((30 * 12)
                                + 4))) {
                        // don't need jamb support
                        Jamb.tEdgeHeight = lGtob;
                        if ((LoadBearingJamb == "Left")) {
                            Jamb.LoadBearing = true;
                            NonExpandableFOJambs = FO.lEdgePosition;
                        }

                    }
                    else {
                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                        Jamb.tEdgeHeight = (30 * 12);
                        JambSupport = new clsMember();
                        JambSupport.bEdgeHeight = 0;
                        JambSupport.CL = FO.lEdgePosition;
                        if ((LoadBearingJamb == "Left")) {
                            JambSupport.LoadBearing = true;
                            NonExpandableFOJambs = JambSupport.CL;
                        }

                        JambSupport.tEdgeHeight = lGtob;
                        JambSupport.Length = lGtob;
                        JambSupport.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        JambSupport.rEdgePosition = (JambSupport.CL
                                    - (JambSupport.Width / 2));
                        WallColumns.Add;
                        JambSupport;
                    }

                    Jamb.Length = Jamb.tEdgeHeight;
                    Jamb.Size = "8"" Receiver Cee";
                    Jamb.Width = 2.5;
                    Jamb.CL = FO.lEdgePosition;
                    Jamb.rEdgePosition = (Jamb.CL
                                - (Jamb.Width / 2));
                    if ((Jamb.LoadBearing == false)) {
                        FO.FOMaterials.Add;
                        Jamb;
                    }
                    else {
                        WallColumns.Add;
                        Jamb;
                    }

                }
                else {
                    LoadBearingJamb = "Right";
                    Jamb = new clsMember();
                    Jamb.bEdgeHeight = 0;
                    if ((rGtob
                                < ((30 * 12)
                                + 4))) {
                        // don't need jamb support
                        Jamb.tEdgeHeight = rGtob;
                        Jamb.LoadBearing = true;
                        NonExpandableFOJambs = FO.rEdgePosition;
                    }
                    else {
                        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                        Jamb.tEdgeHeight = (30 * 12);
                        JambSupport = new clsMember();
                        JambSupport.bEdgeHeight = 0;
                        JambSupport.LoadBearing = true;
                        JambSupport.tEdgeHeight = rGtob;
                        JambSupport.Length = rGtob;
                        JambSupport.CL = FO.rEdgePosition;
                        JambSupport.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        JambSupport.rEdgePosition = (JambSupport.CL
                                    - (JambSupport.Width / 2));
                        WallColumns.Add;
                        JambSupport;
                        NonExpandableFOJambs = JambSupport.CL;
                    }

                    Jamb.Length = Jamb.tEdgeHeight;
                    Jamb.Size = "8"" Receiver Cee";
                    Jamb.Width = 2.5;
                    Jamb.CL = FO.rEdgePosition;
                    Jamb.rEdgePosition = (Jamb.CL
                                - (Jamb.Width / 2));
                    if ((Jamb.LoadBearing == false)) {
                        FO.FOMaterials.Add;
                        Jamb;
                    }
                    else {
                        WallColumns.Add;
                        Jamb;
                    }

                }

                // TODO: Exit Function: Warning!!! Need to return the value
                return;
            }
            else {
                NonExpandableFOJambs = StartPos;
            }

        }

    }

}

// ''''''''''''' Adds FO Jambs for FOs without a wall location
// Window Jambs will default to 7'2" Jambs w/ Header and Stool
FieldLocateFOCalc(b: clsBuilding) {
    let FO: clsFO;
    let Jamb: clsMember;
    let Purlin: clsMember;
    for (FO in b.fieldlocateFOs) {
        if ((FO.FOType == "PDoor")) {
            // Do Nothing - no additional steel for  PDoors
        }
        else if ((FO.FOType == "Window")) {
            // Add left and right jamb
            Jamb = new clsMember();
            Jamb.mType = "FO Material";
            Jamb.Size = "8"" Receiver Cee";
            Jamb.Width = 2.5;
            Jamb.Length = 86;
            Jamb.mType = "FO Receiver Jamb";
            Jamb.Placement = "FO Jamb";
            Jamb.Qty = 1;
            FO.FOMaterials.Add;
            Jamb;
            Jamb = new clsMember();
            Jamb.mType = "FO Material";
            Jamb.Size = "8"" Receiver Cee";
            Jamb.Width = 2.5;
            Jamb.Length = 86;
            Jamb.mType = "FO Receiver Jamb";
            Jamb.Placement = "FO Jamb";
            Jamb.Qty = 1;
            FO.FOMaterials.Add;
            Jamb;
            // Add Header and Stool
            // Header
            Jamb = new clsMember();
            Jamb.mType = "FO Material";
            Jamb.Size = "8"" C Purlin";
            Jamb.Width = 2.5;
            Jamb.Length = FO.Width;
            Jamb.mType = "FO Header";
            Jamb.Placement = Jamb.mType;
            FO.FOMaterials.Add;
            Jamb;
            // Stool
            Jamb = new clsMember();
            Jamb.mType = "FO Material";
            Jamb.Size = "8"" C Purlin";
            Jamb.Width = 2.5;
            Jamb.Length = FO.Width;
            Jamb.mType = "FO Stool";
            Jamb.Placement = Jamb.mType;
            FO.FOMaterials.Add;
            Jamb;
            b.WeldClips = (b.WeldClips + 6);
        }

    }

}

// ''''''''''''' Adds FO Jambs if not already added in Column Calc
FOJambsCalc(b: clsBuilding, eWall: string) {
    let FO: clsFO;
    let Column: clsMember;
    let AllColumnsValid: boolean;
    let ColumnCollection: Collection;
    let FOCollection: Collection;
    let ColIndex: number;
    let Jamb: clsMember;
    let Purlin: clsMember;
    let MiscFOColumnReplacement: boolean;
    let OHDoorColumnReplacement: boolean;
    let WindowColumnReplacement: boolean;
    let Member: clsMember;
    let WeldClips: clsMiscItem;
    let ReplacedColLocation: number;
    let RightJambExists: boolean;
    let LeftJambExists: boolean;
    let RightSupportExists: boolean;
    let LeftSupportExists: boolean;
    let rGtob: number;
    let lGtob: number;
    switch (eWall) {
        case "e1":
            ColumnCollection = b.e1Columns;
            FOCollection = b.e1FOs;
            break;
        case "s2":
            ColumnCollection = b.s2Columns;
            FOCollection = b.s2FOs;
            break;
        case "e3":
            ColumnCollection = b.e3Columns;
            FOCollection = b.e3FOs;
            break;
        case "s4":
            ColumnCollection = b.s4Columns;
            FOCollection = b.s4FOs;
            break;
    }

    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Endwalls
    if (((eWall == "e1")
                || (eWall == "e3"))) {
        for (FO in FOCollection) {
            switch (FO.FOType) {
                case "OHDoor":
                    RightJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            RightJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.rEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            RightJambExists = true;
                        }

                    }

                    // Check if Left Jamb already Exists
                    LeftJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            LeftJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.lEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            LeftJambExists = true;
                        }

                    }

                    // if Right Jamb doesn't exist, create jamb
                    if ((RightJambExists == false)) {
                        rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                        if ((rGtob
                                    <= ((30 * 12)
                                    + 4))) {
                            // create full height jamb
                            Jamb = new clsMember();
                            Jamb.CL = FO.rEdgePosition;
                            Jamb.bEdgeHeight = 0;
                            Jamb.tEdgeHeight = rGtob;
                            Jamb.Length = rGtob;
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.mType = "FO Receiver Jamb";
                            FO.FOMaterials.Add;
                            Jamb;
                        }
                        else {
                            // create 30'4" jamb
                            Jamb = new clsMember();
                            Jamb.CL = FO.rEdgePosition;
                            Jamb.bEdgeHeight = 0;
                            Jamb.tEdgeHeight = ((30 * 12)
                                        + 4);
                            Jamb.Length = ((30 * 12)
                                        + 4);
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.mType = "FO Receiver Jamb";
                            FO.FOMaterials.Add;
                            Jamb;
                            // Check if load bearing column already exists
                            RightSupportExists = false;
                            for (Column in ColumnCollection) {
                                if (((Column.CL
                                            >= (FO.rEdgePosition - 12))
                                            && (Column.CL
                                            <= (FO.rEdgePosition + 12)))) {
                                    RightSupportExists = true;
                                }

                            }

                            if ((RightSupportExists == true)) {
                                // Do Nothing
                            }
                            else {
                                Jamb = new clsMember();
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.bEdgeHeight = 0;
                                Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL);
                                Jamb.Length = Jamb.tEdgeHeight;
                                Jamb.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                "NonExpandable";
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.mType = "FO Support Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                        }

                    }

                    // if Left Jamb doesn't exist, create jamb
                    if ((LeftJambExists == false)) {
                        lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                        if ((lGtob
                                    <= ((30 * 12)
                                    + 4))) {
                            // create full height jamb
                            Jamb = new clsMember();
                            Jamb.CL = FO.lEdgePosition;
                            Jamb.bEdgeHeight = 0;
                            Jamb.tEdgeHeight = lGtob;
                            Jamb.Length = lGtob;
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.mType = "FO Receiver Jamb";
                            FO.FOMaterials.Add;
                            Jamb;
                        }
                        else {
                            // create 30'4" jamb
                            Jamb = new clsMember();
                            Jamb.CL = FO.lEdgePosition;
                            Jamb.bEdgeHeight = 0;
                            Jamb.tEdgeHeight = ((30 * 12)
                                        + 4);
                            Jamb.Length = ((30 * 12)
                                        + 4);
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.mType = "FO Receiver Jamb";
                            FO.FOMaterials.Add;
                            Jamb;
                            // Check if load bearing column already exists
                            LeftSupportExists = false;
                            for (Column in ColumnCollection) {
                                if (((Column.CL
                                            >= (FO.lEdgePosition - 12))
                                            && (Column.CL
                                            <= (FO.lEdgePosition + 12)))) {
                                    LeftSupportExists = true;
                                }

                            }

                            if ((LeftSupportExists == true)) {
                                // Do Nothing
                            }
                            else {
                                Jamb = new clsMember();
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.bEdgeHeight = 0;
                                Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL);
                                Jamb.Length = Jamb.tEdgeHeight;
                                Jamb.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                "NonExpandable";
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.mType = "FO Support Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                        }

                    }

                    // Create Header
                    Jamb = new clsMember();
                    Jamb.mType = "FO Material";
                    Jamb.Size = "8"" C Purlin";
                    Jamb.Width = 2.5;
                    Jamb.rEdgePosition = FO.rEdgePosition;
                    Jamb.Length = (FO.lEdgePosition - FO.rEdgePosition);
                    Jamb.tEdgeHeight = FO.tEdgeHeight;
                    Jamb.CL = 0;
                    Jamb.mType = "FO Header";
                    FO.FOMaterials.Add;
                    Jamb;
                    break;
                case "Window":
                    RightJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            RightJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.rEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            RightJambExists = true;
                        }

                    }

                    // Check if Left Jamb already Exists
                    LeftJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            LeftJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.lEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            LeftJambExists = true;
                        }

                    }

                    // if Right Jamb doesn't exist, create jamb
                    if ((RightJambExists == false)) {
                        Jamb = new clsMember();
                        Jamb.mType = "FO Material";
                        Jamb.Size = "8"" Receiver Cee";
                        Jamb.Width = 2.5;
                        Jamb.CL = FO.rEdgePosition;
                        Jamb.tEdgeHeight = FO.tEdgeHeight;
                        Jamb.bEdgeHeight = FO.bEdgeHeight;
                        Jamb.Length = (Jamb.tEdgeHeight - Jamb.bEdgeHeight);
                        Jamb.rEdgePosition = (Jamb.CL
                                    - (Jamb.Width / 2));
                        Jamb.mType = "FO Receiver Jamb";
                        FO.FOMaterials.Add;
                        Jamb;
                    }

                    // if Right Jamb doesn't exist, create jamb
                    if ((LeftJambExists == false)) {
                        Jamb = new clsMember();
                        Jamb.mType = "FO Material";
                        Jamb.Size = "8"" Receiver Cee";
                        Jamb.Width = 2.5;
                        Jamb.CL = FO.lEdgePosition;
                        Jamb.tEdgeHeight = FO.tEdgeHeight;
                        Jamb.bEdgeHeight = FO.bEdgeHeight;
                        Jamb.Length = (Jamb.tEdgeHeight - Jamb.bEdgeHeight);
                        Jamb.rEdgePosition = (Jamb.CL
                                    - (Jamb.Width / 2));
                        Jamb.mType = "FO Receiver Jamb";
                        FO.FOMaterials.Add;
                        Jamb;
                    }

                    // Add Header and Stool
                    // Header
                    Jamb = new clsMember();
                    Jamb.mType = "FO Material";
                    Jamb.Size = "8"" C Purlin";
                    Jamb.Width = 2.5;
                    Jamb.rEdgePosition = FO.rEdgePosition;
                    Jamb.Length = FO.Width;
                    Jamb.tEdgeHeight = FO.tEdgeHeight;
                    Jamb.bEdgeHeight = FO.bEdgeHeight;
                    Jamb.mType = "FO Header";
                    FO.FOMaterials.Add;
                    Jamb;
                    // Stool
                    Jamb = new clsMember();
                    Jamb.mType = "FO Material";
                    Jamb.Size = "8"" C Purlin";
                    Jamb.Width = 2.5;
                    Jamb.rEdgePosition = FO.rEdgePosition;
                    Jamb.Length = FO.Width;
                    Jamb.tEdgeHeight = FO.bEdgeHeight;
                    Jamb.bEdgeHeight = FO.bEdgeHeight;
                    Jamb.mType = "FO Stool";
                    FO.FOMaterials.Add;
                    Jamb;
                    break;
                case "MiscFO":
                    RightJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            RightJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.rEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            RightJambExists = true;
                        }

                    }

                    // Check if Left Jamb already Exists
                    LeftJambExists = false;
                    for (Member in FO.FOMaterials) {
                        if ((Member.CL == FO.rEdgePosition)) {
                            LeftJambExists = true;
                        }

                    }

                    for (Member in ColumnCollection) {
                        if (((Member.CL == FO.lEdgePosition)
                                    && (Member.LoadBearing == false))) {
                            LeftJambExists = true;
                        }

                    }

                    switch (FO.StructuralSteelOption) {
                        case "Full Height Jambs w/ Header & Stool":
                            if ((RightJambExists == false)) {
                                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                                if ((rGtob
                                            <= ((30 * 12)
                                            + 4))) {
                                    // create full height jamb
                                    Jamb = new clsMember();
                                    Jamb.CL = FO.rEdgePosition;
                                    Jamb.bEdgeHeight = 0;
                                    Jamb.tEdgeHeight = rGtob;
                                    Jamb.Length = rGtob;
                                    Jamb.Size = "8"" Receiver Cee";
                                    Jamb.Width = 2.5;
                                    Jamb.rEdgePosition = (Jamb.CL
                                                - (Jamb.Width / 2));
                                    Jamb.mType = "FO Receiver Jamb";
                                    FO.FOMaterials.Add;
                                    Jamb;
                                }
                                else {
                                    // create FO sized jamb w/ support
                                    Jamb = new clsMember();
                                    Jamb.CL = FO.rEdgePosition;
                                    Jamb.bEdgeHeight = FO.bEdgeHeight;
                                    Jamb.tEdgeHeight = FO.tEdgeHeight;
                                    Jamb.Length = (Jamb.tEdgeHeight - Jamb.bEdgeHeight);
                                    Jamb.Size = "8"" Receiver Cee";
                                    Jamb.Width = 2.5;
                                    Jamb.rEdgePosition = (Jamb.CL
                                                - (Jamb.Width / 2));
                                    Jamb.mType = "FO Receiver Jamb";
                                    FO.FOMaterials.Add;
                                    Jamb;
                                    // Check if load bearing column already exists
                                    RightSupportExists = false;
                                    for (Column in ColumnCollection) {
                                        if (((Column.CL
                                                    >= (FO.rEdgePosition - 12))
                                                    && (Column.CL
                                                    <= (FO.rEdgePosition + 12)))) {
                                            RightSupportExists = true;
                                        }

                                    }

                                    if ((RightSupportExists == true)) {
                                        // Do Nothing
                                    }
                                    else {
                                        Jamb = new clsMember();
                                        Jamb.CL = FO.rEdgePosition;
                                        Jamb.bEdgeHeight = 0;
                                        Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL);
                                        Jamb.Length = Jamb.tEdgeHeight;
                                        Jamb.SetSize;
                                        b;
                                        "Column";
                                        eWall;
                                        30;
                                        "NonExpandable";
                                        Jamb.rEdgePosition = (Jamb.CL
                                                    - (Jamb.Width / 2));
                                        Jamb.mType = "FO Support Jamb";
                                        FO.FOMaterials.Add;
                                        Jamb;
                                    }

                                }

                            }

                            // if Left Jamb doesn't exist, create jamb
                            if ((LeftJambExists == false)) {
                                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                                if ((lGtob
                                            <= ((30 * 12)
                                            + 4))) {
                                    // create full height jamb
                                    Jamb = new clsMember();
                                    Jamb.CL = FO.lEdgePosition;
                                    Jamb.bEdgeHeight = 0;
                                    Jamb.tEdgeHeight = lGtob;
                                    Jamb.Length = lGtob;
                                    Jamb.Size = "8"" Receiver Cee";
                                    Jamb.Width = 2.5;
                                    Jamb.rEdgePosition = (Jamb.CL
                                                - (Jamb.Width / 2));
                                    Jamb.mType = "FO Receiver Jamb";
                                    FO.FOMaterials.Add;
                                    Jamb;
                                }
                                else {
                                    // create FO sized jamb w/ support
                                    Jamb = new clsMember();
                                    Jamb.CL = FO.lEdgePosition;
                                    Jamb.bEdgeHeight = FO.bEdgeHeight;
                                    Jamb.tEdgeHeight = FO.tEdgeHeight;
                                    Jamb.Length = (Jamb.tEdgeHeight - Jamb.bEdgeHeight);
                                    Jamb.Size = "8"" Receiver Cee";
                                    Jamb.Width = 2.5;
                                    Jamb.rEdgePosition = (Jamb.CL
                                                - (Jamb.Width / 2));
                                    Jamb.mType = "FO Receiver Jamb";
                                    FO.FOMaterials.Add;
                                    Jamb;
                                    // Check if load bearing column already exists
                                    LeftSupportExists = false;
                                    for (Column in ColumnCollection) {
                                        if (((Column.CL
                                                    >= (FO.lEdgePosition - 12))
                                                    && (Column.CL
                                                    <= (FO.lEdgePosition + 12)))) {
                                            LeftSupportExists = true;
                                        }

                                    }

                                    if ((LeftSupportExists == true)) {
                                        // Do Nothing
                                    }
                                    else {
                                        Jamb = new clsMember();
                                        Jamb.CL = FO.lEdgePosition;
                                        Jamb.bEdgeHeight = 0;
                                        Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL);
                                        Jamb.Length = Jamb.tEdgeHeight;
                                        Jamb.SetSize;
                                        b;
                                        "Column";
                                        eWall;
                                        30;
                                        "NonExpandable";
                                        Jamb.rEdgePosition = (Jamb.CL
                                                    - (Jamb.Width / 2));
                                        Jamb.mType = "FO Support Jamb";
                                        FO.FOMaterials.Add;
                                        Jamb;
                                    }

                                }

                            }

                            // header and stool
                            Purlin = new clsMember();
                            Purlin.Length = FO.Width;
                            Purlin.Size = "8"" C Purlin";
                            Purlin.mType = "FO Material";
                            if ((FO.tEdgeHeight == 86)) {
                                Purlin.Qty = 1;
                            }
                            else {
                                Purlin.Qty = 2;
                            }

                            Purlin.mType = "FO Header";
                            FO.FOMaterials.Add;
                            Purlin;
                            break;
                        case "7'2"" Jambs w/ Header & Stool":
                            if ((RightJambExists == false)) {
                                // add jambs if they weren't already added as a column replacement
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            if ((LeftJambExists == false)) {
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            // header and stool
                            Purlin = new clsMember();
                            Purlin.Length = FO.Width;
                            Purlin.Size = "8"" C Purlin";
                            Purlin.mType = "FO Material";
                            if ((FO.tEdgeHeight == 86)) {
                                Purlin.Qty = 1;
                            }
                            else {
                                Purlin.Qty = 2;
                            }

                            Purlin.mType = "FO Header";
                            FO.FOMaterials.Add;
                            Purlin;
                            break;
                        case "7'2"" Jambs w/ Stool":
                            if ((RightJambExists == false)) {
                                // add jambs if they weren't already added as a column replacement
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            if ((LeftJambExists == false)) {
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            // stool
                            Purlin = new clsMember();
                            Purlin.Length = FO.Width;
                            Purlin.Size = "8"" C Purlin";
                            Purlin.mType = "FO Material";
                            if ((FO.tEdgeHeight == 86)) {
                                Purlin.Qty = 1;
                            }
                            else {
                                Purlin.Qty = 2;
                            }

                            Purlin.mType = "FO Header";
                            FO.FOMaterials.Add;
                            Purlin;
                            break;
                        case "7'2"" Jambs":
                            if ((RightJambExists == false)) {
                                // add jambs if they weren't already added as a column replacement
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            if ((LeftJambExists == false)) {
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (2 + (7 * 12));
                                Jamb.tEdgeHeight = (2 + (7 * 12));
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            break;
                        case "5' Jambs w/ Header & Stool":
                            if ((RightJambExists == false)) {
                                // add jambs if they weren't already added as a column replacement
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (5 * 12);
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            if ((LeftJambExists == false)) {
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = (5 * 12);
                                Jamb.mType = "FO Receiver Jamb";
                                FO.FOMaterials.Add;
                                Jamb;
                            }

                            // header and stool
                            Purlin = new clsMember();
                            Purlin.Length = FO.Width;
                            Purlin.Size = "8"" C Purlin";
                            Purlin.mType = "FO Material";
                            if ((FO.tEdgeHeight == 86)) {
                                Purlin.Qty = 1;
                            }
                            else {
                                Purlin.Qty = 2;
                            }

                            Purlin.mType = "FO Header";
                            FO.FOMaterials.Add;
                            Purlin;
                            break;
                    }

                    break;
            }

        }

    }

    // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Sidewalls
    if (((eWall == "s2")
                || (eWall == "s4"))) {
        for (FO in FOCollection) {
            switch (FO.FOType) {
                case "pDoor":
                    for (Column in ColumnCollection) {
                        // if within 6" of column CL
                        if (((Column.CL
                                    > (FO.rEdgePosition - 6))
                                    && (Column.CL
                                    < (FO.lEdgePosition + 6)))) {
                            // error condition
                            if ((MsgBox((FO.Description + " has been found to be located within 6"" of a column! Relocate this personnel door before proceeding. Continue anyways?"), System.Windows.Forms.MessageBoxButtons.OKCancel, "FO Placement Error") == 7)) {

                            }

                            return;
                        }

                        Column;
                        "OHDoor";
                        for (ColIndex = ColumnCollection.Count; (ColIndex <= 1); ColIndex = (ColIndex + -1)) {
                            Column = ColumnCollection[ColIndex];
                            // if within 1' of a column CL
                            if (((Column.CL
                                        > (FO.rEdgePosition - (1 * 12)))
                                        && (Column.CL
                                        < (FO.lEdgePosition + (1 * 12))))) {
                                // error condition
                                if ((MsgBox((FO.Description + " has been found to be located within 1' of a column! Relocate or resize this overhead door before proceeding. Continue anyways?"), (vbYesNo + vbCritical), "FO Placement Error") == 7)) {

                                }

                                // Exit Sub
                            }

                            ColIndex;
                            // create jambs
                            Jamb = new clsMember();
                            Jamb.mType = "FO Material";
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.CL = FO.rEdgePosition;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL);
                            Jamb.tEdgeHeight = Jamb.Length;
                            Jamb.bEdgeHeight = 0;
                            FO.FOMaterials.Add;
                            Jamb;
                            Jamb = new clsMember();
                            Jamb.mType = "FO Material";
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.CL = FO.lEdgePosition;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL);
                            Jamb.tEdgeHeight = Jamb.Length;
                            Jamb.bEdgeHeight = 0;
                            FO.FOMaterials.Add;
                            Jamb;
                            // OHDoor Header'''''''''''''''''''''''''''''''''
                            Jamb = new clsMember();
                            Jamb.mType = "FO Material";
                            Jamb.Size = "8"" C Purlin";
                            Jamb.Width = 2.5;
                            Jamb.rEdgePosition = FO.rEdgePosition;
                            Jamb.Length = (FO.lEdgePosition - FO.rEdgePosition);
                            Jamb.tEdgeHeight = FO.tEdgeHeight;
                            Jamb.CL = 0;
                            FO.FOMaterials.Add;
                            Jamb;
                            "Window";
                            for (ColIndex = ColumnCollection.Count; (ColIndex <= 1); ColIndex = (ColIndex + -1)) {
                                Column = ColumnCollection[ColIndex];
                                // if within 1' of a column CL
                                if (((Column.CL > FO.rEdgePosition)
                                            && (Column.CL < FO.lEdgePosition))) {
                                    // error condition
                                    if ((MsgBox((FO.Description + " has been found to intersect a sidewall column! Relocate or resize this window before proceeding. Continue anyways?"), (vbYesNo + vbCritical), "FO Placement Error") == 7)) {

                                    }

                                    // Exit Sub
                                }

                                ColIndex;
                                // create jambs
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.rEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = FO.Height;
                                Jamb.tEdgeHeight = FO.tEdgeHeight;
                                Jamb.bEdgeHeight = FO.bEdgeHeight;
                                FO.FOMaterials.Add;
                                Jamb;
                                Jamb = new clsMember();
                                Jamb.mType = "FO Material";
                                Jamb.Size = "8"" Receiver Cee";
                                Jamb.Width = 2.5;
                                Jamb.CL = FO.lEdgePosition;
                                Jamb.rEdgePosition = (Jamb.CL
                                            - (Jamb.Width / 2));
                                Jamb.Length = FO.Height;
                                Jamb.tEdgeHeight = FO.tEdgeHeight;
                                Jamb.bEdgeHeight = FO.bEdgeHeight;
                                FO.FOMaterials.Add;
                                Jamb;
                                "MiscFO";
                                for (ColIndex = ColumnCollection.Count; (ColIndex <= 1); ColIndex = (ColIndex + -1)) {
                                    Column = ColumnCollection[ColIndex];
                                    // if within 6" of a column CL
                                    if (((Column.CL
                                                > (FO.rEdgePosition - (1 * 6)))
                                                && (Column.CL
                                                < (FO.lEdgePosition + (1 * 6))))) {
                                        // error condition
                                        if ((MsgBox((FO.Description + " has been found to be intersecting a sidewall column! Relocate or resize this misc. FO opening before proceeding. Continue anyways?"), (vbYesNo + vbCritical), "FO Placement Error") == 7)) {

                                        }

                                        // Exit Sub
                                    }

                                    ColIndex;
                                    // create jambs
                                    // add structural steel depending on options selected in input field
                                    switch (FO.StructuralSteelOption) {
                                        case "Full Height Jambs w/ Header & Stool":
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.rEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL);
                                            Jamb.tEdgeHeight = Jamb.Length;
                                            Jamb.bEdgeHeight = 0;
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.lEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL);
                                            Jamb.tEdgeHeight = Jamb.Length;
                                            Jamb.bEdgeHeight = 0;
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            // header and stool
                                            Purlin = new clsMember();
                                            Purlin.Length = FO.Width;
                                            Purlin.Size = "8"" C Purlin";
                                            Purlin.mType = "FO Material";
                                            if ((FO.tEdgeHeight == 86)) {
                                                Purlin.Qty = 1;
                                            }
                                            else {
                                                Purlin.Qty = 2;
                                            }

                                            FO.FOMaterials.Add;
                                            Purlin;
                                            // weld clips
                                            // b.WeldClips = b.WeldClips + 10
                                            // Set WeldClips = New clsMiscItem
                                            // WeldClips.Quantity = 10
                                            // WeldClips.Name = "Weld Clips"
                                            // FO.FOMaterials.Add WeldClips
                                            break;
                                        case "7'2"" Jambs w/ Header & Stool":
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.rEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.lEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            // header and stool
                                            Purlin = new clsMember();
                                            Purlin.Length = FO.Width;
                                            Purlin.Size = "8"" C Purlin";
                                            Purlin.mType = "FO Material";
                                            if ((FO.tEdgeHeight == 86)) {
                                                Purlin.Qty = 1;
                                            }
                                            else {
                                                Purlin.Qty = 2;
                                            }

                                            FO.FOMaterials.Add;
                                            Purlin;
                                            // weld clips
                                            // b.WeldClips = b.WeldClips + 6
                                            // Set WeldClips = New clsMiscItem
                                            // WeldClips.Quantity = 6
                                            // WeldClips.Name = "Weld Clips"
                                            // FO.FOMaterials.Add WeldClips
                                            break;
                                        case "7'2"" Jambs w/ Stool":
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.rEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.lEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            // stool
                                            Purlin = new clsMember();
                                            Purlin.Length = FO.Width;
                                            Purlin.Size = "8"" C Purlin";
                                            Purlin.mType = "FO Material";
                                            if ((FO.tEdgeHeight == 86)) {
                                                Purlin.Qty = 1;
                                            }
                                            else {
                                                Purlin.Qty = 2;
                                            }

                                            FO.FOMaterials.Add;
                                            Purlin;
                                            // weld clips
                                            // b.WeldClips = b.WeldClips + 4
                                            // Set WeldClips = New clsMiscItem
                                            // WeldClips.Quantity = 4
                                            // WeldClips.Name = "Weld Clips"
                                            // FO.FOMaterials.Add WeldClips
                                            break;
                                        case "7'2"" Jambs":
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.rEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.lEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (2 + (7 * 12));
                                            Jamb.tEdgeHeight = (2 + (7 * 12));
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            // weld clips
                                            // b.WeldClips = b.WeldClips + 2
                                            // Set WeldClips = New clsMiscItem
                                            // WeldClips.Quantity = 2
                                            // WeldClips.Name = "Weld Clips"
                                            // FO.FOMaterials.Add WeldClips
                                            break;
                                        case "5' Jambs w/ Header & Stool":
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.rEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (5 * 12);
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            Jamb = new clsMember();
                                            Jamb.mType = "FO Material";
                                            Jamb.Size = "8"" Receiver Cee";
                                            Jamb.Width = 2.5;
                                            Jamb.CL = FO.lEdgePosition;
                                            Jamb.rEdgePosition = (Jamb.CL
                                                        - (Jamb.Width / 2));
                                            Jamb.Length = (5 * 12);
                                            FO.FOMaterials.Add;
                                            Jamb;
                                            // header and stool
                                            Purlin = new clsMember();
                                            Purlin.Length = FO.Width;
                                            Purlin.Size = "8"" C Purlin";
                                            Purlin.mType = "FO Material";
                                            if ((FO.tEdgeHeight == 86)) {
                                                Purlin.Qty = 1;
                                            }
                                            else {
                                                Purlin.Qty = 2;
                                            }

                                            FO.FOMaterials.Add;
                                            Purlin;
                                            // weld clips
                                            // b.WeldClips = b.WeldClips + 8
                                            // Set WeldClips = New clsMiscItem
                                            // WeldClips.Quantity = 8
                                            // WeldClips.Name = "Weld Clips"
                                            // FO.FOMaterials.Add WeldClips
                                            break;
                                    }

                                    FO;
                                    // '''''''''''''''''''' Sub used to sort arrays in ascending order. Currently only used for Endwall centerline calc as of 10/4/2021 at 8:00 PM EST --------------
                                    QuickSort((<void>(arr)), Variant, (<number>(first)), (<number>(last)));
                                    let vCentreVal: Object;
                                    (<void>(vTemp));
                                    let lTempLow: number;
                                    let lTempHi: number;
                                    lTempLow = first;
                                    lTempHi = last;
                                    vCentreVal = arr((first + last), 2);
                                    while ((lTempLow <= lTempHi)) {
                                        while (((arr(lTempLow) < vCentreVal)
                                                    && (lTempLow < last))) {
                                            lTempLow = (lTempLow + 1);
                                        }

                                        while (((vCentreVal < arr(lTempHi))
                                                    && (lTempHi > first))) {
                                            lTempHi = (lTempHi - 1);
                                        }

                                        if ((lTempLow <= lTempHi)) {
                                            //  Swap values
                                            vTemp = arr(lTempLow);
                                            arr(lTempLow) = arr(lTempHi);
                                            arr(lTempHi) = vTemp;
                                            //  Move to next positions
                                            lTempLow = (lTempLow + 1);
                                            lTempHi = (lTempHi - 1);
                                        }

                                    }

                                    if ((first < lTempHi)) {
                                        QuickSort;
                                    }

                                    arr;
                                    first;
                                    lTempHi;
                                    if ((lTempLow < last)) {
                                        QuickSort;
                                    }

                                    arr;
                                    lTempLow;
                                    last;
                                    ReverseArray((<void>(vArray)), Variant);
                                    // Reverse the order of an array, so if it's already sorted
                                    // from smallest to largest, it will now be sorted from
                                    // largest to smallest.
                                    let vTemp: Object;
                                    let i: number;
                                    let iUpper: number;
                                    let iMidPt: number;
                                    iUpper = UBound(vArray);
                                    iMidPt = (UBound(vArray) - LBound(vArray));
                                    (2 + LBound(vArray));
                                    for (i = LBound(vArray); (i <= iMidPt); i++) {
                                        vTemp = vArray(iUpper);
                                        vArray(iUpper) = vArray(i);
                                        vArray(i) = vTemp;
                                        iUpper = (iUpper - 1);
                                    }

                                    NewExpandableEndwallColumnsGen((<clsBuilding>(b)), (<string>(eWall)), (<number>(EndwallColumnCLs())), Optional, (<number>(NewColNum)), Optional, (<boolean>(Reiterate)));
                                    let ColNum: number;
                                    let MaxHorizontalDistance: number;
                                    let ColLocation: number[];
                                    let Column: clsMember;
                                    let DistanceToPreviousColumn: number;
                                    let DistanceToNextColumn: number;
                                    let i: number;
                                    MaxHorizontalDistance = (60 / Sqr(((b.rPitch / 12) | (2 + 1))));
                                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                                    if ((NewColNum != 0)) {
                                        ColNum = NewColNum;
                                    }
                                    else {
                                        if ((b.rShape == "Gable")) {
                                            if ((b.bWidth <= 80)) {
                                                ColNum = 0;
                                            }
                                            else if (((b.bWidth > 80)
                                                        && (b.bWidth
                                                        < (MaxHorizontalDistance * 2)))) {
                                                ColNum = 1;
                                            }
                                            else if ((b.bWidth
                                                        >= (MaxHorizontalDistance * 2))) {
                                                ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                                            }

                                        }
                                        else if ((b.rShape == "Single Slope")) {
                                            if ((b.bWidth < MaxHorizontalDistance)) {
                                                ColNum = 0;
                                            }
                                            else if ((b.bWidth > MaxHorizontalDistance)) {
                                                ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                                            }

                                        }

                                        // lower Col Num by 1 on first iteration to check for marginal cases
                                        // some column widths (to be determined) will require less columns, this will check those cases
                                        if ((ColNum > 0)) {
                                            ColNum = (ColNum - 1);
                                        }

                                    }

                                    // first, evenly space columns along the width of the building to adjust later; add to array
                                    let ColLocation: Object;
                                    ColLocation[0] = 0;
                                    ColLocation[(ColNum + 1)] = (b.bWidth * 12);
                                    switch (ColNum) {
                                        case 1:
                                            ColLocation[1] = (b.bWidth / (2 * 12));
                                            break;
                                        case 2:
                                            ColLocation[1] = (b.bWidth / (3 * 12));
                                            ColLocation[2] = (b.bWidth / (3 * (12 * 2)));
                                            break;
                                        case 3:
                                            ColLocation[1] = (b.bWidth / (4 * 12));
                                            ColLocation[2] = (b.bWidth / (4 * (12 * 2)));
                                            ColLocation[3] = (b.bWidth / (4 * (12 * 3)));
                                            break;
                                        case 4:
                                            ColLocation[1] = (b.bWidth / (5 * 12));
                                            ColLocation[2] = (b.bWidth / (5 * (12 * 2)));
                                            ColLocation[3] = (b.bWidth / (5 * (12 * 3)));
                                            ColLocation[4] = (b.bWidth / (5 * (12 * 4)));
                                            break;
                                    }

                                    // loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
                                    for (i = 1; (i <= ColNum); i++) {
                                        if ((ConflictingEndwallOHDoor(ColLocation[i], b, eWall) == true)) {
                                            ColLocation[i] = NearestEndwallLocation(ColLocation[i], b, ,, eWall);
                                        }

                                    }

                                    // ''''''''''''''check for No Interior Columns
                                    if ((ColNum == 0)) {
                                        // '''''''''''''Distance between Columns
                                        DistanceToPreviousColumn = Abs((ColLocation[0] - ColLocation[1]));
                                        // '''''''''''''Estimate COlumn widths
                                        // get first width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[0]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[0] - ColLocation[1]));
                                        // subtract half of first width
                                        DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                                        // get second width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[1]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[0] - ColLocation[1]));
                                        // subtract half of second width
                                        DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                                        if ((DistanceToPreviousColumn
                                                    > (MaxHorizontalDistance * 12))) {
                                            ColLocation;
                                            EndwallColumnCLs;
                                            NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs(), (ColNum + 1), true);
                                            return;
                                        }

                                    }

                                    // ''''''''''''''check Interior Columns
                                    // check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
                                    for (i = 1; (i <= ColNum); i++) {
                                        // get distance to next column to make sure it does NOT exceed max rafter length
                                        // if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
                                        // estimate column widths to get accurate distances
                                        // '''''''''''''Distance to PREVIOUS Column
                                        if (((ColLocation[i]
                                                    > (b.bWidth * (12 / 2)))
                                                    && ((ColLocation[(i - 1)]
                                                    < (b.bWidth * (12 / 2)))
                                                    && (b.rShape == "Gable")))) {
                                            DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                                                            - ColLocation[i]));
                                        }
                                        else {
                                            DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                                        }

                                        // '''''''''''''Estimate COlumn widths
                                        // get first width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[(i - 1)] - ColLocation[i]));
                                        // subtract half of width
                                        DistanceToPreviousColumn = (DistanceToPreviousColumn
                                                    - (Column.Width / 2));
                                        // get second width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[(i - 1)]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[(i - 1)] - ColLocation[i]));
                                        // subtract width if sidewall column, or half of width otherwise
                                        if (((i - 1)
                                                    == 0)) {
                                            DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                                        }
                                        else {
                                            DistanceToPreviousColumn = (DistanceToPreviousColumn
                                                        - (Column.Width / 2));
                                        }

                                        // '''''''''''''Distance to NEXT Column
                                        if (((ColLocation[i]
                                                    < (b.bWidth * (12 / 2)))
                                                    && ((ColLocation[(i + 1)]
                                                    > (b.bWidth * (12 / 2)))
                                                    && (b.rShape == "Gable")))) {
                                            DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                                                            - ColLocation[i]));
                                        }
                                        else {
                                            DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                                        }

                                        // '''''''''''''Estimate COlumn widths
                                        // get first width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[(i + 1)] - ColLocation[i]));
                                        // subtract half of width
                                        DistanceToNextColumn = (DistanceToNextColumn
                                                    - (Column.Width / 2));
                                        // get second width
                                        Column = new clsMember();
                                        Column.Length = b.DistanceToRoof("e1", ColLocation[(i + 1)]);
                                        Column.tEdgeHeight = Column.Length;
                                        Column.SetSize;
                                        b;
                                        "Column";
                                        "Interior";
                                        Abs((ColLocation[(i + 1)] - ColLocation[i]));
                                        // subtract width if sidewall column, or half of width otherwise
                                        if (((i + 1)
                                                    == UBound(ColLocation[]))) {
                                            DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
                                        }
                                        else {
                                            DistanceToNextColumn = (DistanceToNextColumn
                                                        - (Column.Width / 2));
                                        }

                                        // check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
                                        if (((DistanceToPreviousColumn
                                                    > (MaxHorizontalDistance * 12))
                                                    || (DistanceToNextColumn
                                                    > (MaxHorizontalDistance * 12)))) {
                                            // Debug.Print "columns too far apart"
                                            // CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
                                            if ((NearestEndwallLocation(ColLocation[i], b, "Alternate") != ColLocation[i])) {
                                                ColLocation[i] = NearestEndwallLocation(ColLocation[i], b, "Alternate");
                                                // '''''''''''''Distance to PREVIOUS Column
                                                if (((ColLocation[i]
                                                            > (b.bWidth * (12 / 2)))
                                                            && ((ColLocation[(i - 1)]
                                                            < (b.bWidth * (12 / 2)))
                                                            && (b.rShape == "Gable")))) {
                                                    DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                                                                    - ColLocation[i]));
                                                }
                                                else {
                                                    DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                                                }

                                                // '''''''''''''Estimate COlumn widths
                                                // get first width
                                                Column = new clsMember();
                                                Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                                                Column.tEdgeHeight = Column.Length;
                                                Column.SetSize;
                                                b;
                                                "Column";
                                                "Interior";
                                                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                                                // subtract half of width
                                                DistanceToPreviousColumn = (DistanceToPreviousColumn
                                                            - (Column.Width / 2));
                                                // get second width
                                                Column = new clsMember();
                                                Column.Length = b.DistanceToRoof("e1", ColLocation[(i - 1)]);
                                                Column.tEdgeHeight = Column.Length;
                                                Column.SetSize;
                                                b;
                                                "Column";
                                                "Interior";
                                                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                                                // subtract width if sidewall column, or half of width otherwise
                                                if (((i - 1)
                                                            == 0)) {
                                                    DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                                                }
                                                else {
                                                    DistanceToPreviousColumn = (DistanceToPreviousColumn
                                                                - (Column.Width / 2));
                                                }

                                                // '''''''''''''Distance to NEXT Column
                                                if (((ColLocation[i]
                                                            < (b.bWidth * (12 / 2)))
                                                            && ((ColLocation[(i + 1)]
                                                            > (b.bWidth * (12 / 2)))
                                                            && (b.rShape == "Gable")))) {
                                                    DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                                                                    - ColLocation[i]));
                                                }
                                                else {
                                                    DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                                                }

                                                // '''''''''''''Estimate COlumn widths
                                                // get first width
                                                Column = new clsMember();
                                                Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                                                Column.tEdgeHeight = Column.Length;
                                                Column.SetSize;
                                                b;
                                                "Column";
                                                "Interior";
                                                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                                                // subtract half of width
                                                DistanceToNextColumn = (DistanceToNextColumn
                                                            - (Column.Width / 2));
                                                // get second width
                                                Column = new clsMember();
                                                Column.Length = b.DistanceToRoof("e1", ColLocation[(i + 1)]);
                                                Column.tEdgeHeight = Column.Length;
                                                Column.SetSize;
                                                b;
                                                "Column";
                                                "Interior";
                                                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                                                // subtract width if sidewall column, or half of width otherwise
                                                if (((i + 1)
                                                            == UBound(ColLocation[]))) {
                                                    DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
                                                }
                                                else {
                                                    DistanceToNextColumn = (DistanceToNextColumn
                                                                - (Column.Width / 2));
                                                }

                                            }

                                            if (((DistanceToPreviousColumn
                                                        > (MaxHorizontalDistance * 12))
                                                        || (DistanceToNextColumn
                                                        > (MaxHorizontalDistance * 12)))) {
                                                ColLocation;
                                                EndwallColumnCLs();
                                                NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs(), (ColNum + 1), true);
                                                return;
                                            }

                                            //     ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
                                            //     Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
                                            //     'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
                                            //         EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
                                            //         StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
                                            //         PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
                                            //         DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
                                            //         DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
                                            //         If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
                                            //             Erase ColLocation
                                            //             Call IntColumnsGen(b, ColNum + 1)
                                            //             Exit Sub
                                            //         End If
                                        }

                                    }

                                    let EndwallColumnCLs: Object;
                                    for (i = 0; (i <= UBound(ColLocation)); i++) {
                                        EndwallColumnCLs[i] = ColLocation[i];
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
// ''''''''''''' generates endwall column centerlines, optimizing for like groupings of girt segments, symetric spacing, and variable center column requirements.
EndwallColumnCLCalc(b: clsBuilding, eWall: string) {
    let TwentyFootQty: number;
    // Warning!!! Optional parameters not supported
    let TwentyFiveQty: number;
    let ThirtyQty: number;
    let MinSegs: number;
    let Girt: clsMember;
    let Column: clsMember;
    let tempGirtSpan: number;
    let EndwallColumnCLs: number[];
    let tempEndwallCLs: number[];
    let ColCount: number;
    let SpanCount: number;
    let i: number;
    let GirtSpan: number;
    let ColNum: number;
    let TotalSegmentGroupLength: number;
    let HalfWallSegCount: number;
    let EndwallSecondHalfCLs: number[];
    // Dim PartialSegmentTotal As Integer
    let PreviousSegment: number;
    let NextSegment: number;
    let LargestSegmentSize: number;
    let CenterGirtLength: number;
    let LoadBearingColumn: boolean;
    let EndwallGirts: Collection = new Collection();
    let FO: clsFO;
    let WallColumns: Collection;
    let FOs: Collection;
    let DistanceToS2: number;
    let tempColLocation: number;
    let j: number;
    let LongerDistance: number;
    let IntColumn: clsMember;
    let IdealSpan: number;
    let StartCol: clsMember;
    let StartPos: number;
    let EndPos: number;
    let tempPos: number;
    let MaxHorizontalDistance: number;
    let DistanceToNextCol: number;
    let DistanceToPrevCol: number;
    let NextColumn: clsMember;
    let NewColumn: clsMember;
    let StartPosRight: number;
    let StartPosLeft: number;
    let CenterFO: boolean;
    let lGtob: number;
    let rGtob: number;
    let Jamb: clsMember;
    let JambSupport: clsMember;
    i = 0;
    if ((eWall == "e1")) {
        WallColumns = b.e1Columns;
        FOs = b.e1FOs;
    }
    else if ((eWall == "e3")) {
        WallColumns = b.e3Columns;
        FOs = b.e3FOs;
    }

    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Expandable Endwall Columns
    // create load-bearing columns identical to interior columns
    if (b.ExpandableEndwall(eWall)) {
        MaxHorizontalDistance = ((60 / Sqr(((b.rPitch / 12) | (2 + 1))))
                    * 12);
        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
        if ((EstSht.Range("BayNum").Value <= 1)) {
            // single bay has no interior columns
            let EndwallColumnCLs: Object;
            NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs[]);
        }
        else {
            ColCount = b.InteriorColumns.Count;
            let EndwallColumnCLs: Object;
            for (i = 1; (i <= ColCount); i++) {
                IntColumn = b.InteriorColumns(i);
                EndwallColumnCLs[(i - 1)] = IntColumn.CL;
            }

        }

        if ((eWall == "e1")) {
            QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs));
        }
        else if ((eWall == "e3")) {
            for (i = 0; (i <= UBound(EndwallColumnCLs)); i++) {
                EndwallColumnCLs[i] = ((b.bWidth * 12)
                            - EndwallColumnCLs[i]);
            }

            QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs));
            ReverseArray(EndwallColumnCLs);
        }

        for (i = 0; (i <= UBound(EndwallColumnCLs)); i++) {
            // Check MiscFOs and Windows that interfere with Load Bearing Columns
            for (FO in FOs) {
                if ((((FO.FOType == "Window")
                            || (FO.FOType == "MiscFO"))
                            && ((FO.rEdgePosition < EndwallColumnCLs[i])
                            && (FO.lEdgePosition > EndwallColumnCLs[i])))) {
                    // FO is in the way
                    // check closest jamb
                    if ((Abs((FO.rEdgePosition - EndwallColumnCLs[i])) < Abs((FO.lEdgePosition - EndwallColumnCLs[i])))) {
                        tempColLocation = (FO.rEdgePosition - 12);
                    }
                    else {
                        tempColLocation = (FO.lEdgePosition + 12);
                    }

                    DistanceToNextCol = Abs((tempColLocation - EndwallColumnCLs[(i + 1)]));
                    DistanceToPrevCol = Abs((tempColLocation - EndwallColumnCLs[(i - 1)]));
                    if (((DistanceToNextCol < MaxHorizontalDistance)
                                && (DistanceToPrevCol < MaxHorizontalDistance))) {
                        EndwallColumnCLs[i] = tempColLocation;
                    }
                    else {
                        // check other jamb
                        if ((Abs((FO.rEdgePosition - EndwallColumnCLs[i])) > Abs((FO.lEdgePosition - EndwallColumnCLs[i])))) {
                            tempColLocation = (FO.rEdgePosition - 12);
                        }
                        else {
                            tempColLocation = (FO.lEdgePosition + 12);
                        }

                        DistanceToNextCol = Abs((tempColLocation - EndwallColumnCLs[(i + 1)]));
                        DistanceToPrevCol = Abs((tempColLocation - EndwallColumnCLs[(i - 1)]));
                        if (((DistanceToNextCol < MaxHorizontalDistance)
                                    && (DistanceToPrevCol < MaxHorizontalDistance))) {
                            EndwallColumnCLs[i] = tempColLocation;
                        }
                        else {
                            // make both jambs load bearing
                            let Preserve: Object;
                            EndwallColumnCLs[(UBound(EndwallColumnCLs) + 1)];
                            EndwallColumnCLs[i] = (FO.rEdgePosition - 12);
                            EndwallColumnCLs[UBound(EndwallColumnCLs)] = (FO.lEdgePosition + 12);
                            // Re-Sort
                            if ((eWall == "e1")) {
                                QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs));
                            }
                            else if ((eWall == "e3")) {
                                QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs));
                                ReverseArray(EndwallColumnCLs);
                                for (j = 0; (j <= UBound(EndwallColumnCLs)); j++) {
                                    EndwallColumnCLs[j] = ((b.bWidth * 12)
                                                - EndwallColumnCLs[j]);
                                }

                            }

                        }

                    }

                }

            }

        }

        // set column variables, types, sizes, etc.
        for (i = 0; (i <= UBound(EndwallColumnCLs)); i++) {
            // find larger distance to neighboring columns to use in lookup tables
            // s2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
            if ((i == UBound(EndwallColumnCLs))) {
                LongerDistance = Abs((EndwallColumnCLs[i] - EndwallColumnCLs[(i - 1)]));
            }
            else if ((i == 0)) {
                LongerDistance = Abs((EndwallColumnCLs[i] - EndwallColumnCLs[(i + 1)]));
            }
            else {
                LongerDistance = Application.WorksheetFunction.Max(Abs((EndwallColumnCLs[i] - EndwallColumnCLs[(i - 1)])), Abs((EndwallColumnCLs[i] - EndwallColumnCLs[(i + 1)])));
            }

            Column = new clsMember();
            Column.mType = "Column";
            Column.CL = EndwallColumnCLs[i];
            Column.LoadBearing = true;
            if ((eWall == "e1")) {
                if ((b.rShape == "Single Slope")) {
                    if ((i == 0)) {
                        Column.Length = (((b.bWidth * 12)
                                    * (b.rPitch / 12))
                                    + (b.bHeight * 12));
                        Column.CL = 0;
                    }
                    else if ((i == UBound(EndwallColumnCLs))) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = (b.bWidth * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof("e1", Column.CL);
                    }

                }
                else {
                    // Gable
                    if ((i == 0)) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = 0;
                    }
                    else if ((i == UBound(EndwallColumnCLs))) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = (b.bWidth * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof("e1", Column.CL);
                    }

                }

            }
            else {
                // e3
                if ((b.rShape == "Single Slope")) {
                    if ((i == 0)) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = 0;
                    }
                    else if ((i == UBound(EndwallColumnCLs))) {
                        Column.Length = (((b.bWidth * 12)
                                    * (b.rPitch / 12))
                                    + (b.bHeight * 12));
                        Column.CL = (b.bWidth * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof(eWall, Column.CL);
                    }

                }
                else {
                    // Gable roof
                    if ((i == 0)) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = 0;
                    }
                    else if ((i == UBound(EndwallColumnCLs))) {
                        Column.Length = (b.bHeight * 12);
                        Column.CL = (b.bWidth * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof(eWall, Column.CL);
                    }

                }

            }

            Column.tEdgeHeight = Column.Length;
            Column.SetSize;
            b;
            "Column";
            "Interior";
            LongerDistance;
            if ((Column.CL == 0)) {
                Column.CL = (Column.Width / 2);
            }
            else if ((Column.CL
                        == (b.bWidth * 12))) {
                Column.CL = ((b.bWidth * 12)
                            - (Column.Width / 2));
            }

            Column.rEdgePosition = (Column.CL
                        - (Column.Width / 2));
            WallColumns.Add;
            Column;
        }

        // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Non Load Bearing Columns
        // Add non-load bearing columns in spaces greater than 30'
        ColCount = WallColumns.Count;
        // first check for largest possible space spanning gable roof to add column
        if ((b.rShape == "Gable")) {
            for (i = 1; (i
                        <= (ColCount - 1)); i++) {
                Column = WallColumns[i];
                NextColumn = WallColumns[(i + 1)];
                if (((Column.CL
                            < (b.bWidth * (12 / 2)))
                            && ((NextColumn.CL
                            > (b.bWidth * (12 / 2)))
                            && (Abs((Column.CL - NextColumn.CL)) > (30 * 12))))) {
                    NewColumn = new clsMember();
                    NewColumn.CL = (b.bWidth * (12 / 2));
                    NewColumn.LoadBearing = false;
                    NewColumn.bEdgeHeight = 0;
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL);
                    NewColumn.Length = NewColumn.tEdgeHeight;
                    if ((NewColumn.Length
                                < ((30 * 12)
                                + 4))) {
                        NewColumn.Size = "8"" C Purlin";
                        NewColumn.Width = 2.5;
                    }
                    else {
                        NewColumn.SetSize;
                        b;
                        "Column";
                        eWall;
                        GirtSpan;
                        "NonExpandable";
                    }

                    NewColumn.rEdgePosition = (NewColumn.CL
                                - (NewColumn.Width / 2));
                    WallColumns.Add;
                    NewColumn;
                    (i + 1);
                    ColCount = (ColCount + 1);
                }

            }

        }

        for (i = 1; (i
                    <= (ColCount - 1)); i++) {
            Column = WallColumns[i];
            NextColumn = WallColumns[(i + 1)];
            if (((b.rShape == "Single Slope")
                        && (eWall == "e1"))) {
                if ((Abs((Column.CL - NextColumn.CL)) > (30 * 12))) {
                    tempGirtSpan = (Abs((Column.CL - NextColumn.CL)) / 2);
                    GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", true);
                    NewColumn = new clsMember();
                    NewColumn.CL = (NextColumn.CL - GirtSpan);
                    NewColumn.LoadBearing = false;
                    NewColumn.bEdgeHeight = 0;
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL);
                    NewColumn.Length = NewColumn.tEdgeHeight;
                    if ((NewColumn.Length
                                < ((30 * 12)
                                + 4))) {
                        NewColumn.Size = "8"" C Purlin";
                        NewColumn.Width = 2.5;
                    }
                    else {
                        NewColumn.SetSize;
                        b;
                        "Column";
                        eWall;
                        GirtSpan;
                        "NonExpandable";
                    }

                    NewColumn.rEdgePosition = (NewColumn.CL
                                - (NewColumn.Width / 2));
                    WallColumns.Add;
                    NewColumn;
                }

            }
            else if (((b.rShape == "Single Slope")
                        && (eWall == "e3"))) {
                if ((Abs((Column.CL - NextColumn.CL)) > (30 * 12))) {
                    tempGirtSpan = (Abs((Column.CL - NextColumn.CL)) / 2);
                    GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", true);
                    NewColumn = new clsMember();
                    NewColumn.CL = (Column.CL + GirtSpan);
                    NewColumn.LoadBearing = false;
                    NewColumn.bEdgeHeight = 0;
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL);
                    NewColumn.Length = NewColumn.tEdgeHeight;
                    if ((NewColumn.Length
                                < ((30 * 12)
                                + 4))) {
                        NewColumn.Size = "8"" C Purlin";
                        NewColumn.Width = 2.5;
                    }
                    else {
                        NewColumn.SetSize;
                        b;
                        "Column";
                        eWall;
                        GirtSpan;
                        "NonExpandable";
                    }

                    NewColumn.rEdgePosition = (NewColumn.CL
                                - (NewColumn.Width / 2));
                    WallColumns.Add;
                    NewColumn;
                }

            }
            else {
                // Gable roofs
                Column = WallColumns[i];
                NextColumn = WallColumns[(i + 1)];
                if ((Column.CL
                            < (b.bWidth * (12 / 2)))) {
                    if ((Abs((Column.CL - NextColumn.CL)) > (30 * 12))) {
                        tempGirtSpan = (Abs((Column.CL - NextColumn.CL)) / 2);
                        GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", true);
                        NewColumn = new clsMember();
                        NewColumn.CL = (NextColumn.CL - GirtSpan);
                        NewColumn.LoadBearing = false;
                        NewColumn.bEdgeHeight = 0;
                        NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL);
                        NewColumn.Length = NewColumn.tEdgeHeight;
                        if ((NewColumn.Length
                                    < ((30 * 12)
                                    + 4))) {
                            NewColumn.Size = "8"" C Purlin";
                            NewColumn.Width = 2.5;
                        }
                        else {
                            NewColumn.SetSize;
                            b;
                            "Column";
                            eWall;
                            GirtSpan;
                            "NonExpandable";
                        }

                        NewColumn.rEdgePosition = (NewColumn.CL
                                    - (NewColumn.Width / 2));
                        WallColumns.Add;
                        NewColumn;
                    }

                }
                else if ((Abs((Column.CL - NextColumn.CL)) > (30 * 12))) {
                    tempGirtSpan = (Abs((Column.CL - NextColumn.CL)) / 2);
                    GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", true);
                    NewColumn = new clsMember();
                    NewColumn.CL = (Column.CL + GirtSpan);
                    NewColumn.LoadBearing = false;
                    NewColumn.bEdgeHeight = 0;
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL);
                    NewColumn.Length = NewColumn.tEdgeHeight;
                    if ((NewColumn.Length
                                < ((30 * 12)
                                + 4))) {
                        NewColumn.Size = "8"" C Purlin";
                        NewColumn.Width = 2.5;
                    }
                    else {
                        NewColumn.SetSize;
                        b;
                        "Column";
                        eWall;
                        GirtSpan;
                        "NonExpandable";
                    }

                    NewColumn.rEdgePosition = (NewColumn.CL
                                - (NewColumn.Width / 2));
                    WallColumns.Add;
                    NewColumn;
                }

            }

        }

        let PrevColumn: clsMember;
        let NextDistance: number;
        let PrevDistance: number;
        for (i = 1; (i <= WallColumns.Count); i++) {
            MaxHorizontalDistance = (30 * 12);
            NextDistance = (b.bWidth * 12);
            PrevDistance = 0;
            Column = WallColumns[i];
            if ((Column.LoadBearing == false)) {
                for (j = 1; (j <= WallColumns.Count); j++) {
                    if ((j != WallColumns.Count)) {
                        NextColumn = WallColumns[(j + 1)];
                    }

                    if ((j != 1)) {
                        PrevColumn = WallColumns[(j - 1)];
                    }

                    if ((j == WallColumns.Count)) {
                        NextDistance = (b.bWidth * 12);
                    }
                    else if (((Abs((Column.CL - NextColumn.CL)) < Abs((Column.CL - NextDistance)))
                                && (NextColumn.CL > Column.CL))) {
                        NextDistance = NextColumn.CL;
                    }

                    if ((j == 1)) {
                        PrevDistance = 0;
                    }
                    else if (((Abs((Column.CL - PrevColumn.CL)) < Abs((Column.CL - PrevDistance)))
                                && (PrevColumn.CL < Column.CL))) {
                        PrevDistance = PrevColumn.CL;
                    }

                }

                // Check MiscFOs and Windows that interfere with Load Bearing Columns
                for (FO in FOs) {
                    if (((FO.rEdgePosition < Column.CL)
                                && (FO.lEdgePosition > Column.CL))) {
                        // FO is in the way
                        // if OHDoor or MiscFO w/ full height jambs, remove column
                        if (((FO.FOType == "OHDoor")
                                    || ((FO.FOType == "MiscFO")
                                    && FO.StructuralSteelOption))) {
                            "*Full Height*";
                            WallColumns.Remove;
                            i;
                            break;
                        }

                        // check closest jamb
                        if ((Abs((FO.rEdgePosition - Column.CL)) < Abs((FO.lEdgePosition - Column.CL)))) {
                            tempColLocation = FO.rEdgePosition;
                        }
                        else {
                            tempColLocation = FO.lEdgePosition;
                        }

                        DistanceToNextCol = Abs((tempColLocation - NextDistance));
                        if ((i != 1)) {
                            DistanceToPrevCol = Abs((tempColLocation - PrevDistance));
                        }
                        else {
                            DistanceToPrevCol = Column.CL;
                        }

                        if (((DistanceToNextCol < MaxHorizontalDistance)
                                    && (DistanceToPrevCol < MaxHorizontalDistance))) {
                            Column.CL = tempColLocation;
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                            Column.Length = Column.tEdgeHeight;
                            if ((Column.Length
                                        > ((30 * 12)
                                        + 4))) {
                                Column.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                "NonExpandable";
                            }
                            else {
                                Column.Size = "8"" Receiver Cee";
                                Column.Width = 2.5;
                            }

                        }
                        else {
                            // check other jamb
                            if ((Abs((FO.rEdgePosition - Column.CL)) > Abs((FO.lEdgePosition - Column.CL)))) {
                                tempColLocation = FO.rEdgePosition;
                            }
                            else {
                                tempColLocation = FO.lEdgePosition;
                            }

                            DistanceToNextCol = Abs((tempColLocation - NextDistance));
                            if ((i != 1)) {
                                DistanceToPrevCol = Abs((tempColLocation - PrevDistance));
                            }
                            else {
                                DistanceToPrevCol = Column.CL;
                            }

                            if (((DistanceToNextCol < MaxHorizontalDistance)
                                        && (DistanceToPrevCol < MaxHorizontalDistance))) {
                                Column.CL = tempColLocation;
                                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                                Column.Length = Column.tEdgeHeight;
                                if ((Column.Length
                                            > ((30 * 12)
                                            + 4))) {
                                    Column.SetSize;
                                    b;
                                    "Column";
                                    eWall;
                                    30;
                                    "NonExpandable";
                                }
                                else {
                                    Column.Size = "8"" Receiver Cee";
                                    Column.Width = 2.5;
                                }

                            }
                            else {
                                // make another extra column at both edges
                                Column.CL = FO.rEdgePosition;
                                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                                Column.Length = Column.tEdgeHeight;
                                if ((Column.Length
                                            > ((30 * 12)
                                            + 4))) {
                                    Column.SetSize;
                                    b;
                                    "Column";
                                    eWall;
                                    30;
                                    "NonExpandable";
                                }
                                else {
                                    Column.Size = "8"" Receiver Cee";
                                    Column.Width = 2.5;
                                }

                                Column = new clsMember();
                                Column.CL = FO.lEdgePosition;
                                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                                Column.Length = Column.tEdgeHeight;
                                if ((Column.Length
                                            > ((30 * 12)
                                            + 4))) {
                                    Column.SetSize;
                                    b;
                                    "Column";
                                    eWall;
                                    30;
                                    "NonExpandable";
                                }
                                else {
                                    Column.Size = "8"" Receiver Cee";
                                    Column.Width = 2.5;
                                }

                                WallColumns.Add;
                                Columns;
                                (i + 1);
                            }

                        }

                    }

                }

            }

        }

    }
    else {
        // ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Non-Expandable Endwall Columns
        MaxHorizontalDistance = ((30 / Sqr(((b.rPitch / 12) | (2 + 1))))
                    * 12);
        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
        // largest/ideal girt length for wall
        if ((MaxHorizontalDistance >= (25 * 12))) {
            IdealSpan = (25 * 12);
        }
        else if ((MaxHorizontalDistance >= 20)) {
            IdealSpan = (20 * 12);
        }
        else {
            IdealSpan = (MaxHorizontalDistance * 12);
        }

        if ((b.rShape == "Single Slope")) {
            if ((eWall == "e1")) {
                StartPos = 0;
                EndPos = IdealSpan;
            }
            else {
                StartPos = (b.bWidth * 12);
                EndPos = ((b.bWidth * 12)
                            - IdealSpan);
            }

            if ((eWall == "e1")) {
                while ((EndPos
                            < (b.bWidth * 12))) {
                    if (((((b.bWidth * 12)
                                - StartPos)
                                < (IdealSpan * 1.5))
                                && (((b.bWidth * 12)
                                - StartPos)
                                < IdealSpan))) {
                        IdealSpan = (IdealSpan - 60);
                    }

                    if ((FOs.Count == 0)) {
                        tempPos = StartPos;
                    }
                    else {
                        tempPos = NonExpandableFOJambs(b, eWall, StartPos, MaxHorizontalDistance, IdealSpan, 1);
                    }

                    if ((tempPos == StartPos)) {
                        // no FOs interfered with ideal location, add new column
                        Column = new clsMember();
                        Column.bEdgeHeight = 0;
                        Column.CL = (StartPos + IdealSpan);
                        Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                        Column.Length = Column.tEdgeHeight;
                        Column.LoadBearing = true;
                        Column.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        Column.rEdgePosition = (Column.CL
                                    - (Column.Width / 2));
                        WallColumns.Add;
                        Column;
                        tempPos = Column.CL;
                    }

                    StartPos = tempPos;
                    EndPos = (tempPos + IdealSpan);
                    while ((EndPos > 0)) {
                        if ((StartPos
                                    < (IdealSpan * 1.5))) {
                            IdealSpan = (IdealSpan - 60);
                        }

                        if ((FOs.Count == 0)) {
                            tempPos = StartPos;
                        }
                        else {
                            tempPos = NonExpandableFOJambs(b, eWall, StartPos, MaxHorizontalDistance, IdealSpan, -1);
                        }

                        if ((tempPos == StartPos)) {
                            // no FOs interfered with ideal location, add new column
                            Column = new clsMember();
                            Column.bEdgeHeight = 0;
                            Column.CL = (StartPos - IdealSpan);
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                            Column.Length = Column.tEdgeHeight;
                            Column.LoadBearing = true;
                            Column.SetSize;
                            b;
                            "Column";
                            eWall;
                            30;
                            Column.rEdgePosition = (Column.CL
                                        - (Column.Width / 2));
                            WallColumns.Add;
                            Column;
                            tempPos = Column.CL;
                        }

                        StartPos = tempPos;
                        EndPos = (tempPos - IdealSpan);
                    }

                    // Gable Roof
                    for (FO in FOs) {
                        // Check if an FO is in the center of the endwall, if so, it MUST have both full height jambs and supports if necessary since it displaces the center column
                        if (((FO.rEdgePosition
                                    < (b.bWidth * (12 / 2)))
                                    && (FO.lEdgePosition
                                    > (b.bWidth * (12 / 2))))) {
                            StartPosRight = FO.rEdgePosition;
                            StartPosLeft = FO.lEdgePosition;
                            CenterFO = true;
                            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition);
                            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition);
                            // left jamb
                            Jamb = new clsMember();
                            Jamb.bEdgeHeight = 0;
                            if ((lGtob
                                        < ((30 * 12)
                                        + 4))) {
                                // don't need jamb support
                                Jamb.tEdgeHeight = lGtob;
                                Jamb.LoadBearing = true;
                            }
                            else {
                                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                                Jamb.tEdgeHeight = (30 * 12);
                                JambSupport = new clsMember();
                                JambSupport.bEdgeHeight = 0;
                                JambSupport.CL = FO.lEdgePosition;
                                JambSupport.LoadBearing = true;
                                JambSupport.tEdgeHeight = lGtob;
                                JambSupport.Length = lGtob;
                                JambSupport.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                JambSupport.rEdgePosition = (JambSupport.CL
                                            - (JambSupport.Width / 2));
                                WallColumns.Add;
                                JambSupport;
                            }

                            Jamb.Length = Jamb.tEdgeHeight;
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.CL = FO.lEdgePosition;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            FO.FOMaterials.Add;
                            Jamb;
                            // right jamb
                            Jamb = new clsMember();
                            Jamb.bEdgeHeight = 0;
                            if ((rGtob
                                        < ((30 * 12)
                                        + 4))) {
                                // don't need jamb support
                                Jamb.tEdgeHeight = rGtob;
                                Jamb.LoadBearing = true;
                            }
                            else {
                                // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                                Jamb.tEdgeHeight = (30 * 12);
                                JambSupport = new clsMember();
                                JambSupport.bEdgeHeight = 0;
                                JambSupport.CL = FO.rEdgePosition;
                                JambSupport.LoadBearing = true;
                                JambSupport.tEdgeHeight = rGtob;
                                JambSupport.Length = rGtob;
                                JambSupport.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                JambSupport.rEdgePosition = (JambSupport.CL
                                            - (JambSupport.Width / 2));
                                WallColumns.Add;
                                JambSupport;
                            }

                            Jamb.Length = Jamb.tEdgeHeight;
                            Jamb.Size = "8"" Receiver Cee";
                            Jamb.Width = 2.5;
                            Jamb.CL = FO.rEdgePosition;
                            Jamb.rEdgePosition = (Jamb.CL
                                        - (Jamb.Width / 2));
                            FO.FOMaterials.Add;
                            Jamb;
                        }

                    }

                    if ((CenterFO == false)) {
                        Column = new clsMember();
                        Column.bEdgeHeight = 0;
                        Column.CL = (b.bWidth * (12 / 2));
                        Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                        Column.Length = Column.tEdgeHeight;
                        Column.LoadBearing = true;
                        Column.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        Column.rEdgePosition = (Column.CL
                                    - (Column.Width / 2));
                        WallColumns.Add;
                        Column;
                        StartPosRight = (b.bWidth * (12 / 2));
                        StartPosLeft = (b.bWidth * (12 / 2));
                    }

                    if ((eWall == "e1")) {
                        EndPos = (StartPosRight - IdealSpan);
                    }
                    else {
                        EndPos = (StartPosRight - IdealSpan);
                    }

                    // First side of Gable roof; going right
                    while ((EndPos > 0)) {
                        if (((StartPosRight
                                    < (IdealSpan * 1.5))
                                    && (StartPosRight > IdealSpan))) {
                            IdealSpan = (IdealSpan - 60);
                        }

                        if ((FOs.Count == 0)) {
                            tempPos = StartPosRight;
                        }
                        else {
                            tempPos = NonExpandableFOJambs(b, eWall, StartPosRight, MaxHorizontalDistance, IdealSpan, -1);
                        }

                        if ((tempPos == StartPosRight)) {
                            // no FOs interfered with ideal location, add new column
                            Column = new clsMember();
                            Column.bEdgeHeight = 0;
                            Column.CL = (StartPosRight - IdealSpan);
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                            Column.Length = Column.tEdgeHeight;
                            Column.LoadBearing = true;
                            Column.SetSize;
                            b;
                            "Column";
                            eWall;
                            30;
                            Column.rEdgePosition = (Column.CL
                                        - (Column.Width / 2));
                            WallColumns.Add;
                            Column;
                            tempPos = Column.CL;
                        }

                        StartPosRight = tempPos;
                        EndPos = (tempPos - IdealSpan);
                        Debug.Print;
                        ("Created Column #: " + i);
                        i = (i + 1);
                        // Other side of Gable roof; going to the left
                        // reset ideal span
                        if ((MaxHorizontalDistance >= (25 * 12))) {
                            IdealSpan = (25 * 12);
                        }
                        else if ((MaxHorizontalDistance >= 20)) {
                            IdealSpan = (20 * 12);
                        }
                        else {
                            IdealSpan = (MaxHorizontalDistance * 12);
                        }

                        if ((eWall == "e1")) {
                            EndPos = (StartPosLeft + IdealSpan);
                        }
                        else {
                            EndPos = (StartPosLeft + IdealSpan);
                        }

                        while ((EndPos
                                    < (b.bWidth * 12))) {
                            if (((((b.bWidth * 12)
                                        - StartPosLeft)
                                        < (IdealSpan * 1.5))
                                        && (((b.bWidth * 12)
                                        - StartPosLeft)
                                        > IdealSpan))) {
                                IdealSpan = (IdealSpan - 60);
                            }

                            if ((FOs.Count == 0)) {
                                tempPos = StartPosLeft;
                            }
                            else {
                                tempPos = NonExpandableFOJambs(b, eWall, StartPosLeft, MaxHorizontalDistance, IdealSpan, 1);
                            }

                            if ((tempPos == StartPosLeft)) {
                                // no FOs interfered with ideal location, add new column
                                Column = new clsMember();
                                Column.bEdgeHeight = 0;
                                Column.CL = (StartPosLeft + IdealSpan);
                                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL);
                                Column.Length = Column.tEdgeHeight;
                                Column.LoadBearing = true;
                                Column.SetSize;
                                b;
                                "Column";
                                eWall;
                                30;
                                Column.rEdgePosition = (Column.CL
                                            - (Column.Width / 2));
                                WallColumns.Add;
                                Column;
                                tempPos = Column.CL;
                            }

                            StartPosLeft = tempPos;
                            EndPos = (tempPos + IdealSpan);
                            Debug.Print;
                            ("Created Column #: " + i);
                            i = (i + 1);
                        }

                        // ''''''''''''''Create Corner columns for Non-Expandable Endwalls
                        Column = new clsMember();
                        Column.bEdgeHeight = 0;
                        Column.CL = 0;
                        Column.tEdgeHeight = b.DistanceToRoof(eWall, 0);
                        Column.Length = Column.tEdgeHeight;
                        Column.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        Column.LoadBearing = true;
                        Column.CL = (Column.Width / 2);
                        Column.rEdgePosition = 0;
                        WallColumns.Add;
                        Column;
                        Column = new clsMember();
                        Column.bEdgeHeight = 0;
                        Column.CL = (b.bWidth * 12);
                        Column.tEdgeHeight = b.DistanceToRoof(eWall, (b.bWidth * 12));
                        Column.Length = Column.tEdgeHeight;
                        Column.SetSize;
                        b;
                        "Column";
                        eWall;
                        30;
                        Column.LoadBearing = true;
                        Column.CL = ((b.bWidth * 12)
                                    - (Column.Width / 2));
                        Column.rEdgePosition = (Column.CL
                                    - (Column.Width / 2));
                        WallColumns.Add;
                        Column;
                    }

                }

                BaseAngleTrimGen((<clsBuilding>(b)));
                let BaseAngleCollection: Collection;
                let BaseAngleTrimLength: number;
                let FO: clsFO;
                let Member: clsMember;
                let StartPos: number;
                let EndPos: number;
                let NextFOEdge: number;
                let NextStartPos: number;
                let AngleNetLength: number;
                let ReceiverCNetLength: number;
                let BaseOnly: boolean;
                let OHWidth: number;
                let Qty: number;
                // Endwall 1 - check if partial, excluded, or gable only
                if ((b.WallStatus("e1") == "Include")) {
                    if ((EstSht.Range("e1_LinerPanels") == "None")) {
                        BaseOnly = true;
                    }
                    else {
                        BaseOnly = false;
                    }

                    for (FO in b.e1FOs) {
                        if ((FO.FOType == "OHDoor")) {
                            OHWidth = (OHWidth + FO.Width);
                        }

                    }

                    BaseAngleTrimLength = ((b.bWidth * 12)
                                - OHWidth);
                    if (BaseOnly) {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                    }
                    else {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                        ReceiverCNetLength = (ReceiverCNetLength + BaseAngleTrimLength);
                    }

                }

                OHWidth = 0;
                BaseAngleTrimLength = 0;
                // Endwall 3 - check if partial, excluded, or gable only
                if ((b.WallStatus("e3") == "Include")) {
                    if ((EstSht.Range("e3_LinerPanels") == "None")) {
                        BaseOnly = true;
                    }
                    else {
                        BaseOnly = false;
                    }

                    for (FO in b.e3FOs) {
                        if ((FO.FOType == "OHDoor")) {
                            OHWidth = (OHWidth + FO.Width);
                        }

                    }

                    BaseAngleTrimLength = ((b.bWidth * 12)
                                - OHWidth);
                    if (BaseOnly) {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                    }
                    else {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                        ReceiverCNetLength = (ReceiverCNetLength + BaseAngleTrimLength);
                    }

                }

                OHWidth = 0;
                BaseAngleTrimLength = 0;
                // Sidewall 2 - check if partial, excluded, or gable only
                if ((b.WallStatus("s2") == "Include")) {
                    if ((EstSht.Range("s2_LinerPanels") == "None")) {
                        BaseOnly = true;
                    }
                    else {
                        BaseOnly = false;
                    }

                    for (FO in b.s2FOs) {
                        if ((FO.FOType == "OHDoor")) {
                            OHWidth = (OHWidth + FO.Width);
                        }

                    }

                    BaseAngleTrimLength = ((b.bLength * 12)
                                - OHWidth);
                    if (BaseOnly) {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                    }
                    else {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                        ReceiverCNetLength = (ReceiverCNetLength + BaseAngleTrimLength);
                    }

                }

                OHWidth = 0;
                BaseAngleTrimLength = 0;
                // Sidewall 4 - check if partial, excluded, or gable only
                if ((b.WallStatus("s4") == "Include")) {
                    if ((EstSht.Range("s4_LinerPanels") == "None")) {
                        BaseOnly = true;
                    }
                    else {
                        BaseOnly = false;
                    }

                    for (FO in b.s4FOs) {
                        if ((FO.FOType == "OHDoor")) {
                            OHWidth = (OHWidth + FO.Width);
                        }

                    }

                    BaseAngleTrimLength = ((b.bLength * 12)
                                - OHWidth);
                    if (BaseOnly) {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                    }
                    else {
                        AngleNetLength = (AngleNetLength + BaseAngleTrimLength);
                        ReceiverCNetLength = (ReceiverCNetLength + BaseAngleTrimLength);
                    }

                }

                if ((AngleNetLength > (25 * 12))) {
                    Qty = Application.WorksheetFunction.RoundDown((AngleNetLength / (25 * 12)), 0);
                    AngleNetLength = (AngleNetLength
                                - (Qty * (25 * 12)));
                    Member = new clsMember();
                    Member.Length = (25 * 12);
                    Member.Qty = Qty;
                    Member.Size = "2x4 Base Angle";
                    b.BaseAngleTrim.Add;
                    Member;
                }

                if ((AngleNetLength > (20 * 12))) {
                    Qty = Application.WorksheetFunction.RoundUp((AngleNetLength / (20 * 12)), 0);
                    AngleNetLength = (AngleNetLength
                                - (Qty * (20 * 12)));
                    Member = new clsMember();
                    Member.Size = "2x4 Base Angle";
                    Member.Length = (20 * 12);
                    Member.Qty = Qty;
                    b.BaseAngleTrim.Add;
                    Member;
                }

                if ((ReceiverCNetLength > (30 * 12))) {
                    Qty = Application.WorksheetFunction.RoundDown((ReceiverCNetLength / (30 * 12)), 0);
                    ReceiverCNetLength = (ReceiverCNetLength
                                - (Qty * (30 * 12)));
                    Member = new clsMember();
                    Member.Size = "8"" Receiver Cee";
                    Member.Length = (30 * 12);
                    Member.Qty = Qty;
                    b.BaseAngleTrim.Add;
                    Member;
                }

                if ((ReceiverCNetLength > (25 * 12))) {
                    Qty = Application.WorksheetFunction.RoundDown((ReceiverCNetLength / (25 * 12)), 0);
                    ReceiverCNetLength = (ReceiverCNetLength
                                - (Qty * (25 * 12)));
                    Member = new clsMember();
                    Member.Size = "8"" Receiver Cee";
                    Member.Length = (25 * 12);
                    Member.Qty = Qty;
                    b.BaseAngleTrim.Add;
                    Member;
                }

                if ((ReceiverCNetLength > (20 * 12))) {
                    Qty = Application.WorksheetFunction.RoundUp((ReceiverCNetLength / (20 * 12)), 0);
                    ReceiverCNetLength = (ReceiverCNetLength
                                - (Qty * (20 * 12)));
                    Member = new clsMember();
                    Member.Size = "8"" Receiver Cee";
                    Member.Length = (20 * 12);
                    Member.Qty = Qty;
                    b.BaseAngleTrim.Add;
                    Member;
                }

            }

            OverhangExtensionMembersGen((<clsBuilding>(b)));
            let Member: clsMember;
            let NewMember: clsMember;
            let CopyMember: clsMember;
            let LinerPanels: boolean;
            let e1Overhang: boolean;
            let e1Extension: boolean;
            let e3Overhang: boolean;
            let e3Extension: boolean;
            let s2Overhang: boolean;
            let s2Extension: boolean;
            let s4Overhang: boolean;
            let s4Extension: boolean;
            let Rafterlines: number;
            let Pitch: number;
            let Angle: number;
            let DistanceToLower: number;
            let DistanceToLengthen: number;
            let ExtensionHeight: number;
            let ExtensionWidth: number;
            let i: number;
            let RafterSize: string;
            let RafterWidth: number;
            let StartPos: number;
            let BayLength: number;
            let lEdgeStart: number;
            let rEdgeMax: number;
            let lEdgeMax: number;
            let rEdgeStart: number;
            let tEdgeMax: number;
            let bEdgeStart: number;
            let HorizontalDistance: number;
            let TotalSlopeLength: number;
            let RafterNum: number;
            let Size: string;
            let Width: number;
            // check for liner panels, overhangs, extension, soffit
            if ((EstSht.Range("Roof_LinerPanels").Value != "None")) {
                LinerPanels = true;
            }

            if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
                e1Overhang = true;
            }

            if ((EstSht.Range("e1_GableExtension").Value > 0)) {
                e1Extension = true;
            }

            if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
                e3Overhang = true;
            }

            if ((EstSht.Range("e3_GableExtension").Value > 0)) {
                e3Extension = true;
            }

            if ((EstSht.Range("s2_EaveOverhang").Value > 0)) {
                s2Overhang = true;
            }

            if ((EstSht.Range("s2_EaveExtension").Value > 0)) {
                s2Extension = true;
            }

            if ((EstSht.Range("s4_EaveOverhang").Value > 0)) {
                s4Overhang = true;
            }

            if ((EstSht.Range("s4_EaveExtension").Value > 0)) {
                s4Extension = true;
            }

            if (s2Extension) {
                // '''''''''''''''''''''''''''''''''s2 Extension
                Pitch = b.s2ExtensionPitch;
                Rafterlines = (EstSht.Range("BayNum").Value + 1);
                for (i = 1; (i <= Rafterlines); i++) {
                    if ((i == 1)) {
                        // '''' e1 rafter line
                        RafterSize = b.e1Rafters(1).Size;
                        RafterWidth = b.e1Rafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = (b.bWidth * 12);
                        NewMember.Length = (EstSht.Range("s2_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = (b.bHeight * 12);
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                    + Sqr((NewMember.Length
                                        | ((2
                                        - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                        | 2))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        ExtensionWidth = (NewMember.RafterLeftEdge
                                    - (b.bWidth * 12));
                        NewMember.SetSize;
                        b;
                        "Rafter";
                        "interior";
                        ExtensionWidth;
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        ExtensionHeight = NewMember.bEdgeHeight;
                        // '''''''''''''''''''''''''''''''''''''
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s2 Extension Rafter at Bay " + i);
                        b.e1Rafters.Add;
                        NewMember;
                        if (b.s2e1ExtensionIntersection) {
                            CopyMember = new clsMember();
                            CopyMember.Length = NewMember.Length;
                            CopyMember.Size = NewMember.Size;
                            CopyMember.tEdgeHeight = NewMember.tEdgeHeight;
                            CopyMember.bEdgeHeight = NewMember.bEdgeHeight;
                            CopyMember.rEdgePosition = NewMember.rEdgePosition;
                            CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge;
                            CopyMember.Width = NewMember.Width;
                            CopyMember.mType = "Extension Rafter";
                            CopyMember.Placement = "s2e1 Extension Intersection Rafter";
                            b.e1Rafters.Add;
                            CopyMember;
                        }

                    }
                    else if ((i < Rafterlines)) {
                        RafterSize = b.intRafters(1).Size;
                        RafterWidth = b.intRafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = (b.bWidth * 12);
                        NewMember.Length = (EstSht.Range("s2_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = (b.bHeight * 12);
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                    + Sqr((NewMember.Length
                                        | ((2
                                        - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                        | 2))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.SetSize;
                        b;
                        "Rafter";
                        "interior";
                        ExtensionWidth;
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s2 Extension Rafter at Bay " + i);
                        b.intRafters.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        // e3 rafter
                        RafterSize = b.e3Rafters(1).Size;
                        RafterWidth = b.e3Rafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = (0 - ExtensionWidth);
                        NewMember.Length = (EstSht.Range("s2_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = (b.bHeight * 12);
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.RafterLeftEdge = 0;
                        NewMember.SetSize;
                        b;
                        "Rafter";
                        "interior";
                        ExtensionWidth;
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s2 Extension Rafter at Bay " + i);
                        b.e3Rafters.Add;
                        NewMember;
                        if (b.s2e3ExtensionIntersection) {
                            CopyMember = new clsMember();
                            CopyMember.Length = NewMember.Length;
                            CopyMember.Size = NewMember.Size;
                            CopyMember.tEdgeHeight = NewMember.tEdgeHeight;
                            CopyMember.bEdgeHeight = NewMember.bEdgeHeight;
                            CopyMember.rEdgePosition = NewMember.rEdgePosition;
                            CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge;
                            CopyMember.Width = NewMember.Width;
                            CopyMember.mType = "Extension Rafter";
                            CopyMember.Placement = "s2e3 Extension Intersection Rafter";
                            b.e3Rafters.Add;
                            CopyMember;
                        }

                    }

                    // EXTENSION COLUMNS
                    NewMember = new clsMember();
                    if ((i == 1)) {
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                    + (ExtensionWidth + NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.RafterLeftEdge
                                    - (NewMember.Width / 2));
                        NewMember.rEdgePosition = (NewMember.RafterLeftEdge - NewMember.Width);
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s2 Extension Column at Bay " + i);
                        b.e1Columns.Add;
                        NewMember;
                        // SET BUILDING VARIABLE
                        // extension width is to inside of column
                        // extension height is to top of extension column
                        b.s2ExtensionWidth = (NewMember.lEdgePosition
                                    - (b.bWidth * 12));
                        b.s2ExtensionHeight = ExtensionHeight;
                        if (b.s2e1ExtensionIntersection) {
                            NewMember = new clsMember();
                            NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                            NewMember.SetSize;
                            b;
                            "Column";
                            "Interior";
                            ExtensionWidth;
                            NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                        + (ExtensionWidth + NewMember.Width));
                            NewMember.Length = (ExtensionHeight + DistanceToLower);
                            NewMember.CL = (NewMember.RafterLeftEdge
                                        - (NewMember.Width / 2));
                            NewMember.rEdgePosition = (NewMember.RafterLeftEdge - NewMember.Width);
                            NewMember.mType = "e1 Extension Column";
                            NewMember.Placement = "s2e1 Extension Intersection Column";
                            b.e1Columns.Add;
                            NewMember;
                        }

                    }
                    else if ((i < Rafterlines)) {
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                    + (ExtensionWidth + NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.RafterLeftEdge
                                    - (NewMember.Width / 2));
                        NewMember.rEdgePosition = (NewMember.RafterLeftEdge - NewMember.Width);
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s2 Extension Column at Bay " + i);
                        b.InteriorColumns.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.rEdgePosition = (0
                                    - (ExtensionWidth - NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.rEdgePosition
                                    + (NewMember.Width / 2));
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s2 Extension Column at Bay " + i);
                        b.e3Columns.Add;
                        NewMember;
                        if (b.s2e3ExtensionIntersection) {
                            NewMember = new clsMember();
                            NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                            NewMember.SetSize;
                            b;
                            "Column";
                            "Interior";
                            ExtensionWidth;
                            NewMember.rEdgePosition = (0
                                        - (ExtensionWidth - NewMember.Width));
                            NewMember.Length = (ExtensionHeight + DistanceToLower);
                            NewMember.CL = (NewMember.rEdgePosition
                                        + (NewMember.Width / 2));
                            NewMember.mType = "e3 Extension Column";
                            NewMember.Placement = "s2e3 Extension Intersection Column";
                            b.e3Columns.Add;
                            NewMember;
                        }

                    }

                }

            }

            let RightEdge: number;
            let TopEdge: number;
            let OverhangWidth: number;
            let OverhangHeight: number;
            if (s2Overhang) {
                // '''''''''''''''''''''''''''''''''s2 Overhang (always goes down)
                if ((b.s2ExtensionWidth > 0)) {
                    Pitch = b.s2ExtensionPitch;
                    RightEdge = ((b.bWidth * 12)
                                + b.s2ExtensionWidth);
                    TopEdge = b.s2ExtensionHeight;
                }
                else {
                    Pitch = b.rPitch;
                    RightEdge = (b.bWidth * 12);
                    TopEdge = (b.bHeight * 12);
                }

                Rafterlines = (EstSht.Range("BayNum").Value + 1);
                for (i = 1; (i <= Rafterlines); i++) {
                    if ((i == 1)) {
                        // '''' e1 rafter line
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = RightEdge;
                        NewMember.Length = (EstSht.Range("s2_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = TopEdge;
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                        | ((2
                                        - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                        | 2))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        OverhangWidth = (NewMember.RafterLeftEdge
                                    - (b.bWidth * 12));
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        OverhangHeight = NewMember.bEdgeHeight;
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = "s2 Stub Rafter at Bay 1";
                        b.e1Rafters.Add;
                        NewMember;
                    }
                    else if ((i < Rafterlines)) {
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = RightEdge;
                        NewMember.Length = (EstSht.Range("s2_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = TopEdge;
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                        | ((2
                                        - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                        | 2))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = ("s2 Stub Rafter at Bay " + i);
                        b.intRafters.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        // e3 rafter
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.RafterLeftEdge = (b.s2ExtensionWidth * -1);
                        NewMember.Length = (EstSht.Range("s2_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        NewMember.tEdgeHeight = TopEdge;
                        NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                    - (Sin(Angle) * NewMember.Length));
                        NewMember.rEdgePosition = (NewMember.RafterLeftEdge - Sqr((NewMember.Length
                                        | ((2
                                        - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                        | 2))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Angle = (Atn((Pitch / 12))
                                    * (NewMember.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (NewMember.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        DistanceToLengthen = Sqr((DistanceToLower
                                        | ((2 + NewMember.Width)
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                        NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = ("s2 Stub Rafter at Bay " + i);
                        b.e3Rafters.Add;
                        NewMember;
                    }

                }

                // ''''''''''''''''''add eave struts for s2 eave overhang
                StartPos = 0;
                for (i = 1; (i <= EstSht.Range("BayNum").Value); i++) {
                    BayLength = (EstSht.Range("Bay1_Length").offset((i - 1), 0).Value * 12);
                    NewMember = new clsMember();
                    NewMember.mType = "Eave Strut";
                    NewMember.rEdgePosition = StartPos;
                    NewMember.Length = BayLength;
                    NewMember.tEdgeHeight = OverhangHeight;
                    if ((b.rPitch == 1)) {
                        NewMember.Size = "8"" C Purlin";
                    }
                    else if ((EstSht.Range("s2_EaveOverhangSoffit").Value == "Yes")) {
                        NewMember.Size = ("8"" "
                                    + (b.rPitch + ":12 double up eave strut"));
                    }
                    else {
                        NewMember.Size = ("8"" "
                                    + (b.rPitch + ":12 single up eave strut"));
                    }

                    NewMember.Placement = "s2 Overhang Eave Strut";
                    b.RoofPurlins.Add;
                    NewMember;
                    StartPos = (StartPos + BayLength);
                }

            }

            // '''''''' s4
            if (s4Extension) {
                // '''''''''''''''''''''''''''''''''s4 Extension
                Pitch = b.s4ExtensionPitch;
                Rafterlines = (EstSht.Range("BayNum").Value + 1);
                for (i = 1; (i <= Rafterlines); i++) {
                    if ((i == 1)) {
                        // '''' e1 rafter line
                        RafterSize = b.e1Rafters(1).Size;
                        RafterWidth = b.e1Rafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.RafterLeftEdge = 0;
                        NewMember.Length = (EstSht.Range("s4_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = (b.bHeight * 12);
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (0 - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.rEdgePosition * -1);
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.bEdgeHeight;
                            // '''''''''''''''''''''''''''''''''''''
                        }
                        else {
                            NewMember.bEdgeHeight = ((b.bHeight * 12)
                                        + (b.bWidth * (12
                                        * (b.rPitch / 12))));
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (0 - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.rEdgePosition * -1);
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.tEdgeHeight;
                            // '''''''''''''''''''''''''''''''''''''
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s4 Extension Rafter at Bay " + i);
                        b.e1Rafters.Add;
                        NewMember;
                        if (b.s4e1ExtensionIntersection) {
                            CopyMember = new clsMember();
                            CopyMember.Length = NewMember.Length;
                            CopyMember.Size = NewMember.Size;
                            CopyMember.tEdgeHeight = NewMember.tEdgeHeight;
                            CopyMember.bEdgeHeight = NewMember.bEdgeHeight;
                            CopyMember.rEdgePosition = NewMember.rEdgePosition;
                            CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge;
                            CopyMember.Width = NewMember.Width;
                            CopyMember.mType = "Extension Rafter";
                            CopyMember.Placement = "s4e1 Extension Intersection Rafter";
                            b.e1Rafters.Add;
                            CopyMember;
                        }

                    }
                    else if ((i < Rafterlines)) {
                        RafterSize = b.intRafters(1).Size;
                        RafterWidth = b.intRafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.RafterLeftEdge = 0;
                        NewMember.Length = (EstSht.Range("s4_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = (b.bHeight * 12);
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (0 - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.rEdgePosition * -1);
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.bEdgeHeight;
                        }
                        else {
                            NewMember.bEdgeHeight = ((b.bHeight * 12)
                                        + (b.bWidth * (12
                                        * (b.rPitch / 12))));
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (0 - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.rEdgePosition * -1);
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.tEdgeHeight;
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s4 Extension Rafter at Bay " + i);
                        b.intRafters.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        // e3 rafter
                        RafterSize = b.e3Rafters(1).Size;
                        RafterWidth = b.e3Rafters(1).Width;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = (b.bWidth * 12);
                        NewMember.Length = (EstSht.Range("s4_EaveExtension").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = (b.bHeight * 12);
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.RafterLeftEdge
                                        - (b.bWidth * 12));
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.bEdgeHeight;
                        }
                        else {
                            NewMember.bEdgeHeight = ((b.bHeight * 12)
                                        + (b.bWidth * (12
                                        * (b.rPitch / 12))));
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            ExtensionWidth = (NewMember.RafterLeftEdge
                                        - (b.bWidth * 12));
                            NewMember.SetSize;
                            b;
                            "Rafter";
                            "interior";
                            ExtensionWidth;
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            ExtensionHeight = NewMember.tEdgeHeight;
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Extension Rafter";
                        NewMember.Placement = ("s4 Extension Rafter at Bay " + i);
                        b.e3Rafters.Add;
                        NewMember;
                        if (b.s4e3ExtensionIntersection) {
                            CopyMember = new clsMember();
                            CopyMember.Length = NewMember.Length;
                            CopyMember.Size = NewMember.Size;
                            CopyMember.tEdgeHeight = NewMember.tEdgeHeight;
                            CopyMember.bEdgeHeight = NewMember.bEdgeHeight;
                            CopyMember.rEdgePosition = NewMember.rEdgePosition;
                            CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge;
                            CopyMember.Width = NewMember.Width;
                            CopyMember.mType = "Extension Rafter";
                            CopyMember.Placement = "s4e3 Extension Intersection Rafter";
                            b.e3Rafters.Add;
                            CopyMember;
                        }

                    }

                    // EXTENSION COLUMNS
                    NewMember = new clsMember();
                    if ((i == 1)) {
                        // add to e1 columns
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.rEdgePosition = (0
                                    - (ExtensionWidth - NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.rEdgePosition
                                    + (NewMember.Width / 2));
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s4 Extension Column at Bay " + i);
                        b.e1Columns.Add;
                        NewMember;
                        // SET BUILDING VARIABLE
                        b.s4ExtensionWidth = NewMember.rEdgePosition;
                        b.s4ExtensionHeight = NewMember.tEdgeHeight;
                        if (b.s4e1ExtensionIntersection) {
                            NewMember = new clsMember();
                            NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                            NewMember.SetSize;
                            b;
                            "Column";
                            "Interior";
                            ExtensionWidth;
                            NewMember.rEdgePosition = (0
                                        - (ExtensionWidth - NewMember.Width));
                            NewMember.Length = (ExtensionHeight + DistanceToLower);
                            NewMember.CL = (NewMember.rEdgePosition
                                        + (NewMember.Width / 2));
                            NewMember.mType = "e1 Extension Column";
                            NewMember.Placement = "s4e1 Extension Intersection Column";
                            b.e1Columns.Add;
                            NewMember;
                        }

                    }
                    else if ((i < Rafterlines)) {
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.rEdgePosition = (0
                                    - (ExtensionWidth - NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.rEdgePosition
                                    + (NewMember.Width / 2));
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s4 Extension Column at Bay " + i);
                        b.InteriorColumns.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                        NewMember.SetSize;
                        b;
                        "Column";
                        "Interior";
                        ExtensionWidth;
                        NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                    + (ExtensionWidth + NewMember.Width));
                        NewMember.Length = (ExtensionHeight + DistanceToLower);
                        NewMember.CL = (NewMember.RafterLeftEdge
                                    - (NewMember.Width / 2));
                        NewMember.rEdgePosition = (NewMember.RafterLeftEdge - NewMember.Width);
                        NewMember.mType = "Extension Column";
                        NewMember.Placement = ("s4 Extension Column at Bay " + i);
                        b.e3Columns.Add;
                        NewMember;
                        if (b.s4e3ExtensionIntersection) {
                            NewMember = new clsMember();
                            NewMember.tEdgeHeight = (ExtensionHeight + DistanceToLower);
                            NewMember.SetSize;
                            b;
                            "Column";
                            "Interior";
                            ExtensionWidth;
                            NewMember.RafterLeftEdge = ((b.bWidth * 12)
                                        + (ExtensionWidth + NewMember.Width));
                            NewMember.Length = (ExtensionHeight + DistanceToLower);
                            NewMember.CL = (NewMember.RafterLeftEdge
                                        - (NewMember.Width / 2));
                            NewMember.rEdgePosition = (NewMember.RafterLeftEdge - NewMember.Width);
                            NewMember.mType = "e3 Extension Column";
                            NewMember.Placement = "s4e3 Extension Intersection Column";
                            b.e3Columns.Add;
                            NewMember;
                        }

                    }

                }

            }

            let LeftEdge: number;
            let BottomEdge: number;
            if (s4Overhang) {
                // '''''''''''''''''''''''''''''''''s4 Overhang (goes up for single slope)
                if ((b.s4ExtensionWidth < 0)) {
                    Pitch = b.s4ExtensionPitch;
                    LeftEdge = b.s4ExtensionWidth;
                    if ((b.rShape == "Gable")) {
                        TopEdge = b.s4ExtensionHeight;
                    }
                    else {
                        BottomEdge = b.s4ExtensionHeight;
                    }

                }
                else {
                    Pitch = b.rPitch;
                    LeftEdge = 0;
                    if ((b.rShape == "Gable")) {
                        TopEdge = (b.bHeight * 12);
                    }
                    else {
                        BottomEdge = ((b.bHeight * 12)
                                    + ((b.bWidth * 12)
                                    * (b.rPitch / 12)));
                    }

                }

                Rafterlines = (EstSht.Range("BayNum").Value + 1);
                for (i = 1; (i <= Rafterlines); i++) {
                    if ((i == 1)) {
                        // '''' e1 rafter line
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.RafterLeftEdge = LeftEdge;
                        NewMember.Length = (EstSht.Range("s4_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = TopEdge;
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (LeftEdge - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition * -1);
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.bEdgeHeight;
                        }
                        else {
                            NewMember.bEdgeHeight = BottomEdge;
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (LeftEdge - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition * -1);
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.tEdgeHeight;
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = ("s4 Stub Rafter at Bay " + i);
                        b.e1Rafters.Add;
                        NewMember;
                    }
                    else if ((i < Rafterlines)) {
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.RafterLeftEdge = LeftEdge;
                        NewMember.Length = (EstSht.Range("s4_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = TopEdge;
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (LeftEdge - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition * -1);
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.bEdgeHeight;
                        }
                        else {
                            NewMember.bEdgeHeight = BottomEdge;
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.rEdgePosition = (LeftEdge - Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition * -1);
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.tEdgeHeight;
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = ("s4 Stub Rafter at Bay " + i);
                        b.intRafters.Add;
                        NewMember;
                    }
                    else if ((i == Rafterlines)) {
                        // e3 rafter
                        RafterSize = "W8x10";
                        RafterWidth = 8;
                        NewMember = new clsMember();
                        NewMember.Size = RafterSize;
                        NewMember.Width = RafterWidth;
                        NewMember.rEdgePosition = RightEdge;
                        NewMember.Length = (EstSht.Range("s4_EaveOverhang").Value * 12);
                        Angle = Atn((Pitch / 12));
                        if ((b.rShape == "Gable")) {
                            NewMember.tEdgeHeight = TopEdge;
                            NewMember.bEdgeHeight = (NewMember.tEdgeHeight
                                        - (Sin(Angle) * NewMember.Length));
                            NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition
                                        - (b.bWidth * 12));
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.bEdgeHeight;
                        }
                        else {
                            NewMember.bEdgeHeight = BottomEdge;
                            NewMember.tEdgeHeight = (NewMember.bEdgeHeight
                                        + (Sin(Angle) * NewMember.Length));
                            NewMember.RafterLeftEdge = (NewMember.rEdgePosition + Sqr((NewMember.Length
                                            | ((2
                                            - (NewMember.tEdgeHeight - NewMember.bEdgeHeight))
                                            | 2))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            OverhangWidth = (NewMember.rEdgePosition
                                        - (b.bWidth * 12));
                            Angle = (Atn((Pitch / 12))
                                        * (NewMember.Width / 2));
                            DistanceToLower = Sqr((Angle
                                            | ((2
                                            + (NewMember.Width / 2))
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            DistanceToLengthen = Sqr((DistanceToLower
                                            | ((2 + NewMember.Width)
                                            | 2)));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            NewMember.tEdgeHeight = (NewMember.tEdgeHeight - DistanceToLower);
                            NewMember.bEdgeHeight = (NewMember.bEdgeHeight - DistanceToLower);
                            OverhangHeight = NewMember.tEdgeHeight;
                        }

                        NewMember.Length = (NewMember.Length + DistanceToLengthen);
                        NewMember.mType = "Overhang Stub Rafter";
                        NewMember.Placement = ("s4 Stub Rafter at Bay " + i);
                        b.e3Rafters.Add;
                        NewMember;
                    }

                }

                // ''''''''''''''''''add eave struts for s4 eave overhang
                StartPos = 0;
                for (i = EstSht.Range("BayNum").Value; (i <= 1); i = (i + -1)) {
                    BayLength = (EstSht.Range("Bay1_Length").offset((i - 1), 0).Value * 12);
                    NewMember = new clsMember();
                    NewMember.mType = "Eave Strut";
                    NewMember.rEdgePosition = StartPos;
                    NewMember.Length = BayLength;
                    NewMember.tEdgeHeight = ExtensionHeight;
                    if ((b.rShape == "Gable")) {
                        if ((b.rPitch == 1)) {
                            NewMember.Size = "8"" C Purlin";
                        }
                        else if ((EstSht.Range("s4_EaveOverhangSoffit").Value == "Yes")) {
                            NewMember.Size = ("8"" "
                                        + (b.rPitch + ":12 double up eave strut"));
                        }
                        else {
                            NewMember.Size = ("8"" "
                                        + (b.rPitch + ":12 single up eave strut"));
                        }

                    }
                    else if ((b.rPitch == 1)) {
                        NewMember.Size = "8"" C Purlin";
                    }
                    else if ((EstSht.Range("s2_EaveOverhangSoffit").Value == "Yes")) {
                        NewMember.Size = ("8"" "
                                    + (b.rPitch + ":12 double down eave strut"));
                    }
                    else {
                        NewMember.Size = ("8"" "
                                    + (b.rPitch + ":12 single down eave strut"));
                    }

                    NewMember.Placement = "s4 Overhang Eave Strut";
                    b.RoofPurlins.Add;
                    NewMember;
                    StartPos = (StartPos + BayLength);
                }

            }

            // '''''''' e1
            // if the endwall was non-expandable, the rafter needs to be changed back to a C Purlin
            // only rafters with extension attached need to be changed, so if not extension intersections, leave as Receiver Cees
            if ((e1Overhang || e1Extension)) {
                for (i = 1; (i <= b.e1Rafters.Count); i++) {
                    Member = b.e1Rafters(i);
                    if (((b.WallStatus("e1") == "No")
                                && (((Member.rEdgePosition < 0)
                                && b.s4e1ExtensionIntersection)
                                || (((Member.RafterLeftEdge
                                > (b.bWidth * 12))
                                && b.s2e1ExtensionIntersection)
                                || ((Member.rEdgePosition >= 0)
                                && (Member.RafterLeftEdge
                                <= (b.bWidth * 12))))))) {
                        if ((Member.Size == "8"" Receiver Cee")) {
                            Member.Size = "8"" C Purlin";
                        }
                        else if ((Member.Size == "10"" Receiver Cee")) {
                            Member.Size = "10"" C Purlin";
                        }

                    }

                }

            }

            // ''''''''''''''''''''''''''''''ENDWALL OVERHANGS AND EXTENSIONS
            // e1
            // Add rafters and columns; essentially copied from interior columns and rafters in terms of positioning, coordinates, and size
            if (e1Extension) {
                // If b.InteriorColumns.Count = 0 Then
                // if no interior columns have been made, create e1 extension columns and rafters
                EndwallExtensionColumnsGen(b, "e1");
                RafterGen(b, "e1 Extension");
                for (Member in b.e1ExtensionMembers) {
                    b.e1Columns.Add;
                    Member;
                }

                // Else
                //     For i = 1 To b.intRafters.Count
                //         Set Member = b.intRafters(i)
                //         If Member.Size = "8"" Receiver Cee" Then
                //             Member.Size = "8"" C Purlin"
                //         End If
                //         If Member.mType Like "*Overhang Stub Rafter*" Then
                //             'Do not copy
                //         Else
                //             Set NewMember = New clsMember
                //             NewMember.rEdgePosition = Member.rEdgePosition
                //             NewMember.RafterLeftEdge = Member.RafterLeftEdge
                //             NewMember.Length = Member.Length
                //             NewMember.tEdgeHeight = Member.tEdgeHeight
                //             NewMember.bEdgeHeight = Member.bEdgeHeight
                //             NewMember.mType = "e1 Extension Rafter"
                //             ''''''''''''''''''''''''''''''''''''''''''''''''''''NEED TO DEFINE EXTENSION RAFTER SIZE BASED ON EXPANDABLE, etc.
                //             NewMember.Size = Member.Size
                //             NewMember.Width = Member.Width
                //             NewMember.Placement = "e1 Extension Rafter"
                //             b.e1Rafters.Add NewMember
                //         End If
                //     Next i
                for (i = 1; (i <= b.InteriorColumns.Count); i++) {
                    // adding extensions
                    Member = b.InteriorColumns(i);
                    if (((Member.LoadBearing == true)
                                && Member.mType)) {
                        "*Extension*";
                        NewMember = new clsMember();
                        NewMember.rEdgePosition = Member.rEdgePosition;
                        NewMember.Length = Member.Length;
                        NewMember.CL = Member.CL;
                        NewMember.tEdgeHeight = Member.tEdgeHeight;
                        NewMember.bEdgeHeight = Member.bEdgeHeight;
                        NewMember.mType = "e1 Extension Column";
                        NewMember.Size = Member.Size;
                        NewMember.Width = Member.Width;
                        NewMember.Placement = "e1 Extension Column";
                        b.e1Columns.Add;
                        NewMember;
                    }

                }

                // End If
            }

            if (e1Overhang) {
                // get bay length to determine rafter size
                // if e1 Extension, check Extension bay length
                if ((b.e1Extension > (25 * 12))) {
                    Size = "10"" Receiver Cee";
                    Width = 10;
                }
                else if ((b.e1Extension > 0)) {
                    Size = "8"" Receiver Cee";
                    Width = 8;
                }
                else if ((EstSht.Range("Bay1_Length").Value > 25)) {
                    Size = "10"" Receiver Cee";
                    Width = 10;
                }
                else {
                    Size = "8"" Receiver Cee";
                    Width = 8;
                }

                if ((b.rShape == "Single Slope")) {
                    // set max / min temporary values
                    rEdgeMax = (0 + 20);
                    lEdgeStart = ((b.bWidth * 12)
                                - 20);
                    // get starting and maximum values from e1 Columns
                    for (Member in b.e1Columns) {
                        if (((Member.CL < rEdgeMax)
                                    && !Member.mType)) {
                            "*Extension*";
                            rEdgeMax = Member.lEdgePosition;
                            tEdgeMax = b.DistanceToRoof("e1", Member.lEdgePosition);
                        }
                        else if (((Member.CL > lEdgeStart)
                                    && !Member.mType)) {
                            "*Extension*";
                            lEdgeStart = Member.rEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e1", Member.rEdgePosition);
                        }

                    }

                    for (Member in b.e1Rafters) {
                        if (((Member.rEdgePosition > 0)
                                    && ((Member.RafterLeftEdge
                                    < (b.bWidth * 12))
                                    && !Member.mType))) {
                            ("*Extension*"
                                        & !Member.mType);
                            "*Overhang*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                }
                else {
                    // Gable Roof
                    // first slope of roof from left to right
                    // get starting and maximum values from e1 Columns
                    for (Member in b.e1Columns) {
                        if (((Member.CL
                                    > ((b.bWidth * 12)
                                    - 16))
                                    && ((Member.CL
                                    < (b.bWidth * 12))
                                    && !Member.mType))) {
                            ("*Extension*"
                                        & !Member.mType);
                            "*Overhang*";
                            lEdgeStart = Member.rEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e1", Member.rEdgePosition);
                        }

                    }

                    rEdgeMax = (b.bWidth * (12 / 2));
                    tEdgeMax = b.DistanceToRoof("e1", (b.bWidth * (12 / 2)));
                    for (Member in b.e1Rafters) {
                        if (((Member.rEdgePosition
                                    >= (b.bWidth * (12 / 2)))
                                    && ((Member.lEdgePosition
                                    < (b.bWidth * 12))
                                    && !Member.mType))) {
                            "*Extension*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                }

                // Create Overhang Rafters
                for (i = 1; (i <= RafterNum); i++) {
                    Member = new clsMember();
                    Member.Size = Size;
                    Member.Width = Width;
                    Member.Placement = "e1 Overhang Rafter";
                    Member.bEdgeHeight = bEdgeStart;
                    Member.RafterLeftEdge = lEdgeStart;
                    if ((i != RafterNum)) {
                        Member.Length = (30 * 12);
                        HorizontalDistance = (Member.Length / Sqr(((b.rPitch / 12) | (2 + 1))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Member.rEdgePosition = (lEdgeStart - HorizontalDistance);
                        Member.tEdgeHeight = b.DistanceToRoof("e1", (lEdgeStart - HorizontalDistance));
                    }
                    else {
                        Member.Length = (TotalSlopeLength - (30 * (12
                                    * (i - 1))));
                        Member.rEdgePosition = rEdgeMax;
                        Member.tEdgeHeight = tEdgeMax;
                    }

                    bEdgeStart = Member.tEdgeHeight;
                    lEdgeStart = Member.rEdgePosition;
                    Angle = (Atn((b.rPitch / 12))
                                * (Member.Width / 2));
                    DistanceToLower = Sqr((Angle
                                    | ((2
                                    + (Member.Width / 2))
                                    | 2)));
                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                    Member.bEdgeHeight = (Member.bEdgeHeight - DistanceToLower);
                    Member.tEdgeHeight = (Member.tEdgeHeight - DistanceToLower);
                    b.e1Rafters.Add;
                    Member;
                }

                if ((b.rShape == "Gable")) {
                    //  do second slope of roof
                    TotalSlopeLength = 0;
                    // get starting and maximum values from e1 Columns
                    // get starting and maximum values from e1 Columns
                    for (Member in b.e1Columns) {
                        if (((Member.CL > 0)
                                    && ((Member.CL < 16)
                                    && !Member.mType))) {
                            "*Extension*";
                            rEdgeStart = Member.lEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e1", Member.lEdgePosition);
                        }

                    }

                    lEdgeMax = (b.bWidth * (12 / 2));
                    tEdgeMax = b.DistanceToRoof("e1", (b.bWidth * (12 / 2)));
                    for (Member in b.e1Rafters) {
                        if (((Member.lEdgePosition
                                    <= (b.bWidth * (12 / 2)))
                                    && ((Member.rEdgePosition > 0)
                                    && !Member.Placement))) {
                            ("*Overhang*"
                                        & !Member.mType);
                            ("*Extension*"
                                        & !Member.mType);
                            "*Overhang*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                    for (i = 1; (i <= RafterNum); i++) {
                        Member = new clsMember();
                        Member.Size = Size;
                        Member.Width = Width;
                        Member.Placement = "e1 Overhang Rafter";
                        Member.bEdgeHeight = bEdgeStart;
                        Member.rEdgePosition = rEdgeStart;
                        if ((i != RafterNum)) {
                            Member.Length = (30 * 12);
                            HorizontalDistance = (Member.Length / Sqr(((b.rPitch / 12) | (2 + 1))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            Member.RafterLeftEdge = (rEdgeStart + HorizontalDistance);
                            Member.tEdgeHeight = b.DistanceToRoof("e1", (rEdgeStart + HorizontalDistance));
                        }
                        else {
                            Member.Length = (TotalSlopeLength - (30 * (12
                                        * (i - 1))));
                            Member.RafterLeftEdge = lEdgeMax;
                            Member.tEdgeHeight = tEdgeMax;
                        }

                        bEdgeStart = Member.tEdgeHeight;
                        rEdgeStart = Member.RafterLeftEdge;
                        Angle = (Atn((b.rPitch / 12))
                                    * (Member.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (Member.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Member.bEdgeHeight = (Member.bEdgeHeight - DistanceToLower);
                        Member.tEdgeHeight = (Member.tEdgeHeight - DistanceToLower);
                        b.e1Rafters.Add;
                        Member;
                    }

                }

            }

            // '''''''' e3
            // if the endwall was non-expandable, the rafter needs to be changed back to a C Purlin
            if ((e3Overhang || e3Extension)) {
                for (i = 1; (i <= b.e3Rafters.Count); i++) {
                    Member = b.e3Rafters(i);
                    if (((b.WallStatus("e3") == "No")
                                && (((Member.rEdgePosition < 0)
                                && b.s4e3ExtensionIntersection)
                                || (((Member.RafterLeftEdge
                                > (b.bWidth * 12))
                                && b.s2e3ExtensionIntersection)
                                || ((Member.rEdgePosition >= 0)
                                && (Member.RafterLeftEdge
                                <= (b.bWidth * 12))))))) {
                        if ((Member.Size == "8"" Receiver Cee")) {
                            Member.Size = "8"" C Purlin";
                        }
                        else if ((Member.Size == "10"" Receiver Cee")) {
                            Member.Size = "10"" C Purlin";
                        }

                    }

                }

            }

            // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''e3
            if (e3Extension) {
                // If b.InteriorColumns.Count = 0 Then
                // if no interior columns have been made, create e3 extension columns and rafters
                EndwallExtensionColumnsGen(b, "e3");
                RafterGen(b, "e3 Extension");
                for (Member in b.e3ExtensionMembers) {
                    b.e3Columns.Add;
                    Member;
                }

                // End If
                // Else
                //     For i = 1 To b.intRafters.Count
                //         Set Member = b.intRafters(i)
                //         Set NewMember = New clsMember
                //         NewMember.rEdgePosition = (b.bWidth * 12) - Member.RafterLeftEdge
                //         NewMember.RafterLeftEdge = (b.bWidth * 12) - Member.rEdgePosition
                //         NewMember.Length = Member.Length
                //         NewMember.tEdgeHeight = Member.tEdgeHeight
                //         NewMember.bEdgeHeight = Member.bEdgeHeight
                //         NewMember.mType = "e3 Extension Rafter"
                //         NewMember.Size = Member.Size
                //         NewMember.Width = Member.Width
                //         NewMember.Placement = "e3 Extension Rafter"
                //         b.e3Rafters.Add NewMember
                //     Next i
                for (i = 1; (i <= b.InteriorColumns.Count); i++) {
                    Member = b.InteriorColumns(i);
                    if (((Member.LoadBearing == true)
                                && Member.mType)) {
                        "*Extension*";
                        NewMember = new clsMember();
                        NewMember.CL = ((b.bWidth * 12)
                                    - Member.CL);
                        NewMember.rEdgePosition = (Member.CL
                                    - (Member.Width / 2));
                        NewMember.Length = Member.Length;
                        NewMember.tEdgeHeight = Member.tEdgeHeight;
                        NewMember.bEdgeHeight = Member.bEdgeHeight;
                        NewMember.mType = "e3 Extension Column";
                        NewMember.Size = Member.Size;
                        NewMember.Width = Member.Width;
                        NewMember.Placement = "e3 Extension Column";
                        b.e3Columns.Add;
                        NewMember;
                    }

                }

                // End If
            }

            // e3 Overhang
            // Add rafters and columns
            if (e3Overhang) {
                TotalSlopeLength = 0;
                // get bay length to determine rafter size
                // if e1 Extension, check Extension bay length
                if ((b.e3Extension > (25 * 12))) {
                    Size = "10"" Receiver Cee";
                    Width = 10;
                }
                else if ((b.e3Extension > 0)) {
                    Size = "8"" Receiver Cee";
                    Width = 8;
                }
                else if ((EstSht.Range("Bay1_Length").offset(EstSht.Range("BayNum").Value, 0).Value > 25)) {
                    Size = "10"" Receiver Cee";
                    Width = 10;
                }
                else {
                    Size = "8"" Receiver Cee";
                    Width = 8;
                }

                if ((b.rShape == "Single Slope")) {
                    // get starting and maximum values from e1 Columns
                    for (Member in b.e3Columns) {
                        if (((Member.CL > 0)
                                    && ((Member.CL < 16)
                                    && !Member.mType))) {
                            "*Extension*";
                            rEdgeStart = Member.lEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e3", Member.lEdgePosition);
                        }

                        if (((Member.CL
                                    < (b.bWidth * 12))
                                    && ((Member.CL
                                    > ((b.bWidth * 12)
                                    - 16))
                                    && !Member.mType))) {
                            "*Extension*";
                            lEdgeMax = Member.rEdgePosition;
                            tEdgeMax = b.DistanceToRoof("e3", Member.rEdgePosition);
                        }

                    }

                    for (Member in b.e3Rafters) {
                        if (((Member.rEdgePosition > 0)
                                    && ((Member.lEdgePosition
                                    < (b.bWidth * 12))
                                    && !Member.mType))) {
                            ("*Extension*"
                                        & !Member.mType);
                            "*Overhang*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                }
                else {
                    // first slope of roof from right to left
                    // get starting and maximum values from e3 Columns
                    for (Member in b.e3Columns) {
                        if (((Member.CL > 0)
                                    && ((Member.CL < 16)
                                    && !Member.mType))) {
                            "*Extension*";
                            rEdgeStart = Member.lEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e3", Member.lEdgePosition);
                        }

                    }

                    lEdgeMax = (b.bWidth * (12 / 2));
                    tEdgeMax = b.DistanceToRoof("e3", (b.bWidth * (12 / 2)));
                    for (Member in b.e3Rafters) {
                        if (((Member.rEdgePosition > 0)
                                    && ((Member.RafterLeftEdge
                                    <= (b.bWidth * (12 / 2)))
                                    && !Member.mType))) {
                            ("*Extension*"
                                        & !Member.mType);
                            "*Overhang*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                }

                for (i = 1; (i <= RafterNum); i++) {
                    Member = new clsMember();
                    Member.Size = Size;
                    Member.Width = Width;
                    Member.Placement = "e3 Overhang Rafter";
                    Member.bEdgeHeight = bEdgeStart;
                    Member.rEdgePosition = rEdgeStart;
                    if ((i != RafterNum)) {
                        Member.Length = (30 * 12);
                        HorizontalDistance = (Member.Length / Sqr(((b.rPitch / 12) | (2 + 1))));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Member.RafterLeftEdge = (rEdgeStart + HorizontalDistance);
                        Member.tEdgeHeight = b.DistanceToRoof("e3", (rEdgeStart + HorizontalDistance));
                    }
                    else {
                        Member.Length = (TotalSlopeLength - (30 * (12
                                    * (i - 1))));
                        Member.RafterLeftEdge = lEdgeMax;
                        Member.tEdgeHeight = tEdgeMax;
                    }

                    bEdgeStart = Member.tEdgeHeight;
                    rEdgeStart = Member.RafterLeftEdge;
                    Angle = (Atn((b.rPitch / 12))
                                * (Member.Width / 2));
                    DistanceToLower = Sqr((Angle
                                    | ((2
                                    + (Member.Width / 2))
                                    | 2)));
                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                    Member.bEdgeHeight = (Member.bEdgeHeight - DistanceToLower);
                    Member.tEdgeHeight = (Member.tEdgeHeight - DistanceToLower);
                    b.e3Rafters.Add;
                    Member;
                }

                if ((b.rShape == "Gable")) {
                    //  do second slope of roof from left to right
                    // get starting and maximum values from e1 Columns
                    TotalSlopeLength = 0;
                    for (Member in b.e3Columns) {
                        if (((Member.CL
                                    > ((b.bWidth * 12)
                                    - 16))
                                    && ((Member.CL
                                    < (b.bWidth * 12))
                                    && !Member.mType))) {
                            "*Extension*";
                            lEdgeStart = Member.rEdgePosition;
                            bEdgeStart = b.DistanceToRoof("e3", Member.rEdgePosition);
                        }

                    }

                    rEdgeMax = (b.bWidth * (12 / 2));
                    tEdgeMax = b.DistanceToRoof("e3", (b.bWidth * (12 / 2)));
                    for (Member in b.e3Rafters) {
                        if (((Member.RafterLeftEdge
                                    < (b.bWidth * 12))
                                    && ((Member.rEdgePosition
                                    >= (b.bWidth * (12 / 2)))
                                    && !Member.Placement))) {
                            ("*Overhang*"
                                        & !Member.mType);
                            "*Extension*";
                            TotalSlopeLength = (TotalSlopeLength + Member.Length);
                        }

                    }

                    RafterNum = Application.WorksheetFunction.RoundUp((TotalSlopeLength / (30 * 12)), 0);
                    for (i = 1; (i <= RafterNum); i++) {
                        Member = new clsMember();
                        Member.Size = Size;
                        Member.Width = Width;
                        Member.Placement = "e3 Overhang Rafter";
                        Member.bEdgeHeight = bEdgeStart;
                        Member.RafterLeftEdge = lEdgeStart;
                        if ((i != RafterNum)) {
                            Member.Length = (30 * 12);
                            HorizontalDistance = (Member.Length / Sqr(((b.rPitch / 12) | (2 + 1))));
                            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                            Member.rEdgePosition = (lEdgeStart - HorizontalDistance);
                            Member.tEdgeHeight = b.DistanceToRoof("e3", (lEdgeStart - HorizontalDistance));
                        }
                        else {
                            Member.Length = (TotalSlopeLength - (30 * (12
                                        * (i - 1))));
                            Member.rEdgePosition = rEdgeMax;
                            Member.tEdgeHeight = tEdgeMax;
                        }

                        bEdgeStart = Member.tEdgeHeight;
                        lEdgeStart = Member.rEdgePosition;
                        Angle = (Atn((b.rPitch / 12))
                                    * (Member.Width / 2));
                        DistanceToLower = Sqr((Angle
                                        | ((2
                                        + (Member.Width / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        Member.bEdgeHeight = (Member.bEdgeHeight - DistanceToLower);
                        Member.tEdgeHeight = (Member.tEdgeHeight - DistanceToLower);
                        b.e3Rafters.Add;
                        Member;
                    }

                }

            }

        }

        AdjustEndwallColumns((<clsBuilding>(b)), (<string>(eWall)));
        let ColumnCollection: Collection;
        let RafterCollection: Collection;
        let Column: clsMember;
        let Rafter: clsMember;
        let RafterWidth: number;
        let tEdgeDifference: number;
        let Angle: number;
        let DistanceToLower: number;
        let WedgeDistance: number;
        let FirstColWidth: number;
        let LastColWidth: number;
        switch (eWall) {
            case "e1":
                RafterCollection = b.e1Rafters;
                ColumnCollection = b.e1Columns;
                FirstColWidth = b.s4ColumnWidth;
                LastColWidth = b.s2ColumnWidth;
                break;
            case "e3":
                RafterCollection = b.e3Rafters;
                ColumnCollection = b.e3Columns;
                FirstColWidth = b.s2ColumnWidth;
                LastColWidth = b.s4ColumnWidth;
                break;
            case "Int":
                RafterCollection = b.intRafters;
                ColumnCollection = b.InteriorColumns;
                FirstColWidth = b.s4ColumnWidth;
                LastColWidth = b.s2ColumnWidth;
                break;
        }

        for (Column in ColumnCollection) {
            if (((Column.lEdgePosition
                        != (b.bWidth * 12))
                        && ((Column.rEdgePosition != 0)
                        && (Column.LoadBearing == true)))) {
                // extend each load bearing column to account for angle cut
                Column.tEdgeHeight = (Column.tEdgeHeight
                            + ((Column.Width / 2)
                            * (b.rPitch / 12)));
                Column.Length = Column.tEdgeHeight;
            }
            else if ((Column.LoadBearing == false)) {
                // non load bearing columns need to be lowered to bottom of rafter
                for (Rafter in RafterCollection) {
                    if (((Rafter.rEdgePosition <= Column.CL)
                                && (Rafter.RafterLeftEdge >= Column.CL))) {
                        RafterWidth = Rafter.Width;
                    }

                }

                Angle = (Atn((b.rPitch / 12))
                            * (RafterWidth / 2));
                tEdgeDifference = Sqr((Angle
                                | ((2
                                + (RafterWidth / 2))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                Column.tEdgeHeight = ((Column.tEdgeHeight
                            - (tEdgeDifference * 2))
                            + ((Column.Width / 2)
                            * (b.rPitch / 12)));
                Column.Length = ((Column.Length
                            - (tEdgeDifference * 2))
                            + ((Column.Width / 2)
                            * (b.rPitch / 12)));
                if (((Column.LoadBearing == true)
                            && ((Column.CL
                            == (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable")))) {
                    Column.Placement = (Column.Placement + ("cut Vee for center column at "
                                + (Application.WorksheetFunction.Round(Angle, 2) + " degree angles, ")));
                }
                else {
                    Column.Placement = (Column.Placement + ("cut at "
                                + (Application.WorksheetFunction.Round(Angle, 2) + " degree angle, ")));
                }

            }
            else if ((Column.lEdgePosition
                        == (b.bWidth * 12))) {
                if (((b.rShape == "Single Slope")
                            && (eWall == "e3"))) {
                    // skip high side eave columns
                }
                else {
                    // corner columns on expandable endwalls need to be longer to account for angle cut
                    // non-expandable endwalls only need to be longer if distance is greater than 4"
                    WedgeDistance = (LastColWidth
                                * (b.rPitch / 12));
                    if ((eWall != "Int")) {
                        if ((b.ExpandableEndwall(eWall)
                                    || (WedgeDistance > 4))) {
                            Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                            Column.Length = (Column.Length + WedgeDistance);
                        }

                    }
                    else {
                        Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                        Column.Length = (Column.Length + WedgeDistance);
                    }

                }

            }
            else if ((Column.rEdgePosition == 0)) {
                if (((b.rShape == "Single Slope")
                            && ((eWall == "e1")
                            || (eWall == "Int")))) {
                    // skip high side eave columns
                }
                else {
                    // corner columns on expandable endwalls need to be longer to account for angle cut
                    // non-expandable endwalls only need to be longer if distance is greater than 4"
                    WedgeDistance = (FirstColWidth
                                * (b.rPitch / 12));
                    if ((eWall != "Int")) {
                        if ((b.ExpandableEndwall(eWall)
                                    || (WedgeDistance > 4))) {
                            Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                            Column.Length = (Column.Length + WedgeDistance);
                        }

                    }
                    else {
                        Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                        Column.Length = (Column.Length + WedgeDistance);
                    }

                }

            }

        }

    }

}

RemoveEndwallColumns(b: clsBuilding, eWall: string) {
    let FOCollection: Collection;
    let ColumnCollection: Collection;
    let NearestMemberRight: number;
    let NearestMemberLeft: number;
    let tempNearestMember: number;
    let Column: clsMember;
    let FO: clsFO;
    let Jamb: Object;
    let ColIndex: number;
    let tempColumn: clsMember;
    switch (eWall) {
        case "e1":
            FOCollection = b.e1FOs;
            ColumnCollection = b.e1Columns;
            break;
        case "e3":
            FOCollection = b.e3FOs;
            ColumnCollection = b.e3Columns;
            break;
    }

    for (ColIndex = 1; (ColIndex
                <= (ColumnCollection.Count - 1)); ColIndex++) {
        Column = ColumnCollection[ColIndex];
        if ((Column.LoadBearing == false)) {
            for (FO in FOCollection) {
                if (((FO.rEdgePosition < Column.CL)
                            && (FO.lEdgePosition > Column.CL))) {
                    Column.DeleteFlag = true;
                }

            }

            if ((Column.DeleteFlag == false)) {
                NearestMemberLeft = (b.bWidth * 12);
                NearestMemberRight = 0;
                for (tempColumn in ColumnCollection) {
                    if ((((tempColumn.CL > Column.CL)
                                && (tempColumn.CL < NearestMemberLeft))
                                && (tempColumn.DeleteFlag == false))) {
                        NearestMemberLeft = tempColumn.CL;
                    }

                    if ((((tempColumn.CL < Column.CL)
                                && (tempColumn.CL > NearestMemberRight))
                                && (tempColumn.DeleteFlag == false))) {
                        NearestMemberRight = tempColumn.CL;
                    }

                }

                for (FO in FOCollection) {
                    for (Jamb in FO.FOMaterials) {
                        if ((Jamb.clsType == "Member")) {
                            if (((Jamb.CL > Column.CL)
                                        && (Jamb.CL < NearestMemberLeft))) {
                                NearestMemberLeft = Jamb.CL;
                            }

                            if (((Jamb.CL < Column.CL)
                                        && (Jamb.CL > NearestMemberRight))) {
                                NearestMemberRight = Jamb.CL;
                            }

                        }

                    }

                }

                if ((Abs((NearestMemberLeft - NearestMemberRight)) < (30 * 12))) {
                    ColumnCollection[ColIndex].DeleteFlag = true;
                }
                else {
                    ColumnCollection[ColIndex].DeleteFlag = false;
                }

            }

        }

    }

    for (ColIndex = ColumnCollection.Count; (ColIndex <= 1); ColIndex = (ColIndex + -1)) {
        if ((ColumnCollection[ColIndex].DeleteFlag == true)) {
            ColumnCollection.Remove(ColIndex);
        }

    }

}
CutListOutput(Collection: Collection, Label: string) {
    let LastRow: number;
    let Member: clsMember;
    let SteelSht: Worksheet;
    let FullMemberSht: Worksheet;
    let FO: clsFO;
    let item: Object;
    let j: number;
    let UnitPrice: number;
    let UnitMeasure: string;
    let UnitValue: number;
    let PriceTbl: ListObject;
    // ''''''''''''''''''Full Member List Sheet
    FullMemberSht = ThisWorkbook.Sheets("Optimized Cut List");
    if ((FullMemberSht.Range("E4").Value == "")) {
        LastRow = 4;
    }
    else {
        LastRow = FullMemberSht.Range("E3").End(xlDown).offset(1, 0).Row;
    }

    j = 1;
    // Call DuplicateMaterialRemoval(Collection, "Steel")
    for (Member in Collection) {
        // With...
        // Formatting
        Member.Placement.Range(("C" + LastRow)).Value = ImperialMeasurementFormat(Member.Length);
        FullMemberSht.Range(("A" + LastRow)).Value = ImperialMeasurementFormat(Member.Length);
        xlLeft.Rows[LastRow].RowHeight = 30;
        xlRight.Range(("E" + LastRow)).HorizontalAlignment = 30;
        true.Range(("C" + LastRow)).HorizontalAlignment = 30;
        XlLineStyle.xlContinuous.Range(("A" + LastRow), ("E" + LastRow)).Font.Bold = 30;
        Range(("A" + LastRow), ("E" + LastRow)).Borders(xlEdgeBottom).LineStyle = 30;
        j = (j + 1);
        LastRow = (LastRow + 1);
        for (item in Member.ComponentMembers) {
            // With...
            if ((item.Placement == "")) {
                Range(("D" + LastRow)).Value = item.mType;
                FullMemberSht.Range(("B" + LastRow)).Value = item.Size;
            }
            else {
                Range(("D" + LastRow)).Value = item.Placement;
            }

            Range(("E" + LastRow)).Value = ImperialMeasurementFormat(item.Length);
            LastRow = (LastRow + 1);
        }

    }

}

// ' function returns string of the nearest available Member Size
NearestMemberSize() {
    (<number>(Direction));
    (<string>(MemberType));
    (<boolean>(NumericOutput));
    // DESCRIPTION: Function returns the nearest value to a target
    // INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
    //  and an optional direction variable described below.
    // OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
    //  Set equal to -1 to find the closest value below your target
    //  set equal to 1 to find the closest value above your target
    let t: Object;
    let u: Object;
    let Members: Object;
    let Member: Object;
    let mSize: number;
    let NearestMemberSizeString: string;
    let UniqueMemberType: string;
    //
    if ((MemberType == "C Purlin")) {
        Members = Array((20 * 12), (25 * 12), (30 * 12));
    }
    else if ((MemberType == "TS")) {
        Members = Array((20 * 12), (30 * 12), (40 * 12));
    }
    else if ((MemberType == "W Beam")) {
        Members = Array((20 * 12), (25 * 12), (30 * 12), (35 * 12), (40 * 12), (45 * 12), (50 * 12), (60 * 12));
    }

    t = 1.79769313486231E+308;
    // initialize
    for (Member in Members) {
        if (IsNumeric(Member)) {
            u = Abs((Member - Length));
            if (((Direction > 0)
                        && (Member >= Length))) {
                // only report if closer number is greater than the target
                if ((u < t)) {
                    t = u;
                    mSize = Member;
                }

            }
            else if (((Direction < 0)
                        && (Member <= Length))) {
                // only report if closer number is less than the target
                if ((u < t)) {
                    t = u;
                    mSize = Member;
                }

            }
            else if ((Direction == 0)) {
                if ((u < t)) {
                    t = u;
                    mSize = Member;
                }

            }

        }

    }

    // return available Member name
    NearestMemberSizeString = MaterialsListGen.ImperialMeasurementFormat(mSize);
    // output
    if ((NumericOutput == false)) {
        NearestMemberSize = NearestMemberSizeString;
    }
    else if ((NumericOutput == true)) {
        NearestMemberSize = mSize;
    }

}

SteelPriceOutput(Collection: Collection, Label: string, FOMode: boolean) {
    let LastRow: number;
    // Warning!!! Optional parameters not supported
    let Member: clsMember;
    let SteelSht: Worksheet;
    let FullMemberSht: Worksheet;
    let FO: clsFO;
    let item: Object;
    let j: number;
    let UnitPrice: number;
    let UnitMeasure: string;
    let UnitValue: number;
    let PriceTbl: ListObject;
    // ''''''''''''''''''''''''''Steel Material OUtput Sheet
    SteelSht = ThisWorkbook.Sheets("Structural Steel Price List");
    PriceTbl = ThisWorkbook.Worksheets("Master Price List").ListObjects("SteelPriceListTbl");
    if ((SteelSht.Range("A4").Value == "")) {
        LastRow = 4;
    }
    else {
        LastRow = SteelSht.Range("A3").End(xlDown).offset(1, 0).Row;
    }

    if ((FOMode == false)) {
        DuplicateMaterialRemoval(Collection, "Steel");
        for (Member in Collection) {
            // lookup price information in Price Table
            // check for errors
            if (((IsError(Application.VLookup(Member.Size, PriceTbl.Range, 2, false)) == true)
                        && !Member.Size)) {
                ("W*"
                            & !Member.Size);
                "*eave strut*";
                UnitPrice = 0;
                UnitMeasure = "Unknown";
                UnitValue = 0;
            }
            else {
                // successful lookup
                if (Member.Size) {
                    "W*";
                    UnitPrice = Application.WorksheetFunction.VLookup("W--x--", PriceTbl.Range, 2, false);
                    UnitMeasure = "per lb";
                    UnitValue = ((Member.Length / 12)
                                * (Member.Size.Substring((Member.Size.Length - 2)) * UnitPrice));
                }
                else if (Member.Size) {
                    "*eave strut*";
                    UnitPrice = Application.WorksheetFunction.VLookup("Eave Struts", PriceTbl.Range, 2, false);
                    UnitMeasure = "per ft";
                    UnitValue = ((Member.Length / 12)
                                * UnitPrice);
                }
                else {
                    UnitPrice = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 2, false);
                    UnitMeasure = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 3, false);
                    if ((UnitMeasure == "per ft")) {
                        UnitValue = (UnitPrice
                                    * (Member.Length / 12));
                    }
                    else if ((UnitMeasure == "per lb")) {
                        UnitValue = (UnitPrice
                                    * ((Member.Length / 12)
                                    * Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 4, false)));
                    }

                }

            }

            // With...
            if (((UnitPrice == 0)
                        || (UnitValue == 0))) {
                "Unknown".Range(("H" + LastRow)).Value = "Item not found";
                Member.Size.Range(("C" + LastRow)).Value = ImperialMeasurementFormat(Member.Length);
                Member.Qty.Range(("B" + LastRow)).Value = ImperialMeasurementFormat(Member.Length);
                SteelSht.Range(("A" + LastRow)).Value = ImperialMeasurementFormat(Member.Length);
                "Unknown".Range(("G" + LastRow)).Value = "Item not found";
                "Unknown".Range(("F" + LastRow)).Value = "Item not found";
                Range(("E" + LastRow)).Value = "Item not found";
            }
            else {
                UnitValue.Range(("H" + LastRow)).Value = (UnitValue * Member.Qty);
                UnitMeasure.Range(("G" + LastRow)).Value = (UnitValue * Member.Qty);
                UnitPrice.Range(("F" + LastRow)).Value = (UnitValue * Member.Qty);
                Range(("E" + LastRow)).Value = (UnitValue * Member.Qty);
            }

            LastRow = (LastRow + 1);
        }

    }
    else {
        for (FO in Collection) {
            Label = (FO.FOType + (" " + FO.Wall));
            DuplicateMaterialRemoval(FO.FOMaterials, "Steel");
            for (item in FO.FOMaterials) {
                if ((item.clsType == "Member")) {
                    Member = item;
                    // check for errors
                    if ((IsError(Application.VLookup(Member.Size, PriceTbl.Range, 2, false)) == true)) {
                        UnitPrice = "Unknown";
                        UnitMeasure = "Unknown";
                        UnitValue = "Item Not Found";
                    }
                    else {
                        // successful lookup
                        if (Member.Size) {
                            "W*";
                            UnitPrice = Application.WorksheetFunction.VLookup("W--x--", PriceTbl.Range, 3, false);
                            UnitMeasure = "per lb";
                            UnitValue = ((Member.Length / 12)
                                        * (Member.Size.Substring((Member.Size.Length - 2)) * UnitPrice));
                        }
                        else {
                            UnitPrice = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 2, false);
                            UnitMeasure = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 3, false);
                            if ((UnitMeasure == "per ft")) {
                                UnitValue = (UnitPrice
                                            * (Member.Length / 12));
                            }
                            else if ((UnitMeasure == "per lb")) {
                                UnitValue = (UnitPrice
                                            * (Member.Length / 12));
                            }

                        }

                    }

                    // With...
                    UnitMeasure.Range(("G" + LastRow)).Value = (UnitValue * Member.Qty);
                    UnitPrice.Range(("F" + LastRow)).Value = (UnitValue * Member.Qty);
                    ImperialMeasurementFormat(Member.Length).Range(("E" + LastRow)).Value = (UnitValue * Member.Qty);
                    Label.Range(("D" + LastRow)).Value = (UnitValue * Member.Qty);
                    Member.Size.Range(("C" + LastRow)).Value = (UnitValue * Member.Qty);
                    Member.Qty.Range(("B" + LastRow)).Value = (UnitValue * Member.Qty);
                    SteelSht.Range(("A" + LastRow)).Value = (UnitValue * Member.Qty);
                    LastRow = (LastRow + 1);
                }

            }

        }

    }

}

RoofPurlinGen(b: clsBuilding) {
    let RafterCollection: Collection;
    let InteriorColumnCollection: Collection;
    let s2EaveStrutCollection: Collection;
    let s4EaveStrutCollection: Collection;
    let RoofPurlinCollection: Collection;
    let Purlins: number[];
    let Rafter: clsMember;
    let IntColumn: clsMember;
    let EaveStrut: clsMember;
    let Purlin: clsMember;
    let RafterNum: number;
    let BayNum: number;
    let i: number;
    let j: number;
    let StartPos: number;
    let MaxPos: number;
    let BayLength: number;
    let Overhang: boolean;
    RoofPurlinCollection = b.RoofPurlins;
    InteriorColumnCollection = b.InteriorColumns;
    RafterCollection = new Collection();
    s2EaveStrutCollection = new Collection();
    s4EaveStrutCollection = new Collection();
    // combine rafter collections
    for (Rafter in b.e1Rafters) {
        RafterCollection.Add;
        Rafter;
    }

    for (Rafter in b.e3Rafters) {
        RafterCollection.Add;
        Rafter;
    }

    for (Rafter in b.intRafters) {
        RafterCollection.Add;
        Rafter;
    }

    // get Eave Struts into new collection
    for (EaveStrut in b.s2Girts) {
        if ((EaveStrut.mType == "Eave Strut")) {
            s2EaveStrutCollection.Add;
            EaveStrut;
        }

    }

    for (EaveStrut in b.s4Girts) {
        if ((EaveStrut.mType == "Eave Strut")) {
            s4EaveStrutCollection.Add;
            EaveStrut;
        }

    }

    if ((b.rShape == "Gable")) {
        // calculate 1 side first, then duplicate it.
        RafterNum = Application.WorksheetFunction.RoundUp(((b.RafterLength - 12)
                        / 60), 0);
        let Purlins: number[,];
        // Eave Strut (1) is already made w/ girts
    }
    else {
        RafterNum = Application.WorksheetFunction.RoundUp(((b.RafterLength / 60)
                        - 1), 0);
        let Purlins: number[,];
        // Eave Struts (2) are already made w/ girts
    }

    Purlins[0] = 0;
    for (i = 1; (i <= UBound(Purlins)); i++) {
        if ((i != UBound(Purlins))) {
            Purlins[i] = (60 + (60 * i));
        }
        else {
            Purlins[i] = (b.RafterLength - 12);
        }

    }

    BayNum = EstSht.Range("BayNum").Value;
    StartPos = 0;
    // extend roof through endwall overhangs and extension
    // eave overhangs and extensions will be handled separately, along with any intersections
    // Notes:
    // First, Set start pos either at 0 or (negative) e1 extension
    // set end pos either to b.width or e3 extension
    // Next, if overhang is present, adjust starting pos or ending pos
    // Create roof purlin
    // if eave strut, handle case
    // next purlin
    for (i = 0; (i <= UBound(Purlins)); i++) {
        // for each row of purlins
        for (j = 1; (j <= BayNum); j++) {
            // for each bay
            Overhang = false;
            BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
            // define bay length
            // ^^^ orientation is from the relative "perspective" of s2
            // adjust based on overhangs and extensions
            if ((j == 1)) {
                if ((b.e1Extension > 0)) {
                    // Create extra purlin for e1 Extension section
                    BayLength = (EstSht.Range("e1_GableExtension").Value * 12);
                    StartPos = ((EstSht.Range("e1_GableExtension").Value * 12)
                                * -1);
                    if ((b.e1Overhang > 0)) {
                        BayLength = (BayLength
                                    + (EstSht.Range("e1_GableOverhang").Value * 12));
                        StartPos = (StartPos
                                    - (EstSht.Range("e1_GableOverhang").Value * 12));
                    }

                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = StartPos;
                    Purlin.tEdgeHeight = Purlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if ((BayLength > (25 * 12))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                    }
                    else {
                        // eave strut at s2 e1 gable extension
                        Purlin.mType = "Eave Strut";
                        if ((b.rPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else if (b.e1GableExtensionSoffit) {
                            Purlin.Size = ("8"" "
                                        + (b.rPitch + ":12 double up eave strut"));
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.rPitch + ":12 single up eave strut"));
                        }

                        Purlin.tEdgeHeight = (b.bHeight * 12);
                        Purlin.bEdgeHeight = (b.bHeight * 12);
                    }

                    Purlin.Placement = ("endwall 1 extension roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + """ above eave strut"));
                    if ((i == 0)) {
                        b.s2Girts.Add;
                        Purlin;
                    }
                    else {
                        RoofPurlinCollection.Add;
                        Purlin;
                        b.WeldClips = (b.WeldClips + 2);
                    }

                    BayLength = 0;
                    StartPos = 0;
                }
                else if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
                    BayLength = (BayLength
                                + (EstSht.Range("e1_GableOverhang").Value * 12));
                    StartPos = ((EstSht.Range("e1_GableOverhang").Value * 12)
                                * -1);
                    Overhang = true;
                }

            }
            else if ((j == BayNum)) {
                if ((EstSht.Range("e3_GableExtension").Value > 0)) {
                    // create extra purlin for e3 Extension Section
                    BayLength = (EstSht.Range("e3_GableExtension").Value * 12);
                    if ((b.e1Overhang > 0)) {
                        BayLength = (BayLength
                                    + (EstSht.Range("e3_GableOverhang").Value * 12));
                    }

                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = (b.bLength * 12);
                    Purlin.tEdgeHeight = Purlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if ((BayLength > (25 * 12))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                    }
                    else {
                        // eave strut at s2 e3 gable extension
                        Purlin.mType = "Eave Strut";
                        if ((b.rPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else if (b.e3GableExtensionSoffit) {
                            Purlin.Size = ("8"" "
                                        + (b.rPitch + ":12 double up eave strut"));
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.rPitch + ":12 single up eave strut"));
                        }

                        Purlin.tEdgeHeight = (b.bHeight * 12);
                        Purlin.bEdgeHeight = (b.bHeight * 12);
                    }

                    Purlin.Placement = ("endwall 3 extension roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + """ above eave strut"));
                    if ((i == 0)) {
                        b.s2Girts.Add;
                        Purlin;
                    }
                    else {
                        RoofPurlinCollection.Add;
                        Purlin;
                        b.WeldClips = (b.WeldClips + 2);
                    }

                    BayLength = 0;
                }
                else if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
                    BayLength = (BayLength
                                + (EstSht.Range("e3_GableOverhang").Value * 12));
                    Overhang = true;
                }

            }

            if ((BayLength == 0)) {
                BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
            }

            Purlin = new clsMember();
            Purlin.mType = "Roof Purlin";
            Purlin.Length = BayLength;
            Purlin.rEdgePosition = StartPos;
            Purlin.tEdgeHeight = Purlins[i];
            if (((BayLength > (25 * 12))
                        && (Overhang == false))) {
                Purlin.Size = "10"" C Purlin";
            }
            else {
                Purlin.Size = "8"" C Purlin";
            }

            Purlin.Placement = ("sidewall 2 roof purlin, "
                        + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + (""" above eave strut, bay number " + j)));
            RoofPurlinCollection.Add;
            Purlin;
            b.WeldClips = (b.WeldClips + 2);
            StartPos = (StartPos + BayLength);
        }

    }

    StartPos = 0;
    if ((b.rShape == "Gable")) {
        // duplicate purlins going opposite direction
        for (i = 0; (i <= UBound(Purlins)); i++) {
            // for each row of purlins
            for (j = BayNum; (j <= 1); j = (j + -1)) {
                // for each bay, going from e3 to e1
                Overhang = false;
                BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
                // define bay length
                // ^^^ orientation is from the relative "perspective" of s4
                // adjust based on overhangs and extensions
                if ((j == BayNum)) {
                    if ((EstSht.Range("e3_GableExtension").Value > 0)) {
                        // Create extra purlin for e3 extension
                        BayLength = (EstSht.Range("e3_GableExtension").Value * 12);
                        StartPos = ((EstSht.Range("e3_GableExtension").Value * 12)
                                    * -1);
                        if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
                            BayLength = (BayLength
                                        + (EstSht.Range("e3_GableOverhang").Value * 12));
                            StartPos = (StartPos
                                        - (EstSht.Range("e3_GableOverhang").Value * 12));
                        }

                        Purlin = new clsMember();
                        Purlin.mType = "Roof Purlin";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = StartPos;
                        Purlin.tEdgeHeight = Purlins[i];
                        if ((i != 0)) {
                            // normal purlins
                            if ((BayLength > (25 * 12))) {
                                Purlin.Size = "10"" C Purlin";
                            }
                            else {
                                Purlin.Size = "8"" C Purlin";
                            }

                        }
                        else {
                            // eave strut at s4 e3 gable extension
                            Purlin.mType = "Eave Strut";
                            if ((b.rPitch == 1)) {
                                Purlin.Size = "8"" C Purlin";
                            }
                            else if (b.e3GableExtensionSoffit) {
                                Purlin.Size = ("8"" "
                                            + (b.rPitch + ":12 double up eave strut"));
                            }
                            else {
                                Purlin.Size = ("8"" "
                                            + (b.rPitch + ":12 single up eave strut"));
                            }

                            Purlin.tEdgeHeight = (b.bHeight * 12);
                            Purlin.bEdgeHeight = (b.bHeight * 12);
                        }

                        Purlin.Placement = ("endwall 3 extension roof purlin, "
                                    + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + """ above eave strut"));
                        if ((i == 0)) {
                            b.s4Girts.Add;
                            Purlin;
                        }
                        else {
                            RoofPurlinCollection.Add;
                            Purlin;
                            b.WeldClips = (b.WeldClips + 2);
                        }

                        BayLength = 0;
                        StartPos = 0;
                    }
                    else if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
                        BayLength = (BayLength
                                    + (EstSht.Range("e3_GableOverhang").Value * 12));
                        StartPos = ((EstSht.Range("e3_GableOverhang").Value * 12)
                                    * -1);
                        Overhang = true;
                    }

                }
                else if ((j == 1)) {
                    if ((EstSht.Range("e1_GableExtension").Value > 0)) {
                        // Create extra purlin for e1 Extension
                        BayLength = (EstSht.Range("e1_GableExtension").Value * 12);
                        if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
                            BayLength = (BayLength
                                        + (EstSht.Range("e1_GableOverhang").Value * 12));
                        }

                        Purlin = new clsMember();
                        Purlin.mType = "Roof Purlin";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = (b.bLength * 12);
                        Purlin.tEdgeHeight = Purlins[i];
                        if ((i != 0)) {
                            // normal purlins
                            if ((BayLength > (25 * 12))) {
                                Purlin.Size = "10"" C Purlin";
                            }
                            else {
                                Purlin.Size = "8"" C Purlin";
                            }

                        }
                        else {
                            // eave strut at s4 e1 gable extension
                            Purlin.mType = "Eave Strut";
                            if ((b.rPitch == 1)) {
                                Purlin.Size = "8"" C Purlin";
                            }
                            else if (b.e1GableExtensionSoffit) {
                                Purlin.Size = ("8"" "
                                            + (b.rPitch + ":12 double up eave strut"));
                            }
                            else {
                                Purlin.Size = ("8"" "
                                            + (b.rPitch + ":12 single up eave strut"));
                            }

                            Purlin.tEdgeHeight = (b.bHeight * 12);
                            Purlin.bEdgeHeight = (b.bHeight * 12);
                        }

                        Purlin.Placement = ("endwall 1 extension roof purlin, "
                                    + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + """ above eave strut"));
                        if ((i == 0)) {
                            b.s4Girts.Add;
                            Purlin;
                        }
                        else {
                            RoofPurlinCollection.Add;
                            Purlin;
                            b.WeldClips = (b.WeldClips + 2);
                        }

                        BayLength = 0;
                    }
                    else if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
                        BayLength = (BayLength
                                    + (EstSht.Range("e1_GableOverhang").Value * 12));
                        Overhang = true;
                    }

                }

                if ((BayLength == 0)) {
                    // start from wall line if overhang hasn't been defined above
                    BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
                }

                Purlin = new clsMember();
                Purlin.mType = "Roof Purlin";
                Purlin.Length = BayLength;
                Purlin.rEdgePosition = StartPos;
                Purlin.tEdgeHeight = Purlins[i];
                if (((BayLength > (25 * 12))
                            && (Overhang == false))) {
                    Purlin.Size = "10"" C Purlin";
                }
                else {
                    Purlin.Size = "8"" C Purlin";
                }

                Purlin.Placement = ("sidewall 2 roof purlin, "
                            + (Application.WorksheetFunction.RoundUp(Purlins[i], 2) + (""" above eave strut, bay number " + j)));
                RoofPurlinCollection.Add;
                Purlin;
                b.WeldClips = (b.WeldClips + 2);
                StartPos = (StartPos + BayLength);
            }

        }

    }

    // ''''''''''''''''''''''''''''''''''''''''''Special Case:
    // Single slope building w/ e1 or e3 Gable Extension needs Eave struts for Extended section
    if ((b.rShape == "Single Slope")) {
        if ((b.e1Extension > 0)) {
            // '''''''''''''''''''''''''s4 high side eave gable e1 extension strut
            Purlin = new clsMember();
            Purlin.mType = "Eave Strut";
            Purlin.Length = (EstSht.Range("e1_GableExtension").Value * 12);
            if ((b.e1Overhang > 0)) {
                Purlin.Length = (Purlin.Length
                            + (EstSht.Range("e1_GableOverhang").Value * 12));
            }

            Purlin.rEdgePosition = (b.bLength * 12);
            Purlin.tEdgeHeight = ((b.bHeight * 12)
                        + (b.bWidth * (12
                        * (b.rPitch / 12))));
            Purlin.bEdgeHeight = ((b.bHeight * 12)
                        + (b.bWidth * (12
                        * (b.rPitch / 12))));
            if ((b.rPitch == 1)) {
                Purlin.Size = "8"" C Purlin";
            }
            else if (b.e1GableExtensionSoffit) {
                Purlin.Size = ("8"" "
                            + (b.rPitch + ":12 double down eave strut"));
            }
            else {
                Purlin.Size = ("8"" "
                            + (b.rPitch + ":12 single down eave strut"));
            }

            Purlin.Placement = ("sidewall 4 e1 gable extension, "
                        + (Application.WorksheetFunction.RoundUp(Purlin.tEdgeHeight, 2) + """ above eave strut "));
            Debug.Print;
            "single slope s4 high side eave e1 gable extension strut created";
            Debug.Print;
            (EaveStrutCount + 1);
            // RoofPurlinCollection.Add Purlin
            b.s4Girts.Add;
            Purlin;
        }

        if ((b.e3Extension > 0)) {
            // '''''''''''''''''''''''''''''''''''''''''s4 high side eave gable e3 extension strut
            Purlin = new clsMember();
            Purlin.mType = "Eave Strut";
            Purlin.Length = (EstSht.Range("e3_GableExtension").Value * 12);
            if ((b.e3Overhang > 0)) {
                Purlin.Length = (Purlin.Length
                            + (EstSht.Range("e3_GableOverhang").Value * 12));
            }

            Purlin.rEdgePosition = (Purlin.Length * -1);
            Purlin.tEdgeHeight = ((b.bHeight * 12)
                        + (b.bWidth * (12
                        * (b.rPitch / 12))));
            Purlin.bEdgeHeight = ((b.bHeight * 12)
                        + (b.bWidth * (12
                        * (b.rPitch / 12))));
            if ((b.rPitch == 1)) {
                Purlin.Size = "8"" C Purlin";
            }
            else if (b.e3GableExtensionSoffit) {
                Purlin.Size = ("8"" "
                            + (b.rPitch + ":12 double down eave strut"));
            }
            else {
                Purlin.Size = ("8"" "
                            + (b.rPitch + ":12 single down eave strut"));
            }

            Purlin.Placement = ("sidewall 2 e3 extension, "
                        + (Application.WorksheetFunction.RoundUp(Purlin.tEdgeHeight, 2) + """ above eave strut "));
            Debug.Print;
            "single slope s4 high side eave e4 gable extension strut created";
            Debug.Print;
            (EaveStrutCount + 1);
            // RoofPurlinCollection.Add Purlin
            b.s4Girts.Add;
            Purlin;
        }

    }

    let s2ExtensionPurlinNum: number;
    let s2ExtensionPurlins: number[];
    let s2ExtnesionLength: number;
    let s4ExtensionPurlinNum: number;
    let s4ExtensionPurlins: number[];
    let s4ExtnesionLength: number;
    // '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Eave Extensions
    // '''''''s2 extension
    if ((EstSht.Range("s2_EaveExtension").Value > 0)) {
        s2ExtensionPurlinNum = Application.WorksheetFunction.RoundUp(((b.s2ExtensionRafterLength / 60)
                        - 1), 0);
        let s2ExtensionPurlins: number[,];
        // Eave Strut (1) is already made
        // s2ExtensionLength = b.s2EaveExtensionBuildingLength
        s2ExtensionPurlins[0] = 0;
        // extension eave strut
        for (i = 1; (i <= UBound(s2ExtensionPurlins)); i++) {
            if ((i != UBound(s2ExtensionPurlins))) {
                s2ExtensionPurlins[i] = (60 + (60 * i));
            }
            else {
                s2ExtensionPurlins[i] = (b.s2ExtensionRafterLength - 12);
            }

        }

        for (i = 0; (i <= UBound(s2ExtensionPurlins)); i++) {
            StartPos = 0;
            MaxPos = (b.bLength * 12);
            for (j = 1; (j <= BayNum); j++) {
                BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
                //             If EstSht.Range("e1_GableOverhang").Value > 0 And j = 1 Then
                //                 BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
                //                 StartPos = -EstSht.Range("e1_GableOverhang").Value
                //             ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = BayNum Then
                //                 BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
                //             End If
                if ((i != 0)) {
                    // normal purlins
                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = StartPos;
                    Purlin.tEdgeHeight = s2ExtensionPurlins[i];
                    if (((BayLength > (25 * 12))
                                && (Overhang == false))) {
                        Purlin.Size = "10"" C Purlin";
                    }
                    else {
                        Purlin.Size = "8"" C Purlin";
                    }

                    Purlin.Placement = ("sidewall 2 eave extension roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(s2ExtensionPurlins[i], 2) + (""" above eave strut, bay number " + j)));
                    RoofPurlinCollection.Add;
                    Purlin;
                    b.WeldClips = (b.WeldClips + 2);
                }
                else {
                    Purlin = new clsMember();
                    Purlin.mType = "Eave Strut";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = StartPos;
                    Purlin.tEdgeHeight = s2ExtensionPurlins[i];
                    if ((b.s2ExtensionPitch == 1)) {
                        Purlin.Size = "8"" C Purlin";
                    }
                    else {
                        Purlin.Size = ("8"" "
                                    + (b.s2ExtensionPitch + ": 12 double up eave strut"));
                    }

                    Purlin.Placement = ("sidewall 2 eave extension eave strut, "
                                + (Application.WorksheetFunction.RoundUp(s2ExtensionPurlins[i], 2) + (""" above eave strut, bay number " + j)));
                    RoofPurlinCollection.Add;
                    Purlin;
                }

                StartPos = (StartPos + BayLength);
            }

            if (((EstSht.Range("e1_GableExtension").Value > 0)
                        && (b.s2e1ExtensionIntersection == true))) {
                BayLength = (EstSht.Range("e1_GableExtension").Value * 12);
                Purlin = new clsMember();
                Purlin.mType = "Roof Purlin";
                Purlin.Length = BayLength;
                Purlin.rEdgePosition = ((EstSht.Range("e1_GableExtension").Value * 12)
                            * -1);
                Purlin.tEdgeHeight = s2ExtensionPurlins[i];
                if ((i != 0)) {
                    // normal purlins
                    if (((BayLength > (25 * 12))
                                && (Overhang == false))) {
                        Purlin.Size = "10"" C Purlin";
                    }
                    else {
                        Purlin.Size = "8"" C Purlin";
                    }

                    b.WeldClips = (b.WeldClips + 2);
                }
                else {
                    Purlin.mType = "Eave Strut";
                    if ((b.s2ExtensionPitch == 1)) {
                        Purlin.Size = "8"" C Purlin";
                    }
                    else {
                        Purlin.Size = ("8"" "
                                    + (b.s2ExtensionPitch + ": 12 double up eave strut"));
                    }

                    Debug.Print;
                    "eave strut e1/s2 intersection created";
                    Debug.Print;
                    (EaveStrutCount + 1);
                }

                Purlin.Placement = ("sidewall 2 eave extension e1 intersection roof purlin, "
                            + (Application.WorksheetFunction.RoundUp(s2ExtensionPurlins[i], 2) + """ above extension eave strut"));
                RoofPurlinCollection.Add;
                Purlin;
            }

            if (((EstSht.Range("e3_GableExtension").Value > 0)
                        && (b.s2e3ExtensionIntersection == true))) {
                BayLength = (EstSht.Range("e3_GableExtension").Value * 12);
                Purlin = new clsMember();
                Purlin.mType = "Roof Purlin";
                Purlin.Length = BayLength;
                Purlin.rEdgePosition = (b.bLength * 12);
                Purlin.tEdgeHeight = s2ExtensionPurlins[i];
                if ((i != 0)) {
                    // normal purlins
                    if (((BayLength > (25 * 12))
                                && (Overhang == false))) {
                        Purlin.Size = "10"" C Purlin";
                    }
                    else {
                        Purlin.Size = "8"" C Purlin";
                    }

                    b.WeldClips = (b.WeldClips + 2);
                }
                else {
                    Purlin.mType = "Eave Strut";
                    if ((b.s2ExtensionPitch == 1)) {
                        Purlin.Size = "8"" C Purlin";
                    }
                    else {
                        Purlin.Size = ("8"" "
                                    + (b.s2ExtensionPitch + ": 12 double up eave strut"));
                    }

                    Debug.Print;
                    "eave strut e3/s2 intersection created";
                    Debug.Print;
                    (EaveStrutCount + 1);
                }

                Purlin.Placement = ("sidewall 2 eave extension e3 intersection roof purlin, "
                            + (Application.WorksheetFunction.RoundUp(s2ExtensionPurlins[i], 2) + """ above extension eave strut"));
                RoofPurlinCollection.Add;
                Purlin;
            }

        }

    }

    // '''''''s4 extension
    if ((b.rShape == "Gable")) {
        if ((EstSht.Range("s4_EaveExtension").Value > 0)) {
            s4ExtensionPurlinNum = Application.WorksheetFunction.RoundUp(((b.s4ExtensionRafterLength / 60)
                            - 1), 0);
            let s4ExtensionPurlins: number[,];
            // Eave Strut (1) is already made w/ girts
            // s4ExtensionLength = b.s4EaveExtensionBuildingLength
            s4ExtensionPurlins[0] = 0;
            // extension eave strut
            for (i = 1; (i <= UBound(s4ExtensionPurlins)); i++) {
                if ((i != UBound(s4ExtensionPurlins))) {
                    s4ExtensionPurlins[i] = (60 + (60 * i));
                }
                else {
                    s4ExtensionPurlins[i] = (b.s4ExtensionRafterLength - 12);
                }

            }

            for (i = 0; (i <= UBound(s4ExtensionPurlins)); i++) {
                StartPos = 0;
                MaxPos = b.bLength;
                for (j = BayNum; (j <= 1); j = (j + -1)) {
                    BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
                    //             If EstSht.Range("e1_GableOverhang").Value > 0 And j = BayNum Then
                    //                 BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
                    //                 StartPos = -EstSht.Range("s4_EaveOverhang").Value
                    //             ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = 1 Then
                    //                 BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
                    //             End If
                    if ((i != 0)) {
                        Purlin = new clsMember();
                        Purlin.mType = "Roof Purlin";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = StartPos;
                        Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        Purlin.Placement = ("sidewall 4 eave extension roof purlin, "
                                    + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + (""" above eave strut, bay number " + j)));
                        RoofPurlinCollection.Add;
                        Purlin;
                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin = new clsMember();
                        Purlin.mType = "Eave Strut";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = StartPos;
                        Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double up eave strut"));
                        }

                        Purlin.Placement = ("sidewall 2 eave extension eave strut, "
                                    + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + (""" above eave strut, bay number " + j)));
                        RoofPurlinCollection.Add;
                        Purlin;
                    }

                    StartPos = (StartPos + BayLength);
                }

                if (((EstSht.Range("e1_GableExtension").Value > 0)
                            && (b.s4e1ExtensionIntersection == true))) {
                    BayLength = (EstSht.Range("e1_GableExtension").Value * 12);
                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = (b.bLength * 12);
                    Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin.mType = "Eave Strut";
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double up eave strut"));
                        }

                        Debug.Print;
                        "eave strut e1/s4 intersection created";
                        Debug.Print;
                        (EaveStrutCount + 1);
                    }

                    Purlin.Placement = ("sidewall 4 eave extension e1 intersection roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + """ above extension eave strut"));
                    RoofPurlinCollection.Add;
                    Purlin;
                }

                if (((EstSht.Range("e3_GableExtension").Value > 0)
                            && (b.s4e3ExtensionIntersection == true))) {
                    BayLength = (EstSht.Range("e3_GableExtension").Value * 12);
                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = ((EstSht.Range("e3_GableExtension").Value * 12)
                                * -1);
                    Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin.mType = "Eave Strut";
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double up eave strut"));
                        }

                        Debug.Print;
                        "eave strut e1/s4 intersection created";
                        Debug.Print;
                        (EaveStrutCount + 1);
                    }

                    Purlin.Placement = ("sidewall 4 eave extension e3 intersection roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + """ above extension eave strut"));
                    RoofPurlinCollection.Add;
                    Purlin;
                }

            }

        }

    }
    else {
        // ''''''''''''''''''''''''''''''''''''''''''''''''''''s4 extension single slope
        if ((EstSht.Range("s4_EaveExtension").Value > 0)) {
            s4ExtensionPurlinNum = Application.WorksheetFunction.RoundUp(((b.s4ExtensionRafterLength / 60)
                            - 1), 0);
            let s4ExtensionPurlins: number[,];
            // Eave Strut (1) is already made w/ girts
            // s4ExtensionLength = b.s4EaveExtensionBuildingLength
            s4ExtensionPurlins[0] = b.s4ExtensionRafterLength;
            // extension eave strut
            for (i = 1; (i <= UBound(s4ExtensionPurlins)); i++) {
                if ((i != UBound(s4ExtensionPurlins))) {
                    s4ExtensionPurlins[i] = (b.s4ExtensionRafterLength - (60 + (60 * i)));
                }
                else {
                    s4ExtensionPurlins[i] = 12;
                }

            }

            for (i = 0; (i <= UBound(s4ExtensionPurlins)); i++) {
                StartPos = 0;
                MaxPos = b.bLength;
                for (j = BayNum; (j <= 1); j = (j + -1)) {
                    BayLength = (EstSht.Range("Bay1_Length").offset((j - 1), 0).Value * 12);
                    //             If EstSht.Range("e1_GableOverhang").Value > 0 And j = BayNum Then
                    //                 BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
                    //                 StartPos = -EstSht.Range("s4_EaveOverhang").Value
                    //             ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = 1 Then
                    //                 BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
                    //             End If
                    if ((i != 0)) {
                        Purlin = new clsMember();
                        Purlin.mType = "Roof Purlin";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = StartPos;
                        Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        Purlin.Placement = ("sidewall 4 eave extension roof purlin, "
                                    + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + (""" above high side eave, bay number " + j)));
                        RoofPurlinCollection.Add;
                        Purlin;
                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin = new clsMember();
                        Purlin.mType = "Eave Strut";
                        Purlin.Length = BayLength;
                        Purlin.rEdgePosition = StartPos;
                        Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double down eave strut"));
                        }

                        Purlin.Placement = ("sidewall 2 eave extension eave strut, "
                                    + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + (""" above high side eave, bay number " + j)));
                        RoofPurlinCollection.Add;
                        Purlin;
                    }

                    StartPos = (StartPos + BayLength);
                }

                if (((EstSht.Range("e1_GableExtension").Value > 0)
                            && (b.s4e1ExtensionIntersection == true))) {
                    BayLength = (EstSht.Range("e1_GableExtension").Value * 12);
                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = (b.bLength * 12);
                    Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin.mType = "Eave Strut";
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double down eave strut"));
                        }

                        Debug.Print;
                        "eave strut e1/s4 intersection created";
                        Debug.Print;
                        (EaveStrutCount + 1);
                    }

                    Purlin.Placement = ("sidewall 4 eave extension e1 intersection roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + """ above high side eave"));
                    RoofPurlinCollection.Add;
                    Purlin;
                }

                if (((EstSht.Range("e3_GableExtension").Value > 0)
                            && (b.s4e3ExtensionIntersection == true))) {
                    BayLength = (EstSht.Range("e3_GableExtension").Value * 12);
                    Purlin = new clsMember();
                    Purlin.mType = "Roof Purlin";
                    Purlin.Length = BayLength;
                    Purlin.rEdgePosition = ((EstSht.Range("e3_GableExtension").Value * 12)
                                * -1);
                    Purlin.tEdgeHeight = s4ExtensionPurlins[i];
                    if ((i != 0)) {
                        // normal purlins
                        if (((BayLength > (25 * 12))
                                    && (Overhang == false))) {
                            Purlin.Size = "10"" C Purlin";
                        }
                        else {
                            Purlin.Size = "8"" C Purlin";
                        }

                        b.WeldClips = (b.WeldClips + 2);
                    }
                    else {
                        Purlin.mType = "Eave Strut";
                        if ((b.s4ExtensionPitch == 1)) {
                            Purlin.Size = "8"" C Purlin";
                        }
                        else {
                            Purlin.Size = ("8"" "
                                        + (b.s4ExtensionPitch + ": 12 double down eave strut"));
                        }

                        Debug.Print;
                        "eave strut e3/s4 intersection created";
                        Debug.Print;
                        (EaveStrutCount + 1);
                    }

                    Purlin.Placement = ("sidewall 4 eave extension e3 intersection roof purlin, "
                                + (Application.WorksheetFunction.RoundUp(s4ExtensionPurlins[i], 2) + """ above high side eave"));
                    RoofPurlinCollection.Add;
                    Purlin;
                }

            }

        }

    }

}

// adjusts heights and lengths for FO Jambs depending on girts and rafters
AdjustFOMembers(b: clsBuilding, eWall: string) {
    let FOCollection: Collection;
    let GirtCollection: Collection;
    let RafterCollection: Collection;
    let FO: clsFO;
    let item: Object;
    let Jamb: clsMember;
    let Girt: clsMember;
    let tempNearestGirtAbove: number;
    let tempNearestGirtBelow: number;
    let i: number;
    let Rafter: clsMember;
    let RafterWidth: number;
    let DistanceToRafter: number;
    let tEdgeDifference: number;
    switch (eWall) {
        case "e1":
            FOCollection = b.e1FOs;
            GirtCollection = b.e1Girts;
            RafterCollection = b.e1Rafters;
            break;
        case "s2":
            FOCollection = b.s2FOs;
            GirtCollection = b.s2Girts;
            break;
        case "e3":
            FOCollection = b.e3FOs;
            GirtCollection = b.e3Girts;
            RafterCollection = b.e3Rafters;
            break;
        case "s4":
            FOCollection = b.s4FOs;
            GirtCollection = b.s4Girts;
            break;
    }

    for (FO in FOCollection) {
        for (item in FO.FOMaterials) {
            // Extend jambs to nearest horizontal girt so that windows and MiscFOs aren't floating
            if (((item.clsType == "Member")
                        && ((FO.FOType == "Window")
                        || (FO.FOType == "MiscFO")))) {
                Jamb = item;
                if (((Jamb.CL != 0)
                            && (Jamb.Length != b.DistanceToRoof(eWall, Jamb.CL)))) {
                    // 'horizontal jambs weren't given CenterLines
                    // if Jamb Lenght isn't already full height, check that it touches the nearest girt; adjust if necessary
                    if (((Jamb.tEdgeHeight == (2 + (7 * 12)))
                                && (FO.tEdgeHeight <= (2 + (7 * 12))))) {
                        // if jamb and FO are less than or equal to 7'2", don't extend above girt
                        // check that bottom edge goes to building slab
                        Jamb.bEdgeHeight = 0;
                        Jamb.Length = Jamb.tEdgeHeight;
                    }
                    else {
                        DistanceToRafter = b.DistanceToRoof(eWall, Jamb.CL);
                        tempNearestGirtAbove = DistanceToRafter;
                        tempNearestGirtBelow = 0;
                        for (Girt in GirtCollection) {
                            if (((Girt.tEdgeHeight > FO.tEdgeHeight)
                                        && ((Girt.tEdgeHeight - FO.tEdgeHeight)
                                        < (tempNearestGirtAbove - FO.tEdgeHeight)))) {
                                tempNearestGirtAbove = Girt.tEdgeHeight;
                            }

                            if (((Girt.bEdgeHeight < FO.bEdgeHeight)
                                        && ((FO.bEdgeHeight - Girt.bEdgeHeight)
                                        < (FO.bEdgeHeight - tempNearestGirtBelow)))) {
                                tempNearestGirtBelow = Girt.bEdgeHeight;
                            }

                        }

                        if (((tempNearestGirtAbove - tempNearestGirtBelow)
                                    <= ((30 * 12)
                                    + 4))) {
                            Jamb.tEdgeHeight = tempNearestGirtAbove;
                            Jamb.bEdgeHeight = tempNearestGirtBelow;
                            Jamb.Length = (tempNearestGirtAbove - tempNearestGirtBelow);
                        }

                    }

                }

            }

        }

    }

    // Reduce full height jambs or jambs that connect to rafters
    if (((eWall == "e1")
                || (eWall == "e3"))) {
        for (FO in FOCollection) {
            for (item in FO.FOMaterials) {
                if ((item.clsType == "Member")) {
                    Jamb = item;
                    if (((Jamb.CL != 0)
                                && (Jamb.tEdgeHeight == b.DistanceToRoof(eWall, Jamb.CL)))) {
                        // if jamb goes all the way to the ceiling and isn't load bearing, then it needs to be reduced in length to the centerline of the new rafter position
                        for (Rafter in RafterCollection) {
                            if (((Rafter.rEdgePosition < Jamb.CL)
                                        && (Rafter.RafterLeftEdge > Jamb.CL))) {
                                RafterWidth = Rafter.Width;
                            }

                        }

                        tEdgeDifference = Sqr((((RafterWidth / 2)
                                        * (b.rPitch / 12))
                                        | ((2
                                        + (RafterWidth / 2))
                                        | 2)));
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                        // + ((RafterWidth / 2) * (b.rPitch / 12))
                        // Jamb.tEdgeHeight = Jamb.tEdgeHeight - tEdgeDifference
                        Jamb.tEdgeHeight = ((Jamb.tEdgeHeight
                                    - (tEdgeDifference * 2))
                                    + ((Jamb.Width / 2)
                                    * (b.rPitch / 12)));
                        Jamb.Length = ((Jamb.Length
                                    - (tEdgeDifference * 2))
                                    + ((Jamb.Width / 2)
                                    * (b.rPitch / 12)));
                        Jamb.Placement = (Jamb.Placement + ("cut at "
                                    + (Application.WorksheetFunction.Round(Atn((b.rPitch / 12)), 2) + " degree angle, ")));
                    }

                }

            }

        }

    }

}
RafterGen(b: clsBuilding, eWall: string) {
    let ColumnCollection: Collection;
    let FOCollection: Collection;
    let RafterCollection: Collection;
    let item: Object;
    let Rafters: number[,];
    let ColIndex: number;
    let Column: clsMember;
    let Member: clsMember;
    let RafterMember: clsMember;
    let FO: clsFO;
    let StartPos: number;
    let EndPos: number;
    let MidPos: number;
    let NextStartPos: number;
    let MaxDistance: number;
    let tempNearestColumn: number;
    let i: number;
    let PrevLocation: number;
    let IntColumnMode: boolean;
    let RafterType: string;
    let RafterPlacement: string;
    let largestWidth: number;
    let largestSize: string;
    let SecondDimension: number;
    let tempSecondDimension: number;
    let DistanceToLower: number;
    let DistanceToLengthen: number;
    let AngleCut: boolean;
    let Angle: number;
    let FirstColWidth: number;
    let LastColWidth: number;
    switch (eWall) {
        case "e1":
            ColumnCollection = b.e1Columns;
            FOCollection = b.e1FOs;
            RafterCollection = b.e1Rafters;
            IntColumnMode = false;
            if (b.ExpandableEndwall(eWall)) {
                RafterType = "W-Beam";
            }
            else if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
                RafterType = "8"" C Purlin";
            }
            else if ((EstSht.Range("Bay1_Length").Value > 25)) {
                RafterType = SteelLookupSht.Range("NonExpandableEndwallRaftersWithLargeBay").Value;
            }
            else {
                RafterType = SteelLookupSht.Range("NonExpandableEndwallRaftersWithNormalBay").Value;
            }

            RafterPlacement = (eWall + " endwall rafter");
            break;
        case "e3":
            ColumnCollection = b.e3Columns;
            FOCollection = b.e3FOs;
            RafterCollection = b.e3Rafters;
            IntColumnMode = false;
            if (b.ExpandableEndwall(eWall)) {
                RafterType = "W-Beam";
            }
            else if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
                RafterType = "8"" C Purlin";
            }
            else if ((EstSht.Range("Bay1_Length").offset((EstSht.Range("BayNum").Value - 1), 0).Value > 25)) {
                RafterType = SteelLookupSht.Range("NonExpandableEndwallRaftersWithLargeBay").Value;
            }
            else {
                RafterType = SteelLookupSht.Range("NonExpandableEndwallRaftersWithNormalBay").Value;
            }

            RafterPlacement = (eWall + " endwall rafter");
            break;
        case "int":
            ColumnCollection = b.InteriorColumns;
            RafterCollection = b.intRafters;
            eWall = "e1";
            RafterType = "W-Beam";
            IntColumnMode = true;
            RafterPlacement = "main rafter line";
            break;
        case "e1 Extension":
            ColumnCollection = b.e1ExtensionMembers;
            RafterCollection = b.e1Rafters;
            eWall = "e1";
            RafterType = "W-Beam";
            IntColumnMode = true;
            RafterPlacement = "e1 Extension Rafter";
            break;
        case "e3 Extension":
            ColumnCollection = b.e3ExtensionMembers;
            RafterCollection = b.e3Rafters;
            eWall = "e3";
            RafterType = "W-Beam";
            IntColumnMode = true;
            RafterPlacement = "e3 Extension Rafter";
            break;
    }

    // set actual start and end points for building using corner column widths
    for (Column in ColumnCollection) {
        if ((Column.rEdgePosition == 0)) {
            StartPos = Column.Width;
            FirstColWidth = Column.Width;
        }
        else if ((Column.lEdgePosition
                    == (b.bWidth * 12))) {
            MaxDistance = ((b.bWidth * 12)
                        - Column.Width);
            LastColWidth = Column.Width;
        }

    }

    // Rafters(0) = StartPos
    for (i = 0; (i <= 25); i++) {
        // not used, I just need to loop enough times to generate all the rafters. there should be a better way to do this. NEEDS FIX
        if ((EndPos < MaxDistance)) {
            tempNearestColumn = 1.79769313486231E+308;
            // loop through columns, find nearest rEdgePosition
            for (Column in ColumnCollection) {
                if (((Abs((Column.rEdgePosition - StartPos)) < Abs((tempNearestColumn - StartPos)))
                            && ((Column.rEdgePosition > StartPos)
                            && (Column.LoadBearing == true)))) {
                    tempNearestColumn = Column.rEdgePosition;
                    NextStartPos = Column.lEdgePosition;
                }

            }

            if ((IntColumnMode == false)) {
                // Only Endwalls have FOs
                for (FO in FOCollection) {
                    for (item in FO.FOMaterials) {
                        if ((item.clsType == "Member")) {
                            Member = item;
                            if ((Member.LoadBearing == true)) {
                                if (((Abs((Member.CL - StartPos)) < Abs((tempNearestColumn - StartPos)))
                                            && ((Member.rEdgePosition > StartPos)
                                            && (Member.tEdgeHeight == b.DistanceToRoof(eWall, Member.CL))))) {
                                    tempNearestColumn = Member.rEdgePosition;
                                    NextStartPos = Member.lEdgePosition;
                                }

                            }

                        }

                    }

                }

            }

            // check that girt edge does not exceed building width
            if ((tempNearestColumn <= MaxDistance)) {
                EndPos = tempNearestColumn;
            }
            else {
                EndPos = MaxDistance;
            }

            if (((EndPos
                        > (b.bWidth * (12 / 2)))
                        && ((StartPos
                        < (b.bWidth * (12 / 2)))
                        && (b.rShape == "Gable")))) {
                // if the next position is on the other side of the Gable Roof, add a midpoint and create a rafter that goes to midpoint
                RafterMember = new clsMember();
                RafterMember.mType = (RafterPlacement + " Rafter");
                MidPos = (b.bWidth * (12 / 2));
                RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, StartPos);
                RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, MidPos);
                RafterMember.rEdgePosition = StartPos;
                RafterMember.RafterLeftEdge = MidPos;
                RafterMember.Length = Sqr(((MidPos - StartPos)
                                | ((2
                                + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // Set Size to nearest COLUMN (not FO jambs)
                if ((RafterType != "W-Beam")) {
                    RafterMember.Width = RafterType.Substring(0, ((RafterType.IndexOf(" ", 0) + 1)
                                    - 2));
                    RafterMember.Size = RafterType;
                }
                else {
                    // To find size, use the distance between lEdge (higher) and rEdge (lower)
                    RafterMember.SetSize;
                    b;
                    "Rafter";
                    eWall;
                    Abs((EndPos - StartPos));
                }

                RafterMember.Placement = (RafterPlacement + (", "
                            + (RafterMember.Length + "' long")));
                RafterCollection.Add;
                RafterMember;
                // Second Rafter across gable peak
                RafterMember = new clsMember();
                RafterMember.mType = (RafterPlacement + " Rafter");
                RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, EndPos);
                RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, MidPos);
                RafterMember.rEdgePosition = MidPos;
                RafterMember.RafterLeftEdge = EndPos;
                RafterMember.Length = Sqr(((EndPos - MidPos)
                                | ((2
                                + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // Set Size to nearest COLUMN (not FO jambs)
                if ((RafterType != "W-Beam")) {
                    RafterMember.Width = RafterType.Substring(0, ((RafterType.IndexOf(" ", 0) + 1)
                                    - 2));
                    RafterMember.Size = RafterType;
                }
                else {
                    // To find size, use the distance between lEdge (higher) and rEdge (lower)
                    RafterMember.SetSize;
                    b;
                    "Rafter";
                    eWall;
                    Abs((EndPos - StartPos));
                }

                RafterMember.Placement = (RafterPlacement + (", "
                            + (RafterMember.Length + "' long")));
                RafterCollection.Add;
                RafterMember;
                StartPos = NextStartPos;
            }
            else {
                RafterMember = new clsMember();
                RafterMember.mType = (RafterPlacement + " Rafter");
                RafterMember.rEdgePosition = StartPos;
                RafterMember.RafterLeftEdge = EndPos;
                // If StartPos <= 27 Then StartPos = 0
                // If EndPos >= b.bWidth * 12 - 27 Then EndPos = b.bWidth * 12
                if ((((EndPos
                            > (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable"))
                            || ((b.rShape == "Single Slope")
                            && (eWall == "e1")))) {
                    // on the far side of a Gable Roof, tEdge and bEdge are switched
                    RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, StartPos);
                    RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, EndPos);
                }
                else {
                    RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, StartPos);
                    RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, EndPos);
                }

                // Set Size to nearest COLUMN (not FO jambs)
                if ((RafterType == "10"" Receiver Cee")) {
                    RafterMember.Width = 10;
                    RafterMember.Size = RafterType;
                }
                else if ((RafterType == "8"" Receiver Cee")) {
                    RafterMember.Width = 8;
                    RafterMember.Size = RafterType;
                }
                else if ((RafterType == "8"" C Purlin")) {
                    RafterMember.Width = 8;
                    RafterMember.Size = RafterType;
                }
                else if ((RafterType == "10"" C Purlin")) {
                    RafterMember.Width = 10;
                    RafterMember.Size = RafterType;
                }
                else {
                    RafterMember.SetSize;
                    b;
                    "Rafter";
                    eWall;
                    Abs((StartPos - EndPos));
                }

                RafterMember.Length = Sqr(((EndPos - StartPos)
                                | ((2
                                + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                RafterMember.Placement = (RafterPlacement + (", "
                            + (RafterMember.Length + "' long")));
                RafterCollection.Add;
                RafterMember;
                StartPos = NextStartPos;
                // ''''''' lower all rafters so that the top edge will be the actual building height
                // ''''''' AND '''''' Make rafters that connect to corners or center of gable building longer
                // formulas:
                // Distance to lower rafters = SQUARE ROOT(  ((width/2)(pitch/12))^2    *    (width/2)^2    )
                // Distance to lengthen rafters = (width/2)(pitch/12) per corner/peak
                Angle = (Atn((b.rPitch / 12))
                            * (RafterMember.Width / 2));
                DistanceToLower = Sqr((Angle
                                | ((2
                                + (RafterMember.Width / 2))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                DistanceToLengthen = Sqr((DistanceToLower
                                | ((2
                                - (RafterMember.Width / 2))
                                | 2)));
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
                RafterMember.bEdgeHeight = (RafterMember.bEdgeHeight - DistanceToLower);
                RafterMember.tEdgeHeight = (RafterMember.tEdgeHeight - DistanceToLower);
                if ((RafterMember.bEdgeHeight
                            < (b.bHeight * 12))) {
                    RafterMember.Length = (RafterMember.Length + DistanceToLengthen);
                    AngleCut = true;
                }

                // Single Slope
                if ((RafterMember.tEdgeHeight
                            >= ((b.bHeight * 12)
                            + ((b.bWidth * b.rPitch)
                            - DistanceToLower)))) {
                    RafterMember.Length = (RafterMember.Length + DistanceToLengthen);
                    AngleCut = true;
                }
                else if ((RafterMember.tEdgeHeight
                            >= ((b.bHeight * 12)
                            + ((b.bWidth / (2 * b.rPitch))
                            - DistanceToLower)))) {
                    RafterMember.Length = (RafterMember.Length + DistanceToLengthen);
                    AngleCut = true;
                }

                if (AngleCut) {
                    RafterMember.Placement = (RafterMember.Placement + (", cut at "
                                + (Application.WorksheetFunction.Round(Atn((b.rPitch / 12)), 2) + " angle, ")));
                }

            }

            i;
            if ((eWall == "int")) {
                for (RafterMember in RafterCollection) {
                    RafterMember.Qty = (EstSht.Range("BayNum").Value - 1);
                    // Debug.Print "x Start: " & RafterMember.rEdgePosition & ", y Start: " & RafterMember.bEdgeHeight & ", y End: " & RafterMember.tEdgeHeight & ", length: " & RafterMember.Length
                }

            }

        }

        // Specify Eave Strut types
        EaveStrutTypes((<clsBuilding>(b)), (<string>(eWall)));
        let EaveStruts: Collection;
        let Member: clsMember;
        let LinerPanels: boolean;
        let e1Overhang: boolean;
        let e1Extension: boolean;
        let e3Overhang: boolean;
        let e3Extension: boolean;
        let s2Overhang: boolean;
        let s2Extension: boolean;
        let s4Overhang: boolean;
        let s4Extension: boolean;
        // check for liner panels, overhangs, extension, soffit
        if ((EstSht.Range("Roof_LinerPanels").Value != "None")) {
            LinerPanels = true;
        }

        if ((EstSht.Range("e1_GableOverhang").Value > 0)) {
            e1Overhang = true;
        }

        if ((EstSht.Range("e1_GableExtension").Value > 0)) {
            e1Extension = true;
        }

        if ((EstSht.Range("e3_GableOverhang").Value > 0)) {
            e3Overhang = true;
        }

        if ((EstSht.Range("e3_GableExtension").Value > 0)) {
            e3Extension = true;
        }

        if ((EstSht.Range("s2_EaveOverhang").Value > 0)) {
            s2Overhang = true;
        }

        if ((EstSht.Range("s2_EaveExtension").Value > 0)) {
            s2Extension = true;
        }

        if ((EstSht.Range("s4_EaveOverhang").Value > 0)) {
            s4Overhang = true;
        }

        if ((EstSht.Range("s4_EaveExtension").Value > 0)) {
            s4Extension = true;
        }

        switch (eWall) {
            case "s2":
                EaveStruts = b.s2Girts;
                for (Member in EaveStruts) {
                    if ((Member.mType == "Eave Strut")) {
                        // if Overhang, extend eave strut
                        if (((Member.rEdgePosition
                                    + (Member.Length
                                    > ((b.bLength * 12)
                                    - 15)))
                                    && ((e3Overhang == true)
                                    && (e3Extension == false)))) {
                            Member.Length = (Member.Length
                                        + (EstSht.Range("e3_GableOverhang").Value * 12));
                            if ((b.rPitch == 1)) {
                                Member.Size = "8"" C Purlin";
                            }
                            else {
                                Member.Size = ("8"" "
                                            + (b.rPitch + ":12 "));
                            }

                            if ((b.e3GableOverhangSoffit
                                        || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                        && (b.rPitch != 1)))) {
                                Member.Size = (Member.Size + "double up eave strut");
                            }
                            else {
                                Member.Size = (Member.Size + "single up eave strut");
                            }

                            Debug.Print;
                            "s2 regular eave strut created";
                            Debug.Print;
                            (EaveStrutCount + 1);
                        }
                        else if (((Member.rEdgePosition < 15)
                                    && ((e1Overhang == true)
                                    && (e1Extension == false)))) {
                            Member.rEdgePosition = ((EstSht.Range("e1_GableOverhang").Value * 12)
                                        * -1);
                            Member.Length = (Member.Length
                                        + (EstSht.Range("e1_GableOverhang").Value * 12));
                            if ((b.rPitch == 1)) {
                                Member.Size = "8"" C Purlin";
                            }
                            else {
                                Member.Size = ("8"" "
                                            + (b.rPitch + ":12 "));
                            }

                            if ((b.e1GableOverhangSoffit
                                        || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                        && (b.rPitch != 1)))) {
                                Member.Size = (Member.Size + "double up eave strut");
                            }
                            else if ((b.rPitch == 1)) {
                                // do nothing
                            }
                            else {
                                Member.Size = (Member.Size + "single up eave strut");
                            }

                            Debug.Print;
                            "s2 regular eave strut created";
                            Debug.Print;
                            (EaveStrutCount + 1);
                        }
                        else {
                            // no overhangs, eave strut determined only by normal factors
                            if ((b.rPitch == 1)) {
                                Member.Size = "8"" C Purlin";
                            }
                            else {
                                Member.Size = ("8"" "
                                            + (b.rPitch + ":12 "));
                            }

                            if (LinerPanels) {
                                Member.Size = (Member.Size + "double up eave strut");
                            }
                            else if ((b.rPitch != 1)) {
                                Member.Size = (Member.Size + "single up eave strut");
                            }

                            Debug.Print;
                            "s2 regular eave strut created";
                            Debug.Print;
                            (EaveStrutCount + 1);
                        }

                    }

                }

                break;
            case "s4":
                EaveStruts = b.s4Girts;
                for (Member in EaveStruts) {
                    if ((Member.mType == "Eave Strut")) {
                        if ((b.rShape == "Gable")) {
                            // if Overhang, extend eave strut
                            if (((Member.rEdgePosition
                                        + (Member.Length
                                        > ((b.bLength * 12)
                                        - 15)))
                                        && ((e1Overhang == true)
                                        && (e1Extension == false)))) {
                                Member.Length = (Member.Length
                                            + (EstSht.Range("e1_GableOverhang").Value * 12));
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if ((b.e1GableOverhangSoffit
                                            || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                            && (b.rPitch != 1)))) {
                                    Member.Size = (Member.Size + "double up eave strut");
                                }
                                else {
                                    Member.Size = (Member.Size + "single up eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }
                            else if (((Member.rEdgePosition < 15)
                                        && ((e3Overhang == true)
                                        && (e1Extension == false)))) {
                                Member.rEdgePosition = ((EstSht.Range("e3_GableOverhang").Value * 12)
                                            * -1);
                                Member.Length = (Member.Length
                                            + (EstSht.Range("e3_GableOverhang").Value * 12));
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if ((b.e3GableOverhangSoffit
                                            || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                            && (b.rPitch != 1)))) {
                                    Member.Size = (Member.Size + "double up eave strut");
                                }
                                else {
                                    Member.Size = (Member.Size + "single up eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }
                            else {
                                // no overhangs, eave strut determined only by normal factors
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if (LinerPanels) {
                                    Member.Size = (Member.Size + "double up eave strut");
                                }
                                else if ((b.rPitch != 1)) {
                                    Member.Size = (Member.Size + "single up eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }

                        }
                        else {
                            // Single Slope --> struts will be single/double down
                            // if Overhang, extend eave strut
                            if (((Member.rEdgePosition
                                        + (Member.Length
                                        > ((b.bLength * 12)
                                        - 15)))
                                        && ((e1Overhang == true)
                                        && (e1Extension == false)))) {
                                Member.Length = (Member.Length
                                            + (EstSht.Range("e1_GableOverhang").Value * 12));
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if ((b.e1GableOverhangSoffit
                                            || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                            && (b.rPitch != 1)))) {
                                    Member.Size = (Member.Size + "double down eave strut");
                                }
                                else {
                                    Member.Size = (Member.Size + "single down eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }
                            else if (((Member.rEdgePosition < 15)
                                        && ((e3Overhang == true)
                                        && (e1Extension == false)))) {
                                Member.rEdgePosition = ((EstSht.Range("e3_GableOverhang").Value * 12)
                                            * -1);
                                Member.Length = (Member.Length
                                            + (EstSht.Range("e3_GableOverhang").Value * 12));
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if ((b.e3GableOverhangSoffit
                                            || ((EstSht.Range("Roof_LinerPanels").Value != "None")
                                            && (b.rPitch != 1)))) {
                                    Member.Size = (Member.Size + "double down eave strut");
                                }
                                else {
                                    Member.Size = (Member.Size + "single down eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }
                            else {
                                // no overhangs, eave strut determined only by normal factors
                                if ((b.rPitch == 1)) {
                                    Member.Size = "8"" C Purlin";
                                }
                                else {
                                    Member.Size = ("8"" "
                                                + (b.rPitch + ":12 "));
                                }

                                if (LinerPanels) {
                                    Member.Size = (Member.Size + "double down eave strut");
                                }
                                else if ((b.rPitch != 1)) {
                                    Member.Size = (Member.Size + "single down eave strut");
                                }

                                Debug.Print;
                                "s4 regular eave strut created";
                                Debug.Print;
                                (EaveStrutCount + 1);
                            }

                        }

                    }

                }

                break;
        }

        (<number>(ClosestWallGirt((<void>(Height)), Variant, Optional, (<number>(Direction)))));
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
        Purlins = Array(86, 146, 206, 266, 326, 386, 446, 506, 566, 626, 686, 746, 806, 866, 926, 986, 1046, 1106, 1166);
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
                        ClosestWallGirt = Purlin;
                    }

                }
                else if (((Direction < 0)
                            && (Purlin <= Height))) {
                    // only report if closer number is less than the target
                    if ((u < t)) {
                        t = u;
                        ClosestWallGirt = Purlin;
                    }

                }
                else if ((Direction == 0)) {
                    if ((u < t)) {
                        t = u;
                        ClosestWallGirt = Purlin;
                    }

                }

            }

        }

        // ''''''''' Creates Wall Collection with all relevant items, calculates girt lengths
        // '''''''''''''not finished
        EndwallGirtLengthCalc((<clsBuilding>(b)), Optional, (<string>(eWall)));
        let ColumnCollection: Collection;
        let FOCollection: Collection;
        let GirtsCollection: Collection;
        let RafterCollection: Collection;
        let Column: clsMember;
        let Girt: clsMember;
        let FO: clsFO;
        let FOMaterial: Collection;
        let Member: clsMember;
        let Rafter: clsMember;
        let item: Object;
        let Points: number[];
        let Girts: number[];
        let GirtNum: number;
        let LowestPoint: number;
        let WallStatus: string;
        let TotalHeight: number;
        let StartPos: number;
        let EndPos: number;
        let MaxDistance: number;
        let NextIntersection: number;
        let GirtIndex: number;
        let GirtMiddle: number;
        let GirtHeight: number;
        let WallLength: number;
        let DistanceToLower: number;
        let RafterGirtAdjustment: number;
        let RafterWidth: number;
        let i: number;
        let Angle: number;
        let Girts: Object;
        switch (eWall) {
            case "e1":
                ColumnCollection = b.e1Columns;
                FOCollection = b.e1FOs;
                GirtsCollection = b.e1Girts;
                RafterCollection = b.e1Rafters;
                break;
            case "s2":
                ColumnCollection = b.s2Columns;
                FOCollection = b.s2FOs;
                GirtsCollection = b.s2Girts;
                break;
            case "e3":
                ColumnCollection = b.e3Columns;
                FOCollection = b.e3FOs;
                GirtsCollection = b.e3Girts;
                RafterCollection = b.e1Rafters;
                break;
            case "s4":
                ColumnCollection = b.s4Columns;
                FOCollection = b.s4FOs;
                GirtsCollection = b.s4Girts;
                break;
        }

        WallStatus = b.WallStatus(eWall);
        LowestPoint = (b.LengthAboveFinishedFloor(eWall) * 12);
        // in
        // check for excluded wall
        if ((WallStatus == "Exclude")) {
            if (((eWall == "e1")
                        || (eWall == "e3"))) {
                return;
            }
            else {
                LowestPoint = (b.bHeight * 12);
            }

        }

        // get highest point of wall
        if (((eWall == "s2")
                    || ((eWall == "s4")
                    && (b.rShape == "Gable")))) {
            TotalHeight = (b.bHeight * 12);
        }
        else if (((eWall == "s4")
                    && (b.rShape == "Single Slope"))) {
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth * b.rPitch));
        }
        else if ((b.rShape == "Single Slope")) {
            // in
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth * b.rPitch));
        }
        else {
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth / (2 * b.rPitch)));
        }

        // base case, normal building in 5' increments after 12'
        if ((LowestPoint == 0)) {
            Girts[0] = 86;
            Girts[1] = 146;
            i = 2;
        }
        else {
            // Partial Walls or "gable only" walls starting above 86"
            Girts[0] = LowestPoint;
            i = 1;
            // partial walls starting lower than 86"
            if ((LowestPoint < 86)) {
                Girts[1] = 86;
                Girts[2] = 146;
                i = 3;
            }
            else if (((LowestPoint > 86)
                        && (LowestPoint < 146))) {
                // partial walls starting above 86" but below 146"
                Girts[1] = 146;
                i = 2;
            }
            else if ((WallStatus != "Exclude")) {
                // partial walls above 146"
                Girts[1] = ClosestWallGirt((LowestPoint + 60), -1);
                i = 2;
            }

        }

        GirtNum = i;
        // Add girt heights to array, if taller than building, value = 0
        for (i = i; (i <= 20); i++) {
            if (((Girts[(i - 1)] + (60 < TotalHeight))
                        && (Girts[(i - 1)] > 0))) {
                Girts[i] = (Girts[(i - 1)] + 60);
                GirtNum = (GirtNum + 1);
            }

        }

        if (((eWall == "s2")
                    || (eWall == "s4"))) {
            if ((WallStatus == "Exclude")) {
                GirtNum = 0;
                let Preserve: Object;
                Girts[GirtNum];
            }
            else {
                Girts[GirtNum] = TotalHeight;
                let Preserve: Object;
                Girts[GirtNum];
            }

        }
        else {
            let Preserve: Object;
            Girts[(GirtNum - 1)];
            GirtNum = (GirtNum - 1);
        }

        // get length of wall depending on wall selection
        if (((eWall == "e1")
                    || (eWall == "e3"))) {
            WallLength = (b.bWidth * 12);
        }
        else {
            WallLength = (b.bLength * 12);
        }

        // get rafter width and distance rafters were lowered
        if (((eWall == "e1")
                    || (eWall == "e3"))) {
            RafterWidth = 20;
            for (Rafter in RafterCollection) {
                if ((Rafter.Width < RafterWidth)) {
                    RafterWidth = Rafter.Width;
                }

            }

            Angle = (Atn((b.rPitch / 12)) * (180 / 3.14159265358979));
            DistanceToLower = ((RafterWidth / 2)
                        / (Cos(Angle) * (180 / 3.14159265358979)));
            Angle = (Atn((b.rPitch / 12))
                        * (RafterWidth / 2));
            DistanceToLower = Sqr((Angle
                            | ((2
                            + (RafterWidth / 2))
                            | 2)));
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            RafterGirtAdjustment = ((DistanceToLower / b.rPitch)
                        * 12);
        }
        else {
            RafterGirtAdjustment = 0;
        }

        // Run through each girt row;
        // set starting value = 0 for begining of row;
        // set ending value = first intersecting wall item;
        // create member, add to girt collection
        for (i = 0; (i <= GirtNum); i++) {
            EndPos = 0;
            // girts under eave height start at 0
            // s2 will only have girts up to eave height
            // s4 will have girts above eave height on single slope
            if (((Girts[i]
                        <= (b.bHeight * 12))
                        || ((Girts[i] <= b.DistanceToRoof(eWall, 0))
                        && ((b.rShape == "Single Slope")
                        && (eWall == "s4"))))) {
                StartPos = 0;
                while ((EndPos < WallLength)) {
                    StartPos = EndPos;
                    EndPos = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts[i]);
                    if ((EndPos > WallLength)) {
                        // for s2 and s4 since sidewall column collection doesn't include endpoints
                        EndPos = WallLength;
                    }

                    Girt = new clsMember();
                    Girt.mType = "Girt";
                    Girt.Size = "8"" C Purlin";
                    Girt.bEdgeHeight = Girts[i];
                    Girt.tEdgeHeight = Girts[i];
                    Girt.rEdgePosition = StartPos;
                    Girt.Length = (EndPos - StartPos);
                    if ((((Girts[i]
                                == (b.bHeight * 12))
                                && (eWall == "s2"))
                                || ((eWall == "s4")
                                && ((Girts[i]
                                == (b.bHeight * 12))
                                && (b.rShape == "Gable"))))) {
                        Girt.mType = "Eave Strut";
                    }
                    else if (((eWall == "s4")
                                && ((Girts[i]
                                == ((b.bHeight * 12)
                                + (b.bWidth * (12
                                * (b.rPitch / 12)))))
                                && (b.rShape == "Single Slope")))) {
                        Girt.mType = "Eave Strut";
                    }

                    Girt.Placement = ("girt screwline row "
                                + ((i + 1) + (" at "
                                + (Girts[i] + (" inches, wall "
                                + (eWall + (", start "
                                + (StartPos + (" inches, end "
                                + (EndPos + (" inches, length " + Girt.Length)))))))))));
                    GirtsCollection.Add;
                    Girt;
                    ("e1" | eWall) = "e3";
                    eWall = "e3";
                    // for endwalls girts above eave height start at roof location at height
                    switch (b.rShape) {
                        case "Single Slope":
                            if ((eWall == "e1")) {
                                StartPos = 0;
                                EndPos = 0;
                                MaxDistance = (b.DistanceFromCorner(eWall, Girts[i]) - RafterGirtAdjustment);
                                while ((EndPos < MaxDistance)) {
                                    StartPos = EndPos;
                                    NextIntersection = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts[i]);
                                    // for endpoing on gable roof, columns past max distance should be ignored and replaced with max distance across gable
                                    if ((NextIntersection > MaxDistance)) {
                                        EndPos = MaxDistance;
                                    }
                                    else {
                                        EndPos = NextIntersection;
                                    }

                                    Girt = new clsMember();
                                    Girt.mType = "Girt";
                                    Girt.Size = "8"" C Purlin";
                                    Girt.bEdgeHeight = Girts[i];
                                    Girt.tEdgeHeight = Girts[i];
                                    // NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                                    Girt.rEdgePosition = StartPos;
                                    Girt.Length = (EndPos - StartPos);
                                    // Girt.Width = Girt.Length 'ONE OF THESE SHOULD BE REMOVED
                                    Girt.Placement = ("girt screwline row "
                                                + ((i + 1) + (" at "
                                                + (Girts[i] + (" inches, wall "
                                                + (eWall + (", start "
                                                + (StartPos + (" inches, end "
                                                + (EndPos + (" inches, length " + Girt.Length)))))))))));
                                    GirtsCollection.Add;
                                    Girt;
                                    // e3 single slope: 0' is lowest point
                                    eWall = "e3";
                                    StartPos = (b.DistanceFromCorner(eWall, Girts[i]) + RafterGirtAdjustment);
                                    EndPos = StartPos;
                                    MaxDistance = (b.bWidth * 12);
                                    while ((EndPos < WallLength)) {
                                        StartPos = EndPos;
                                        NextIntersection = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts[i]);
                                        if ((NextIntersection > MaxDistance)) {
                                            EndPos = MaxDistance;
                                        }
                                        else {
                                            EndPos = NextIntersection;
                                        }

                                        Girt = new clsMember();
                                        Girt.mType = "Girt";
                                        Girt.Size = "8"" C Purlin";
                                        Girt.bEdgeHeight = Girts[i];
                                        Girt.tEdgeHeight = Girts[i];
                                        // NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                                        Girt.rEdgePosition = StartPos;
                                        Girt.Length = (EndPos - StartPos);
                                        // Girt.Width = Girt.Length 'ONE OF THESE SHOULD BE REMOVED
                                        Girt.Placement = ("girt screwline row "
                                                    + ((i + 1) + (" at "
                                                    + (Girts[i] + (" inches, wall "
                                                    + (eWall + (", start "
                                                    + (StartPos + (" inches, end "
                                                    + (EndPos + (" inches, length " + Girt.Length)))))))))));
                                        GirtsCollection.Add;
                                        Girt;
                                    }

                                    "Gable";
                                    StartPos = (b.DistanceFromCorner(eWall, Girts[i]) + RafterGirtAdjustment);
                                    EndPos = StartPos;
                                    MaxDistance = (WallLength
                                                - (b.DistanceFromCorner(eWall, Girts[i]) - RafterGirtAdjustment));
                                    while ((EndPos < MaxDistance)) {
                                        StartPos = EndPos;
                                        EndPos = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts[i]);
                                        if ((EndPos > MaxDistance)) {
                                            EndPos = MaxDistance;
                                        }

                                        Girt = new clsMember();
                                        Girt.mType = "Girt";
                                        Girt.Size = "8"" C Purlin";
                                        Girt.bEdgeHeight = Girts[i];
                                        Girt.tEdgeHeight = Girts[i];
                                        // NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                                        Girt.rEdgePosition = StartPos;
                                        Girt.Length = (EndPos - StartPos);
                                        Girt.Width = Girt.Length;
                                        // ONE OF THESE SHOULD BE REMOVED
                                        Girt.Placement = ("girt screwline row "
                                                    + (i + (" at "
                                                    + (Girts[i] + (" inches, wall "
                                                    + (eWall + (", start "
                                                    + (StartPos + (" inches, end "
                                                    + (EndPos + (" inches, length " + Girt.Length)))))))))));
                                        GirtsCollection.Add;
                                        Girt;
                                    }

                                }

                                i;
                                // Step through girts, check if it's in the middle of an FO, then remove
                                for (GirtIndex = GirtsCollection.Count; (GirtIndex <= 1); GirtIndex = (GirtIndex + -1)) {
                                    Girt = GirtsCollection[GirtIndex];
                                    GirtMiddle = (Girt.rEdgePosition
                                                + (Girt.Length / 2));
                                    // since each girt starts and ends at a column/jamb, middle points are either valid or invalid
                                    GirtHeight = Girt.tEdgeHeight;
                                    for (FO in FOCollection) {
                                        if (((GirtHeight > FO.bEdgeHeight)
                                                    && ((GirtHeight < FO.tEdgeHeight)
                                                    && ((GirtMiddle > FO.rEdgePosition)
                                                    && (GirtMiddle < FO.lEdgePosition))))) {
                                            GirtsCollection.Remove(GirtIndex);
                                        }

                                    }

                                }

                                // ---------DEBUG--------
                                // For Each Girt In GirtsCollection
                                // Debug.Print Girt.Placement
                                // Next Girt
                            }

                            (<number>(NextHorizontalGirtIntersection((<clsBuilding>(b)), (<Collection>(Columns)), (<Collection>(FOs)), (<number>(start)), (<string>(Wall)), (<number>(Height)))));
                            let Member: clsMember;
                            let FO: clsFO;
                            let item: Object;
                            let tempNearestIntersection: number;
                            tempNearestIntersection = 1.79769313486231E+308;
                            if (((Wall == "e1")
                                        || (Wall == "e3"))) {
                                for (Member in Columns) {
                                    if ((((Member.CL - start)
                                                < (tempNearestIntersection - start))
                                                && ((Member.CL > start)
                                                && ((Member.CL > 15)
                                                && (Member.CL
                                                < ((b.bWidth * 12)
                                                - 15)))))) {
                                        tempNearestIntersection = Member.CL;
                                    }

                                }

                            }
                            else {
                                // s2 and s4 use bLength
                                for (Member in Columns) {
                                    if ((((Member.CL - start)
                                                < (tempNearestIntersection - start))
                                                && ((Member.CL > start)
                                                && ((Member.CL > 15)
                                                && (Member.CL
                                                < ((b.bLength * 12)
                                                - 15)))))) {
                                        tempNearestIntersection = Member.CL;
                                    }

                                }

                            }

                            for (FO in FOs) {
                                for (item in FO.FOMaterials) {
                                    if ((item.clsType == "Member")) {
                                        Member = item;
                                        if ((((Member.CL - start)
                                                    < (tempNearestIntersection - start))
                                                    && (Member.CL > start))) {
                                            // check if FO is along girt height OR if member intersects with girt height
                                            if ((((FO.bEdgeHeight < Height)
                                                        && (FO.tEdgeHeight > Height))
                                                        || ((Member.bEdgeHeight < Height)
                                                        && (Member.tEdgeHeight > Height)))) {
                                                tempNearestIntersection = Member.CL;
                                            }

                                        }

                                    }

                                }

                                if ((((FO.rEdgePosition - start)
                                            < (tempNearestIntersection - start))
                                            && (FO.rEdgePosition > start))) {
                                    // check if FO is along girt height OR if member intersects with girt height
                                    if (((FO.bEdgeHeight < Height)
                                                && (FO.tEdgeHeight > Height))) {
                                        tempNearestIntersection = FO.rEdgePosition;
                                    }

                                }
                                else if ((((FO.lEdgePosition - start)
                                            < (tempNearestIntersection - start))
                                            && (FO.lEdgePosition > start))) {
                                    // check if FO is along girt height OR if member intersects with girt height
                                    if (((FO.bEdgeHeight < Height)
                                                && (FO.tEdgeHeight > Height))) {
                                        tempNearestIntersection = FO.lEdgePosition;
                                    }

                                }

                            }

                            NextHorizontalGirtIntersection = tempNearestIntersection;
                            break;
                    }

                }

                TestGirGen((<clsBuilding>(b)));
                // start w/ column collection
                // start w/ FO collection
                // start w/ building
                // find out how many girt lines exist
                // # of gerts = 2 + RoundDown((TotalHeight - 12)/5)
                // create array of column intersections for each girt line
                // turn array into collection of spans
                // create array of all FO intersections for each girt line
                // remove negative space for FODoors and MISCFOs from collection of spans
                // MAYBE: add skirts/headers & jambs for MISCFOs and Windows
                // RESULT: collection of spans includes:
                // girts
                // MAYBE: skirts/headers/jambs
                // placement description
                // PLACEMENT NAMING CONVENTION:
                // "Screwline height etc.
                // "Segment #1, #2, #3, etc.
                // try to describe startpoint and endpoint
                // "Window #1 Skirt, Window #1 Header, Window #1 jamb(s), etc.
                // combine collection of spans for each wall, pass to BPP
                // RETURNS: collection of ORDERED members, with "combinedmembers" collection inside each
                // CombinedMembers includes:
                // placement description
                // span length
                // questions:
                // are skirts/headers/jambs the same material as the girts?
                // do gerts extend above the bHeight? - yes
                // Single Slope Building ~ 70':
                // Endwall Columns = 4; 2 corners, 2 inside
                // Need 1 interior column
                // Should it line up with the shorter of the 2 endwall columns? or the longer? or neither?
            }

            EndwallExtensionColumnsGen((<clsBuilding>(b)), (<string>(eWall)), Optional, (<number>(NewColNum)), Optional, (<boolean>(Reiterate)));
            // This sub is ONLY used in buildings with 1 bay which also have endwall Extensions, since there are no columns to copy
            // The calculation is the same as the Interior Column gen
            let ColumnCollection: Collection;
            let Column: clsMember;
            let e1CenterColumn: boolean;
            let e3CenterColumn: boolean;
            let RafterNum: number;
            let i: number;
            let j: number;
            let ColLocation: number[];
            let MaxHorizontalDistance: number;
            let MinHorizontalDistance: number;
            let StartWidth: number;
            let EndWidth: number;
            let PrevWidth: number;
            let DistanceToPreviousColumn: number;
            let DistanceToNextColumn: number;
            let LargerDistance: number;
            let ColNum: number;
            if ((eWall == "e1")) {
                ColumnCollection = b.e1ExtensionMembers;
            }
            else if ((eWall == "e3")) {
                ColumnCollection = b.e3ExtensionMembers;
            }

            RafterNum = b.s2Columns.Count;
            // find horizontal distance equal to 60' rafter for this building plus maximum and minimum column thicknesses
            MaxHorizontalDistance = (60 / Sqr(((b.rPitch / 12) | (2 + 1))));
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            MinHorizontalDistance = (60 / Sqr(((b.rPitch / 12) | (2 + 1))));
            // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
            if ((Reiterate == true)) {
                ColNum = NewColNum;
            }
            else {
                if ((ColNum == 0)) {
                    if ((b.rShape == "Gable")) {
                        if ((b.bWidth <= 80)) {
                            ColNum = 0;
                        }
                        else if (((b.bWidth > 80)
                                    && (b.bWidth
                                    < (MaxHorizontalDistance * 2)))) {
                            ColNum = 1;
                        }
                        else if ((b.bWidth
                                    >= (MaxHorizontalDistance * 2))) {
                            ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                        }

                    }
                    else if ((b.rShape == "Single Slope")) {
                        if ((b.bWidth < MaxHorizontalDistance)) {
                            ColNum = 0;
                        }
                        else if ((b.bWidth > MaxHorizontalDistance)) {
                            ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                        }

                    }

                }

                // lower Col Num by 1 on first iteration to check for marginal cases
                // some column widths (to be determined) will require less columns, this will check those cases
                if ((ColNum > 0)) {
                    ColNum = (ColNum - 1);
                }

            }

            // first, evenly space columns along the width of the building to adjust later; add to array
            let ColLocation: Object;
            ColLocation[0] = 0;
            ColLocation[(ColNum + 1)] = (b.bWidth * 12);
            switch (ColNum) {
                case 1:
                    ColLocation[1] = (b.bWidth / (2 * 12));
                    break;
                case 2:
                    ColLocation[1] = (b.bWidth / (3 * 12));
                    ColLocation[2] = (b.bWidth / (3 * (12 * 2)));
                    break;
                case 3:
                    ColLocation[1] = (b.bWidth / (4 * 12));
                    ColLocation[2] = (b.bWidth / (4 * (12 * 2)));
                    ColLocation[3] = (b.bWidth / (4 * (12 * 3)));
                    break;
                case 4:
                    ColLocation[1] = (b.bWidth / (5 * 12));
                    ColLocation[2] = (b.bWidth / (5 * (12 * 2)));
                    ColLocation[3] = (b.bWidth / (5 * (12 * 3)));
                    ColLocation[4] = (b.bWidth / (5 * (12 * 4)));
                    break;
            }

            if ((eWall == "e3")) {
                for (i = 0; (i <= UBound(ColLocation[])); i++) {
                    ColLocation[i] = ((b.bWidth * 12)
                                - ColLocation[i]);
                }

            }

            // loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
            for (i = 1; (i <= ColNum); i++) {
                if ((ConflictingEndwallOHDoor(ColLocation[i], b) == true)) {
                    ColLocation[i] = NearestEndwallLocation(ColLocation[i], b);
                }

            }

            // ''''''''''''''check for No Interior Columns
            if ((ColNum == 0)) {
                // '''''''''''''Distance between Columns
                DistanceToPreviousColumn = Abs((ColLocation[0] - ColLocation[1]));
                // '''''''''''''Estimate COlumn widths
                // get first width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof(eWall, ColLocation[0]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[0] - ColLocation[1]));
                // subtract half of first width
                DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                // get second width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof(eWall, ColLocation[1]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[0] - ColLocation[1]));
                // subtract half of second width
                DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                if ((DistanceToPreviousColumn
                            > (MaxHorizontalDistance * 12))) {
                    ColLocation;
                    EndwallExtensionColumnsGen(b, eWall, (ColNum + 1), true);
                    return;
                }

            }

            // ''''''''''''''check Interior Columns
            // check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
            for (i = 1; (i <= ColNum); i++) {
                // get distance to next column to make sure it does NOT exceed max rafter length
                // if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
                // estimate column widths to get accurate distances
                // '''''''''''''Distance to PREVIOUS Column
                if (((ColLocation[i]
                            > (b.bWidth * (12 / 2)))
                            && ((ColLocation[(i - 1)]
                            < (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable")))) {
                    DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                                    - ColLocation[i]));
                }
                else {
                    DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                }

                // '''''''''''''Estimate COlumn widths
                // get first width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof(eWall, ColLocation[i]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                // subtract half of width
                DistanceToPreviousColumn = (DistanceToPreviousColumn
                            - (Column.Width / 2));
                // get second width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof(eWall, ColLocation[(i - 1)]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                // subtract width if sidewall column, or half of width otherwise
                if (((i - 1)
                            == 0)) {
                    DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                }
                else {
                    DistanceToPreviousColumn = (DistanceToPreviousColumn
                                - (Column.Width / 2));
                }

                // '''''''''''''Distance to NEXT Column
                if (((ColLocation[i]
                            < (b.bWidth * (12 / 2)))
                            && ((ColLocation[(i + 1)]
                            > (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable")))) {
                    DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                                    - ColLocation[i]));
                }
                else {
                    DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                }

                // '''''''''''''Estimate COlumn widths
                // get first width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                // subtract half of width
                DistanceToNextColumn = (DistanceToNextColumn
                            - (Column.Width / 2));
                // get second width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof(eWall, ColLocation[(i + 1)]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                // subtract width if sidewall column, or half of width otherwise
                if (((i + 1)
                            == UBound(ColLocation[]))) {
                    DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
                }
                else {
                    DistanceToNextColumn = (DistanceToNextColumn
                                - (Column.Width / 2));
                }

                // check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
                if (((DistanceToPreviousColumn
                            > (MaxHorizontalDistance * 12))
                            || (DistanceToNextColumn
                            > (MaxHorizontalDistance * 12)))) {
                    // Debug.Print "columns too far apart"
                    // CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
                    if ((NearestEndwallLocation(ColLocation[i], b, "Alternate") != ColLocation[i])) {
                        ColLocation[i] = NearestEndwallLocation(ColLocation[i], b, "Alternate");
                        // '''''''''''''Distance to PREVIOUS Column
                        if (((ColLocation[i]
                                    > (b.bWidth * (12 / 2)))
                                    && ((ColLocation[(i - 1)]
                                    < (b.bWidth * (12 / 2)))
                                    && (b.rShape == "Gable")))) {
                            DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                                            - ColLocation[i]));
                        }
                        else {
                            DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                        }

                        // '''''''''''''Estimate COlumn widths
                        // get first width
                        Column = new clsMember();
                        Column.Length = b.DistanceToRoof(eWall, ColLocation[i]);
                        Column.tEdgeHeight = Column.Length;
                        Column.SetSize;
                        b;
                        "Column";
                        "Interior";
                        Abs((ColLocation[(i - 1)] - ColLocation[i]));
                        // subtract half of width
                        DistanceToPreviousColumn = (DistanceToPreviousColumn
                                    - (Column.Width / 2));
                        // get second width
                        Column = new clsMember();
                        Column.Length = b.DistanceToRoof(eWall, ColLocation[(i - 1)]);
                        Column.tEdgeHeight = Column.Length;
                        Column.SetSize;
                        b;
                        "Column";
                        "Interior";
                        Abs((ColLocation[(i - 1)] - ColLocation[i]));
                        // subtract width if sidewall column, or half of width otherwise
                        if (((i - 1)
                                    == 0)) {
                            DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                        }
                        else {
                            DistanceToPreviousColumn = (DistanceToPreviousColumn
                                        - (Column.Width / 2));
                        }

                        // '''''''''''''Distance to NEXT Column
                        if (((ColLocation[i]
                                    < (b.bWidth * (12 / 2)))
                                    && ((ColLocation[(i + 1)]
                                    > (b.bWidth * (12 / 2)))
                                    && (b.rShape == "Gable")))) {
                            DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                                            - ColLocation[i]));
                        }
                        else {
                            DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                        }

                        // '''''''''''''Estimate COlumn widths
                        // get first width
                        Column = new clsMember();
                        Column.Length = b.DistanceToRoof(eWall, ColLocation[i]);
                        Column.tEdgeHeight = Column.Length;
                        Column.SetSize;
                        b;
                        "Column";
                        "Interior";
                        Abs((ColLocation[(i + 1)] - ColLocation[i]));
                        // subtract half of width
                        DistanceToNextColumn = (DistanceToNextColumn
                                    - (Column.Width / 2));
                        // get second width
                        Column = new clsMember();
                        Column.Length = b.DistanceToRoof(eWall, ColLocation[(i + 1)]);
                        Column.tEdgeHeight = Column.Length;
                        Column.SetSize;
                        b;
                        "Column";
                        "Interior";
                        Abs((ColLocation[(i + 1)] - ColLocation[i]));
                        // subtract width if sidewall column, or half of width otherwise
                        if (((i + 1)
                                    == UBound(ColLocation[]))) {
                            DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
                        }
                        else {
                            DistanceToNextColumn = (DistanceToNextColumn
                                        - (Column.Width / 2));
                        }

                    }

                    if (((DistanceToPreviousColumn
                                > (MaxHorizontalDistance * 12))
                                || (DistanceToNextColumn
                                > (MaxHorizontalDistance * 12)))) {
                        ColLocation;
                        IntColumnsGen(b, (ColNum + 1), true);
                        return;
                    }

                    //     ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
                    //     Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
                    //     'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
                    //         EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
                    //         StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
                    //         PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
                    //         DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
                    //         DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
                    //         If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
                    //             Erase ColLocation
                    //             Call IntColumnsGen(b, ColNum + 1)
                    //             Exit Sub
                    //         End If
                }

            }

            // debugging
            for (i = 0; (i
                        <= (ColNum + 1)); i++) {
                Debug.Print;
                ("Column #: "
                            + (i + (", "
                            + ((ColLocation[i] / 12) + ("' from s2, Rafter Line " + i)))));
            }

            // set column variables, types, sizes, etc.
            for (i = 0; (i
                        <= (ColNum + 1)); i++) {
                // find larger distance to neighboring columns to use in lookup tables
                // s2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
                if ((i
                            == (ColNum + 1))) {
                    LargerDistance = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                }
                else if ((i == 0)) {
                    LargerDistance = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                }
                else {
                    LargerDistance = Application.WorksheetFunction.Max(Abs((ColLocation[i] - ColLocation[(i - 1)])), Abs((ColLocation[i] - ColLocation[(i + 1)])));
                }

                Column = new clsMember();
                Column.mType = (eWall + " Extension Column");
                Column.CL = ColLocation[i];
                Column.LoadBearing = true;
                Column.Qty = 1;
                Column.Placement = (eWall + " Extension Column");
                if ((b.rShape == "Single Slope")) {
                    if ((i == 0)) {
                        Column.Length = (((b.bWidth * 12)
                                    * (b.rPitch / 12))
                                    + (b.bHeight * 12));
                    }
                    else if ((i
                                == (ColNum + 1))) {
                        Column.Length = (b.bHeight * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof(eWall, Column.CL);
                    }

                }
                else {
                    // Gable
                    if ((i == 0)) {
                        Column.Length = (b.bHeight * 12);
                    }
                    else if ((i
                                == (ColNum + 1))) {
                        Column.Length = (b.bHeight * 12);
                    }
                    else {
                        Column.Length = b.DistanceToRoof(eWall, Column.CL);
                    }

                }

                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                LargerDistance;
                if ((Column.CL == 0)) {
                    Column.CL = (Column.Width / 2);
                }
                else if ((Column.CL
                            == (b.bWidth * 12))) {
                    Column.CL = ((b.bWidth * 12)
                                - (Column.Width / 2));
                }

                Column.rEdgePosition = (Column.CL
                            - (Column.Width / 2));
                Column.Placement = (eWall + " Extension Column");
                ColumnCollection.Add;
                Column;
            }

        }

    }

}

IntColumnsGen(b: clsBuilding, NewColNum: number, Reiterate: boolean) {
    let e1ColumnCollection: Collection;
    // Warning!!! Optional parameters not supported
    // Warning!!! Optional parameters not supported
    let e3ColumnCollection: Collection;
    let Column: clsMember;
    let e1CenterColumn: boolean;
    let e3CenterColumn: boolean;
    let RafterNum: number;
    let i: number;
    let j: number;
    let ColLocation: number[];
    let MaxHorizontalDistance: number;
    let MinHorizontalDistance: number;
    let StartWidth: number;
    let EndWidth: number;
    let PrevWidth: number;
    let DistanceToPreviousColumn: number;
    let DistanceToNextColumn: number;
    let LargerDistance: number;
    let ColNum: number;
    // check for building with only 1 bay (no main rafter lines)
    if (((EstSht.Range("BayNum").Value == 1)
                && (NewColNum == 0))) {
        return;
    }

    RafterNum = b.s2Columns.Count;
    // find horizontal distance equal to 60' rafter for this building plus maximum and minimum column thicknesses
    MaxHorizontalDistance = (60 / Sqr(((b.rPitch / 12) | (2 + 1))));
    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
    MinHorizontalDistance = (60 / Sqr(((b.rPitch / 12) | (2 + 1))));
    // TODO: Warning!!! The operator should be an XOR ^ instead of an OR, but not available in CodeDOM
    if ((Reiterate == true)) {
        ColNum = NewColNum;
    }
    else {
        if ((ColNum == 0)) {
            if ((b.rShape == "Gable")) {
                if ((b.bWidth <= 80)) {
                    ColNum = 0;
                }
                else if (((b.bWidth > 80)
                            && (b.bWidth
                            < (MaxHorizontalDistance * 2)))) {
                    ColNum = 1;
                }
                else if ((b.bWidth
                            >= (MaxHorizontalDistance * 2))) {
                    ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                }

            }
            else if ((b.rShape == "Single Slope")) {
                if ((b.bWidth < MaxHorizontalDistance)) {
                    ColNum = 0;
                }
                else if ((b.bWidth > MaxHorizontalDistance)) {
                    ColNum = (Application.WorksheetFunction.RoundUp((b.bWidth / MaxHorizontalDistance), 0) - 1);
                }

            }

        }

        // lower Col Num by 1 on first iteration to check for marginal cases
        // some column widths (to be determined) will require less columns, this will check those cases
        if ((ColNum > 0)) {
            ColNum = (ColNum - 1);
        }

    }

    // first, evenly space columns along the width of the building to adjust later; add to array
    let ColLocation: Object;
    ColLocation[0] = 0;
    ColLocation[(ColNum + 1)] = (b.bWidth * 12);
    switch (ColNum) {
        case 1:
            ColLocation[1] = (b.bWidth / (2 * 12));
            break;
        case 2:
            ColLocation[1] = (b.bWidth / (3 * 12));
            ColLocation[2] = (b.bWidth / (3 * (12 * 2)));
            break;
        case 3:
            ColLocation[1] = (b.bWidth / (4 * 12));
            ColLocation[2] = (b.bWidth / (4 * (12 * 2)));
            ColLocation[3] = (b.bWidth / (4 * (12 * 3)));
            break;
        case 4:
            ColLocation[1] = (b.bWidth / (5 * 12));
            ColLocation[2] = (b.bWidth / (5 * (12 * 2)));
            ColLocation[3] = (b.bWidth / (5 * (12 * 3)));
            ColLocation[4] = (b.bWidth / (5 * (12 * 4)));
            break;
    }

    // loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
    for (i = 1; (i <= ColNum); i++) {
        if ((ConflictingEndwallOHDoor(ColLocation[i], b) == true)) {
            ColLocation[i] = NearestEndwallLocation(ColLocation[i], b);
        }

    }

    // ''''''''''''''check for No Interior Columns
    if ((ColNum == 0)) {
        // '''''''''''''Distance between Columns
        DistanceToPreviousColumn = Abs((ColLocation[0] - ColLocation[1]));
        // '''''''''''''Estimate COlumn widths
        // get first width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[0]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[0] - ColLocation[1]));
        // subtract half of first width
        DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
        // get second width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[1]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[0] - ColLocation[1]));
        // subtract half of second width
        DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
        if ((DistanceToPreviousColumn
                    > (MaxHorizontalDistance * 12))) {
            ColLocation;
            IntColumnsGen(b, (ColNum + 1), true);
            return;
        }

    }

    // ''''''''''''''check Interior Columns
    // check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
    for (i = 1; (i <= ColNum); i++) {
        // get distance to next column to make sure it does NOT exceed max rafter length
        // if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
        // estimate column widths to get accurate distances
        // '''''''''''''Distance to PREVIOUS Column
        if (((ColLocation[i]
                    > (b.bWidth * (12 / 2)))
                    && ((ColLocation[(i - 1)]
                    < (b.bWidth * (12 / 2)))
                    && (b.rShape == "Gable")))) {
            DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                            - ColLocation[i]));
        }
        else {
            DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
        }

        // '''''''''''''Estimate COlumn widths
        // get first width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[(i - 1)] - ColLocation[i]));
        // subtract half of width
        DistanceToPreviousColumn = (DistanceToPreviousColumn
                    - (Column.Width / 2));
        // get second width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[(i - 1)]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[(i - 1)] - ColLocation[i]));
        // subtract width if sidewall column, or half of width otherwise
        if (((i - 1)
                    == 0)) {
            DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
        }
        else {
            DistanceToPreviousColumn = (DistanceToPreviousColumn
                        - (Column.Width / 2));
        }

        // '''''''''''''Distance to NEXT Column
        if (((ColLocation[i]
                    < (b.bWidth * (12 / 2)))
                    && ((ColLocation[(i + 1)]
                    > (b.bWidth * (12 / 2)))
                    && (b.rShape == "Gable")))) {
            DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                            - ColLocation[i]));
        }
        else {
            DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
        }

        // '''''''''''''Estimate COlumn widths
        // get first width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[(i + 1)] - ColLocation[i]));
        // subtract half of width
        DistanceToNextColumn = (DistanceToNextColumn
                    - (Column.Width / 2));
        // get second width
        Column = new clsMember();
        Column.Length = b.DistanceToRoof("e1", ColLocation[(i + 1)]);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[(i + 1)] - ColLocation[i]));
        // subtract width if sidewall column, or half of width otherwise
        if (((i + 1)
                    == UBound(ColLocation[]))) {
            DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
        }
        else {
            DistanceToNextColumn = (DistanceToNextColumn
                        - (Column.Width / 2));
        }

        // check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
        if (((DistanceToPreviousColumn
                    > (MaxHorizontalDistance * 12))
                    || (DistanceToNextColumn
                    > (MaxHorizontalDistance * 12)))) {
            // Debug.Print "columns too far apart"
            // CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
            if ((NearestEndwallLocation(ColLocation[i], b, "Alternate") != ColLocation[i])) {
                ColLocation[i] = NearestEndwallLocation(ColLocation[i], b, "Alternate");
                // '''''''''''''Distance to PREVIOUS Column
                if (((ColLocation[i]
                            > (b.bWidth * (12 / 2)))
                            && ((ColLocation[(i - 1)]
                            < (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable")))) {
                    DistanceToPreviousColumn = Abs(((b.bWidth * (12 / 2))
                                    - ColLocation[i]));
                }
                else {
                    DistanceToPreviousColumn = Abs((ColLocation[i] - ColLocation[(i - 1)]));
                }

                // '''''''''''''Estimate COlumn widths
                // get first width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                // subtract half of width
                DistanceToPreviousColumn = (DistanceToPreviousColumn
                            - (Column.Width / 2));
                // get second width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof("e1", ColLocation[(i - 1)]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i - 1)] - ColLocation[i]));
                // subtract width if sidewall column, or half of width otherwise
                if (((i - 1)
                            == 0)) {
                    DistanceToPreviousColumn = (DistanceToPreviousColumn - Column.Width);
                }
                else {
                    DistanceToPreviousColumn = (DistanceToPreviousColumn
                                - (Column.Width / 2));
                }

                // '''''''''''''Distance to NEXT Column
                if (((ColLocation[i]
                            < (b.bWidth * (12 / 2)))
                            && ((ColLocation[(i + 1)]
                            > (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable")))) {
                    DistanceToNextColumn = Abs(((b.bWidth * (12 / 2))
                                    - ColLocation[i]));
                }
                else {
                    DistanceToNextColumn = Abs((ColLocation[i] - ColLocation[(i + 1)]));
                }

                // '''''''''''''Estimate COlumn widths
                // get first width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof("e1", ColLocation[i]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                // subtract half of width
                DistanceToNextColumn = (DistanceToNextColumn
                            - (Column.Width / 2));
                // get second width
                Column = new clsMember();
                Column.Length = b.DistanceToRoof("e1", ColLocation[(i + 1)]);
                Column.tEdgeHeight = Column.Length;
                Column.SetSize;
                b;
                "Column";
                "Interior";
                Abs((ColLocation[(i + 1)] - ColLocation[i]));
                // subtract width if sidewall column, or half of width otherwise
                if (((i + 1)
                            == UBound(ColLocation[]))) {
                    DistanceToNextColumn = (DistanceToNextColumn - Column.Width);
                }
                else {
                    DistanceToNextColumn = (DistanceToNextColumn
                                - (Column.Width / 2));
                }

            }

            if (((DistanceToPreviousColumn
                        > (MaxHorizontalDistance * 12))
                        || (DistanceToNextColumn
                        > (MaxHorizontalDistance * 12)))) {
                ColLocation;
                IntColumnsGen(b, (ColNum + 1), true);
                return;
            }

            //     ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
            //     Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
            //     'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
            //         EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
            //         StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
            //         PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
            //         DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
            //         DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
            //         If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
            //             Erase ColLocation
            //             Call IntColumnsGen(b, ColNum + 1)
            //             Exit Sub
            //         End If
        }

    }

    // debugging
    for (i = 0; (i
                <= (ColNum + 1)); i++) {
        Debug.Print;
        ("Column #: "
                    + (i + (", "
                    + ((ColLocation[i] / 12) + ("' from s2, Rafter Line " + i)))));
    }

    // use temporary columns to find sidewall col widths first
    if ((b.rShape == "Single Slope")) {
        // Sidewall 2 Column Width:
        Column = new clsMember();
        Column.Length = (b.bHeight * 12);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[ColNum] - ColLocation[(ColNum + 1)]));
        b.s2ColumnWidth = Column.Width;
        // Sidewall 4 Column Width:
        Column = new clsMember();
        Column.Length = ((b.bWidth * b.rPitch)
                    + (b.bHeight * 12));
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[0] - ColLocation[1]));
        b.s4ColumnWidth = Column.Width;
    }
    else {
        // Sidewall 2 Column Width:
        Column = new clsMember();
        Column.Length = (b.bHeight * 12);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[ColNum] - ColLocation[(ColNum + 1)]));
        b.s2ColumnWidth = Column.Width;
        // Sidewall 4 Column Width:
        Column = new clsMember();
        Column.Length = (b.bHeight * 12);
        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        Abs((ColLocation[0] - ColLocation[1]));
        b.s4ColumnWidth = Column.Width;
    }

    // set column variables, types, sizes, etc.
    for (i = 0; (i
                <= (ColNum + 1)); i++) {
        // find larger distance to neighboring columns to use in lookup tables
        // s2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
        if ((i
                    == (ColNum + 1))) {
            LargerDistance = Abs((ColLocation[i] - ColLocation[(i - 1)]));
        }
        else if ((i == 0)) {
            LargerDistance = Abs((ColLocation[i] - ColLocation[(i + 1)]));
        }
        else {
            LargerDistance = Application.WorksheetFunction.Max(Abs((ColLocation[i] - ColLocation[(i - 1)])), Abs((ColLocation[i] - ColLocation[(i + 1)])));
        }

        Column = new clsMember();
        Column.mType = "Column";
        Column.CL = ColLocation[i];
        Column.LoadBearing = true;
        Column.Qty = RafterNum;
        Column.Placement = ("main rafter line interior column number " + i);
        if ((b.rShape == "Single Slope")) {
            if ((i == 0)) {
                Column.Length = (((b.bWidth * 12)
                            * (b.rPitch / 12))
                            + (b.bHeight * 12));
            }
            else if ((i
                        == (ColNum + 1))) {
                Column.Length = (b.bHeight * 12);
            }
            else {
                Column.Length = b.DistanceToRoof("e1", Column.CL);
            }

        }
        else {
            // Gable
            if ((i == 0)) {
                Column.Length = (b.bHeight * 12);
            }
            else if ((i
                        == (ColNum + 1))) {
                Column.Length = (b.bHeight * 12);
            }
            else {
                Column.Length = b.DistanceToRoof("e1", Column.CL);
            }

        }

        Column.tEdgeHeight = Column.Length;
        Column.SetSize;
        b;
        "Column";
        "Interior";
        LargerDistance;
        if ((Column.CL == 0)) {
            Column.CL = (Column.Width / 2);
        }
        else if ((Column.CL
                    == (b.bWidth * 12))) {
            Column.CL = ((b.bWidth * 12)
                        - (Column.Width / 2));
        }

        Column.rEdgePosition = (Column.CL
                    - (Column.Width / 2));
        Column.Placement = (Column.Size + (" interior column, "
                    + (Column.Length + "' long")));
        b.InteriorColumns.Add;
        Column;
    }

}

// given column along rafterline, return the estimated width of column based on height and distance to nearest columns
MinimumInteriorColumnWidth(b: clsBuilding, ColIndex: number, Columns: number[]): number {
    let ColHeight: number;
    let PrevColDistance: number;
    let NextColDistance: number;
    let MaxColDistance: number;
    let ColumnType: string;
    let WidthCell: Range;
    let HeightCell: Range;
    let Depth: string;
    let Width: string;
    let i: number;
    let LookupTbl: ListObject;
    LookupTbl = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl");
    if ((ColIndex == UBound(Columns))) {
        NextColDistance = 0;
    }
    else {
        NextColDistance = (Columns((ColIndex + 1)) - Columns(ColIndex));
    }

    ColHeight = b.DistanceToRoof("e1", Columns(ColIndex));
    for (WidthCell in LookupTbl.ListColumns(1).DataBodyRange) {
        if (((((NextColDistance / 12)
                    >= WidthCell.Value)
                    && ((NextColDistance / 12)
                    < WidthCell.offset(1, 0).Value))
                    || (NextColDistance < (30 * 12)))) {
            break;
        }

    }

    for (i = 2; (i
                <= (LookupTbl.HeaderRowRange.Cells.Count - 2)); i++) {
        if (((((ColHeight / 12)
                    >= LookupTbl.HeaderRowRange(1, i).Value)
                    && ((ColHeight / 12)
                    < LookupTbl.HeaderRowRange(1, (i + 1)).Value))
                    || ((ColHeight / 12)
                    <= 30))) {
            break;
        }

    }

    ColumnType = WidthCell.offset(0, (i - 1)).Value;
    Width = ColumnType.Substring(0, ((ColumnType.IndexOf("x", 0) + 1)
                    - 1)).Substring((ColumnType.Substring(0, ((ColumnType.IndexOf("x", 0) + 1)
                        - 1)).Length - (ColumnType.Substring(0, ((ColumnType.IndexOf("x", 0) + 1)
                        - 1)).Length - 1)));
    return number.Parse(Width);
}

// function returns nearest endwall location that does not conflict with an OHDoor
NearestEndwallLocation(Location: number, b: clsBuilding, Alternate: string, eWall: string): number {
    let Column: clsMember;
    // Warning!!! Optional parameters not supported
    // Warning!!! Optional parameters not supported
    let e1Column: clsMember;
    let e3Column: clsMember;
    let e1ColLocation: number;
    let e3ColLocation: number;
    let e1tempNearestLocation: number;
    let e3tempNearestLocation: number;
    let FO: clsFO;
    let iterationLocation: number;
    let AlternateEdge: boolean;
    if ((Alternate == "Alternate")) {
        AlternateEdge = true;
    }
    else {
        AlternateEdge = false;
    }

    e1tempNearestLocation = 1.79769313486231E+308;
    // initialize
    e3tempNearestLocation = 1.79769313486231E+308;
    // initialize
    if ((eWall != "e3")) {
        // get closest OHDoor Edge for e1; only called if columns don't work
        for (FO in b.e1FOs) {
            if ((FO.FOType == "OHDoor")) {
                if ((AlternateEdge == false)) {
                    if (((FO.rEdgePosition < Location)
                                && (FO.lEdgePosition > Location))) {
                        if ((Abs((FO.rEdgePosition - Location)) < Abs((FO.lEdgePosition - Location)))) {
                            e1tempNearestLocation = FO.rEdgePosition;
                        }
                        else {
                            e1tempNearestLocation = FO.lEdgePosition;
                        }

                    }

                }
                else if ((AlternateEdge == true)) {
                    if ((((FO.rEdgePosition - 1)
                                < Location)
                                && (FO.lEdgePosition + (1 > Location)))) {
                        if ((Abs((FO.rEdgePosition - Location)) < Abs((FO.lEdgePosition - Location)))) {
                            e1tempNearestLocation = FO.lEdgePosition;
                        }
                        else {
                            e1tempNearestLocation = FO.rEdgePosition;
                        }

                    }

                }

            }

        }

    }

    if ((eWall != "e1")) {
        // get closest OHDoor Edge for e3; only called if columns don't work
        for (FO in b.e3FOs) {
            if ((FO.FOType == "OHDoor")) {
                if ((AlternateEdge == false)) {
                    if (((((b.bWidth * 12)
                                - FO.rEdgePosition)
                                > Location)
                                && (((b.bWidth * 12)
                                - FO.lEdgePosition)
                                < Location))) {
                        if ((Abs((b.bWidth
                                        - (FO.rEdgePosition - Location))) < Abs((b.bWidth
                                        - (FO.lEdgePosition - Location))))) {
                            e3tempNearestLocation = ((b.bWidth * 12)
                                        - FO.rEdgePosition);
                        }
                        else {
                            e3tempNearestLocation = ((b.bWidth * 12)
                                        - FO.lEdgePosition);
                        }

                    }

                }
                else if ((AlternateEdge == true)) {
                    if (((((b.bWidth * 12)
                                - FO.rEdgePosition) + (1 > Location))
                                && (((b.bWidth * 12)
                                - (FO.lEdgePosition - 1))
                                < Location))) {
                        if ((Abs(((b.bWidth * 12)
                                        - (FO.rEdgePosition - Location))) < Abs(((b.bWidth * 12)
                                        - (FO.lEdgePosition - Location))))) {
                            e3tempNearestLocation = ((b.bWidth * 12)
                                        - FO.lEdgePosition);
                        }
                        else {
                            e3tempNearestLocation = ((b.bWidth * 12)
                                        - FO.rEdgePosition);
                        }

                    }

                }

            }

        }

    }

    if ((eWall == "e1")) {
        NearestEndwallLocation = e1tempNearestLocation;
    }
    else if ((eWall == "e3")) {
        NearestEndwallLocation = e3tempNearestLocation;
    }
    else if (ConflictingEndwallOHDoor(e1tempNearestLocation, b)) {
        NearestEndwallLocation = e3tempNearestLocation;
    }
    else if (ConflictingEndwallOHDoor(e3tempNearestLocation, b)) {
        NearestEndwallLocation = e1tempNearestLocation;
    }
    else if ((Abs((e1tempNearestLocation - Location)) < Abs((e3tempNearestLocation - Location)))) {
        NearestEndwallLocation = e1tempNearestLocation;
    }
    else {
        NearestEndwallLocation = e3tempNearestLocation;
    }

    if ((NearestEndwallLocation == 1.79769313486231E+308)) {
        NearestEndwallLocation = Location;
    }

}

// Returns TRUE if a location has matching endwall columns on BOTH ends of the building
MatchingEndwallColumn(Location: number, b: clsBuilding): boolean {
    let Column: clsMember;
    Column = new clsMember();
    let e1MatchingEndwallColumn: boolean;
    let e3MatchingEndwallColumn: boolean;
    e1MatchingEndwallColumn = false;
    e3MatchingEndwallColumn = false;
    for (Column in b.e1Columns) {
        if ((Column.CL == Location)) {
            e1MatchingEndwallColumn = true;
        }

    }

    for (Column in b.e3Columns) {
        if ((Column.CL
                    == ((b.bWidth * 12)
                    - Location))) {
            e3MatchingEndwallColumn = true;
        }

    }

    if (((e1MatchingEndwallColumn == true)
                && (e3MatchingEndwallColumn == true))) {
        MatchingEndwallColumn = true;
    }
    else {
        MatchingEndwallColumn = false;
    }

}
//Returns TRUE if location conflicts with an OHDoor on either Endwall
ConflictingEndwallOHDoor(Location: number, b: clsBuilding, eWall: string): boolean {
    let FO: clsFO;
    // Warning!!! Optional parameters not supported
    let e1Conflict: boolean;
    let e3Conflict: boolean;
    if ((eWall != "e3")) {
        for (FO in b.e1FOs) {
            if ((FO.FOType == "OHDoor")) {
                if (((Location > FO.rEdgePosition)
                            && (Location < FO.lEdgePosition))) {
                    e1Conflict = true;
                }
                else {
                    e1Conflict = false;
                }

            }

        }

    }

    if ((eWall != "e1")) {
        for (FO in b.e3FOs) {
            if ((FO.FOType == "OHDoor")) {
                if (((Location
                            < ((b.bWidth * 12)
                            - FO.rEdgePosition))
                            && (Location
                            > ((b.bWidth * 12)
                            - FO.lEdgePosition)))) {
                    e3Conflict = true;
                }
                else {
                    e3Conflict = false;
                }

            }

        }

    }

    if (((e1Conflict == true)
                || (e3Conflict == true))) {
        ConflictingEndwallOHDoor = true;
    }
    else {
        ConflictingEndwallOHDoor = false;
    }

}
    AdjustSidewallColumns(b: clsBuilding, eWall: string) {
        let ColumnCollection: Collection;
        let IntColumnCollection: Collection;
        let NearestColumn: clsMember;
        let Column: clsMember;
        let IntColumn: clsMember;
        let Index: number;
        let WedgeDistance: number;
        switch (eWall) {
            case "s2":
                ColumnCollection = b.s2Columns;
                IntColumnCollection = b.InteriorColumns;
                for (Index = 1; (Index <= IntColumnCollection.Count); Index++) {
                    IntColumn = IntColumnCollection[Index];
                    if ((IntColumn.CL
                                > ((b.bWidth * 12)
                                - 15))) {
                        // 15 is half the widest possible column
                        NearestColumn = IntColumn;
                        // IntColumnCollection.Remove (Index)
                        break;
                    }

                }

                b.s2ColumnWidth = NearestColumn.Width;
                // Calculate angle cut for s2 columns
                // increase size of each column to account for angle cut
                WedgeDistance = (b.s2ColumnWidth
                            * (b.rPitch / 12));
                for (Column in ColumnCollection) {
                    Column.Size = NearestColumn.Size;
                    Column.Width = NearestColumn.Width;
                    Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                    Column.Length = (Column.Length + WedgeDistance);
                }

                // Calculate angle cut for s2 columns
                WedgeDistance = (b.s2ColumnWidth
                            * (b.rPitch / 12));
                break;
            case "s4":
                ColumnCollection = b.s4Columns;
                IntColumnCollection = b.InteriorColumns;
                for (Index = 1; (Index <= IntColumnCollection.Count); Index++) {
                    IntColumn = IntColumnCollection[Index];
                    if ((IntColumn.CL < 15)) {
                        // 15 is half the widest column possible
                        NearestColumn = IntColumn;
                        // IntColumnCollection.Remove (Index)
                        break;
                    }

                }

                b.s4ColumnWidth = NearestColumn.Width;
                // Calculate angle cut for s2 columns
                // IF GABLE: increase size of each column to account for angle cut
                WedgeDistance = (b.s2ColumnWidth
                            * (b.rPitch / 12));
                for (Column in ColumnCollection) {
                    Column.Size = NearestColumn.Size;
                    Column.Width = NearestColumn.Width;
                    if ((b.rShape == "Gable")) {
                        Column.tEdgeHeight = (Column.tEdgeHeight + WedgeDistance);
                        Column.Length = (Column.Length + WedgeDistance);
                    }

                }

                break;
        }

    }
    DisplayDrawingInfo(Placement: number) {
        // Dim CallingShape As String
        // CallingShape = Application.Caller
        // If CallingShape Like "*" & "Straight Connector" & "*" Then
        //     Exit Sub
        // End If
        // MsgBox "Length: " & Application.Round(CDbl(CallingShape), 2) & """"
        MsgBox;
        ("Length: "
                    + (ImperialMeasurementFormat(Placement) + "'"));
    }













































//DrawDimension 8263 - 8376 of Structural Steel Materials Gen Module in VBA doesn't convert automatically




















































ArrayRemoveDups() {
    let nFirst: number;
    let nLast: number;
    let i: number;
    let item: string;
    let arrTemp: Object;
    let Coll: Collection = new Collection();
    // Get First and Last Array Positions
    nFirst = LBound(MyArray);
    nLast = UBound(MyArray);
    let arrTemp: Object;
    for (i = nFirst; (i <= nLast); i++) {
        arrTemp[i] = MyArray(i);
    }

    // Populate Temporary Collection
    // TODO: On Error Resume Next Warning!!!: The statement is not translatable
    for (i = nFirst; (i <= nLast); i++) {
        Coll.Add;
        arrTemp[i];
        arrTemp[i].ToString();
    }

    Err.Clear;
    // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
    // Resize Array
    nLast = (Coll.Count
                + (nFirst - 1));
    let arrTemp: Object;
    for (i = nFirst; (i <= nLast); i++) {
        arrTemp[i] = Coll[((i - nFirst)
                    + 1)];
    }

    // Output Array
    return arrTemp;
}
DrawItems(b: clsBuilding) {
    let ColumnCollection: Collection;
    let FOCollection: Collection;
    let GirtsCollection: Collection;
    let RafterCollection: Collection;
    let IntColumnCollection: Collection;
    let OverhangCollection: Collection;
    let ExtensionCollection: Collection;
    let Member: clsMember;
    let FO: clsFO;
    let item: Object;
    let eWall: string;
    let TotalHeight: number;
    let MaxHeight: number;
    let lEdgePosition: number;
    let x1: number;
    let y1: number;
    let i: number;
    let ColumnWidth: number;
    let mString: string;
    let Length: string;
    let FloorplanHeight: number;
    let xf: number;
    let yf: number;
    let BayNum: number;
    let BayStart: number;
    let j: number;
    let WeldPlate: clsMiscItem;
    let Plate: clsMiscItem;
    let IntDimensionHeight: number;
    let MyShape: Shape;
    let s2ExtensionEdge: number;
    let s4ExtensionEdge: number;
    let DrawSht: Worksheet;
    let DimensionOffset: number;
    // Call TestDimension(b)
    // delete and remake sheet
    Application.DisplayAlerts = false;
    if (SheetExists("Wall Drawings", ThisWorkbook)) {
        ThisWorkbook.Worksheets("Wall Drawings").Delete;
    }

    Application.DisplayAlerts = true;
    ThisWorkbook.Sheets.Add(ThisWorkbook.Sheets("Structural Steel Price List")).Name = "Wall Drawings";
    // TODO: Labeled Arguments not supported. Argument: 1 := 'After'
    ThisWorkbook.Worksheets("Wall Drawings").Activate;
    ActiveWindow.DisplayGridlines = false;
    ActiveWindow.Zoom = 40;
    DrawSht = ThisWorkbook.ActiveSheet;
    // Set Max Building height
    if ((b.rShape == "Single Slope")) {
        // in
        MaxHeight = ((b.bHeight * 12)
                    + (b.bWidth * b.rPitch));
    }
    else {
        MaxHeight = ((b.bHeight * 12)
                    + (b.bWidth / (2 * b.rPitch)));
    }

    // Set floorplan adjustment
    FloorplanHeight = ((b.bLength * 12)
                + b.e3Extension);
    xf = ((b.bWidth * 12)
                + (b.s4Extension + 350));
    // 350 as buffer
    yf = (FloorplanHeight + 350);
    //  350 as buffer
    DrawSht.Range("A1:A12").RowHeight = ((350
                + (yf + b.e1Extension))
                / 12);
    DrawSht.Range("A1:Z1").ColumnWidth = ((xf + b.s2Extension)
                / 26);
    DrawSht.Range("A13:A22").RowHeight = ((MaxHeight + 350)
                / 10);
    DrawSht.Range("A23:A32").RowHeight = ((MaxHeight + 350)
                / 10);
    DrawSht.Range("A33:A42").RowHeight = ((MaxHeight + 350)
                / 10);
    DrawSht.Range("A43:A52").RowHeight = ((MaxHeight + 350)
                / 10);
    // Draw Floorplan label
    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
    // With...
    // .Name = eWall
    RGB(150, 150, 150).Line.Weight = 3;
    RGB(0, 0, 0).Line.ForeColor.RGB = 3;
    75.Fill.ForeColor.RGB = 3;
    200.Height = 3;
    (yf
                - (FloorplanHeight - 100.Width)) = 3;
    12.5.Top = 3;
    MyShape.Left = 3;
    // With...
    36.HorizontalAlignment = xlVAlignCenter;
    true.Characters.Font.Size = xlVAlignCenter;
    "Floorplan".Characters.Font.Bold = xlVAlignCenter;
    MyShape.TextFrame.Characters.Text = xlVAlignCenter;
    // draw floorplan outline
    IntDimensionHeight = 0;
    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
    // With...
    msoSendToBack.Fill.Transparency = 0.5;
    (b.bWidth * 12.Line.Transparency) = RGB(200, 200, 200).ZOrder;
    (b.bLength * 12).Width = RGB(200, 200, 200).ZOrder;
    (yf
                - (b.bLength * 12.Height)) = RGB(200, 200, 200).ZOrder;
    (xf
                - (b.bWidth * 12.Top)) = RGB(200, 200, 200).ZOrder;
    MyShape.Left = RGB(200, 200, 200).ZOrder;
    // e1 Wall
    MyShape = DrawSht.Shapes.AddLine(xf, yf, (((b.bWidth * 12)
                    * -1)
                    + xf), yf);
    // With...
    MyShape.Line.ForeColor.RGB = 0.5;
    // e1 Dimension
    if (((b.e1Extension
                + (b.e1Overhang > 80))
                && (b.e1Extension
                + (b.e1Overhang < 150)))) {
        DimensionOffset = -75;
    }

    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
    // With...
    0.5.Fill.Transparency = 1;
    RGB(255, 255, 255).Line.Weight = 1;
    RGB(255, 255, 255).Line.ForeColor.RGB = 1;
    (b.bWidth * 12.Fill.ForeColor.RGB) = 1;
    100.Width = 1;
    (yf + (100 + DimensionOffset).Height) = 1;
    (((b.bWidth * 12)
                * -1)
                + xf.Top) = 1;
    MyShape.Left = 1;
    // With...
    xlHAlignCenter.VerticalAlignment = RGB(0, 20, 132);
    24.HorizontalAlignment = RGB(0, 20, 132);
    true.Characters.Font.Size = RGB(0, 20, 132);
    ("Endwall 1 " + ("
" + ImperialMeasurementFormat((b.bWidth * 12)).Characters.Font.Bold)) = RGB(0, 20, 132);
    MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
    // vertical dimension lines
    MyShape = DrawSht.Shapes.AddLine(xf, (yf + (100 + DimensionOffset)), xf, (yf + (150 + DimensionOffset)));
    // With...
    RGB(0, 20, 132).Weight = msoLineDash;
    MyShape.Line.ForeColor.RGB = msoLineDash;
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + xf), (yf + (100 + DimensionOffset)), (((b.bWidth * 12)
                    * -1)
                    + xf), (yf + (150 + DimensionOffset)));
    // With...
    RGB(0, 20, 132).Weight = msoLineDash;
    MyShape.Line.ForeColor.RGB = msoLineDash;
    // horizontal dimension lines
    MyShape = DrawSht.Shapes.AddLine(xf, (yf + (125 + DimensionOffset)), (xf
                    - (b.bWidth * (12 / 3))), (yf + (125 + DimensionOffset)));
    // With...
    msoArrowheadLong.BeginArrowheadStyle = msoArrowheadWide;
    msoLineDash.BeginArrowheadLength = msoArrowheadWide;
    0.5.DashStyle = msoArrowheadWide;
    RGB(0, 20, 132).Weight = msoArrowheadWide;
    MyShape.Line.ForeColor.RGB = msoArrowheadWide;
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + xf), (yf + (125 + DimensionOffset)), (((b.bWidth * 12)
                    * -1)
                    + (xf
                    + (b.bWidth * (12 / 3)))), (yf + (125 + DimensionOffset)));
    // With...
    msoArrowheadLong.BeginArrowheadStyle = msoArrowheadWide;
    msoLineDash.BeginArrowheadLength = msoArrowheadWide;
    0.5.DashStyle = msoArrowheadWide;
    RGB(0, 20, 132).Weight = msoArrowheadWide;
    MyShape.Line.ForeColor.RGB = msoArrowheadWide;
    // e3 Wall
    MyShape = DrawSht.Shapes.AddLine(xf, (yf
                    - (b.bLength * 12)), (((b.bWidth * 12)
                    * -1)
                    + xf), (yf
                    - (b.bLength * 12)));
    // With...
    MyShape.Line.ForeColor.RGB = 0.5;
    // s4 Wall
    MyShape = DrawSht.Shapes.AddLine(xf, yf, xf, (yf
                    - (b.bLength * 12)));
    // With...
    MyShape.Line.ForeColor.RGB = 0.5;
    // s2 Dimension
    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
    // With...
    0.5.Fill.Transparency = 1;
    RGB(255, 255, 255).Line.Weight = 1;
    RGB(255, 255, 255).Line.ForeColor.RGB = 1;
    190.Fill.ForeColor.RGB = 1;
    (b.bLength * 12.Width) = 1;
    (yf
                - (b.bLength * 12.Height)) = 1;
    (((b.bWidth * 12)
                * -1)
                + (xf - 300.Top)) = 1;
    MyShape.Left = 1;
    // With...
    xlHAlignRight.VerticalAlignment = RGB(0, 20, 132);
    24.HorizontalAlignment = RGB(0, 20, 132);
    true.Characters.Font.Size = RGB(0, 20, 132);
    ("Sidewall 2 " + ("
" + ImperialMeasurementFormat((b.bLength * 12)).Characters.Font.Bold)) = RGB(0, 20, 132);
    MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
    // vertical dimension lines
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + (xf - 150)), yf, (((b.bWidth * 12)
                    * -1)
                    + (xf - 150)), (yf
                    - (b.bLength * (12 / 3))));
    // With...
    msoArrowheadLong.BeginArrowheadStyle = msoArrowheadWide;
    msoLineDash.BeginArrowheadLength = msoArrowheadWide;
    0.5.DashStyle = msoArrowheadWide;
    RGB(0, 20, 132).Weight = msoArrowheadWide;
    MyShape.Line.ForeColor.RGB = msoArrowheadWide;
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + (xf - 150)), (((b.bLength * 12)
                    * -1)
                    + yf), (((b.bWidth * 12)
                    * -1)
                    + (xf - 150)), (((b.bLength * 12)
                    * -1)
                    + (yf
                    + (b.bLength * (12 / 3)))));
    // With...
    msoArrowheadLong.BeginArrowheadStyle = msoArrowheadWide;
    msoLineDash.BeginArrowheadLength = msoArrowheadWide;
    0.5.DashStyle = msoArrowheadWide;
    RGB(0, 20, 132).Weight = msoArrowheadWide;
    MyShape.Line.ForeColor.RGB = msoArrowheadWide;
    // horizontal dimension lines
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + (xf - 175)), yf, (((b.bWidth * 12)
                    * -1)
                    + (xf - 125)), yf);
    // With...
    RGB(0, 20, 132).Weight = msoLineDash;
    MyShape.Line.ForeColor.RGB = msoLineDash;
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + (xf - 175)), (((b.bLength * 12)
                    * -1)
                    + yf), (((b.bWidth * 12)
                    * -1)
                    + (xf - 125)), (((b.bLength * 12)
                    * -1)
                    + yf));
    // With...
    RGB(0, 20, 132).Weight = msoLineDash;
    MyShape.Line.ForeColor.RGB = msoLineDash;
    // s2 Wall
    MyShape = DrawSht.Shapes.AddLine((((b.bWidth * 12)
                    * -1)
                    + xf), yf, (((b.bWidth * 12)
                    * -1)
                    + xf), (yf
                    - (b.bLength * 12)));
    // With...
    MyShape.Line.ForeColor.RGB = 0.5;
    // get Extension start/end points
    for (Member in b.e1Columns) {
        if ((Member.CL < 0)) {
            s4ExtensionEdge = (Member.rEdgePosition * -1);
        }
        else if ((Member.CL
                    > (b.bWidth * 12))) {
            s2ExtensionEdge = (Member.lEdgePosition
                        - (b.bWidth * 12));
        }

    }

    // Interior Columns
    if ((b.InteriorColumns.Count > 0)) {
        let Bay1Start: number;
        IntColumnCollection = b.InteriorColumns;
        BayNum = (EstSht.Range("BayNum").Value - 1);
        BayStart = EstSht.Range("Bay1_Length").Value;
        Bay1Start = BayStart;
        let DimensionsArr: Object;
        let DimensionsArr: Object;
        for (i = IntColumnCollection.Count; (i <= 1); i = (i + -1)) {
            Member = IntColumnCollection[i];
            for (j = 1; (j <= BayNum); j++) {
                if ((j == 1)) {
                    BayStart = EstSht.Range("Bay1_Length").Value;
                }
                else {
                    BayStart = (BayStart + EstSht.Range("Bay1_Length").offset((j - 1), 0).Value);
                }

                MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                // With...
                if (Member.Size) {
                    "*TS*";
                    (yf - ((BayStart * 12)
                                + 2).Height) = 4;
                    Top = 4;
                }
                else {
                    (yf - ((BayStart * 12)
                                + 4).Height) = 8;
                    Top = 8;
                }

                RGB(0, 0, 0).Line.Weight = 1;
                RGB(0, 0, 0).Line.ForeColor.RGB = 1;
                Member.Width.Fill.ForeColor.RGB = 1;
                Width = 1;
                if (((Member.CL < 0)
                            || (Member.CL
                            > (b.bWidth * 12)))) {
                    MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0);
                    MyShape.Line.ForeColor.RGB = RGB(0, 230, 0);
                }

                // Draw Weld Plate
                for (WeldPlate in Member.ComponentMembers) {
                    if ((WeldPlate.clsType == "Weld Plate")) {
                        Plate = WeldPlate;
                        break;
                    }

                }

                // label Weld Plate
                MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                // With...
                75.Fill.Transparency = 1;
                25.Width = 1;
                ((yf
                            - (BayStart * 12))
                            + 15.Height) = 1;
                (xf
                            - (Member.rEdgePosition + Member.Width)).Top = 1;
                MyShape.Left = 1;
                // With...
                xlHAlignRight.VerticalAlignment = RGB(0, 20, 132);
                14.HorizontalAlignment = RGB(0, 20, 132);
                true.Characters.Font.Size = RGB(0, 20, 132);
                (Plate.Width + ("""x"
                            + (Plate.Height + """".Characters.Font.Bold))) = RGB(0, 20, 132);
                MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
                if ((j == 1)) {
                    if (((b.InteriorColumns(i).CL < 15)
                                && (b.InteriorColumns(i).CL > 0))) {
                        // ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                        DimensionsArr[(i - 1)] = Member.rEdgePosition;
                    }
                    else if (((b.InteriorColumns(i).CL
                                > ((b.bWidth * 12)
                                - 15))
                                && (b.InteriorColumns(i).CL
                                < (b.bWidth * 12)))) {
                        // ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                        DimensionsArr[(i - 1)] = Member.lEdgePosition;
                    }
                    else if ((b.InteriorColumns(i).CL < 0)) {
                        // DimensionsArr(i - 1) = -Member.rEdgePosition + b.bWidth * 12
                        s4ExtensionEdge = (Member.rEdgePosition * -1);
                    }
                    else if ((b.InteriorColumns(i).CL
                                > (b.bWidth * 12))) {
                        // DimensionsArr(i - 1) = -(Member.lEdgePosition - b.bWidth * 12)
                        s2ExtensionEdge = (Member.lEdgePosition
                                    - (b.bWidth * 12));
                    }
                    else {
                        // ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                        DimensionsArr[(i - 1)] = ((b.bWidth * 12)
                                    - b.InteriorColumns(i).CL);
                    }

                }

            }

        }

        // Interior column dimension
        DimensionsArr = ArrayRemoveDups(DimensionsArr);
        DrawDimension(b, (xf
                        - (b.bWidth * 12)), (yf
                        - ((Bay1Start * 12)
                        + 50)), (b.bWidth * 12), 50, "Horizontal", 18, DimensionsArr);
    }

    // e1 column dimensions
    let DimensionsArr: Object;
    for (i = b.e1Columns.Count; (i <= 1); i = (i + -1)) {
        if (((b.e1Columns(i).CL < 15)
                    && (b.e1Columns(i).CL > 0))) {
            DimensionsArr[(i - 1)] = 0;
        }
        else if (((b.e1Columns(i).CL
                    > ((b.bWidth * 12)
                    - 15))
                    && (b.e1Columns(i).CL
                    < (b.bWidth * 12)))) {
            DimensionsArr[(i - 1)] = (b.bWidth * 12);
        }
        else if ((b.e1Columns(i).CL < 0)) {
            DimensionsArr[(i - 1)] = ((b.e1Columns(i).rEdgePosition * -1)
                        + (b.bWidth * 12));
        }
        else if ((b.e1Columns(i).CL
                    > (b.bWidth * 12))) {
            DimensionsArr[(i - 1)] = ((b.e1Columns(i).lEdgePosition
                        - (b.bWidth * 12))
                        * -1);
        }
        else if (b.e1Columns(i).mType) {
            "*Extension*";
            // do not add
        }
        else {
            DimensionsArr[(i - 1)] = ((b.bWidth * 12)
                        - b.e1Columns(i).CL);
        }

    }

    for (i = b.e1FOs.Count; (i <= 1); i = (i + -1)) {
        if ((b.e1FOs(i).FOType == "OHDoor")) {
            let Preserve: Object;
            DimensionsArr[(UBound(DimensionsArr) + 2)];
            DimensionsArr[(UBound(DimensionsArr) - 1)] = ((b.bWidth * 12)
                        - (b.e1FOs(i).rEdgePosition + b.e1FOs(i).Width));
            DimensionsArr[UBound(DimensionsArr)] = ((b.bWidth * 12)
                        - b.e1FOs(i).rEdgePosition);
        }

    }

    // if s2 extension, add extension width to all CLs
    if ((b.s2Extension > 0)) {
        for (i = 0; (i <= UBound(DimensionsArr)); i++) {
            DimensionsArr[i] = (DimensionsArr[i] + s2ExtensionEdge);
        }

    }

    DimensionsArr = ArrayRemoveDups(DimensionsArr);
    DrawDimension(b, (xf
                    - ((b.bWidth * 12)
                    - s2ExtensionEdge)), (yf
                    + (b.e1Extension
                    + (b.e1Overhang + 25))), ((b.bWidth * 12)
                    + (s2ExtensionEdge + s4ExtensionEdge)), 50, "Horizontal", 18, DimensionsArr);
    // e3 column dimensions
    let DimensionsArr: Object;
    for (i = b.e3Columns.Count; (i <= 1); i = (i + -1)) {
        if (((b.e3Columns(i).CL < 15)
                    && (b.e3Columns(i).CL > 0))) {
            DimensionsArr[(i - 1)] = 0;
        }
        else if (((b.e3Columns(i).CL
                    > ((b.bWidth * 12)
                    - 15))
                    && (b.e3Columns(i).CL
                    < (b.bWidth * 12)))) {
            DimensionsArr[(i - 1)] = (b.bWidth * 12);
        }
        else if ((b.e3Columns(i).CL < 0)) {
            DimensionsArr[(i - 1)] = b.e3Columns(i).rEdgePosition;
        }
        else if ((b.e3Columns(i).CL
                    > (b.bWidth * 12))) {
            DimensionsArr[(i - 1)] = b.e3Columns(i).lEdgePosition;
        }
        else if (b.e3Columns(i).mType) {
            "*Extension*";
            // do not add
        }
        else {
            DimensionsArr[(i - 1)] = b.e3Columns(i).CL;
        }

    }

    for (i = b.e3FOs.Count; (i <= 1); i = (i + -1)) {
        if ((b.e3FOs(i).FOType == "OHDoor")) {
            let Preserve: Object;
            DimensionsArr[(UBound(DimensionsArr) + 2)];
            DimensionsArr[(UBound(DimensionsArr) - 1)] = (b.e3FOs(i).rEdgePosition + b.e3FOs(i).Width);
            DimensionsArr[UBound(DimensionsArr)] = b.e3FOs(i).rEdgePosition;
        }

    }

    // if s2 extension, add extension width to all CLs
    if ((b.s2ExtensionWidth > 0)) {
        for (i = 0; (i <= UBound(DimensionsArr)); i++) {
            DimensionsArr[i] = (DimensionsArr[i] + s2ExtensionEdge);
        }

    }

    DimensionsArr = ArrayRemoveDups(DimensionsArr);
    DrawDimension(b, (xf
                    - ((b.bWidth * 12)
                    - s2ExtensionEdge)), (yf
                    - ((b.bLength * 12)
                    - (b.e3Extension
                    - (b.e3Overhang - 75)))), ((b.bWidth * 12)
                    + (s2ExtensionEdge + s4ExtensionEdge)), 50, "Horizontal", 18, DimensionsArr);
    // s4 column dimensions
    let DimIndex: number;
    let BayTotal: number;
    i = 1;
    DimIndex = EstSht.Range("BayNum").Value;
    let DimensionsArr: Object;
    if ((b.e1Extension > 0)) {
        let Preserve: Object;
        DimensionsArr[(UBound(DimensionsArr) + 1)];
        DimensionsArr[UBound(DimensionsArr)] = ((b.bLength * 12)
                    + b.e1Extension);
    }

    for (i = i; (i <= EstSht.Range("BayNum").Value); i++) {
        BayTotal = (BayTotal
                    + (EstSht.Range("Bay1_Length").offset((i - 1), 0).Value * 12));
        DimensionsArr[(i - 1)] = ((b.bLength * 12)
                    - BayTotal);
    }

    DimensionsArr[EstSht.Range("BayNum").Value] = (b.bLength * 12);
    for (FO in b.s4FOs) {
        if ((FO.FOType == "OHDoor")) {
            let Preserve: Object;
            DimensionsArr[(UBound(DimensionsArr) + 2)];
            DimensionsArr[(UBound(DimensionsArr) - 1)] = (FO.rEdgePosition + FO.Width);
            DimensionsArr[UBound(DimensionsArr)] = FO.rEdgePosition;
        }

    }

    if ((b.e3Extension > 0)) {
        let Preserve: Object;
        DimensionsArr[(UBound(DimensionsArr) + 1)];
        DimensionsArr[UBound(DimensionsArr)] = (b.e3Extension * -1);
        for (i = 0; (i <= UBound(DimensionsArr)); i++) {
            DimensionsArr[i] = (DimensionsArr[i] + b.e3Extension);
        }

    }

    DrawDimension(b, (xf
                    + (b.s4Extension
                    + (b.s4Overhang + 75))), (yf
                    - (b.e3Extension
                    - (b.bLength * 12))), 50, ((b.bLength * 12)
                    + (b.e1Extension + b.e3Extension)), "Vertical", 18, DimensionsArr);
    // s2 column dimensions
    BayTotal = 0;
    i = 1;
    DimIndex = EstSht.Range("BayNum").Value;
    let DimensionsArr: Object;
    if ((b.e1Extension > 0)) {
        let Preserve: Object;
        DimensionsArr[(UBound(DimensionsArr) + 1)];
        DimensionsArr[UBound(DimensionsArr)] = ((b.bLength * 12)
                    + b.e1Extension);
    }

    for (i = 1; (i <= EstSht.Range("BayNum").Value); i++) {
        BayTotal = (BayTotal
                    + (EstSht.Range("Bay1_Length").offset((i - 1), 0).Value * 12));
        DimensionsArr[(i - 1)] = ((b.bLength * 12)
                    - BayTotal);
    }

    DimensionsArr[EstSht.Range("BayNum").Value] = (b.bLength * 12);
    for (FO in b.s2FOs) {
        if ((FO.FOType == "OHDoor")) {
            let Preserve: Object;
            DimensionsArr[(UBound(DimensionsArr) + 2)];
            DimensionsArr[(UBound(DimensionsArr) - 1)] = (((b.bLength * 12)
                        - FO.lEdgePosition)
                        + FO.Width);
            DimensionsArr[UBound(DimensionsArr)] = ((b.bLength * 12)
                        - FO.lEdgePosition);
        }

    }

    if ((b.e3Extension > 0)) {
        let Preserve: Object;
        DimensionsArr[(UBound(DimensionsArr) + 1)];
        DimensionsArr[UBound(DimensionsArr)] = (b.e3Extension * -1);
        for (i = 0; (i <= UBound(DimensionsArr)); i++) {
            DimensionsArr[i] = (DimensionsArr[i] + b.e3Extension);
        }

    }

    // Call DrawDimension(b, xf - b.bWidth * 12 - b.s2Extension - b.s2Overhang - 75, yf - b.e3Extension - (b.bLength * 12), 50, b.bLength * 12 + b.e1Extension + b.e3Extension, "Vertical", 18, DimensionsArr)
    // ''''''''''''''''''''Extension and Overhang Shaded Areas
    // e1 Extension
    if ((b.e1Extension > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightVertical.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.bWidth * 12.Fill.ForeColor.RGB) = 1.Fill.Patterned;
        b.e1Extension.Width = 1.Fill.Patterned;
        yf.Height = 1.Fill.Patterned;
        (((b.bWidth * 12)
                    * -1)
                    + xf.Top) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e1 Overhang
    if ((b.e1Overhang > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightUpwardDiagonal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.bWidth * 12.Fill.ForeColor.RGB) = 1.Fill.Patterned;
        b.e1Overhang.Width = 1.Fill.Patterned;
        (yf + b.e1Extension.Height) = 1.Fill.Patterned;
        (((b.bWidth * 12)
                    * -1)
                    + xf.Top) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e3 Extension
    if ((b.e3Extension > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightVertical.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.bWidth * 12.Fill.ForeColor.RGB) = 1.Fill.Patterned;
        b.e3Extension.Width = 1.Fill.Patterned;
        (yf
                    - ((b.bLength * 12)
                    - b.e3Extension.Height)) = 1.Fill.Patterned;
        (((b.bWidth * 12)
                    * -1)
                    + xf.Top) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e3 Overhang
    if ((b.e3Overhang > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightUpwardDiagonal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.bWidth * 12.Fill.ForeColor.RGB) = 1.Fill.Patterned;
        b.e3Overhang.Width = 1.Fill.Patterned;
        (yf
                    - ((b.bLength * 12)
                    - (b.e3Extension - b.e3Overhang.Height))) = 1.Fill.Patterned;
        (((b.bWidth * 12)
                    * -1)
                    + xf.Top) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // s2 Extension
    if ((b.s2ExtensionWidth > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        b.s2ExtensionWidth.Fill.ForeColor.RGB = 1.Fill.Patterned;
        (b.bLength * 12.Width) = 1.Fill.Patterned;
        (yf - (b.bLength * 12).Height) = 1.Fill.Patterned;
        (xf
                    - ((b.bWidth * 12)
                    - b.s2ExtensionWidth.Top)) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // s2 Overhang
    if ((b.s2Overhang > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightUpwardDiagonal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        b.s2Overhang.Fill.ForeColor.RGB = 1.Fill.Patterned;
        (b.bLength * 12.Width) = 1.Fill.Patterned;
        (yf - (b.bLength * 12).Height) = 1.Fill.Patterned;
        (xf
                    - ((b.bWidth * 12)
                    - (b.s2ExtensionWidth - b.s2Overhang.Top))) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // s4 Extension
    if (((b.s4ExtensionWidth * -1)
                > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.s4ExtensionWidth.Fill.ForeColor.RGB * -1) = 1.Fill.Patterned;
        (b.bLength * 12.Width) = 1.Fill.Patterned;
        (yf - (b.bLength * 12).Height) = 1.Fill.Patterned;
        xf.Top = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // s4 Overhang
    if ((b.s4Overhang > 0)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightUpwardDiagonal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        b.s4Overhang.Fill.ForeColor.RGB = 1.Fill.Patterned;
        (b.bLength * 12.Width) = 1.Fill.Patterned;
        (yf - (b.bLength * 12).Height) = 1.Fill.Patterned;
        (xf
                    + (b.s4ExtensionWidth.Top * -1)) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e1s2 Extension Intersection
    if ((b.s2e1ExtensionIntersection == true)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        b.s2ExtensionWidth.Fill.ForeColor.RGB = 1.Fill.Patterned;
        b.e1Extension.Width = 1.Fill.Patterned;
        yf.Height = 1.Fill.Patterned;
        (xf
                    - ((b.bWidth * 12)
                    - b.s2ExtensionWidth.Top)) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e1s4 Extension Intersection
    if ((b.s4e1ExtensionIntersection == true)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.s4ExtensionWidth.Fill.ForeColor.RGB * -1) = 1.Fill.Patterned;
        b.e1Extension.Width = 1.Fill.Patterned;
        yf.Height = 1.Fill.Patterned;
        xf.Top = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e3s2 Extension Intersection
    if ((b.s2e3ExtensionIntersection == true)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        b.s2ExtensionWidth.Fill.ForeColor.RGB = 1.Fill.Patterned;
        b.e3Extension.Width = 1.Fill.Patterned;
        (yf
                    - ((b.bLength * 12)
                    - b.e3Extension.Height)) = 1.Fill.Patterned;
        (xf
                    - ((b.bWidth * 12)
                    - b.s2ExtensionWidth.Top)) = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    // e3s4 Extension Intersection
    if ((b.s4e3ExtensionIntersection == true)) {
        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        msoPatternLightHorizontal.ZOrder;
        0.5.Fill.Transparency = 1.Fill.Patterned;
        RGB(255, 255, 255).Line.Weight = 1.Fill.Patterned;
        RGB(0, 230, 0).Line.ForeColor.RGB = 1.Fill.Patterned;
        (b.s4ExtensionWidth.Fill.ForeColor.RGB * -1) = 1.Fill.Patterned;
        b.e3Extension.Width = 1.Fill.Patterned;
        (yf
                    - ((b.bLength * 12)
                    - b.e3Extension.Height)) = 1.Fill.Patterned;
        xf.Top = 1.Fill.Patterned;
        MyShape.Left = 1.Fill.Patterned;
        msoSendToBack;
    }

    for (i = 1; (i <= 4); i++) {
        if ((i == 1)) {
            eWall = "e1";
        }
        else if ((i == 2)) {
            eWall = "s2";
        }
        else if ((i == 3)) {
            eWall = "e3";
        }
        else if ((i == 4)) {
            eWall = "s4";
        }

        switch (eWall) {
            case "e1":
                ColumnCollection = b.e1Columns;
                FOCollection = b.e1FOs;
                GirtsCollection = b.e1Girts;
                RafterCollection = b.e1Rafters;
                OverhangCollection = b.e1OverhangMembers;
                ExtensionCollection = b.e1ExtensionMembers;
                x1 = (((b.bWidth * 12)
                            + 350)
                            + b.s2Extension);
                y1 = ((350
                            + (yf + b.e1Extension))
                            + (MaxHeight + 350));
                break;
            case "s2":
                ColumnCollection = b.s2Columns;
                FOCollection = b.s2FOs;
                GirtsCollection = b.s2Girts;
                OverhangCollection = b.s2OverhangMembers;
                ExtensionCollection = b.s2ExtensionMembers;
                x1 = (((b.bLength * 12)
                            + 350)
                            + b.s2Extension);
                y1 = (y1
                            + (MaxHeight + 350));
                break;
            case "e3":
                ColumnCollection = b.e3Columns;
                FOCollection = b.e3FOs;
                GirtsCollection = b.e3Girts;
                RafterCollection = b.e3Rafters;
                OverhangCollection = b.e3OverhangMembers;
                ExtensionCollection = b.e3ExtensionMembers;
                x1 = (((b.bWidth * 12)
                            + 350)
                            + b.s2Extension);
                y1 = (y1
                            + (MaxHeight + 350));
                break;
            case "s4":
                ColumnCollection = b.s4Columns;
                FOCollection = b.s4FOs;
                GirtsCollection = b.s4Girts;
                OverhangCollection = b.s4OverhangMembers;
                ExtensionCollection = b.s4ExtensionMembers;
                x1 = (((b.bLength * 12)
                            + 350)
                            + b.s2Extension);
                y1 = (y1
                            + (MaxHeight + 350));
                break;
        }

        // get highest point of wall
        if (((eWall == "s2")
                    || ((eWall == "s4")
                    && (b.rShape == "Gable")))) {
            TotalHeight = (b.bHeight * 12);
        }
        else if (((eWall == "s4")
                    && (b.rShape == "Single Slope"))) {
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth * b.rPitch));
        }
        else if ((b.rShape == "Single Slope")) {
            // in
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth * b.rPitch));
        }
        else {
            TotalHeight = ((b.bHeight * 12)
                        + (b.bWidth / (2 * b.rPitch)));
        }

        MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
        // With...
        // .Name = eWall
        RGB(150, 150, 150).Line.Weight = 3;
        RGB(0, 0, 0).Line.ForeColor.RGB = 3;
        75.Fill.ForeColor.RGB = 3;
        75.Height = 3;
        (y1 - MaxHeight.Width) = 3;
        12.5.Top = 3;
        MyShape.Left = 3;
        // With...
        36.HorizontalAlignment = xlVAlignCenter;
        true.Characters.Font.Size = xlVAlignCenter;
        eWall.Characters.Font.Bold = xlVAlignCenter;
        MyShape.TextFrame.Characters.Text = xlVAlignCenter;
        for (Member in ColumnCollection) {
            if (((eWall == "e1")
                        || (eWall == "e3"))) {
                MyShape = DrawSht.Shapes.AddLine(((Member.CL * -1)
                                + x1), ((Member.bEdgeHeight * -1)
                                + y1), ((Member.CL * -1)
                                + x1), ((Member.tEdgeHeight * -1)
                                + y1));
                // With...
                if (Member.Placement) {
                    ("*Extension*" | Member.Placement);
                    "*Overhang*";
                    RGB(0, 230, 0).Transparency = 0.4;
                    MyShape.Line.ForeColor.RGB = 0.4;
                }
                else {
                    MyShape.Line.ForeColor.RGB = RGB(75, 75, 75);
                }

                MyShape.Line.Weight = Member.Width;
                MyShape.Select;
                if ((Member.Length != 0)) {
                    // Selection.Name = Member.Placement
                    MyShape.OnAction = ("'DisplayDrawingInfo "
                                + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                }

                let ExtensionLength: number;
                // Floorplan Columns
                if ((eWall == "e1")) {
                    if ((Member.mType == "e1 Extension Column")) {
                        ExtensionLength = b.e1Extension;
                    }
                    else {
                        ExtensionLength = 0;
                    }

                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    if (Member.Size) {
                        "*TS*";
                        ((yf - 4)
                                    + ExtensionLength.Height) = 4;
                        Top = 4;
                    }
                    else {
                        ((yf - 8)
                                    + ExtensionLength.Height) = 8;
                        Top = 8;
                    }

                    RGB(0, 0, 0).Line.Weight = 1;
                    RGB(0, 0, 0).Line.ForeColor.RGB = 1;
                    Member.Width.Fill.ForeColor.RGB = 1;
                    Width = 1;
                    if (((Member.CL < 0)
                                || ((Member.CL
                                > (b.bWidth * 12))
                                || (Member.mType == "e1 Extension Column")))) {
                        MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0);
                        MyShape.Line.ForeColor.RGB = RGB(0, 230, 0);
                    }

                    // Draw Dimension
                    // get Weld Plate
                    for (WeldPlate in Member.ComponentMembers) {
                        if ((WeldPlate.clsType == "Weld Plate")) {
                            Plate = WeldPlate;
                            break;
                        }

                    }

                    // label Weld Plate
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    75.Fill.Transparency = 1;
                    25.Width = 1;
                    (yf + (15 + ExtensionLength.Height)) = 1;
                    (xf
                                - (Member.rEdgePosition + Member.Width)).Top = 1;
                    MyShape.Left = 1;
                    // With...
                    xlHAlignRight.VerticalAlignment = RGB(0, 20, 132);
                    14.HorizontalAlignment = RGB(0, 20, 132);
                    true.Characters.Font.Size = RGB(0, 20, 132);
                    (Plate.Width + ("""x"
                                + (Plate.Height + """".Characters.Font.Bold))) = RGB(0, 20, 132);
                    MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
                }
                else if ((eWall == "e3")) {
                    if ((Member.mType == "e3 Extension Column")) {
                        ExtensionLength = b.e3Extension;
                    }
                    else {
                        ExtensionLength = 0;
                    }

                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    if (Member.Size) {
                        "*TS*";
                        MyShape.Left = (yf
                                    - ((b.bLength * 12)
                                    - ExtensionLength));
                        Height = 4;
                    }
                    else {
                        Height = 8;
                    }

                    RGB(150, 150, 150).Line.Weight = 1;
                    RGB(0, 0, 0).Line.ForeColor.RGB = 1;
                    Member.Width.Fill.ForeColor.RGB = 1;
                    Width = 1;
                    if (((Member.CL < 0)
                                || ((Member.CL
                                > (b.bWidth * 12))
                                || (Member.mType == "e3 Extension Column")))) {
                        MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0);
                        MyShape.Line.ForeColor.RGB = RGB(0, 230, 0);
                    }

                    // get Weld Plate
                    for (WeldPlate in Member.ComponentMembers) {
                        if ((WeldPlate.clsType == "Weld Plate")) {
                            Plate = WeldPlate;
                            break;
                        }

                    }

                    // label Weld Plate
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    75.Fill.Transparency = 1;
                    25.Width = 1;
                    ((yf
                                - (b.bLength * 12)) + (15 - ExtensionLength.Height)) = 1;
                    (xf - ((b.bWidth * 12)
                                - Member.rEdgePosition).Top) = 1;
                    MyShape.Left = 1;
                    // With...
                    xlHAlignRight.VerticalAlignment = RGB(0, 20, 132);
                    14.HorizontalAlignment = RGB(0, 20, 132);
                    true.Characters.Font.Size = RGB(0, 20, 132);
                    (Plate.Width + ("""x"
                                + (Plate.Height + """".Characters.Font.Bold))) = RGB(0, 20, 132);
                    MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
                }

            }
            else {
                MyShape = DrawSht.Shapes.AddLine(((Member.CL * -1)
                                + x1), ((Member.bEdgeHeight * -1)
                                + y1), ((Member.CL * -1)
                                + x1), ((Member.tEdgeHeight * -1)
                                + y1));
                // With...
                if (Member.Placement) {
                    ("*Extension*" | Member.Placement);
                    "*Overhang*";
                    RGB(0, 230, 0).Transparency = 0.4;
                    MyShape.Line.ForeColor.RGB = 0.4;
                }
                else {
                    MyShape.Line.ForeColor.RGB = RGB(75, 75, 75);
                }

                MyShape.Line.Weight = Member.Width;
                MyShape.Select;
                if ((Member.Length != 0)) {
                    // Selection.Name = Member.Placement
                }

                MyShape.OnAction = ("'DisplayDrawingInfo "
                            + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                ColumnWidth = Member.Width;
                if ((eWall == "s2")) {
                    // Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                    // With MyShape
                    //     .Left = xf - (b.bWidth * 12)
                    //     .Top = yf - Member.CL - Member.Width
                    //     .Height = Member.Width * 2
                    //     .Width = Member.Width * 2
                    //     .Fill.ForeColor.RGB = RGB(0, 0, 0)
                    //     .Line.ForeColor.RGB = RGB(150, 150, 150)
                    //     .Line.Weight = 1
                    // End With
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    0.5.Line.DashStyle = 0.4;
                    RGB(0, 0, 0).Line.Weight = 0.4;
                    RGB(230, 0, 0).Line.ForeColor.RGB = 0.4;
                    (b.bWidth * 12.Fill.ForeColor.RGB) = 0.4;
                    (Member.Width * 2.Width) = 0.4;
                    (yf
                                - (Member.CL - Member.Width.Height)) = 0.4;
                    (xf - (b.bWidth * 12).Top) = 0.4;
                    MyShape.Left = 0.4;
                    // With...
                    2.HorizontalAlignment = xlVAlignCenter;
                    16.Characters.Font.ColorIndex = xlVAlignCenter;
                    true.Characters.Font.Size = xlVAlignCenter;
                    "Main Rafter Line".Characters.Font.Bold = xlVAlignCenter;
                    MyShape.TextFrame.Characters.Text = xlVAlignCenter;
                }
                else if ((eWall == "s4")) {
                    // Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                    // With MyShape
                    //     .Left = xf - Member.Width * 2
                    //     .Top = yf - b.bLength * 12 + Member.CL - Member.Width
                    //     .Height = Member.Width * 2
                    //     .Width = Member.Width * 2
                    //     .Fill.ForeColor.RGB = RGB(0, 0, 0)
                    //     .Line.ForeColor.RGB = RGB(150, 150, 150)
                    //     .Line.Weight = 1
                    // End With
                }

            }

        }

        // endwall columns on s2 and s4
        if ((eWall == "s2")) {
            // With...
            DrawSht.Shapes.AddLine((((b.bLength * 12)
                            * -1)
                            + x1), (0 + y1), (((b.bLength * 12)
                            * -1)
                            + x1), ((TotalHeight * -1)
                            + y1)).Line.ForeColor.RGB = b.e3Columns(1).Width;
            // With...
            DrawSht.Shapes.AddLine((((0 * 12)
                            * -1)
                            + x1), (0 + y1), (0 + x1), ((TotalHeight * -1)
                            + y1)).Line.ForeColor.RGB = b.e1Columns(1).Width;
        }
        else if ((eWall == "s4")) {
            // With...
            DrawSht.Shapes.AddLine((((b.bLength * 12)
                            * -1)
                            + x1), (0 + y1), (((b.bLength * 12)
                            * -1)
                            + x1), ((TotalHeight * -1)
                            + y1)).Line.ForeColor.RGB = b.e1Columns(1).Width;
            // With...
            DrawSht.Shapes.AddLine((((0 * 12)
                            * -1)
                            + x1), (0 + y1), (0 + x1), ((TotalHeight * -1)
                            + y1)).Line.ForeColor.RGB = b.e3Columns(1).Width;
        }

        for (Member in GirtsCollection) {
            MyShape = DrawSht.Shapes.AddLine(((Member.rEdgePosition * -1)
                            + x1), ((Member.bEdgeHeight * -1)
                            + y1), (((Member.rEdgePosition - Member.Length)
                            * -1)
                            + x1), ((Member.tEdgeHeight * -1)
                            + y1));
            // With...
            if (((eWall == "e1")
                        || (eWall == "e3"))) {
                if (((Member.rEdgePosition < 0)
                            || (Member.rEdgePosition
                            + (Member.Length
                            > (b.bWidth * 12))))) {
                    MyShape.Line.ForeColor.RGB = RGB(0, 230, 0);
                }
                else {
                    MyShape.Line.ForeColor.RGB = RGB(150, 150, 150);
                }

            }
            else if (((Member.rEdgePosition < 0)
                        || (Member.rEdgePosition
                        + (Member.Length
                        > (b.bLength * 12))))) {
                MyShape.Line.ForeColor.RGB = RGB(0, 230, 0);
            }
            else {
                MyShape.Line.ForeColor.RGB = RGB(150, 150, 150);
            }

            MyShape.Line.Weight = 2.5;
            if ((Member.bEdgeHeight == 86)) {
                MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                // With...
                msoSendToBack;
                RGB(255, 255, 255).Line.Weight = 1.ZOrder;
                RGB(255, 255, 255).Line.ForeColor.RGB = 1.ZOrder;
                Member.Length.Fill.ForeColor.RGB = 1.ZOrder;
                50.Width = 1.ZOrder;
                ((Member.bEdgeHeight * -1)
                            + (y1 - 50.Height)) = 1.ZOrder;
                (((Member.rEdgePosition - Member.Length)
                            * -1)
                            + x1.Top) = 1.ZOrder;
                MyShape.Left = 1.ZOrder;
                // With...
                xlHAlignCenter.VerticalAlignment = RGB(0, 20, 132);
                24.HorizontalAlignment = RGB(0, 20, 132);
                true.Characters.Font.Size = RGB(0, 20, 132);
                ImperialMeasurementFormat(Member.Length).Characters.Font.Bold = RGB(0, 20, 132);
                MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
            }
            else {
                MyShape.Select;
                Length = ImperialMeasurementFormat(Member.Length);
                // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                MyShape.OnAction = ("'DisplayDrawingInfo "
                            + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
            }

        }

        for (FO in FOCollection) {
            // With...
            DrawSht.Shapes.AddLine(((FO.rEdgePosition * -1)
                            + x1), ((FO.bEdgeHeight * -1)
                            + y1), (((FO.rEdgePosition - FO.Width)
                            * -1)
                            + x1), ((FO.bEdgeHeight * -1)
                            + y1)).Line.ForeColor.RGB = 2.5;
            // With...
            DrawSht.Shapes.AddLine((((FO.rEdgePosition - FO.Width)
                            * -1)
                            + x1), ((FO.bEdgeHeight * -1)
                            + y1), (((FO.rEdgePosition - FO.Width)
                            * -1)
                            + x1), ((FO.tEdgeHeight * -1)
                            + y1)).Line.ForeColor.RGB = 2.5;
            // With...
            DrawSht.Shapes.AddLine((((FO.rEdgePosition - FO.Width)
                            * -1)
                            + x1), ((FO.tEdgeHeight * -1)
                            + y1), ((FO.rEdgePosition * -1)
                            + x1), ((FO.tEdgeHeight * -1)
                            + y1)).Line.ForeColor.RGB = 2.5;
            // With...
            DrawSht.Shapes.AddLine(((FO.rEdgePosition * -1)
                            + x1), ((FO.bEdgeHeight * -1)
                            + y1), ((FO.rEdgePosition * -1)
                            + x1), ((FO.tEdgeHeight * -1)
                            + y1)).Line.ForeColor.RGB = 2.5;
            // Draw Dimension (width x height) of FO
            MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
            // With...
            // TODO: # ... Warning!!! not translated
            RGB(255, 255, 255).Line.Weight = 0;
            RGB(255, 255, 255).Line.ForeColor.RGB = 0;
            FO.Width.Fill.ForeColor.RGB = 0;
            FO.Height.Width = 0;
            ((FO.tEdgeHeight * -1)
                        + y1.Height) = 0;
            (((FO.rEdgePosition - FO.Width)
                        * -1)
                        + x1.Top) = 0;
            MyShape.Left = 0;
            ZOrder;
            msoSendToBack;
            // With...
            xlHAlignCenter.VerticalAlignment = RGB(0, 20, 132);
            18.HorizontalAlignment = RGB(0, 20, 132);
            true.Characters.Font.Size = RGB(0, 20, 132);
            ("W"
                        + (ImperialMeasurementFormat(FO.Width) + (" x H" + ImperialMeasurementFormat(FO.Height).Characters.Font.Bold))) = RGB(0, 20, 132);
            MyShape.TextFrame.Characters.Text = RGB(0, 20, 132);
            // Floorplan View
            if ((eWall == "e1")) {
                MyShape = DrawSht.Shapes.AddLine((xf - FO.rEdgePosition), yf, (xf - FO.lEdgePosition), yf);
                // With...
                MyShape.Line.ForeColor.RGB = 5;
                if ((FO.FOType == "OHDoor")) {
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    RGB(240, 240, 0).Line.Weight = msoLineDash.ZOrder(msoSendToBack);
                    1.Line.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    RGB(255, 255, 255).Fill.Transparency = msoLineDash.ZOrder(msoSendToBack);
                    FO.Width.Fill.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    FO.Height.Width = msoLineDash.ZOrder(msoSendToBack);
                    (xf - FO.lEdgePosition.Height) = msoLineDash.ZOrder(msoSendToBack);
                    (yf - FO.Height.Left) = msoLineDash.ZOrder(msoSendToBack);
                    MyShape.Top = msoLineDash.ZOrder(msoSendToBack);
                    // With...
                    1.HorizontalAlignment = xlVAlignTop;
                    16.Characters.Font.ColorIndex = xlVAlignTop;
                    true.Characters.Font.Size = xlVAlignTop;
                    (FO.FOType + ("
"
                                + (ImperialMeasurementFormat(FO.Width) + ("x" + ImperialMeasurementFormat(FO.Height).Characters.Font.Bold)))) = xlVAlignTop;
                    MyShape.TextFrame.Characters.Text = xlVAlignTop;
                    // Draw Dimension
                }

            }
            else if ((eWall == "e3")) {
                MyShape = DrawSht.Shapes.AddLine((xf
                                - ((b.bWidth * 12)
                                - FO.rEdgePosition)), (yf
                                - (b.bLength * 12)), (xf
                                - ((b.bWidth * 12)
                                - FO.lEdgePosition)), (yf
                                - (b.bLength * 12)));
                // With...
                MyShape.Line.ForeColor.RGB = 5;
                if ((FO.FOType == "OHDoor")) {
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    RGB(240, 240, 0).Line.Weight = msoLineDash.ZOrder(msoSendToBack);
                    1.Line.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    RGB(255, 255, 255).Fill.Transparency = msoLineDash.ZOrder(msoSendToBack);
                    FO.Width.Fill.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    FO.Height.Width = msoLineDash.ZOrder(msoSendToBack);
                    (xf - ((b.bWidth * 12)
                                - FO.rEdgePosition).Height) = msoLineDash.ZOrder(msoSendToBack);
                    (yf
                                - (b.bLength * 12.Left)) = msoLineDash.ZOrder(msoSendToBack);
                    MyShape.Top = msoLineDash.ZOrder(msoSendToBack);
                    // With...
                    1.HorizontalAlignment = xlVAlignBottom;
                    16.Characters.Font.ColorIndex = xlVAlignBottom;
                    true.Characters.Font.Size = xlVAlignBottom;
                    (FO.FOType + ("
"
                                + (ImperialMeasurementFormat(FO.Width) + ("x" + ImperialMeasurementFormat(FO.Height).Characters.Font.Bold)))) = xlVAlignBottom;
                    MyShape.TextFrame.Characters.Text = xlVAlignBottom;
                }

            }
            else if ((eWall == "s2")) {
                MyShape = DrawSht.Shapes.AddLine((xf
                                - (b.bWidth * 12)), (yf - FO.rEdgePosition), (xf
                                - (b.bWidth * 12)), (yf - FO.lEdgePosition));
                // With...
                MyShape.Line.ForeColor.RGB = 5;
                if ((FO.FOType == "OHDoor")) {
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    RGB(240, 240, 0).Line.Weight = msoLineDash.ZOrder(msoSendToBack);
                    1.Line.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    RGB(255, 255, 255).Fill.Transparency = msoLineDash.ZOrder(msoSendToBack);
                    FO.Height.Fill.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    FO.Width.Width = msoLineDash.ZOrder(msoSendToBack);
                    (((b.bWidth * 12)
                                * -1)
                                + xf.Height) = msoLineDash.ZOrder(msoSendToBack);
                    (yf - FO.lEdgePosition.Left) = msoLineDash.ZOrder(msoSendToBack);
                    MyShape.Top = msoLineDash.ZOrder(msoSendToBack);
                    // With...
                    1.HorizontalAlignment = xlVAlignCenter;
                    16.Characters.Font.ColorIndex = xlVAlignCenter;
                    true.Characters.Font.Size = xlVAlignCenter;
                    (FO.FOType + ("
"
                                + (ImperialMeasurementFormat(FO.Width) + ("x" + ImperialMeasurementFormat(FO.Height).Characters.Font.Bold)))) = xlVAlignCenter;
                    MyShape.TextFrame.Characters.Text = xlVAlignCenter;
                }

            }
            else if ((eWall == "s4")) {
                MyShape = DrawSht.Shapes.AddLine(xf, ((yf
                                - (b.bLength * 12))
                                + FO.rEdgePosition), xf, ((yf
                                - (b.bLength * 12))
                                + FO.lEdgePosition));
                // With...
                MyShape.Line.ForeColor.RGB = 5;
                if ((FO.FOType == "OHDoor")) {
                    MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1);
                    // With...
                    RGB(240, 240, 0).Line.Weight = msoLineDash.ZOrder(msoSendToBack);
                    1.Line.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    RGB(255, 255, 255).Fill.Transparency = msoLineDash.ZOrder(msoSendToBack);
                    FO.Height.Fill.ForeColor.RGB = msoLineDash.ZOrder(msoSendToBack);
                    FO.Width.Width = msoLineDash.ZOrder(msoSendToBack);
                    (xf - FO.Height.Height) = msoLineDash.ZOrder(msoSendToBack);
                    ((yf
                                - (b.bLength * 12))
                                + FO.rEdgePosition.Left) = msoLineDash.ZOrder(msoSendToBack);
                    MyShape.Top = msoLineDash.ZOrder(msoSendToBack);
                    // With...
                    1.HorizontalAlignment = xlVAlignCenter;
                    16.Characters.Font.ColorIndex = xlVAlignCenter;
                    true.Characters.Font.Size = xlVAlignCenter;
                    (FO.FOType + ("
"
                                + (ImperialMeasurementFormat(FO.Width) + ("x" + ImperialMeasurementFormat(FO.Height).Characters.Font.Bold)))) = xlVAlignCenter;
                    MyShape.TextFrame.Characters.Text = xlVAlignCenter;
                }

            }

            for (item in FO.FOMaterials) {
                if ((item.clsType == "Member")) {
                    Member = item;
                    if ((Member.CL != 0)) {
                        MyShape = DrawSht.Shapes.AddLine(((Member.CL * -1)
                                        + x1), ((Member.bEdgeHeight * -1)
                                        + y1), ((Member.CL * -1)
                                        + x1), ((Member.tEdgeHeight * -1)
                                        + y1));
                        // With...
                        MyShape.Line.ForeColor.RGB = 2.5;
                        MyShape.OnAction = ("'DisplayDrawingInfo "
                                    + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                    }

                }

            }

        }

        if (((eWall == "e1")
                    || (eWall == "e3"))) {
            for (Member in RafterCollection) {
                if (((Member.rEdgePosition
                            < (b.bWidth * (12 / 2)))
                            && (b.rShape == "Gable"))) {
                    // b = c * cos(a)
                    // a = atn(b.rPitch/12)
                    lEdgePosition = (Member.rEdgePosition
                                + ((Member.tEdgeHeight - Member.bEdgeHeight)
                                / (b.rPitch * 12)));
                    lEdgePosition = Member.RafterLeftEdge;
                    MyShape = DrawSht.Shapes.AddLine(((Member.rEdgePosition * -1)
                                    + x1), ((Member.bEdgeHeight * -1)
                                    + y1), ((lEdgePosition * -1)
                                    + x1), ((Member.tEdgeHeight * -1)
                                    + y1));
                    // With...
                    if (Member.Placement) {
                        ("*Extension*" | Member.Placement);
                        "*Overhang*";
                        RGB(0, 230, 0).Transparency = 0.4;
                        MyShape.Line.ForeColor.RGB = 0.4;
                    }
                    else {
                        MyShape.Line.ForeColor.RGB = RGB(75, 75, 75);
                    }

                    MyShape.Line.Weight = Member.Width;
                    // .DashStyle = msoLineDashDotDot
                    MyShape.Select;
                    // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                    MyShape.OnAction = ("'DisplayDrawingInfo "
                                + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                }
                else if (((b.rShape == "Gable")
                            || ((b.rShape == "Single Slope")
                            && (eWall == "e1")))) {
                    lEdgePosition = (Member.rEdgePosition
                                + (Abs((Member.tEdgeHeight - Member.bEdgeHeight))
                                / (b.rPitch * 12)));
                    lEdgePosition = Member.RafterLeftEdge;
                    MyShape = DrawSht.Shapes.AddLine(((lEdgePosition * -1)
                                    + x1), ((Member.bEdgeHeight * -1)
                                    + y1), ((Member.rEdgePosition * -1)
                                    + x1), ((Member.tEdgeHeight * -1)
                                    + y1));
                    // With...
                    if (Member.Placement) {
                        ("*Extension*" | Member.Placement);
                        "*Overhang*";
                        RGB(0, 230, 0).Transparency = 0.4;
                        MyShape.Line.ForeColor.RGB = 0.4;
                    }
                    else {
                        MyShape.Line.ForeColor.RGB = RGB(75, 75, 75);
                    }

                    MyShape.Line.Weight = Member.Width;
                    // .DashStyle = msoLineDashDotDot
                    MyShape.Select;
                    // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                    MyShape.OnAction = ("'DisplayDrawingInfo "
                                + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                }
                else if ((b.rShape == "Single Slope")) {
                    lEdgePosition = (Member.rEdgePosition
                                + ((Member.tEdgeHeight - Member.bEdgeHeight)
                                / (b.rPitch * 12)));
                    lEdgePosition = Member.RafterLeftEdge;
                    MyShape = DrawSht.Shapes.AddLine(((Member.rEdgePosition * -1)
                                    + x1), ((Member.bEdgeHeight * -1)
                                    + y1), ((lEdgePosition * -1)
                                    + x1), ((Member.tEdgeHeight * -1)
                                    + y1));
                    // With...
                    if (Member.Placement) {
                        ("*Extension*" | Member.Placement);
                        "*Overhang*";
                        RGB(0, 230, 0).Transparency = 0.4;
                        MyShape.Line.ForeColor.RGB = 0.4;
                    }
                    else {
                        MyShape.Line.ForeColor.RGB = RGB(75, 75, 75);
                    }

                    MyShape.Line.Weight = Member.Width;
                    // .DashStyle = msoLineDashDotDot
                    MyShape.Select;
                    // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                    MyShape.OnAction = ("'DisplayDrawingInfo "
                                + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                }

            }

            // Interior Columns
            if ((b.InteriorColumns.Count > 0)) {
                for (Member in IntColumnCollection) {
                    if ((eWall == "e1")) {
                        MyShape = DrawSht.Shapes.AddLine(((Member.CL * -1)
                                        + x1), ((Member.bEdgeHeight * -1)
                                        + y1), ((Member.CL * -1)
                                        + x1), ((Member.tEdgeHeight * -1)
                                        + y1));
                        // With...
                        if (Member.Placement) {
                            ("*Extension*" | Member.Placement);
                            "*Overhang*";
                            RGB(230, 0, 0).Transparency = 0.4;
                            MyShape.Line.ForeColor.RGB = 0.4;
                        }
                        else {
                            RGB(230, 0, 0).Transparency = 0.4;
                            MyShape.Line.ForeColor.RGB = 0.4;
                        }

                        Member.Width.DashStyle = msoLineDash;
                        MyShape.Line.Weight = msoLineDash;
                        MyShape.ZOrder;
                        msoSendToBack;
                        MyShape.Select;
                        // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                        MyShape.OnAction = ("'DisplayDrawingInfo "
                                    + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                    }
                    else {
                        MyShape = DrawSht.Shapes.AddLine(((((b.bWidth * 12)
                                        - Member.CL)
                                        * -1)
                                        + x1), ((Member.bEdgeHeight * -1)
                                        + y1), ((((b.bWidth * 12)
                                        - Member.CL)
                                        * -1)
                                        + x1), ((Member.tEdgeHeight * -1)
                                        + y1));
                        // With...
                        if (Member.Placement) {
                            ("*Extension*" | Member.Placement);
                            "*Overhang*";
                            RGB(230, 0, 0).Transparency = 0.4;
                            MyShape.Line.ForeColor.RGB = 0.4;
                        }
                        else {
                            RGB(230, 0, 0).Transparency = 0.4;
                            MyShape.Line.ForeColor.RGB = 0.4;
                        }

                        Member.Width.DashStyle = msoLineDash;
                        MyShape.Line.Weight = msoLineDash;
                        MyShape.ZOrder;
                        msoSendToBack;
                        MyShape.Select;
                        // If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                        MyShape.OnAction = ("'DisplayDrawingInfo "
                                    + (Application.WorksheetFunction.Round(Member.Length, 4) + "'"));
                    }

                }

            }

        }

    }

    Debug.Print;
    "drawing done";
    DrawSht.PageSetup.FitToPagesWide = 1;
    Application.PrintCommunication = false;
    // With...
    "".FirstPage.CenterFooter.Text = "";
    "".FirstPage.LeftFooter.Text = "";
    "".FirstPage.RightHeader.Text = "";
    "".FirstPage.CenterHeader.Text = "";
    "".FirstPage.LeftHeader.Text = "";
    "".EvenPage.RightFooter.Text = "";
    "".EvenPage.CenterFooter.Text = "";
    "".EvenPage.LeftFooter.Text = "";
    "".EvenPage.RightHeader.Text = "";
    "".EvenPage.CenterHeader.Text = "";
    true.EvenPage.LeftHeader.Text = "";
    true.AlignMarginsHeaderFooter = "";
    false.ScaleWithDocHeaderFooter = "";
    false.DifferentFirstPageHeaderFooter = "";
    xlPrintErrorsDisplayed.OddAndEvenPagesHeaderFooter = "";
    0.PrintErrors = "";
    1.FitToPagesTall = "";
    false.FitToPagesWide = "";
    false.Zoom = "";
    xlDownThenOver.BlackAndWhite = "";
    xlAutomatic.Order = "";
    xlPaperLetter.FirstPageNumber = "";
    false.PaperSize = "";
    xlPortrait.Draft = "";
    false.Orientation = "";
    false.CenterVertically = "";
    600.CenterHorizontally = "";
    xlPrintNoComments.PrintQuality = "";
    false.PrintComments = "";
    false.PrintGridlines = "";
    Application.InchesToPoints(0.2).PrintHeadings = "";
    Application.InchesToPoints(0.2).FooterMargin = "";
    Application.InchesToPoints(0.5).HeaderMargin = "";
    Application.InchesToPoints(0.5).BottomMargin = "";
    Application.InchesToPoints(0.25).TopMargin = "";
    Application.InchesToPoints(0.25).RightMargin = "";
    "".LeftMargin = "";
    "".RightFooter = "";
    "".CenterFooter = "";
    "".LeftFooter = "";
    "".RightHeader = "";
    "".CenterHeader = "";
    DrawSht.PageSetup.LeftHeader = "";
    // TODO: On Error Resume Next Warning!!!: The statement is not translatable
    Application.PrintCommunication = true;
    DrawSht.ResetAllPageBreaks;
    DrawSht.HPageBreaks.Add;
    /* Warning! Labeled Statements are not Implemented */DrawSht.Rows[14];
    DrawSht.Rows[14].PageBreak = xlPageBreakManual;
    DrawSht.HPageBreaks.Add;
    /* Warning! Labeled Statements are not Implemented */DrawSht.Rows[24];
    DrawSht.Rows[24].PageBreak = xlPageBreakManual;
    DrawSht.HPageBreaks.Add;
    /* Warning! Labeled Statements are not Implemented */DrawSht.Rows[34];
    DrawSht.Rows[34].PageBreak = xlPageBreakManual;
    DrawSht.HPageBreaks.Add;
    /* Warning! Labeled Statements are not Implemented */DrawSht.Rows[45];
    DrawSht.Rows[34].PageBreak = xlPageBreakManual;
    DrawSht.HPageBreaks.Add;
    /* Warning! Labeled Statements are not Implemented */DrawSht.Rows[56];
    DrawSht.Rows[34].PageBreak = xlPageBreakManual;
    EstSht.Activate;
}
