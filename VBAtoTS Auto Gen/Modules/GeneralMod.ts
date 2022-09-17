// TODO: Option Explicit ... Warning!!! not translated


EventsEnable() {
    Application.EnableEvents = true;
}

ImportOpen(ImportFile: Workbook) {
    let FileToOpen: Object;
    // User selects file to import
    FileToOpen = Application.GetOpenFilename("Browse for your File  & Import Data", "Excel Files (*.xlsx),*.xlsx");
    // TODO: Labeled Arguments not supported. Argument: 1 := 'Title'
    // TODO: Labeled Arguments not supported. Argument: 2 := 'FileFilter'
    // check valid file
    if ((FileToOpen != false)) {
        ImportFile = Application.Workbooks.Open(FileToOpen);
        // Set Sourcews = Sourcewb.Worksheets(1)
    }
    else {
        return;
    }

}

ImportOHDoorPriceSht() {
    let ImportFile: Workbook;
    let ImportSht: Worksheet;
    let ImportTbl: ListObject;
    let SectionalOHDoorPriceTbl: ListObject;
    let mCell: Range;
    let i: number;
    let j: number;
    let Vendor: string;
    Application.ScreenUpdating = false;
    ImportOpen(ImportFile);
    if ((ImportFile == null)) {
        return;
    }

    ImportSht = ImportFile.Worksheets(1);
    ImportTbl = ImportSht.ListObjects(1);
    SectionalOHDoorPriceTbl = MasterPriceSht.ListObjects("SectionalOHDoorPriceTbl");
    MasterPriceSht.Unprotect;
    "WhiteTruckMafia";
    ImportSht.Unprotect;
    "WhiteTruckMafia";
    SectionalOHDoorPriceTbl.AutoFilter.ShowAllData;
    MasterPriceSht.UsedRange.Rows.Hidden = false;
    ImportTbl.AutoFilter.ShowAllData;
    if ((ImportTbl.DataBodyRange.Cells.Count != SectionalOHDoorPriceTbl.DataBodyRange.Cells.Count)) {
        MsgBox;
        "The import data table does not match the Sectional OH Door Price Table, please enter the new data manually or choose a different file.";
        return;
    }

    for (i = 1; (i <= ImportTbl.ListColumns.Count); i++) {
        if ((i != 1)) {
            for (j = 1; (j <= ImportTbl.ListRows.Count); j++) {
                mCell = ImportTbl.DataBodyRange(j, i);
                if (((mCell.Value != "-")
                            && ((mCell.Value != 0)
                            && (mCell.EntireRow.Hidden == false)))) {
                    Debug.Print;
                    mCell.Value;
                    Debug.Print;
                    mCell.Address;
                    SectionalOHDoorPriceTbl.DataBodyRange(j, i).Value = mCell.Value;
                }

            }

        }

    }

    SectionalOHDoorPriceTbl.AutoFilter.ShowAllData;
    MasterPriceSht.Protect;
    "WhiteTruckMafia";
    ImportSht.Protect;
    "WhiteTruckMafia";
    ImportFile.Close;
    false;
    Application.ScreenUpdating = true;
}

ImportSheetMetalPriceSht() {
    let ImportFile: Workbook;
    let ImportSht: Worksheet;
    let ImportTbl: ListObject;
    let MasterPriceTbl: ListObject;
    let mCell: Range;
    let i: number;
    let j: number;
    let Vendor: string;
    Application.ScreenUpdating = false;
    ImportOpen(ImportFile);
    if ((ImportFile == null)) {
        return;
    }

    ImportSht = ImportFile.Worksheets(1);
    ImportTbl = ImportSht.ListObjects(1);
    MasterPriceTbl = MasterPriceSht.ListObjects("MasterPriceTbl");
    MasterPriceTbl.AutoFilter.ShowAllData;
    ImportTbl.AutoFilter.ShowAllData;
    if ((ImportTbl.DataBodyRange.Cells.Count != MasterPriceTbl.DataBodyRange.Cells.Count)) {
        MsgBox;
        "The import data table does not match the Master Price Table, please enter the new data manually or choose a different file.";
        return;
    }

    Vendor = ImportSht.Range("B1").Value;
    MasterPriceSht.Unprotect;
    "WhiteTruckMafia";
    ImportSht.Unprotect;
    "WhiteTruckMafia";
    MasterPriceTbl.DataBodyRange.AutoFilter;
    MasterPriceTbl.ListColumns.Count;
    Vendor;
    ImportTbl.DataBodyRange.AutoFilter;
    MasterPriceTbl.ListColumns.Count;
    Vendor;
    for (i = 1; (i <= ImportTbl.ListColumns.Count); i++) {
        if (((i != 1)
                    && ((i != 2)
                    && (i < 11)))) {
            for (j = 1; (j <= ImportTbl.ListRows.Count); j++) {
                mCell = ImportTbl.DataBodyRange(j, i);
                if (((mCell.Value != "-")
                            && ((mCell.Value != 0)
                            && (mCell.EntireRow.Hidden == false)))) {
                    Debug.Print;
                    mCell.Value;
                    Debug.Print;
                    mCell.Address;
                    MasterPriceTbl.DataBodyRange(j, i).Value = mCell.Value;
                }

            }

        }

    }

    MasterPriceTbl.AutoFilter.ShowAllData;
    MasterPriceSht.Protect;
    "WhiteTruckMafia";
    ImportSht.Protect;
    "WhiteTruckMafia";
    ImportFile.Close;
    false;
    Application.ScreenUpdating = true;
}

CreateSheetMetalPriceExport() {
    Application.ScreenUpdating = false;
    let MasterPriceTbl: ListObject;
    let ExportSht: Worksheet;
    let NewName: string;
    let NewPriceSht: Worksheet;
    let NewExportFile: Workbook;
    let FirstCol: number;
    let LastCol: number;
    let FirstRow: number;
    let i: number;
    let j: number;
    let mCell: Range;
    let Vendor: string;
    ExportSht = MasterPriceSht;
    Vendor = ExportSht.Range("SelectedVendor").Value;
    NewName = ("SheetMetalPriceSht_"
                + (Vendor + ("_" + Format(Date, "mmddyy"))));
    SaveAs(NewName, ExportSht);
    if ((Dir((ThisWorkbook.path + ("\"
                    + (NewName + ".xlsx")))) == "")) {
        return;
    }

    NewExportFile = Workbooks((NewName + ".xlsx"));
    NewPriceSht = NewExportFile.Worksheets(1);
    MasterPriceTbl = NewPriceSht.ListObjects("MasterPriceTbl");
    FirstCol = (MasterPriceTbl.ListColumns(1).DataBodyRange.Column - 2);
    NewPriceSht.Unprotect("WhiteTruckMafia");
    // delete columns to left
    for (i = 1; (i <= FirstCol); i++) {
        NewPriceSht.Columns[1].Delete;
    }

    // delete columns to right
    // With...
    LastCol = MasterPriceTbl.ListColumns;
    (MasterPriceTbl.ListColumns.Count.DataBodyRange.Column + 2);
    for (i = LastCol; (i <= NewPriceSht.UsedRange.SpecialCells(xlCellTypeVisible).Columns.Count); i++) {
        NewPriceSht.Columns[LastCol].Delete;
    }

    // delete rows above
    for (i = 1; (i
                <= (MasterPriceTbl.HeaderRowRange.Row - 5)); i++) {
        NewPriceSht.Rows[1].Delete;
    }

    NewPriceSht.UsedRange.Rows.Hidden = false;
    NewPriceSht.Range("A1").Value = "Prepared For:";
    NewPriceSht.Range("B1").Value = Vendor;
    NewPriceSht.Range("A1").HorizontalAlignment = xlRight;
    NewPriceSht.Range("B1").HorizontalAlignment = xlLeft;
    NewPriceSht.Range("A1:B1").Locked = true;
    MasterPriceTbl.DataBodyRange.Locked = true;
    MasterPriceTbl.HeaderRowRange.Locked = true;
    for (i = 1; (i <= MasterPriceTbl.ListColumns.Count); i++) {
        if (((i != 1)
                    && ((i != 2)
                    && (i < 11)))) {
            for (j = 1; (j <= MasterPriceTbl.ListRows.Count); j++) {
                mCell = MasterPriceTbl.DataBodyRange(j, i);
                if (((mCell.Value != "-")
                            && (mCell.Value != 0))) {
                    mCell.Locked = false;
                }
                else {
                    mCell.Locked = true;
                }

            }

        }

    }

    MasterPriceTbl.AutoFilter.ShowAllData;
    MasterPriceTbl.DataBodyRange.AutoFilter;
    MasterPriceTbl.ListColumns.Count;
    Vendor;
    NewPriceSht.Protect("WhiteTruckMafia");
    NewExportFile.Close;
    true;
    Application.ScreenUpdating = true;
}

CreateOHDoorPriceExport() {
    Application.ScreenUpdating = false;
    let SectionalOHDoorPriceTbl: ListObject;
    let ExportSht: Worksheet;
    let NewName: string;
    let NewPriceSht: Worksheet;
    let NewExportFile: Workbook;
    let FirstCol: number;
    let LastCol: number;
    let FirstRow: number;
    let i: number;
    ExportSht = MasterPriceSht;
    NewName = ("OHDoorPriceSht_" + Format(Date, "mmddyy"));
    SaveAs(NewName, ExportSht);
    if ((Dir((ThisWorkbook.path + ("\"
                    + (NewName + ".xlsx")))) == "")) {
        return;
    }

    NewExportFile = Workbooks((NewName + ".xlsx"));
    NewPriceSht = NewExportFile.Worksheets(1);
    SectionalOHDoorPriceTbl = NewPriceSht.ListObjects("SectionalOHDoorPriceTbl");
    FirstCol = (SectionalOHDoorPriceTbl.ListColumns(1).DataBodyRange.Column - 2);
    NewPriceSht.Unprotect("WhiteTruckMafia");
    // delete columns to left
    for (i = 1; (i <= FirstCol); i++) {
        NewPriceSht.Columns[1].Delete;
    }

    // delete columns to right
    // With...
    LastCol = SectionalOHDoorPriceTbl.ListColumns;
    (SectionalOHDoorPriceTbl.ListColumns.Count.DataBodyRange.Column + 2);
    for (i = LastCol; (i <= NewPriceSht.UsedRange.SpecialCells(xlCellTypeVisible).Columns.Count); i++) {
        NewPriceSht.Columns[LastCol].Delete;
    }

    // delete rows above
    for (i = 1; (i
                <= (SectionalOHDoorPriceTbl.HeaderRowRange.Row - 5)); i++) {
        NewPriceSht.Rows[1].Delete;
    }

    NewPriceSht.UsedRange.Rows.Hidden = false;
    SectionalOHDoorPriceTbl.DataBodyRange.Locked = false;
    NewPriceSht.Protect("WhiteTruckMafia");
    NewExportFile.Close;
    true;
    Application.ScreenUpdating = true;
}

SaveAs(NewName: string, CopySht: Worksheet) {
    let FName: string;
    let NewBook: Workbook;
    Application.DisplayAlerts = false;
    FName = (ThisWorkbook.path + ("\"
                + (NewName + ".xlsx")));
    NewBook = Workbooks.Add;
    CopySht.Copy;
    /* Warning! Labeled Statements are not Implemented */NewBook.Sheets(1);
    if ((Dir(FName) != "")) {
        MsgBox;
        ("File "
                    + (FName + " already exists. To save a new version, delete the existing file and try again."));
        NewBook.Close;
        false;
        Workbooks.Open(FName);
    }
    else {
        NewBook.SaveAs;
        /* Warning! Labeled Statements are not Implemented */FName;
    }

    Application.DisplayAlerts = true;
}

PrintFloorplan() {
    let DrawSht: Worksheet;
    let LastRow: number;
    DrawSht = ThisWorkbook.Worksheets("Wall Drawings");
    Application.DisplayAlerts = false;
    let Length: number;
    let Width: number;
    let Zoom: number;
    Length = (EstSht.Range("Building_Length")
                + (EstSht.Range("e1_GableExtension").Value + EstSht.Range("e3_GableExtension").Value));
    Width = (EstSht.Range("Building_Width").Value
                + (EstSht.Range("s2_EaveExtension").Value + EstSht.Range("s4_EaveExtension").Value));
    if (((Length <= 40)
                && (Width <= 80))) {
        Zoom = 50;
    }
    else if (((Length <= 70)
                && (Width <= 120))) {
        Zoom = 35;
    }
    else {
        Zoom = 25;
    }

    // Drawings (first page is landscape)
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    xlLandscape.CenterHorizontally = Application.InchesToPoints(0.2);
    Zoom.Orientation = Application.InchesToPoints(0.2);
    0.Zoom = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    DrawSht.PageSetup.FitToPagesWide = Application.InchesToPoints(0.2);
    Application.PrintCommunication = true;
    DrawSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    /* Warning! Labeled Statements are not Implemented */1;
    1;
    Application.DisplayAlerts = true;
}

PrintManagerPackage() {
    let CostSht: Worksheet;
    let LastRow: number;
    PrintEmployeePackage();
    CostSht = ThisWorkbook.Worksheets("Cost Estimate");
    Application.DisplayAlerts = false;
    LastRow = CostSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:F" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    CostSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = false;
    CostSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    Application.DisplayAlerts = true;
}

PrintDescription() {
    let DescriptionSht: Worksheet;
    let LastRow: number;
    DescriptionSht = ThisWorkbook.Worksheets("Project Description");
    Application.DisplayAlerts = false;
    LastRow = DescriptionSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:A" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    DescriptionSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = false;
    DescriptionSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    Application.DisplayAlerts = true;
}

PrintEmployeePackage() {
    let DescriptionSht: Worksheet;
    let SheetMetalSht: Worksheet;
    let SteelSht: Worksheet;
    let DrawSht: Worksheet;
    let MiscMSht: Worksheet;
    let LastRow: number;
    DescriptionSht = ThisWorkbook.Worksheets("Project Description");
    SheetMetalSht = ThisWorkbook.Worksheets("Employee Materials List");
    SteelSht = ThisWorkbook.Worksheets("Structural Steel Materials List");
    MiscMSht = ThisWorkbook.Worksheets("Vendor Misc. Materials");
    DrawSht = ThisWorkbook.Worksheets("Wall Drawings");
    Application.DisplayAlerts = false;
    let Length: number;
    let Width: number;
    let Zoom: number;
    Length = (EstSht.Range("Building_Length")
                + (EstSht.Range("e1_GableExtension").Value + EstSht.Range("e3_GableExtension").Value));
    Width = (EstSht.Range("Building_Width").Value
                + (EstSht.Range("s2_EaveExtension").Value + EstSht.Range("s4_EaveExtension").Value));
    if (((Length <= 50)
                && (Width <= 80))) {
        Zoom = 50;
    }
    else {
        Zoom = 25;
    }

    // Drawings (first page is landscape)
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    xlLandscape.CenterHorizontally = Application.InchesToPoints(0.2);
    Zoom.Orientation = Application.InchesToPoints(0.2);
    0.Zoom = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    DrawSht.PageSetup.FitToPagesWide = Application.InchesToPoints(0.2);
    Application.PrintCommunication = true;
    DrawSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    LastRow = MiscMSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:E" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    MiscMSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = true;
    MiscMSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    LastRow = SteelSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:E" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    SteelSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = true;
    SteelSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    LastRow = SheetMetalSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:E" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    SheetMetalSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = true;
    SheetMetalSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    LastRow = DescriptionSht.Cells[Rows.Count, 1].End(xlUp).Row;
    Application.PrintCommunication = false;
    // With...
    Application.InchesToPoints(0.25).HeaderMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).BottomMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).TopMargin = Application.InchesToPoints(0.2);
    Application.InchesToPoints(0.25).RightMargin = Application.InchesToPoints(0.2);
    true.LeftMargin = Application.InchesToPoints(0.2);
    0.CenterHorizontally = Application.InchesToPoints(0.2);
    1.FitToPagesTall = Application.InchesToPoints(0.2);
    ("A1:A" + LastRow.FitToPagesWide) = Application.InchesToPoints(0.2);
    DescriptionSht.PageSetup.PrintArea = Application.InchesToPoints(0.2);
    Application.PrintCommunication = false;
    DescriptionSht.PrintOut;
    /* Warning! Labeled Statements are not Implemented */1;
    /* Warning! Labeled Statements are not Implemented */true;
    /* Warning! Labeled Statements are not Implemented */false;
    Application.DisplayAlerts = true;
}

SaveAsNewEstimate() {
    // Export doc as .xlsm
    let OS: string;
    OS = Application.OperatingSystem;
    let un: Object;
    un = System.Environment.GetEnvironmentVariable("Username");
    let wb: Workbook;
    wb = ThisWorkbook;
    let EstName: string;
    let FilePath: string;
    let FileOnly: string;
    let PathOnly: string;
    let NewFilePath: string;
    let CustomerFolderPath: string;
    let ClientName: string;
    FilePath = wb.FullName;
    FileOnly = wb.Name;
    PathOnly = FilePath.Substring(0, (FilePath.Length - FileOnly.Length));
    // With...
    ClientName = EstSht.Range;
    "CustomerName".Value;
    CustomerFolderPath = (PathOnly + ClientName);
    EstName = EstSht.Range;
    ("CustomerName".Value + ("_Estimate_" + EstSht.Range));
    ("Building_Width".Value + ("x" + EstSht.Range));
    ("Building_Length".Value + ("x" + EstSht.Range));
    ("Building_Height".Value + ("_"
                + (Now().Month + Now().Year)));
    EstName = ValidWBName(EstName);
    NewFilePath = (CustomerFolderPath + EstName);
    if (((OS.IndexOf("Windows") + 1)
                > 0)) {
        MakeNewFolderPC(CustomerFolderPath);
        wb.SaveAs;
        /* Warning! Labeled Statements are not Implemented */NewFilePath;
        /* Warning! Labeled Statements are not Implemented */xlOpenXMLWorkbookMacroEnabled;
    }
    else {
        MakeNewFolderMAC(ClientName);
        wb.SaveAs;
        /* Warning! Labeled Statements are not Implemented */NewFilePath;
        /* Warning! Labeled Statements are not Implemented */xlOpenXMLWorkbookMacroEnabled;
    }

}

MakeNewFolderPC(path: string) {
    let fso: FileSystemObject = new FileSystemObject();
    // examples for what are the input arguments
    // strDir = "Folder"
    // strPath = "C:\"
    // path = strPath & strDir
    if (!fso.FolderExists(path)) {
        //  doesn't exist, so create the folder
        fso.CreateFolder;
        path;
    }

}

MakeNewFolderMAC(ClientName: string) {
    // Note: This macro uses the FileOrFolderExistsOnYourMac function.
    // Note : Use 1 as second argument for File and 2 for Folder
    // Test if the folder with the name TestFolder is on your desktop
    let FolderPath: string;
    FolderPath = (MacScript("return (path to desktop folder) as string") + ClientName);
    if ((FolderPath.Substring((FolderPath.Length - 1)) == Application.PathSeparator)) {
        MsgBox;
        "Remove the / at the end of the FolderPath";
        return;
    }

    if ((FileOrFolderExistsOnYourMac(FolderPath, 2) == true)) {
        MsgBox;
        "Folder exists.";
    }
    else {
        MkDir;
        MacScript(("return POSIX path of (" + ('\"'
                        + (FolderPath + ('\"' + ")")))));
        MsgBox;
        "Folder not exists but created .";
    }

}

FileOrFolderExistsOnYourMac(FileOrFolderstr: string, FileOrFolder: number): boolean {
    // Ron de Bruin : 13-Dec-2020, for Excel 2016 and higher
    // Function to test if a file or folder exist on your Mac
    // Use 1 as second argument for File and 2 for Folder
    let ScriptToCheckFileFolder: string;
    let FileOrFolderPath: string;
    if ((FileOrFolder == 1)) {
        // File test
        // TODO: On Error Resume Next Warning!!!: The statement is not translatable
        FileOrFolderPath = Dir((FileOrFolderstr + "*"));
        // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
        if (!(FileOrFolderPath == null)) {
            FileOrFolderExistsOnYourMac = true;
        }
        else {
            // folder test
            // TODO: On Error Resume Next Warning!!!: The statement is not translatable
            FileOrFolderPath = Dir((FileOrFolderstr + "*"), vbDirectory);
            // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
            if (!(FileOrFolderPath == null)) {
                FileOrFolderExistsOnYourMac = true;
            }

        }

        (<string>(ValidWBName((<string>(Arg)))));
        let RegEx: Object;
        RegEx = CreateObject("VBScript.RegExp");
        // With...
        Replace(Arg, "");
        RegEx.Pattern = true;
    }

}

BayUpdate(BayRange: Range, BLen: Range, ChangeCell: Range) {
    let BayCell: Range;
    let BayValue: Range;
    let BaySum: number;
    // set bay sum
    BaySum = HiddenSht.Range("TotalBayLength").Value;
    // reset any blanks to 0
    for (BayCell in BayRange) {
        if ((BayCell.Value == "")) {
            BayCell.Value = 0;
        }

    }

    // check for bay length over total building length
    if ((BaySum > BLen.Value)) {
        // error message
        MsgBox;
        "The total bay length cannot exceed the building length! Please correct the data and try again.";
        vbExclamation;
        "Excess Bay Length";
        ChangeCell.Value = 0;
    }

}

MaterialsListCaller() {
    let Confirm: Object;
    let FOCell: Range;
    let ItemCount: number;
    let MissingData: boolean;
    let sqrFootage: number;
    let BayNum: number;
    let Bay1Length: number;
    let LastBayLength: number;
    let Bay1Overhang: number;
    let LastBayOverhang: number;
    // ''Deternine FO trim generation time
    // With...
    // Personell Doors
    for (FOCell in Range(EstSht.Range, "pDoorCell1", EstSht.Range, "pDoorCell12")) {
        // if cell isn't hidden, door size is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // add FO trim pieces to item count
            ItemCount = (ItemCount + 3);
        }

    }

    // Overhead Doors
    for (FOCell in Range(EstSht.Range, "OHDoorCell1", EstSht.Range, "OHDoorCell12")) {
        // if cell isn't hidden, door width is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // add FO trim pieces to item count
            ItemCount = (ItemCount + (3 * 5));
        }

    }

    // Windows
    for (FOCell in Range(EstSht.Range, "WindowCell1", EstSht.Range, "WindowCell12")) {
        // if cell isn't hidden, door width is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // add FO trim pieces to item count
            ItemCount = (ItemCount + 4);
        }

    }

    // Misc Fos
    for (FOCell in Range(EstSht.Range, "MiscFOCell1", EstSht.Range, "MiscFOCell12")) {
        // if cell isn't hidden, door width is entered
        if (((FOCell.EntireRow.Hidden == false)
                    && (FOCell.offset(0, 1).Value != ""))) {
            // add FO trim pieces to item count
            ItemCount = (ItemCount + (4 * 5));
        }

    }

    sqrFootage = EstSht.Range;
    ("Building_Width".Value * EstSht.Range);
    ("Building_Height".Value * EstSht.Range);
    "Building_Length".Value;
    ItemCount = (ItemCount
                + (Application.WorksheetFunction.RoundUp((sqrFootage / 50000), 0) * 10));
    // confirmation message dependent upon item count
    if ((ItemCount > 0)) {
        Confirm = MsgBox(("Would you like to generate a materials list using the information entered? This will take approximately "
                        + (ItemCount + " seconds to complete.")), (System.Windows.Forms.MessageBoxIcon.Information + vbYesNo), "Confirm Materials List Generation");
    }
    else {
        Confirm = MsgBox("Would you like to generate a materials list using the information entered?", (System.Windows.Forms.MessageBoxIcon.Information + vbYesNo), "Confirm Materials List Generation");
    }

    // '' Check for missing information
    // With...
    switch (true) {
        case EstSht.Range:
            "Roof_Shape".Value = "";
            MsgBox;
            "The building's roof shape must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Roof_Shape");
            MissingData = true;
            break;
        case EstSht.Range:
            "Building_Width".Value = "";
            MsgBox;
            "The building's width must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Building_Width");
            MissingData = true;
            break;
        case EstSht.Range:
            "Roof_Pitch".Value = "";
            MsgBox;
            "The building's roof pitch must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Roof_Pitch");
            MissingData = true;
            break;
        case EstSht.Range:
            "Building_Height".Value = "";
            MsgBox;
            "The building's height must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Building_Height");
            MissingData = true;
            break;
        case EstSht.Range:
            "Building_Length".Value = "";
            MsgBox;
            "The building's length must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Building_Length");
            MissingData = true;
            break;
        case EstSht.Range:
            "Wall_pShape".Value = "";
            MsgBox;
            "The building's wall panel shape must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Wall_pShape");
            MissingData = true;
            break;
        case EstSht.Range:
            "Wall_pType".Value = "";
            MsgBox;
            "The building's wall panel type must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Wall_pType");
            MissingData = true;
            break;
        case EstSht.Range:
            "Wall_Color".Value = "";
            MsgBox;
            "The building's wall color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Wall_Color");
            MissingData = true;
            break;
        case EstSht.Range:
            "Roof_pShape".Value = "";
            MsgBox;
            "The building's roof panel shape must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Roof_pShape");
            MissingData = true;
            break;
        case EstSht.Range:
            "Roof_pType".Value = "";
            MsgBox;
            "The building's roof panel type must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Roof_pType");
            MissingData = true;
            break;
        case EstSht.Range:
            "Roof_Color".Value = "";
            MsgBox;
            "The building's roof color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Roof_Color");
            MissingData = true;
            break;
        case EstSht.Range:
            "Rake_tColor".Value = "";
            MsgBox;
            "The building's rake trim color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Rake_tColor");
            MissingData = true;
            break;
        case EstSht.Range:
            "Eave_tColor".Value = "";
            MsgBox;
            "The building's eave trim color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Eave_tColor");
            MissingData = true;
            break;
        case EstSht.Range:
            "OutsideCorner_tColor".Value = "";
            MsgBox;
            "The building's corner trim color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("OutsideCorner_tColor");
            MissingData = true;
            break;
        case EstSht.Range:
            "Base_tColor".Value = "";
            MsgBox;
            "The building's base trim color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("Base_tColor");
            MissingData = true;
            break;
        case EstSht.Range:
            (("e1_Wainscot".Value != "None")
                        & EstSht.Range["e1_Wainscot"].offset(0, 1).Value) = ("" | EstSht.Range);
            "e1_Wainscot".offset(0, 2).Value = "";
            MsgBox;
            "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("e1_Wainscot");
            MissingData = true;
            break;
        case EstSht.Range:
            (("s2_Wainscot".Value != "None")
                        & EstSht.Range["s2_Wainscot"].offset(0, 1).Value) = ("" | EstSht.Range);
            "s2_Wainscot".offset(0, 2).Value = "";
            MsgBox;
            "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("s2_Wainscot");
            MissingData = true;
            break;
        case EstSht.Range:
            (("e3_Wainscot".Value != "None")
                        & EstSht.Range["e3_Wainscot"].offset(0, 1).Value) = ("" | EstSht.Range);
            "e3_Wainscot".offset(0, 2).Value = "";
            MsgBox;
            "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("e3_Wainscot");
            MissingData = true;
            break;
        case EstSht.Range:
            (("s4_Wainscot".Value != "None")
                        & EstSht.Range["s4_Wainscot"].offset(0, 1).Value) = ("" | EstSht.Range);
            "s4_Wainscot".offset(0, 2).Value = "";
            MsgBox;
            "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.";
            vbExclamation;
            "Missing Data";
            Application.GoTo.Range("s4_Wainscot");
            MissingData = true;
            break;
    }

    let i: number;
    let j: number;
    let MissingOHData: boolean;
    // With...
    // Check OHDoor data
    for (i = 1; (i <= 11); i++) {
        if (EstSht.Range) {
            (("OHDoorCell1".offset((i - 1), 1).Value != "")
                        & EstSht.Range);
            "OHDoorCell1".offset((i - 1), 1).EntireRow.Hidden = false;
            MissingOHData = false;
            for (j = 1; (j <= 8); j++) {
                if (EstSht.Range) {
                    "OHDoorCell1".offset((i - 1), (j + 1)).Value = "";
                    MsgBox;
                    "Overhead door information must be entered. Please enter this data before generating a materials list.";
                    vbExclamation;
                    "Missing Data";
                    MissingOHData = true;
                    MissingData = true;
                    Application.GoTo.Range("OHDoorCell1").offset((i - 1), 1);
                    break;
                }

            }

            if ((MissingOHData == true)) {
                break;
            }

            i;
            // Check PDoor data
            for (i = 1; (i <= 11); i++) {
                if (EstSht.Range) {
                    (("pDoorCell1".offset((i - 1), 1).Value != "")
                                & EstSht.Range);
                    "pDoorCell1".offset((i - 1), 1).EntireRow.Hidden = false;
                    MissingOHData = false;
                    for (j = 1; (j <= 7); j++) {
                        if (EstSht.Range) {
                            "pDoorCell1".offset((i - 1), (j + 1)).Value = "";
                            MsgBox;
                            "Personnel door information must be entered. Please enter this data before generating a materials list.";
                            vbExclamation;
                            "Missing Data";
                            MissingOHData = true;
                            MissingData = true;
                            Application.GoTo.Range("pDoorCell1").offset((i - 1), 1);
                            break;
                        }

                    }

                    if ((MissingOHData == true)) {
                        break;
                    }

                    i;
                    // Check Window data
                    for (i = 1; (i <= 11); i++) {
                        if (EstSht.Range) {
                            (("WindowCell1".offset((i - 1), 1).Value != "")
                                        & EstSht.Range);
                            "WindowCell1".offset((i - 1), 1).EntireRow.Hidden = false;
                            MissingOHData = false;
                            for (j = 1; (j <= 5); j++) {
                                if (EstSht.Range) {
                                    "WindowCell1".offset((i - 1), (j + 1)).Value = "";
                                    MsgBox;
                                    "Window information must be entered. Please enter this data before generating a materials list.";
                                    vbExclamation;
                                    "Missing Data";
                                    MissingOHData = true;
                                    MissingData = true;
                                    Application.GoTo.Range("WindowCell1").offset((i - 1), 1);
                                    break;
                                }

                            }

                            if ((MissingOHData == true)) {
                                break;
                            }

                            i;
                            // Check Misc FO data
                            for (i = 1; (i <= 11); i++) {
                                if (EstSht.Range) {
                                    (("MiscFOCell1".offset((i - 1), 1).Value != "")
                                                & EstSht.Range);
                                    "MiscFOCell1".offset((i - 1), 1).EntireRow.Hidden = false;
                                    MissingOHData = false;
                                    for (j = 1; (j <= 7); j++) {
                                        if (EstSht.Range) {
                                            "MiscFOCell1".offset((i - 1), (j + 1)).Value = "";
                                            MsgBox;
                                            "Misc. Framed Opening information must be entered. Please enter this data before generating a materials list.";
                                            vbExclamation;
                                            "Missing Data";
                                            MissingOHData = true;
                                            MissingData = true;
                                            Application.GoTo.Range("MiscFOCell1").offset((i - 1), 1);
                                            break;
                                        }

                                    }

                                    if ((MissingOHData == true)) {
                                        break;
                                    }

                                    i;
                                }

                                // Check that Bay Lenghts + Overhangs don't add up to more than 30'
                                // With...
                                // Get Bay Num
                                BayNum = EstSht.Range;
                                "BayNum".Value;
                                if ((BayNum > 0)) {
                                    // Get Bay 1 Length
                                    Bay1Length = EstSht.Range;
                                    "Bay1_Length".Value;
                                    // Get Last Bay Length
                                    LastBayLength = EstSht.Range;
                                    "Bay1_Length".offset((BayNum - 1), 0);
                                    // Get Bay 1 Overhang
                                    Bay1Overhang = EstSht.Range;
                                    "e1_GableOverhang".Value;
                                    // Get Last Bay Overhang
                                    LastBayOverhang = EstSht.Range;
                                    "e3_GableOverhang".Value;
                                    if ((Bay1Length
                                                + (Bay1Overhang > 30))) {
                                        MsgBox;
                                        ("The combined bay length ("
                                                    + (Bay1Length + ("') and overhang length ("
                                                    + (Bay1Overhang + "') at endwall 1 is longer than 30'. Please enter values less than 30' before generating a materials list."))));
                                        vbExclamation;
                                        "Bay Length Error";
                                        Application.GoTo.Range("e1_GableOverhang");
                                        MissingData = true;
                                    }

                                    if ((LastBayLength
                                                + (LastBayOverhang > 30))) {
                                        MsgBox;
                                        ("The combined bay length ("
                                                    + (LastBayLength + ("') and overhang length ("
                                                    + (LastBayOverhang + "') at endwall 3 is longer than 30'. Please enter values less than 30' before generating a materials list."))));
                                        vbExclamation;
                                        "Bay Length Error";
                                        Application.GoTo.Range("e3_GableOverhang");
                                        MissingData = true;
                                    }

                                }

                                if (((Confirm == System.Windows.Forms.MessageBoxButtons.Yes)
                                            && (MissingData == false))) {
                                    Application.ScreenUpdating = false;
                                    MaterialsListGen.MaterialsListGen;
                                    // screen updating is turned back on at the end of the materials list gen so the view can be changed
                                }

                                ImportProjectDetails();
                                let Sourcewb: Workbook;
                                let Sourcews: Worksheet;
                                let Activewb: Workbook;
                                let FileToOpen: Object;
                                let i: number;
                                let j: number;
                                let Address: string;
                                let BayErrorMsg: string;
                                let OS: string;
                                let FileFound: boolean;
                                let mybook: Workbook;
                                let MyPath: string;
                                let MyScript: string;
                                let MyFiles: string;
                                let MySplit: Object;
                                let N: number;
                                let FName: string;
                                Application.ScreenUpdating = false;
                                Application.Calculation = xlCalculationManual;
                                FileFound = false;
                                OS = Application.OperatingSystem;
                                if (((OS.IndexOf("Windows") + 1)
                                            > 0)) {
                                    Application.ScreenUpdating = false;
                                    Activewb = ThisWorkbook;
                                    // ''''Select File to Import From''''''
                                    FileToOpen = Application.GetOpenFilename("Browse for your File & Import Range", "Excel Files (*.xlsm*),*xlsm*");
                                    // TODO: Labeled Arguments not supported. Argument: 1 := 'Title'
                                    // TODO: Labeled Arguments not supported. Argument: 2 := 'FileFilter'
                                    // Check Valid File and that it's an Estimating Template
                                    if ((FileToOpen != false)) {
                                        Sourcewb = Application.Workbooks.Open(FileToOpen);
                                        FileFound = true;
                                    }

                                }
                                else {
                                    // ''' CODE TO OPEN FILE ON MAC
                                    // TODO: On Error Resume Next Warning!!!: The statement is not translatable
                                    MyPath = MacScript("return (path to documents folder) as String");
                                    // Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"
                                    //  In the following statement, change true to false in the line "multiple
                                    //  selections allowed true" if you do not want to be able to select more
                                    //  than one file. Additionally, if you want to filter for multiple files, change
                                    //  {""com.microsoft.Excel.xls""} to
                                    //  {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
                                    //  if you want to filter on xls and csv files, for example.
                                    MyScript = ("set applescript's text item delimiters to "","" " + ("
" + ("set theFiles to (choose file of type " + (" {""org.openxmlformats.spreadsheetml.sheet.macroenabled""}" + ("with prompt ""Please select a file or files"" default location alias """
                                                + (MyPath + (""" multiple selections allowed false) as string" + ("
" + ("set applescript's text item delimiters to """" " + ("
" + "return theFiles"))))))))));
                                    MyFiles = MacScript(MyScript);
                                    // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
                                    if ((MyFiles != "")) {
                                        // With...
                                        Application.ScreenUpdating = false;
                                        MyFiles = MyFiles.Replace(":", "/");
                                        MyFiles = Replace(MyFiles, "Macintosh HD", "", 1);
                                        // TODO: Labeled Arguments not supported. Argument: 4 := 'Count'
                                        MySplit = MyFiles.Split(",");
                                        for (N = LBound(MySplit); (N <= UBound(MySplit)); N++) {
                                            //  Get the file name only and test to see if it is open.
                                            FName = MySplit[N].Substring((MySplit[N].Length - (MySplit[N].Length - InStrRev(MySplit[N], Application.PathSeparator, ,, 1))));
                                            if ((bIsBookOpen(FName) == false)) {
                                                mybook = null;
                                                // TODO: On Error Resume Next Warning!!!: The statement is not translatable
                                                mybook = Workbooks.Open(MySplit[N]);
                                                // TODO: On Error GoTo 0 Warning!!!: The statement is not translatable
                                                if (!(mybook == null)) {
                                                    Sourcewb = mybook;
                                                    FileFound = true;
                                                    /* Warning! GOTO is not Implemented */}

                                            }

                                        }

                                    }

                                }

                                // TODO: Continue... Warning!!! not translated
                                if ((FileFound == false)) {
                                    MsgBox;
                                    "Please Select Valid File";
                                    Application.ScreenUpdating = true;
                                    Application.Calculation = xlCalculationAutomatic;
                                    /* Warning! GOTO is not Implemented */}

                                // Define Import Sheet and check for Project Details using custom function
                                if (SheetExists("Project Details", Sourcewb)) {
                                    // ''change to named range
                                    Sourcews = Sourcewb.Worksheets("Project Details");
                                }
                                else {
                                    Sourcewb.Close;
                                    false;
                                    MsgBox;
                                    "Please Select Valid File";
                                    Application.ScreenUpdating = true;
                                    Application.Calculation = xlCalculationAutomatic;
                                    /* Warning! GOTO is not Implemented */}

                                // '''''''Import Data; for ranges WITHOUT Change triggers, Turn Events OFF by "FALSE" parameter
                                Application.EnableEvents = false;
                                CopyNamedRange;
                                "Building_Width";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Roof_Pitch";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Building_Height";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Building_Length";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "BayNum";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("BayNum").Value > 0)) {
                                    Application.Calculation = xlCalculationAutomatic;
                                    for (i = 0; (i <= 11); i++) {
                                        CopyNamedRange;
                                        Sourcews.Range("Building_Height").offset((i + 3), 0).Address;
                                        Sourcews;
                                        true;
                                    }

                                    Application.Calculation = xlCalculationManual;
                                }

                                CopyNamedRange;
                                "Roof_Shape";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Wall_pShape";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Wall_pType";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Wall_Color";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "LinerPanels";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("LinerPanels").Value == "Yes")) {
                                    for (i = 0; (i <= 4); i++) {
                                        for (j = 0; (j <= 3); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("e1_LinerPanels").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "Roof_pShape";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Roof_pType";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Roof_Color";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "AlterWalls";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("AlterWalls").Value == "Yes")) {
                                    Sourcews.Unprotect;
                                    "WhiteTruckMafia";
                                    EstSht.Unprotect;
                                    "WhiteTruckMafia";
                                    CopyNamedRange;
                                    "e1_WallStatus";
                                    Sourcews;
                                    true;
                                    if ((Sourcews.Range("e1_WallStatus").Value == "Partial")) {
                                        CopyNamedRange;
                                        Sourcews.Range("e1_WallStatus").offset(0, 2).Address;
                                        Sourcews;
                                        false;
                                    }

                                    CopyNamedRange;
                                    "s2_WallStatus";
                                    Sourcews;
                                    true;
                                    if ((Sourcews.Range("s2_WallStatus").Value == "Partial")) {
                                        CopyNamedRange;
                                        Sourcews.Range("s2_WallStatus").offset(0, 2).Address;
                                        Sourcews;
                                        false;
                                    }

                                    CopyNamedRange;
                                    "e3_WallStatus";
                                    Sourcews;
                                    true;
                                    if ((Sourcews.Range("e3_WallStatus").Value == "Partial")) {
                                        CopyNamedRange;
                                        Sourcews.Range("e3_WallStatus").offset(0, 2).Address;
                                        Sourcews;
                                        false;
                                    }

                                    CopyNamedRange;
                                    "s4_WallStatus";
                                    Sourcews;
                                    true;
                                    if ((Sourcews.Range("s4_WallStatus").Value == "Partial")) {
                                        CopyNamedRange;
                                        Sourcews.Range("s4_WallStatus").offset(0, 2).Address;
                                        Sourcews;
                                        false;
                                    }

                                    CopyNamedRange;
                                    "e1_Expandable";
                                    Sourcews;
                                    false;
                                    CopyNamedRange;
                                    "e3_Expandable";
                                    Sourcews;
                                    false;
                                    Sourcews.Protect;
                                    "WhiteTruckMafia";
                                    EstSht.Protect;
                                    "WhiteTruckMafia";
                                }

                                CopyNamedRange;
                                "All_tColors";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "FO_tColor";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Base_tColor";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Rake_tColor";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Eave_tColor";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "OutsideCorner_tColor";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "Wainscot";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("Wainscot").Value == "Yes")) {
                                    for (i = 0; (i <= 3); i++) {
                                        for (j = 0; (j <= 2); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("e1_Wainscot").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "GutterAndDownspouts";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "PDoorNum";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("PDoorNum").Value > 0)) {
                                    for (i = 0; (i <= 11); i++) {
                                        for (j = 1; (j <= 6); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("pDoorCell1").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "OHDoorNum";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("OHDoorNum").Value > 0)) {
                                    for (i = 0; (i <= 11); i++) {
                                        for (j = 1; (j <= 8); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("OHDoorCell1").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "WindowNum";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("WindowNum").Value > 0)) {
                                    for (i = 0; (i <= 23); i++) {
                                        for (j = 1; (j <= 3); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("WindowCell1").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "MiscFONum";
                                Sourcews;
                                true;
                                if ((Sourcews.Range("MiscFONum").Value > 0)) {
                                    for (i = 0; (i <= 11); i++) {
                                        for (j = 1; (j <= 5); j++) {
                                            CopyNamedRange;
                                            Sourcews.Range("MiscFOCell1").offset(i, j).Address;
                                            Sourcews;
                                            false;
                                        }

                                    }

                                }

                                CopyNamedRange;
                                "WallInsulation";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "RoofInsulation";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "RidgeVentQty";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "TranslucentWallPanelQty";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "SkylightQty";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "RidgeVentType";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "TranslucentWallPanelLength";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "SkylightLength";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "e1_GableOverhang";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "s2_EaveOverhang";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "e3_GableOverhang";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "s4_EaveOverhang";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "e1_GableOverhangSoffit";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "s2_EaveOverhangSoffit";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "e3_GableOverhangSoffit";
                                Sourcews;
                                false;
                                CopyNamedRange;
                                "s4_EaveOverhangSoffit";
                                Sourcews;
                                false;
                                for (i = 0; (i <= 3); i++) {
                                    for (j = 1; (j <= 4); j++) {
                                        Address = EstSht.Range("e1_GableOverhangSoffit").offset(i, j).Address;
                                        CopyNamedRange;
                                        Address;
                                        Sourcews;
                                        false;
                                    }

                                }

                                CopyNamedRange;
                                "e1_GableExtension";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "s2_EaveExtension";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "e3_GableExtension";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "s4_EaveExtension";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "e1_GableExtensionSoffit";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "s2_EaveExtensionSoffit";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "e3_GableExtensionSoffit";
                                Sourcews;
                                true;
                                CopyNamedRange;
                                "s4_EaveExtensionSoffit";
                                Sourcews;
                                true;
                                for (i = 0; (i <= 3); i++) {
                                    for (j = 1; (j <= 4); j++) {
                                        Address = EstSht.Range("e1_GableExtensionSoffit").offset(i, j).Address;
                                        CopyNamedRange;
                                        Address;
                                        Sourcews;
                                        false;
                                    }

                                }

                                Application.EnableEvents = true;
                                Application.ScreenUpdating = true;
                                Application.Calculation = xlCalculationAutomatic;
                                Application.GoTo;
                                /* Warning! Labeled Statements are not Implemented */EstSht.Range("CustomerName");
                                /* Warning! Labeled Statements are not Implemented */true;
                                Sourcewb.Close;
                                false;
                                /* Warning! Labeled Statements are not Implemented */CopyNamedRange((<string>(Name)), (<Worksheet>(Sourcews)), (<boolean>(EnableEvent)));
                                if ((EnableEvent == true)) {
                                    Application.EnableEvents = true;
                                }

                                // TODO: On Error GoTo Warning!!!: The statement is not translatable
                                // Ignores cells that are protected and unavailable(which means they would be on the current sheet as well)
                                // Also avoids errors for old files that are missing named ranges
                                // Sourcews.Range(Name).Copy
                                EstSht.Range(Name).Value = Sourcews.Range(Name).Value;
                                /* Warning! Labeled Statements are not Implemented */Application.EnableEvents = false;
                                (<boolean>(SheetExists((<string>(SheetName)), (<Workbook>(Sourcewb)))));
                                // TODO: On Error Resume Next Warning!!!: The statement is not translatable
                                SheetExists = (Sourcewb.Sheets(SheetName).Index > 0);
                                (<boolean>(bIsBookOpen(/* ref */(<string>(szBookName)))));
                                //  Contributed by Rob Bovey
                                // TODO: On Error Resume Next Warning!!!: The statement is not translatable
                                bIsBookOpen = !(Application.Workbooks(szBookName) == null);
                            }

                        }

                    }

                }

            }

        }

    }

}
