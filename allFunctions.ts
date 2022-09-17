Public Sub AdditionalWeldClips(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub AdjustEndwallColumns(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub AdjustFOMembers(b As clsBuilding, eWall As String)

    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub AdjustSidewallColumns(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function ArrayRemoveDups(MyArray)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub BaseAngleTrimGen(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub BayUpdate(BayRange As Range, BLen As Range, ChangeCell As Range)
    Member of VBAProject.GeneralMod

Public Function bIsBookOpen(szBookName As String) As Boolean
    Member of VBAProject.GeneralMod

Public Sub BPP_Solver(SolvedCollection As Collection, InputCollection As Collection, InputType As String, [FOType As String], [Wall As String])
    Member of VBAProject.JankyBPPSolver

Public Sub CheckDate()
    Member of VBAProject.BETASelfDestructModule

Public Function ClosestRoofPurlin(RafterLength, [Direction As Integer]) As Double
    Member of VBAProject.MaterialsListGen

Public Function ClosestWallGirt(Height, [Direction As Integer]) As Double
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function ClosestWallPurlin(Height, [Direction As Integer], [NonstandardFloorPurlin As Boolean]) As Double
    Member of VBAProject.MaterialsListGen

Public Sub ColTest()
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub CombinePurlins(b As clsBuilding, manualGirtOptimization As Collection, tempGirtsCollection As Collection, NewOptimizedCol As Collection)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub CombineWeldPlates(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function ConflictingEndwallOHDoor(Location As Double, b As clsBuilding, [eWall As String]) As Boolean
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub CopyNamedRange(Name As String, Sourcews As Worksheet, EnableEvent As Boolean)
    Member of VBAProject.GeneralMod

Public Sub CostEstimateGen(PanelCollection As Collection, TrimCollection As Collection, MiscCollection As Collection, b As clsBuilding)
    Member of VBAProject.VendorAndPriceLists

Public Sub CreateOHDoorPriceExport()
    Member of VBAProject.GeneralMod

Public Sub CreateSheetMetalPriceExport()
    Member of VBAProject.GeneralMod

Public Sub CutListOutput(Collection As Collection, Label As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub DescriptionGen(b As clsBuilding)
    Member of VBAProject.VendorAndPriceLists

Public Sub DisplayDrawingInfo(Placement As Double)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub DrawDimension(b As clsBuilding, xLeft As Double, yTop As Double, Width As Double, Height As Double, Direction As String, Font As Integer, Dimensions() As Variant, [Label As String])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub DrawItems(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub DuplicateMaterialRemoval(MaterialCollection As Collection, [CollectionType As String])
    Member of VBAProject.MaterialsListGen

Public EaveStrutCount As Integer
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub EaveStrutTypes(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub EndwallColumnCLCalc(b As clsBuilding, [eWall As String])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub EndwallExtensionColumnsGen(b As clsBuilding, eWall As String, [NewColNum As Integer], [Reiterate As Boolean])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub EndwallGirtLengthCalc(b As clsBuilding, [eWall As String])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub EventsEnable()
    Member of VBAProject.GeneralMod

Public Sub FieldLocateFOCalc(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function FileOrFolderExistsOnYourMac(FileOrFolderstr As String, FileOrFolder As Long) As Boolean
    Member of VBAProject.GeneralMod

Public Sub FOJambsCalc(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function HighLiftSize(FtLength, [Direction As Integer])
    Member of VBAProject.MiscMaterialsGen

Public Function ImperialMeasurementFormat(TotalInches As Double) As String
    Member of VBAProject.MaterialsListGen

Public Sub ImportOHDoorPriceSht()
    Member of VBAProject.GeneralMod

Public Sub ImportOpen(ImportFile As Workbook)
    Member of VBAProject.GeneralMod

Public Sub ImportProjectDetails()
    Member of VBAProject.GeneralMod

Public Sub ImportSheetMetalPriceSht()
    Member of VBAProject.GeneralMod

Public Sub IntColumnsGen(b As clsBuilding, [NewColNum As Integer], [Reiterate As Boolean])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub KillMe()
    Member of VBAProject.BETASelfDestructModule

Public Sub MakeNewFolderMAC(ClientName As String)
    Member of VBAProject.GeneralMod

Public Function MakeNewFolderPC(path As String)
    Member of VBAProject.GeneralMod

Public Function MatchingEndwallColumn(Location As Double, b As clsBuilding) As Boolean
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub MaterialsListCaller()
    Member of VBAProject.GeneralMod

Public Sub MaterialsListGen()
    Member of VBAProject.MaterialsListGen

Public Function MinimumInteriorColumnWidth(b As clsBuilding, ColIndex As Integer, Columns() As Double) As Double
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub MiscMaterialCalc(MiscMaterials As Collection, WriteCell As Range, b As clsBuilding)
    Member of VBAProject.MiscMaterialsGen

Public Sub MoveExtensionOverhangMembers(b As clsBuilding)

    Member of VBAProject.StructuralSteelMaterialsGen

Public Function NearestEndwallLocation(Location As Double, b As clsBuilding, [Alternate As String], [eWall As String]) As Double
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function NearestMemberSize(Length, [Direction As Integer], [MemberType As String], [NumericOutput As Boolean])

    Member of VBAProject.StructuralSteelMaterialsGen

Public Function NearestTrimSize(Length, [Direction As Integer], [UniqueTrimType As String], [NumericOutput As Boolean])
    Member of VBAProject.MaterialsListGen

Public Sub NewExpandableEndwallColumnsGen(b As clsBuilding, eWall As String, EndwallColumnCLs() As Double, [NewColNum As Integer], [Reiterate As Boolean])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function NextHorizontalGirtIntersection(b As clsBuilding, Columns As Collection, FOs As Collection, start As Double, Wall As String, Height As Double) As Double
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function NonExpandableFOJambs(b As clsBuilding, eWall As String, StartPos As Double, MaxDistance As Double, IdealSpan As Double, Direction As Integer) As Double
    Member of VBAProject.StructuralSteelMaterialsGen

Public Const offset_constant = 21
    Member of VBAProject.JankyBPPSolver

Public Sub OverhangExtensionMembersGen(b As clsBuilding)

    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub ParseFOPurlins(tempGirtsCollection As Collection, temp8Receivers As Collection, b As clsBuilding, FOCollection As Collection, manualGirtOptimization As Collection)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub ParseGirts(b As clsBuilding, tempGirtsCollection As Collection, buildingGirts As Collection, EaveStrutCollection As Collection, manualGirtOptimization As Collection)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub PriceListGen(PanelCollection As Collection, TrimCollection As Collection, MiscCollection As Collection)
    Member of VBAProject.VendorAndPriceLists

Public Sub PrintDescription()
    Member of VBAProject.GeneralMod

Public Sub PrintEmployeePackage()
    Member of VBAProject.GeneralMod

Public Sub PrintFloorplan()
    Member of VBAProject.GeneralMod

Public Sub PrintManagerPackage()
    Member of VBAProject.GeneralMod

Public Sub QuickSort(arr, first As Long, last As Long)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub RafterGen(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub RemoveEndwallColumns(b As clsBuilding, eWall As String)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub ReverseArray(vArray)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub RoofPurlinGen(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub SaveAs(NewName As String, CopySht As Worksheet)
    Member of VBAProject.GeneralMod

Public Sub SaveAsNewEstimate()
    Member of VBAProject.GeneralMod

Public Function SheetExists(SheetName As String, Sourcewb As Workbook) As Boolean
    Member of VBAProject.GeneralMod

Public Sub SidewallPanelGen(SidewallPanels As Collection, sWall As String, b As clsBuilding, [FullHeightLinerPanels As Boolean])
    Member of VBAProject.MaterialsListGen

Public Sub SteelMaterialOutput(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public SteelMode As Boolean
    Member of VBAProject.JankyBPPSolver

Public Sub SteelPriceOutput(Collection As Collection, Label As String, [FOMode As Boolean])
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub Test32()
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub TestGirGen(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub TestingSub2(b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Sub VendorMaterialListsGen(PanelCollection As Collection, TrimCollection As Collection, MiscCollection As Collection)
    Member of VBAProject.VendorAndPriceLists

Public Sub WeldPlateGen(RafterLine As String, b As clsBuilding)
    Member of VBAProject.StructuralSteelMaterialsGen

Public Function XLMod(a, b)
    Member of VBAProject.MaterialsListGen

