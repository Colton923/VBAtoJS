// TODO: Option Explicit ... Warning!!! not translated
let Height: number;
let Width: number;
let FOType: string;
let rEdgePosition: number;
let bEdgeHeight: number;
let Wall: string;
let Description: string;
let FOMaterials: Collection;
let StructuralSteelOption: string;
// Public Sub SetWall(WallInput As String)
// Select Case WallInput
// Case "Endwall 1"
//     Wall = "e1"
// Case "Sidewall 2"
//     Wall = "s2"
// Case "Endwall 3"
//     Wall = "e3"
// Case "Sidewall 4"
//     Wall = "s4"
// End Select
// End Sub


    private Class_Initialize() {
        // default bottom edge to floor level
        bEdgeHeight = 0;
        // new FO Materials Collection
        FOMaterials = new Collection();
    }

    public tEdgeHeight(): number {
        return (bEdgeHeight + Height);
    }

    public lEdgePosition(): number {
        return (rEdgePosition + Width);
    }
