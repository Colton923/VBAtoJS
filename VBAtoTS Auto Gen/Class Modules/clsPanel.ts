// TODO: Option Explicit ... Warning!!! not translated
let PanelLength: number;
let PanelMeasurement: string;
let Quantity: number;
let DeleteFlag: boolean;
let PanelShape: string;
let PanelType: string;
let clsType: string;
let PanelColor: string;
Public(<void>(TotalCost));
VariantPublic(<void>(UnitCost));
VariantPublic(<void>(FootageCost));
Variantlet SkipFlag: boolean;
let rEdgePosition: number;
let bEdgeHeight: number;


    private Class_Initialize() {
        clsType = "Panel";
        FootageCost = "N/A";
    }

    public lEdgePosition(): number {
        return (rEdgePosition + (3 * 12));
    }

    public tEdgeHeight(): number {
        return (bEdgeHeight + PanelLength);
    }
