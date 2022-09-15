class Building {
  bLength: number
  bHeight: number
  rPitch: number
  RafterLength: number
  s2RafterSheetLength: number
  s4RafterSheetLength: number
  bWidth: number
  rShape: string
  s2Overhang: number
  s4Overhang: number
  e1Overhang: number
  e3Overhang: number
  s2Extension: number
  s4Extension: number
  e1Extension: number
  e3Extension: number
  e1ExtensionPanelQty: number
  e3ExtensionPanelQty: number
  Gutters: boolean
  BaseTri: boolean
  //endwall wall panel overlaps
  e1WallPanelOverlaps: number
  e3WallPanelOverlaps: number
  //extension pitches
  s2ExtensionPitch: number
  s4ExtensionPitch: number
  //extension Heights
  s2ExtensionHeight: number
  s4ExtensionHeight: number
  //extension widths
  s2ExtensionWidth: number
  s4ExtensionWidth: number
  //Panel Shapes
  wPanelShape: string    //sidewall panel shapes
  rPanelShape: string    //roof panel shapes
  //Panel Types, Colors
  rPanelType: string
  rPanelColo: string
  wPanelType: string
  wPanelColo: string
  //Trim Colors
  RakeTrimColor: string
  OutsideCorner: string
  //soffit booleans
  e1GableOverhangSoffit: boolean
  e3GableOverhangSoffit: boolean
  s2EaveOverhangSoffit: boolean
  s4EaveOverhangSoffit: boolean
  e1GableExtensionSoffi: boolean
  e3GableExtensionSoffi: boolean
  s2EaveExtensionSoffit: boolean
  s4EaveExtensionSoffit: boolean
  // this for totaling eave extension string
  EaveExtLength: number
  //roof panel overage
  bLengthRoofPanelOverage: number
  //Interior Columns Collection
  InteriorColumns: any[]
  s2ColumnWidth: number
  s4ColumnWidth: number
  //'Weld Clips
  WeldClips: number
  //'Structural Steel total cost
  SSTotalCost: number

  //''''''''''''''''''''''''''''''''''''''''''''''''' FO Collections
  e1FOs: any[]
  s2FOs: any[]
  e3FOs: any[]
  s4FOs: any[]
  fieldlocateFOs: any[]
  //''''''''''''''''''''''''''''''''''''''''''''''''' Column Collections
  e1Columns: any[]
  s2Columns: any[]
  e3Columns: any[]
  s4Columns: any[]
  //''''''''''''''''''''''''''''''''''''''''''''''''' Girt Collections
  e1Girts: any[]
  s2Girts: any[]
  e3Girts: any[]
  s4Girts: any[]
  //''''''''''''''''''''''''''''''''''''''''''''''''' Rafter Collections
  e1Rafters: any[]
  intRafters: any[]
  e3Rafters: any[]
  //''''''''''''''''''''''''''''''''''''''''''''''''' Roof Purlin Collection
  RoofPurlins = []

  //''''''''''''''''''''''''''''''''''''''''''''''''' Overhang Members
  e1OverhangMembers = []
  s2OverhangMembers = []
  e3OverhangMembers = []
  s4OverhangMembers = []
  //''''''''''''''''''''''''''''''''''''''''''''''''' Extension Members
  e1ExtensionMembers = []
  s2ExtensionMembers = []
  e3ExtensionMembers = []
  s4ExtensionMembers = []

  //''''''''''''''''''''''''''''''''''''''''''''''''' Base Angle Trim
  BaseAngleTrim = []

  //''''''''''''''''''''''''''''''''''''''''''''''''' Weld Plates
  WeldPlates = []


  constructor() {

}


    RoofLength() {
        return this.bLength * 12 + this.e1Overhang + this.e1Extension + this.e3Overhang + this.e3Extension;
    }
    RoofFtLength () {
        return (this.blength * 12 + this.e1Overhang + this.e1Extension + this.e3Overhang + this.e3Extension) / 12;
    }
    HighSideEaveHeight() {
        return (this.bHeight * 12) + (this.bWidth * this.rPitch);
    }
    s2ExtensionRafterLength() {
        if (this.s2Extension === 0) {
            return 0
        }
        else {
            return (s2Extension / 12 * Math.sqrt(144 +  this.s4ExtensionPitch))
        }
    }
    s4ExtensionRafterLength() {
        if (this.s4Extension === 0) {
            return 0
        }
        else {
            return (s4Extension / 12) * Math.sqrt((12 ^ 2) + (s4ExtensionPitch ^ 2))
        }
    }
    /*
    * Extension Intersections
    * Note: Intersecting extension panels are accounted for as eave extension panels
    * 
    */
    s2e1ExtensionIntersection() {
        //If the input box for s2e1_Intersection = "N/A" or "Exclude" then
        //return false
        //If '' = "Include" then
        //return true
        //Input box from Estimation Sheet Range Key "s2e1_Intersection"
    }
    s2e3ExtensionIntersection() {
        //If the input box for s2e3_Intersection = "N/A" or "Exclude" then
        //return false
        //If '' = "Include" then
        //return true
        //Input box from Estimation Sheet Range Key "s2e3_Intersection"
    }
    s4e1ExtensionIntersection() {
        //If the input box for s4e1_Intersection = "N/A" or "Exclude" then
        //return false
        //If '' = "Include" then
        //return true
        //Input box from Estimation Sheet Range Key "s4e1_Intersection"
    }
    s4e3ExtensionIntersection() {
        //If the input box for s4e3_Intersection = "N/A" or "Exclude" then
        //return false
        //If '' = "Include" then
        //return true
        //Input box from Estimation Sheet Range Key "s4e3_Intersection"
    }
    /*
    * Eave Extension Lengths (from endwall to endwall)
    * The below return errors from this.fn() will be resolved when logic is resolved
    */
    s2EaveExtensionBuildingLength() {
        this.EaveExtLength = (this.bLength * 12) + this.e1Overhang + this.e3Overhang
        if (this.s2e1ExtensionIntersection()) {
            this.EaveExtLength += this.e1Extension

        }
    }
}