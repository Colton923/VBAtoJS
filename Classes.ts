class cBuilding {
	length: number
	width: number
	height: number
	gutters: boolean
	baseTrim: string
	roof: cRoof
	endWall1: cWall
	sideWall1: cWall
	endWall2: cWall
	sideWall2: cWall
	bays?: any[]
	allTrimColor: string
	openingsTrimColor: string
	baseTrimColor: string
	rakeColor: string
	eaveColor: string
	cornerColor: string
	guttersAndDownspouts: boolean
	gutterColor: string
	DownspoutColor: string
	wainscot: boolean
	wainscotTrimColor: string

	constructor(
		length: number,
		width: number,
		height: number,
		roofPitch: string,
		gutters: boolean,
		baseTrim: string,
		roofInsulation: boolean,
		roofInsulationType: string,
		wallInsulation: boolean,
		wallInsulationType: string,
		allTrimColor: string,
		openingsTrimColor: string,
		baseTrimColor: string,
		rakeColor: string,
		eaveColor: string,
		cornerColor: string,
		guttersAndDownspouts: boolean,
		gutterColor: string,
		DownspoutColor: string,
		wainscot: boolean,
		wainscotTrimColor: string
	) {
		this.length = length
		this.width = width
		this.height = height
		this.gutters = gutters
		this.baseTrim = baseTrim
		this.bays = []
		this.roof = new cRoof(roofPitch, length, width, roofInsulation, roofInsulationType)
		this.endWall1 = new cWall(width, height, wallInsulation, wallInsulationType)
		this.sideWall1 = new cWall(length, height, wallInsulation, wallInsulationType)
		this.endWall2 = new cWall(width, height, wallInsulation, wallInsulationType)
		this.sideWall2 = new cWall(length, height, wallInsulation, wallInsulationType)
		this.allTrimColor = allTrimColor
		this.openingsTrimColor = openingsTrimColor
		this.baseTrimColor = baseTrimColor
		this.rakeColor = rakeColor
		this.eaveColor = eaveColor
		this.cornerColor = cornerColor
		this.guttersAndDownspouts = guttersAndDownspouts
		this.gutterColor = gutterColor
		this.DownspoutColor = DownspoutColor
		this.wainscot = wainscot
		this.wainscotTrimColor = wainscotTrimColor
	}
	addBay(length: number) {
		this.bays.push(new cBay(length))
	}
}
class cRoof {
	insulation: boolean
	length: number
	width: number
	roofPitch: string
	insulationType: string
	eaveStruts: any[]
	purlins: any[]
	rafters: any[]
	panels: any[]
	soffits: any[]
	constructor(
		roofPitch: string,
		length: number,
		width: number,
		insulation: boolean,
		insulationType: string
	) {
		this.roofPitch = roofPitch
		this.length = length
		this.width = width
		this.insulation = insulation
		this.insulationType = insulationType
		this.eaveStruts = []
		this.purlins = []
		this.rafters = []
		this.panels = []
		this.soffits = []
	}
	addEaveStrut(size: string, length: number) {
		this.eaveStruts.push(new cEaveStrut(size, length))
	}
	addPurlin(size: string, length: number) {
		this.purlins.push(new cPurlin(size, length))
	}
	addPanel(liner: boolean, length: number, shape: string, type: string, color: string) {
		this.panels.push(new cPanel(liner, length, shape, type, color))
	}
	addRafter(size: string, length: number) {
		this.rafters.push(new cRafter(size, length))
	}
	addSoffit(shape: string, type: string, color: string, trimColor: string) {
		this.soffits.push(new cSoffit(shape, type, color, trimColor))
	}
}
class cWall {
	length: number
	height: number
	columns: any[]
	girts: any[]
	panels: any[]
	insulation: boolean
	insulationType?: string
	expandable?: boolean
	status?: string
	wainscotTrimColor?: string
	wainscotType?: string
	wainscotPanel?: string
	wainscotPanelColor?: string

	constructor(length: number, height: number, insulation: boolean, insulationType: string) {
		this.length = length
		this.height = height
		this.insulation = insulation
		this.insulationType = insulationType
		this.columns = []
		this.girts = []
		this.panels = []
		this.expandable = false
		this.status = ''
	}

	alterHeight(height: number) {
		this.height = height
	}
	alterExpandable(expandable: boolean) {
		this.expandable = expandable
	}
	alterStatus(status: string) {
		this.status = status
	}

	addWainscot(
		wainscotTrimColor: string,
		wainscotType: string,
		wainscotPanel: string,
		wainscotPanelColor: string
	) {
		this.wainscotTrimColor = wainscotTrimColor
		this.wainscotType = wainscotType
		this.wainscotPanel = wainscotPanel
		this.wainscotPanelColor = wainscotPanelColor
	}
	addColumn(size: string, length: number) {
		this.columns.push(new cColumn(size, length))
	}
	addGirt(size: string, length: number) {
		this.girts.push(new cGirt(size, length))
	}
	addPanel(liner: boolean, length: number, shape: string, type: string, color: string) {
		this.panels.push(new cPanel(liner, length, shape, type, color))
	}
}
class cGirt {
	size: string
	length: number

	constructor(size: string, length: number) {
		this.size = size
		this.length = length
	}
}
class cColumn {
	size: string
	length: number

	constructor(size: string, length: number) {
		this.size = size
		this.length = length
	}
}
class cRafter {
	size: string
	length: number

	constructor(size: string, length: number) {
		this.size = size
		this.length = length
	}
}
class cPurlin {
	size: string
	length: number

	constructor(size: string, length: number) {
		this.size = size
		this.length = length
	}
}
class cEaveStrut {
	size: string
	length: number
	constructor(size: string, length: number) {
		this.size = size
		this.length = length
	}
}
class cSoffit {
	shape: string
	type: string
	color: string
	trimColor: string
	constructor(shape: string, type: string, color: string, trimColor: string) {
		this.shape = shape
		this.type = type
		this.color = color
		this.trimColor = trimColor
	}
}
class cFramedOpening {
	width: number
	height: number
	exhaustFanOrLouver: boolean
	weatherHood: boolean
	exhaust: string
	hood: string
	distanceFromLeft: number
	openingType: string

	constructor(
		width: number,
		height: number,
		exhaustFanOrLouver: boolean,
		weatherHood: boolean,
		exhaust: string,
		hood: string,
		distanceFromLeft: number,
		openingType: string
	) {
		this.width = width
		this.height = height
		this.exhaustFanOrLouver = exhaustFanOrLouver
		this.exhaust = exhaust
		this.weatherHood = weatherHood
		this.hood = hood
		this.distanceFromLeft = distanceFromLeft
		this.openingType = openingType
	}
}
class cPanel {
	liner: boolean
	length: number
	shape: string
	type: string
	color: string

	constructor(liner: boolean, length: number, shape: string, type: string, color: string) {
		this.liner = liner
		this.length = length
		this.shape = shape
		this.type = type
		this.color = color
	}
}
class cBay {
	length: number
	constructor(length: number) {
		this.length = length
	}
}

//with functions this class shouldn't be necessary
class cExtension {
	constructor() {}
}
