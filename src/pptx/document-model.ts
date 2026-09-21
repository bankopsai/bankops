/** PPTX document model: plain data classes mirroring the OOXML structure. Node-only plumbing. */

export const DEFAULT_SLIDE_WIDTH = 9144000;
export const DEFAULT_SLIDE_HEIGHT = 6858000;

export class Presentation {
  slideSize: { cx: number; cy: number };
  slides: Slide[];
  slideMasters: SlideMaster[];
  slideLayouts: SlideLayout[];
  theme: Theme | null;
  media: Record<string, any>;
  contentTypes: any[];
  relationships: any[];
  defaultTextStyle: any;
  /** Native chart parts (zip path -> content) injected by the chart pipeline. */
  _chartParts?: Record<string, any>;
  _chartContentTypes?: Array<{ partName: string; contentType: string }>;

  constructor() {
    this.slideSize = { cx: DEFAULT_SLIDE_WIDTH, cy: DEFAULT_SLIDE_HEIGHT };
    this.slides = [];
    this.slideMasters = [];
    this.slideLayouts = [];
    this.theme = null;
    this.media = {};
    this.contentTypes = [];
    this.relationships = [];
    this.defaultTextStyle = null;
  }
}

export class Theme {
  name: string;
  clrScheme: Record<string, string>;
  fontScheme: {
    name: string;
    majorFont: { latin: string; ea: string; cs: string };
    minorFont: { latin: string; ea: string; cs: string };
  };

  constructor(name?: string) {
    this.name = name || 'Office Theme';
    this.clrScheme = {
      name: 'Office',
      dk1: '#000000', lt1: '#FFFFFF',
      dk2: '#44546A', lt2: '#E7E6E6',
      accent1: '#4472C4', accent2: '#ED7D31',
      accent3: '#A5A5A5', accent4: '#FFC000',
      accent5: '#5B9BD5', accent6: '#70AD47',
      hlink: '#0563C1', folHlink: '#954F72'
    };
    this.fontScheme = {
      name: 'Office',
      majorFont: { latin: 'Calibri Light', ea: '', cs: '' },
      minorFont: { latin: 'Calibri', ea: '', cs: '' }
    };
  }
}

export class SlideMaster {
  id: any;
  name: string;
  shapes: Shape[];
  background: any;
  clrMap: Record<string, string>;
  slideLayouts: SlideLayout[];
  theme: Theme | null;
  relationships: any[];

  constructor(id?: any) {
    this.id = id;
    this.name = '';
    this.shapes = [];
    this.background = null;
    this.clrMap = {
      dk1: 'dk1', lt1: 'lt1', dk2: 'dk2', lt2: 'lt2',
      accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
      accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
      hlink: 'hlink', folHlink: 'folHlink'
    };
    this.slideLayouts = [];
    this.theme = null;
    this.relationships = [];
  }
}

export class SlideLayout {
  id: any;
  name: string;
  type: string;
  shapes: Shape[];
  background: any;
  clrMapOvr: any;
  slideMaster: SlideMaster | null;
  relationships: any[];

  constructor(id?: any) {
    this.id = id;
    this.name = '';
    this.type = '';
    this.shapes = [];
    this.background = null;
    this.clrMapOvr = null;
    this.slideMaster = null;
    this.relationships = [];
  }
}

export class Slide {
  id: any;
  rId: string;
  shapes: Shape[];
  background: any;
  clrMapOvr: any;
  slideLayout: SlideLayout | null;
  notes: any;
  transition: any;
  relationships: any[];

  constructor(id?: any) {
    this.id = id;
    this.rId = '';
    this.shapes = [];
    this.background = null;
    this.clrMapOvr = null;
    this.slideLayout = null;
    this.notes = null;
    this.transition = null;
    this.relationships = [];
  }
}

export class Shape {
  id: number;
  name: string;
  type: string;
  hidden: boolean;
  placeholder: any;
  xfrm: { off: { x: number; y: number }; ext: { cx: number; cy: number }; rot: number; flipH: boolean; flipV: boolean };
  geometry: any;
  fill: any;
  line: any;
  effects: any;
  textBody: TextBody | null;
  style: any;
  imageRId: string | null;
  imageOrigSize: { cx: number; cy: number } | null;
  lockAspectRatio: boolean;
  tableData: TableData | null;
  // Optional properties attached by the parser / writer / builders
  shadow?: any;
  defaultTextColor?: string;
  imageUrl?: string;
  _imagePath?: string;
  _exportRId?: string;
  _cropRect?: { l: number; t: number; r: number; b: number };
  _chartIdx?: number;
  _graphicFrameXml?: string;

  constructor() {
    this.id = 0;
    this.name = '';
    this.type = 'sp';
    this.hidden = false;
    this.placeholder = null;
    this.xfrm = { off: { x: 0, y: 0 }, ext: { cx: 0, cy: 0 }, rot: 0, flipH: false, flipV: false };
    this.geometry = { type: 'rect', avLst: [] };
    this.fill = null;
    this.line = null;
    this.effects = null;
    this.textBody = null;
    this.style = null;
    this.imageRId = null;
    this.imageOrigSize = null;
    this.lockAspectRatio = false;
    this.tableData = null;
  }
}

export class Fill {
  type: string;
  color?: any;
  alpha?: number;
  stops?: any[];
  angle?: number;

  constructor(type: string, props?: any) {
    this.type = type;
    if (type === 'solid') {
      this.color = props && props.color || '#FFFFFF';
      this.alpha = props && props.alpha != null ? props.alpha : 100;
    } else if (type === 'gradient') {
      this.stops = props && props.stops || [];
      this.angle = props && props.angle || 0;
    }
  }
}

export class LineProps {
  width: number;
  color: any;
  alpha: number;
  dash: string;
  cap: string;
  join: string;

  constructor(props?: any) {
    this.width = props && props.width || 12700;
    this.color = props && props.color || '#000000';
    this.alpha = props && props.alpha != null ? props.alpha : 100;
    this.dash = props && props.dash || 'solid';
    this.cap = props && props.cap || 'flat';
    this.join = props && props.join || 'round';
  }
}

export class TextBody {
  bodyPr: {
    wrap: string; lIns: number; tIns: number; rIns: number; bIns: number;
    anchor: string; anchorCtr: boolean; autoFit: string;
    vert: string; rot: number;
  };
  paragraphs: Paragraph[];

  constructor() {
    this.bodyPr = {
      wrap: 'square', lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720,
      anchor: 't', anchorCtr: false, autoFit: 'none',
      vert: 'horz', rot: 0
    };
    this.paragraphs = [];
  }
}

export class Paragraph {
  pPr: any;
  runs: Run[];
  endParaRPr: any;

  constructor() {
    this.pPr = { algn: 'l', lvl: 0, indent: 0, marL: 0, spcBef: 0, spcBefPct: null, spcAft: 0, spcAftPct: null, lnSpc: 100, lnSpcType: 'pct', buNone: true, buChar: null, buClr: null, buFont: null, buSzPct: null, buAutoNum: null };
    this.runs = [];
    this.endParaRPr = null;
  }
}

export class Run {
  text: string;
  rPr: any;

  constructor(text?: string) {
    this.text = text || '';
    this.rPr = { sz: 1800, b: false, i: false, u: 'none', strike: 'noStrike', color: null, highlight: null, fontFamily: null, baseline: 0 };
  }
}

export class TableData {
  cols: Array<{ w: number }>;
  rows: TableRow[];
  tblPr: any;

  constructor() {
    this.cols = [];
    this.rows = [];
    this.tblPr = {};
  }
}

export class TableRow {
  h: number;
  cells: TableCell[];

  constructor(h?: number) {
    this.h = h || 0;
    this.cells = [];
  }
}

export class TableCell {
  txBody: TextBody | null;
  tcPr: {
    marL: number; marR: number; marT: number; marB: number;
    anchor: string;
    fill: any;
    borders: { l: any; r: any; t: any; b: any };
  };
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;

  constructor() {
    this.txBody = null;
    this.tcPr = {
      marL: 45720, marR: 45720, marT: 45720, marB: 45720,
      anchor: 't',
      fill: null,
      borders: { l: null, r: null, t: null, b: null }
    };
    this.gridSpan = 1;
    this.rowSpan = 1;
    this.hMerge = false;
    this.vMerge = false;
  }
}

const DM = {
  Presentation, Theme,
  SlideMaster, SlideLayout, Slide,
  Shape, Fill, LineProps,
  TextBody, Paragraph, Run,
  TableData, TableRow, TableCell,
  DEFAULT_SLIDE_WIDTH, DEFAULT_SLIDE_HEIGHT
};

export default DM;
