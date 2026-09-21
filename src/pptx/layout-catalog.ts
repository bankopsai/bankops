// ===== Layout Catalog Data Structures =====
// Plain data classes for layout templates discovered by analyzing reference .pptx files.

export interface EmuBounds { x: number; y: number; cx: number; cy: number; }

// Layout type constants
// (typed loosely so lookups of unknown keys — e.g. CONTENT_TYPES.BULLET_LIST in the
// analyzer — keep the original JS semantics of resolving to undefined)
export const LAYOUT_TYPES: Record<string, string> = {
  QUADRANT: 'quadrant',
  THREE_COLUMN: 'threeColumn',
  TWO_PANEL: 'twoPanel',
  TIMELINE: 'timeline',
  CARD_GRID: 'cardGrid',
  DASHBOARD: 'dashboard',
  CONTENT: 'content',
  TABLE: 'table',
  SECTION_DIVIDER: 'sectionDivider',
  TITLE: 'title',
  IMAGE_HEAVY: 'imageHeavy'
};

// Zone role constants
export const ZONE_ROLES: Record<string, string> = {
  TOP_LEFT: 'topLeft',
  TOP_RIGHT: 'topRight',
  BOTTOM_LEFT: 'bottomLeft',
  BOTTOM_RIGHT: 'bottomRight',
  LEFT: 'left',
  CENTER: 'center',
  RIGHT: 'right',
  TOP: 'top',
  BOTTOM: 'bottom',
  FULL: 'full'
};

// Content type constants
export const CONTENT_TYPES: Record<string, string> = {
  STAT_GRID: 'statGrid',
  CARD_GRID: 'cardGrid',
  ICON_GRID: 'iconGrid',
  TABLE: 'table',
  TIMELINE: 'timeline',
  TEXT: 'text',
  IMAGE: 'image',
  PLACEHOLDER: 'placeholder'
};

// ---- LayoutCatalog ----
// Top-level container for all discovered layout templates
export class LayoutCatalog {
  templates: LayoutTemplate[];
  metadata: { sourceFile: string | null; slideCount: number; analyzedAt: string | null };

  constructor() {
    this.templates = [];
    this.metadata = {
      sourceFile: null,
      slideCount: 0,
      analyzedAt: null
    };
  }

  addTemplate(template: LayoutTemplate): void {
    // Merge with existing template of same type, or add new
    for (let i = 0; i < this.templates.length; i++) {
      if (this.templates[i].type === template.type) {
        // Merge: add slide indices, average zone positions, increment frequency
        const existing = this.templates[i];
        existing.frequency += template.frequency;
        for (let si = 0; si < template.slideIndices.length; si++) {
          if (existing.slideIndices.indexOf(template.slideIndices[si]) === -1) {
            existing.slideIndices.push(template.slideIndices[si]);
          }
        }
        // Average zone bounds with existing
        if (template.zones.length > 0 && existing.zones.length === template.zones.length) {
          for (let zi = 0; zi < existing.zones.length; zi++) {
            const ez = existing.zones[zi].bounds;
            const tz = template.zones[zi].bounds;
            const w = existing.frequency - 1; // weight of existing
            ez.x = Math.round((ez.x * w + tz.x) / existing.frequency);
            ez.y = Math.round((ez.y * w + tz.y) / existing.frequency);
            ez.cx = Math.round((ez.cx * w + tz.cx) / existing.frequency);
            ez.cy = Math.round((ez.cy * w + tz.cy) / existing.frequency);
          }
        } else if (template.zones.length > existing.zones.length) {
          existing.zones = template.zones;
        }
        return;
      }
    }
    this.templates.push(template);
  }

  getTemplate(type: string): LayoutTemplate | null {
    for (let i = 0; i < this.templates.length; i++) {
      if (this.templates[i].type === type) return this.templates[i];
    }
    return null;
  }

  getTypes(): string[] {
    return this.templates.map(function(t) { return t.type; });
  }

  toSummary(): any {
    const summary: any = {
      source: this.metadata.sourceFile,
      slideCount: this.metadata.slideCount,
      layouts: []
    };
    for (let i = 0; i < this.templates.length; i++) {
      const t = this.templates[i];
      const layoutInfo: any = {
        type: t.type,
        frequency: t.frequency,
        slideIndices: t.slideIndices,
        zones: []
      };
      for (let zi = 0; zi < t.zones.length; zi++) {
        const z = t.zones[zi];
        layoutInfo.zones.push({
          role: z.role,
          bounds: {
            xIn: +(z.bounds.x / 914400).toFixed(2),
            yIn: +(z.bounds.y / 914400).toFixed(2),
            wIn: +(z.bounds.cx / 914400).toFixed(2),
            hIn: +(z.bounds.cy / 914400).toFixed(2)
          },
          hasHeaderBar: !!z.headerBar,
          contentType: z.contentType
        });
      }
      summary.layouts.push(layoutInfo);
    }
    return summary;
  }
}

// ---- LayoutTemplate ----
// A discovered layout pattern (e.g. "quadrant" with 4 zones)
export class LayoutTemplate {
  type: string;
  zones: ZoneTemplate[];
  slideIndices: number[];
  frequency: number;
  slideSize: { cx: number; cy: number };

  constructor(type: string) {
    this.type = type;
    this.zones = [];
    this.slideIndices = [];
    this.frequency = 1;
    this.slideSize = { cx: 0, cy: 0 };
  }
}

// ---- ZoneTemplate ----
// A rectangular region within a layout (e.g. "topLeft" quadrant)
export class ZoneTemplate {
  role: string;
  bounds: EmuBounds;
  headerBar: ShapeSpec | null;
  contentType: string | null;
  components: ComponentTemplate[];

  constructor(role: string) {
    this.role = role;
    this.bounds = { x: 0, y: 0, cx: 0, cy: 0 };
    this.headerBar = null; // ShapeSpec or null
    this.contentType = null; // string from CONTENT_TYPES
    this.components = [];
  }
}

// ---- ShapeSpec ----
// Specification for a shape (used for header bars, borders, etc.)
export class ShapeSpec {
  bounds: EmuBounds;
  geometry: string;
  fillColor: any;
  lineColor: any;
  lineWidth: number;
  textStyle: any;

  constructor() {
    this.bounds = { x: 0, y: 0, cx: 0, cy: 0 };
    this.geometry = 'roundRect';
    this.fillColor = null;
    this.lineColor = null;
    this.lineWidth = 0;
    this.textStyle = null; // { font, fontSize, bold, color }
  }
}

// ---- ComponentTemplate ----
// A content component within a zone (stat box, card, etc.)
export class ComponentTemplate {
  type: string;
  bounds: EmuBounds;
  textStyle: any;
  items: any[];

  constructor(type: string) {
    this.type = type;
    this.bounds = { x: 0, y: 0, cx: 0, cy: 0 };
    this.textStyle = null;
    this.items = [];
  }
}

const LC = {
  LAYOUT_TYPES,
  ZONE_ROLES,
  CONTENT_TYPES,
  LayoutCatalog,
  LayoutTemplate,
  ZoneTemplate,
  ShapeSpec,
  ComponentTemplate
};

export default LC;
