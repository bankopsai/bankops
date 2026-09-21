/**
 * SlideLayout — 12-column nested grid engine (Bootstrap/MUI style).
 * Computes absolute EMU bounds for a slide spec. Pure and isomorphic:
 * the browser preview and the PPTX writer consume the same LayoutTree.
 */

import type { Bounds, GridNode, LayoutNodeResult, LayoutTree, SlideSpec } from "./types.js";

const EMU_PER_INCH = 914400;

export interface SlideLayoutConfig {
  slideWidth?: number;        // EMU (default 10in)
  slideHeight?: number;       // EMU (default 7.5in)
  margin?: number;            // inches (default 0.5)
  sectionLabelHeight?: number;
  titleHeight?: number;
  footerHeight?: number;
  defaultGap?: number;        // inches
  headerBarHeight?: number;
  subheaderHeight?: number;
  subfooterHeight?: number;
  columns?: number;
}

export interface LayoutPreset { body: GridNode }

export class SlideLayout {
  slideWidth: number;
  slideHeight: number;
  margin: number;
  sectionLabelHeight: number;
  titleHeight: number;
  footerHeight: number;
  defaultGap: number;
  headerBarHeight: number;
  headerBarGap: number;
  subheaderHeight: number;
  subheaderLineGap: number;
  subheaderGap: number;
  subfooterHeight: number;
  subfooterGap: number;
  columns: number;

  constructor(config: SlideLayoutConfig = {}) {
    const c = config;
    this.slideWidth = c.slideWidth || 9144000;
    this.slideHeight = c.slideHeight || 6858000;
    this.margin = Math.round((c.margin != null ? c.margin : 0.5) * EMU_PER_INCH);
    this.sectionLabelHeight = Math.round((c.sectionLabelHeight != null ? c.sectionLabelHeight : 0.3) * EMU_PER_INCH);
    this.titleHeight = Math.round((c.titleHeight != null ? c.titleHeight : 0.55) * EMU_PER_INCH);
    this.footerHeight = Math.round((c.footerHeight != null ? c.footerHeight : 0.35) * EMU_PER_INCH);
    this.defaultGap = c.defaultGap != null ? c.defaultGap : 0.15;
    this.headerBarHeight = Math.round((c.headerBarHeight != null ? c.headerBarHeight : 0.35) * EMU_PER_INCH);
    this.headerBarGap = Math.round(0.08 * EMU_PER_INCH);
    this.subheaderHeight = c.subheaderHeight != null ? Math.round(c.subheaderHeight * EMU_PER_INCH) : Math.round(0.22 * EMU_PER_INCH);
    this.subheaderLineGap = Math.round(0.02 * EMU_PER_INCH);
    this.subheaderGap = Math.round(0.06 * EMU_PER_INCH);
    this.subfooterHeight = Math.round((c.subfooterHeight != null ? c.subfooterHeight : 0.22) * EMU_PER_INCH);
    this.subfooterGap = Math.round(0.05 * EMU_PER_INCH);
    this.columns = c.columns || 12;
  }

  /** Compute the layout tree for a slide spec. */
  compute(spec: SlideSpec): LayoutTree {
    const result: LayoutTree = { sectionLabelBounds: null, titleBounds: null, footerBounds: null, body: null };

    const x = this.margin;
    let y = this.margin;
    const cx = this.slideWidth - 2 * this.margin;
    let bottom = this.slideHeight - this.margin;

    if (spec.sectionLabel) {
      result.sectionLabelBounds = { x, y, cx, cy: this.sectionLabelHeight };
      y += this.sectionLabelHeight;
    }
    if (spec.title) {
      result.titleBounds = { x, y, cx, cy: this.titleHeight };
      y += this.titleHeight + Math.round(0.1 * EMU_PER_INCH);
    }
    if (spec.footer) {
      result.footerBounds = { x, y: bottom - this.footerHeight, cx, cy: this.footerHeight };
      bottom -= this.footerHeight + Math.round(0.05 * EMU_PER_INCH);
    }

    const bodyArea: Bounds = { x, y, cx, cy: Math.max(0, bottom - y) };
    if (spec.body) {
      result.body = this._layoutNode(spec.body, bodyArea);
    } else {
      result.body = { bounds: bodyArea, children: null, content: null } as unknown as LayoutNodeResult;
    }
    return result;
  }

  _layoutNode(node: GridNode, available: Bounds): LayoutNodeResult {
    if (node.margin || node.marginTop != null || node.marginRight != null || node.marginBottom != null || node.marginLeft != null) {
      const mBase = node.margin || 0;
      const mT = Math.round((node.marginTop != null ? node.marginTop : mBase) * EMU_PER_INCH);
      const mR = Math.round((node.marginRight != null ? node.marginRight : mBase) * EMU_PER_INCH);
      const mB = Math.round((node.marginBottom != null ? node.marginBottom : mBase) * EMU_PER_INCH);
      const mL = Math.round((node.marginLeft != null ? node.marginLeft : mBase) * EMU_PER_INCH);
      available = {
        x: available.x + mL,
        y: available.y + mT,
        cx: Math.max(0, available.cx - mL - mR),
        cy: Math.max(0, available.cy - mT - mB),
      };
    }

    const result: LayoutNodeResult = {
      bounds: { x: available.x, y: available.y, cx: available.cx, cy: available.cy },
      background: node.background || null,
      backgroundOpacity: node.backgroundOpacity != null ? node.backgroundOpacity : null,
      border: node.border || null,
      borderTop: node.borderTop || null,
      borderRight: node.borderRight || null,
      borderBottom: node.borderBottom || null,
      borderLeft: node.borderLeft || null,
      shadow: node.shadow || null,
      childShape: node.childShape || null,
      childShapeColor: node.childShapeColor || null,
      header: node.header || null,
      headerBounds: null,
      subheader: node.subheader || null,
      subheaderBounds: null,
      subheaderLineBounds: null,
      subfooter: node.subfooter || null,
      subfooterBounds: null,
      contentBounds: null,
      content: node.content || null,
      children: null,
      _node: node,
    };

    let topY = available.y;
    let bottomY = available.y + available.cy;

    if (node.header) {
      result.headerBounds = { x: available.x, y: topY, cx: available.cx, cy: this.headerBarHeight };
      topY += this.headerBarHeight + this.headerBarGap;
    }
    if (node.subheader) {
      result.subheaderBounds = { x: available.x, y: topY, cx: available.cx, cy: this.subheaderHeight };
      const lineY = topY + this.subheaderHeight + this.subheaderLineGap;
      result.subheaderLineBounds = { x: available.x, y: lineY, cx: available.cx, cy: 0 };
      topY = lineY + this.subheaderGap;
    }
    if (node.subfooter) {
      result.subfooterBounds = { x: available.x, y: bottomY - this.subfooterHeight, cx: available.cx, cy: this.subfooterHeight };
      bottomY -= this.subfooterHeight + this.subfooterGap;
    }

    const contentArea: Bounds = { x: available.x, y: topY, cx: available.cx, cy: Math.max(0, bottomY - topY) };
    if (node.content) result.contentBounds = contentArea;

    if (!node.children || node.children.length === 0) {
      if (!result.contentBounds) result.contentBounds = contentArea;
      return result;
    }

    const direction = node.direction || "row";
    const gapInches = node.gap != null ? node.gap : this.defaultGap;
    const gap = Math.round(gapInches * EMU_PER_INCH);
    const pBase = node.padding != null ? node.padding : 0;
    const pT = Math.round((node.paddingTop != null ? node.paddingTop : pBase) * EMU_PER_INCH);
    const pR = Math.round((node.paddingRight != null ? node.paddingRight : pBase) * EMU_PER_INCH);
    const pB = Math.round((node.paddingBottom != null ? node.paddingBottom : pBase) * EMU_PER_INCH);
    const pL = Math.round((node.paddingLeft != null ? node.paddingLeft : pBase) * EMU_PER_INCH);

    const inner: Bounds = {
      x: contentArea.x + pL,
      y: contentArea.y + pT,
      cx: Math.max(0, contentArea.cx - pL - pR),
      cy: Math.max(0, contentArea.cy - pT - pB),
    };

    const children = node.children;
    const count = children.length;

    let totalExplicit = 0;
    let autoCount = 0;
    for (const child of children) {
      if (child.span != null && child.span > 0) totalExplicit += child.span;
      else autoCount++;
    }
    const autoSpan = autoCount > 0 ? Math.max(1, (this.columns - totalExplicit) / autoCount) : 0;
    let totalSpan = totalExplicit + autoCount * autoSpan;
    if (totalSpan <= 0) totalSpan = this.columns;

    const totalGapSpace = gap * Math.max(0, count - 1);
    const isRow = direction === "row";
    let availableSize = (isRow ? inner.cx : inner.cy) - totalGapSpace;
    if (availableSize < 0) availableSize = 0;
    const unitSize = availableSize / totalSpan;

    result.children = [];
    let offset = 0;
    for (const child of children) {
      const span = child.span != null && child.span > 0 ? child.span : autoSpan;
      const childSize = Math.round(unitSize * span);
      const childBounds: Bounds = isRow
        ? { x: inner.x + offset, y: inner.y, cx: childSize, cy: inner.cy }
        : { x: inner.x, y: inner.y + offset, cx: inner.cx, cy: childSize };
      offset += childSize + gap;
      result.children.push(this._layoutNode(child, childBounds));
    }
    return result;
  }

  /** All leaf nodes of a computed tree (depth-first). */
  flattenLeaves(layoutTree: LayoutTree): LayoutNodeResult[] {
    const leaves: LayoutNodeResult[] = [];
    this._collectLeaves(layoutTree.body, leaves);
    return leaves;
  }

  _collectLeaves(node: LayoutNodeResult | null, leaves: LayoutNodeResult[]): void {
    if (!node) return;
    if (!node.children || node.children.length === 0) { leaves.push(node); return; }
    for (const child of node.children) this._collectLeaves(child, leaves);
  }

  /** All nodes (containers and leaves) with a `_depth` annotation. */
  flattenAll(layoutTree: LayoutTree, depth = 0): LayoutNodeResult[] {
    const nodes: LayoutNodeResult[] = [];
    this._collectAll(layoutTree.body, nodes, depth);
    return nodes;
  }

  _collectAll(node: LayoutNodeResult | null, nodes: LayoutNodeResult[], depth: number): void {
    if (!node) return;
    (node as any)._depth = depth;
    nodes.push(node);
    if (node.children) for (const child of node.children) this._collectAll(child, nodes, depth + 1);
  }

  // ===== Presets: common body layouts =====
  static presets: Record<string, LayoutPreset> = {
    twoColumn: { body: { direction: "row", children: [{ span: 6 }, { span: 6 }] } },
    threeColumn: { body: { direction: "row", children: [{ span: 4 }, { span: 4 }, { span: 4 }] } },
    quadrant: {
      body: { direction: "row", children: [
        { span: 6, direction: "col", children: [{ span: 6 }, { span: 6 }] },
        { span: 6, direction: "col", children: [{ span: 6 }, { span: 6 }] },
      ] },
    },
    sidebar: { body: { direction: "row", children: [{ span: 8 }, { span: 4 }] } },
    topBottom: { body: { direction: "col", children: [{ span: 4 }, { span: 8 }] } },
    wideTop: {
      body: { direction: "col", children: [
        { span: 3 },
        { span: 9, direction: "row", children: [{ span: 6 }, { span: 6 }] },
      ] },
    },
    dashboard: {
      body: { direction: "col", children: [
        { span: 3, direction: "row", children: [{ span: 3 }, { span: 3 }, { span: 3 }, { span: 3 }] },
        { span: 9, direction: "row", children: [
          { span: 6, direction: "col", children: [{ span: 6 }, { span: 6 }] },
          { span: 6 },
        ] },
      ] },
    },
  };

  /** Deep-cloned preset spec with optional title/sectionLabel/footer. */
  static preset(name: string, overrides?: Pick<SlideSpec, "title" | "sectionLabel" | "footer">): SlideSpec | null {
    const p = SlideLayout.presets[name];
    if (!p) return null;
    const spec: SlideSpec = JSON.parse(JSON.stringify(p));
    if (overrides) {
      if (overrides.title) spec.title = overrides.title;
      if (overrides.sectionLabel) spec.sectionLabel = overrides.sectionLabel;
      if (overrides.footer) spec.footer = overrides.footer;
    }
    return spec;
  }
}

export default SlideLayout;
