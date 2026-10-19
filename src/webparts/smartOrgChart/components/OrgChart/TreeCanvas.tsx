import * as React from 'react';
import * as strings from 'SmartOrgChartWebPartStrings';
import styles from './OrgChart.module.scss';

interface ITreeCanvasProps {
  defaultZoom: number;
  className: string;
  busy: boolean;
  onZoom: (zoom: number) => void;
  onViewport: (el: HTMLDivElement | null) => void;
}

/** A DOM tree on a map-like camera. Camera updates never re-render its cards. */
export class TreeCanvas extends React.Component<ITreeCanvasProps, { active: boolean }> {
  public state = { active: false };
  private viewport = React.createRef<HTMLDivElement>();
  private content = React.createRef<HTMLDivElement>();
  private x = 0;
  private y = 0;
  private zoom = 1;
  private automatic = true;
  private observer: ResizeObserver | undefined;
  private frame = 0;
  private pointers = new Map<number, { x: number; y: number }>();
  private gesture: { x: number; y: number; distance: number } | undefined;
  private moved = false;
  private suppressClickUntil = 0;
  private touchQuery = window.matchMedia('(pointer: coarse)');

  public componentDidMount(): void {
    const el = this.viewport.current;
    const content = this.content.current;
    if (!el || !content) return;
    this.props.onViewport(el);
    el.addEventListener('wheel', this.wheel, { passive: false });
    this.observer = new ResizeObserver(() => {
      if (this.automatic) this.fit(false);
    });
    this.observer.observe(el);
    this.observer.observe(content);
    this.reset();
  }

  public componentDidUpdate(prev: ITreeCanvasProps): void {
    if (prev.defaultZoom !== this.props.defaultZoom) this.reset();
  }

  public componentWillUnmount(): void {
    this.viewport.current?.removeEventListener('wheel', this.wheel);
    this.observer?.disconnect();
    cancelAnimationFrame(this.frame);
    this.props.onViewport(null);
  }

  private paint(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      if (this.content.current) {
        this.content.current.style.transform = `translate(${this.x}px, ${this.y}px) scale(${this.zoom})`;
      }
    });
  }

  private center(zoom: number): void {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c) return;
    this.zoom = Math.max(0.01, Math.min(1.5, zoom));
    this.x = (v.clientWidth - c.offsetWidth * this.zoom) / 2;
    this.y = (v.clientHeight - c.offsetHeight * this.zoom) / 2;
    this.paint();
    this.props.onZoom(this.zoom);
  }

  public reset = (): void => {
    this.automatic = !(this.props.defaultZoom > 0);
    if (this.automatic) this.fit(false);
    else this.center(this.props.defaultZoom);
  };

  public fit = (manual = true): void => {
    const v = this.viewport.current;
    const c = this.content.current;
    if (!v || !c || !c.offsetWidth || !c.offsetHeight) return;
    if (manual) this.automatic = true;
    this.center(Math.min(1, Math.max(1, v.clientWidth - 48) / c.offsetWidth,
      Math.max(1, v.clientHeight - 48) / c.offsetHeight));
  };

  public autoFit = (): void => { if (this.automatic) this.fit(false); };

  public setZoom = (zoom: number, anchorX?: number, anchorY?: number): void => {
    const v = this.viewport.current;
    if (!v || !Number.isFinite(zoom)) return;
    const next = Math.max(0.01, Math.min(1.5, zoom));
    const ax = anchorX === undefined ? v.clientWidth / 2 : anchorX;
    const ay = anchorY === undefined ? v.clientHeight / 2 : anchorY;
    this.x = ax - (ax - this.x) * next / this.zoom;
    this.y = ay - (ay - this.y) * next / this.zoom;
    this.zoom = next;
    this.automatic = false;
    this.paint();
    this.props.onZoom(next);
  };

  public reveal = (card: HTMLElement): void => {
    // Flush pending transforms before measuring screen coordinates.
    const c = this.content.current;
    const v = this.viewport.current;
    if (!v || !c) return;
    c.style.transform = `translate(${this.x}px, ${this.y}px) scale(${this.zoom})`;
    const a = card.getBoundingClientRect();
    const b = v.getBoundingClientRect();
    this.x += b.left + b.width / 2 - (a.left + a.width / 2);
    this.y += b.top + b.height / 2 - (a.top + a.height / 2);
    this.automatic = false;
    this.paint();
  };

  private wheel = (e: WheelEvent): void => {
    if (this.touchQuery.matches && !this.state.active) return;
    if ((e.target as HTMLElement).closest('button, input, a')) return;
    e.preventDefault();
    const viewport = this.viewport.current;
    if (!viewport) return;
    const r = viewport.getBoundingClientRect();
    const delta = e.deltaY * (e.deltaMode === 1 ? 16 : e.deltaMode === 2 ? r.height : 1);
    this.setZoom(this.zoom * Math.exp(-Math.max(-100, Math.min(100, delta)) * 0.002), e.clientX - r.left, e.clientY - r.top);
  };

  private measureGesture(): { x: number; y: number; distance: number } {
    const points: Array<{ x: number; y: number }> = [];
    this.pointers.forEach(p => points.push(p));
    const a = points[0];
    const b = points[1] || a;
    return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2,
      distance: Math.sqrt((a.x - b.x) ** 2 + (a.y - b.y) ** 2) };
  }

  private pointerDown = (e: React.PointerEvent<HTMLDivElement>): void => {
    if ((e.pointerType !== 'mouse' || this.touchQuery.matches) && !this.state.active) return;
    if (e.button !== 0 || (e.target as HTMLElement).closest('button, a, input')) return;
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    this.gesture = this.measureGesture();
    if (this.pointers.size === 1) this.moved = false;
    // Capture only after a drag threshold, so taps still reach the card.
  };

  private pointerMove = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!this.pointers.has(e.pointerId) || !this.gesture) return;
    this.pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
    const next = this.measureGesture();
    const dx = next.x - this.gesture.x;
    const dy = next.y - this.gesture.y;
    if (!this.moved && this.pointers.size < 2 && Math.sqrt(dx * dx + dy * dy) < 5) return;
    e.currentTarget.setPointerCapture(e.pointerId);
    this.moved = true;
    this.automatic = false;
    if (this.gesture.distance > 0 && next.distance > 0) {
      const viewport = this.viewport.current;
      if (!viewport) return;
      const r = viewport.getBoundingClientRect();
      this.setZoom(this.zoom * next.distance / this.gesture.distance, this.gesture.x - r.left, this.gesture.y - r.top);
    }
    this.x += dx;
    this.y += dy;
    this.gesture = next;
    this.paint();
    e.currentTarget.style.cursor = 'grabbing';
  };

  private pointerEnd = (e: React.PointerEvent<HTMLDivElement>): void => {
    if (!this.pointers.delete(e.pointerId)) return;
    if (this.moved) this.suppressClickUntil = Date.now() + 350;
    if (e.currentTarget.hasPointerCapture(e.pointerId)) e.currentTarget.releasePointerCapture(e.pointerId);
    this.gesture = this.pointers.size ? this.measureGesture() : undefined;
    if (!this.pointers.size) e.currentTarget.style.cursor = '';
  };

  private deactivate = (): void => {
    const v = this.viewport.current;
    this.pointers.forEach((_p, id) => { if (v?.hasPointerCapture(id)) v.releasePointerCapture(id); });
    this.pointers.clear();
    this.gesture = undefined;
    if (v) v.style.cursor = '';
    this.setState({ active: false });
  };

  private focusVisible = (e: React.FocusEvent<HTMLDivElement>): void => {
    const card = (e.target as HTMLElement).closest('[data-soc-id]') as HTMLElement | null;
    const v = this.viewport.current;
    if (!card || !v) return;
    const a = card.getBoundingClientRect();
    const b = v.getBoundingClientRect();
    if ((e.target as HTMLElement).matches(':focus-visible') &&
      (a.left < b.left || a.right > b.right || a.top < b.top || a.bottom > b.bottom)) this.reveal(card);
    // Native focus scrolling must not create a second, invisible camera.
    requestAnimationFrame(() => { if (this.viewport.current) {
      this.viewport.current.scrollLeft = 0;
      this.viewport.current.scrollTop = 0;
    } });
  };

  public render(): React.ReactElement {
    return <div ref={this.viewport}
      className={`${this.props.className} ${styles.treeCanvas} ${this.state.active ? styles.canvasActive : ''}`}
      aria-busy={this.props.busy}
      onPointerDown={this.pointerDown} onPointerMove={this.pointerMove}
      onPointerUp={this.pointerEnd} onPointerCancel={this.pointerEnd} onLostPointerCapture={this.pointerEnd}
      onPointerLeave={e => { if (!e.currentTarget.hasPointerCapture(e.pointerId)) this.pointerEnd(e); }}
      onKeyDown={e => { if (e.key === 'Escape') this.deactivate(); }}
      onFocusCapture={this.focusVisible}
      onDragStart={e => e.preventDefault()}
      onClickCapture={e => {
        if (!(e.target as HTMLElement).closest(`.${styles.canvasTouchControls}`) && Date.now() < this.suppressClickUntil) {
          e.preventDefault(); e.stopPropagation();
        }
      }}>
      <div ref={this.content} className={styles.canvasContent}>{this.props.children}</div>
      <div className={styles.canvasTouchControls}>
        <button type="button" aria-pressed={this.state.active}
          onClick={this.state.active ? this.deactivate : () => this.setState({ active: true })}>
          {this.state.active ? strings.Chart_CanvasDone : strings.Chart_CanvasActivate}
        </button>
        {!this.state.active && <span>{strings.Chart_CanvasHint}</span>}
      </div>
      {!this.state.active && <div className={styles.canvasTouchShield} aria-hidden="true" />}
    </div>;
  }
}
