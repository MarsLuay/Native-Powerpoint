export interface PowerPointPresentOptions {
  ownerDocument: Document;
  slideCount: number;
  startIndex: number;
  renderSlide: (index: number) => SVGSVGElement | null;
  onExit?: (lastIndex: number) => void;
}

/**
 * Fullscreen slideshow controller for the Native PowerPoint view. It owns a
 * black letterboxed overlay appended to the document body, renders the current
 * slide SVG sized to fit the screen, and handles keyboard / click navigation.
 * All listeners are torn down on exit so nothing leaks back into the editor.
 */
export class PowerPointPresentController {
  private readonly ownerDocument: Document;
  private readonly slideCount: number;
  private readonly renderSlide: (index: number) => SVGSVGElement | null;
  private readonly onExit?: (lastIndex: number) => void;

  private index: number;
  private overlay: HTMLElement | null = null;
  private stage: HTMLElement | null = null;
  private counterEl: HTMLElement | null = null;
  private disposed = false;
  private readonly cleanups: Array<() => void> = [];

  constructor(options: PowerPointPresentOptions) {
    this.ownerDocument = options.ownerDocument;
    this.slideCount = options.slideCount;
    this.renderSlide = options.renderSlide;
    this.onExit = options.onExit;
    this.index = Math.min(Math.max(0, options.startIndex), Math.max(0, options.slideCount - 1));
  }

  start(): void {
    if (this.disposed || this.slideCount <= 0) return;

    const body = this.ownerDocument.body;
    const overlay = body.createDiv({ cls: 'native-powerpoint-present-overlay' });
    overlay.tabIndex = -1;
    this.overlay = overlay;

    this.stage = overlay.createDiv({ cls: 'native-powerpoint-present-stage' });

    const hud = overlay.createDiv({ cls: 'native-powerpoint-present-hud' });
    this.counterEl = hud.createDiv({ cls: 'native-powerpoint-present-counter' });
    const exitButton = hud.createEl('button', {
      cls: 'native-powerpoint-present-exit',
      text: 'Exit',
      attr: { 'aria-label': 'Exit slideshow (Esc)', title: 'Exit slideshow (Esc)' }
    });
    exitButton.addEventListener('click', (event) => {
      event.stopPropagation();
      this.dispose();
    });

    this.addListener(overlay, 'click', (event) => {
      const target = event.target as HTMLElement | null;
      if (target && target.closest('.native-powerpoint-present-hud')) return;
      this.next();
    });
    this.addListener(overlay, 'contextmenu', (event) => {
      event.preventDefault();
      this.previous();
    });
    this.addListener(this.ownerDocument, 'keydown', (event) => this.handleKeydown(event as KeyboardEvent), true);
    this.addListener(this.ownerDocument, 'fullscreenchange', () => this.handleFullscreenChange());

    this.renderCurrent();

    const requestFullscreen = overlay.requestFullscreen?.bind(overlay);
    if (requestFullscreen) {
      void Promise.resolve(requestFullscreen()).catch(() => {
        // Fullscreen can be rejected (e.g. no user gesture); the overlay still
        // covers the viewport, so presentation continues without it.
      });
    }

    overlay.focus({ preventScroll: true });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;

    for (const cleanup of this.cleanups.splice(0)) {
      cleanup();
    }

    if (this.ownerDocument.fullscreenElement === this.overlay) {
      void Promise.resolve(this.ownerDocument.exitFullscreen?.()).catch(() => {
        // Ignore: the overlay is being removed regardless.
      });
    }

    this.overlay?.remove();
    this.overlay = null;
    this.stage = null;
    this.counterEl = null;

    this.onExit?.(this.index);
  }

  private addListener<K extends keyof DocumentEventMap>(
    target: Document,
    type: K,
    handler: (event: DocumentEventMap[K]) => void,
    capture?: boolean
  ): void;
  private addListener(
    target: HTMLElement,
    type: string,
    handler: (event: Event) => void,
    capture?: boolean
  ): void;
  private addListener(
    target: Document | HTMLElement,
    type: string,
    handler: (event: Event) => void,
    capture = false
  ): void {
    target.addEventListener(type, handler as EventListener, capture);
    this.cleanups.push(() => target.removeEventListener(type, handler as EventListener, capture));
  }

  private handleKeydown(event: KeyboardEvent): void {
    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
      case 'PageDown':
      case ' ':
      case 'Spacebar':
      case 'Enter':
        event.preventDefault();
        this.next();
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
      case 'Backspace':
        event.preventDefault();
        this.previous();
        break;
      case 'Home':
        event.preventDefault();
        this.goTo(0);
        break;
      case 'End':
        event.preventDefault();
        this.goTo(this.slideCount - 1);
        break;
      case 'Escape':
        event.preventDefault();
        this.dispose();
        break;
      default:
        break;
    }
  }

  private handleFullscreenChange(): void {
    if (this.disposed) return;
    // The user pressed Esc / the system exited fullscreen: close the overlay.
    if (this.overlay && this.ownerDocument.fullscreenElement !== this.overlay) {
      this.dispose();
    }
  }

  private next(): void {
    this.goTo(this.index + 1);
  }

  private previous(): void {
    this.goTo(this.index - 1);
  }

  private goTo(index: number): void {
    const clamped = Math.min(Math.max(0, index), this.slideCount - 1);
    if (clamped === this.index) return;
    this.index = clamped;
    this.renderCurrent();
  }

  private renderCurrent(): void {
    if (!this.stage) return;

    this.stage.empty();
    const svg = this.renderSlide(this.index);
    if (svg) {
      svg.addClass('native-powerpoint-present-slide');
      svg.removeAttribute('width');
      svg.removeAttribute('height');
      this.stage.appendChild(svg);
    } else {
      this.stage.createDiv({
        cls: 'native-powerpoint-present-error',
        text: 'This slide could not be rendered for the slideshow.'
      });
    }

    this.counterEl?.setText(`${this.index + 1} / ${this.slideCount}`);
  }
}
