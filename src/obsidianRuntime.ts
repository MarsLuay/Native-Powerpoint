import type { MarkdownRenderChild as ObsidianMarkdownRenderChild } from 'obsidian';

type Cleanup = () => void;

interface ComponentLike {
	load: () => void;
	unload: () => void;
}

interface ObsidianRuntime {
	Notice?: new (message: string, timeout?: number) => unknown;
	Platform?: {
		isMacOS?: boolean;
		isMobile?: boolean;
		isMobileApp?: boolean;
	};
	setIcon?: (parent: HTMLElement, iconId: string) => void;
}

const RUNTIME_KEY = '__docxidianObsidianRuntime';

function getRuntime(): ObsidianRuntime {
	return (globalThis as typeof globalThis & { [RUNTIME_KEY]?: ObsidianRuntime })[RUNTIME_KEY] ?? {};
}

export function configureObsidianRuntime(runtime: ObsidianRuntime): void {
	(globalThis as typeof globalThis & { [RUNTIME_KEY]?: ObsidianRuntime })[RUNTIME_KEY] = runtime;
}

export class Notice {
	constructor(message: string, timeout?: number) {
		const RuntimeNotice = getRuntime().Notice;
		if (RuntimeNotice) {
			return new RuntimeNotice(message, timeout) as Notice;
		}

		console.info(`[Native PowerPoint Doc Editor] ${message}`);
	}
}

export const Platform = {
	get isMacOS(): boolean {
		return getRuntime().Platform?.isMacOS === true;
	},
	get isMobile(): boolean {
		return getRuntime().Platform?.isMobile === true;
	},
	get isMobileApp(): boolean {
		return getRuntime().Platform?.isMobileApp === true;
	},
};

export function setIcon(parent: HTMLElement, iconId: string): void {
	const runtimeSetIcon = getRuntime().setIcon;
	if (runtimeSetIcon) {
		runtimeSetIcon(parent, iconId);
		return;
	}

	parent.textContent = iconId;
}

export class Component {
	children: ComponentLike[] = [];
	cleanups: Cleanup[] = [];
	loaded = false;

	addChild<T extends ComponentLike>(child: T): T {
		this.children.push(child);
		if (this.loaded) {
			child.load();
		}

		return child;
	}

	removeChild<T extends ComponentLike>(child: T): T {
		this.children = this.children.filter(existingChild => existingChild !== child);
		child.unload();
		return child;
	}

	register(cleanup: Cleanup): void {
		this.cleanups.push(cleanup);
	}

	registerEvent(eventRef: { detach?: () => void } | { offref?: () => void }): void {
		this.register(() => {
			if ('detach' in eventRef && typeof eventRef.detach === 'function') {
				eventRef.detach();
			}
			if ('offref' in eventRef && typeof eventRef.offref === 'function') {
				eventRef.offref();
			}
		});
	}

	registerDomEvent(
		el: Window | Document | HTMLElement,
		type: string,
		callback: EventListenerOrEventListenerObject,
		options?: boolean | AddEventListenerOptions,
	): void {
		el.addEventListener(type, callback, options);
		this.register(() => el.removeEventListener(type, callback, options));
	}

	registerInterval(id: number): number {
		this.register(() => window.clearInterval(id));
		return id;
	}

	load(): void {
		if (this.loaded) {
			return;
		}

		this.loaded = true;
		this.onload();
		for (const child of this.children) {
			child.load();
		}
	}

	unload(): void {
		if (!this.loaded) {
			return;
		}

		for (const child of [...this.children].reverse()) {
			child.unload();
		}
		this.children = [];

		for (const cleanup of [...this.cleanups].reverse()) {
			cleanup();
		}
		this.cleanups = [];
		this.loaded = false;
		this.onunload();
	}

	onload(): void {}

	onunload(): void {}
}

export class MarkdownRenderChild extends Component {
	constructor(public containerEl: HTMLElement) {
		super();
	}
}

export type MarkdownRenderChildLike = ObsidianMarkdownRenderChild;
