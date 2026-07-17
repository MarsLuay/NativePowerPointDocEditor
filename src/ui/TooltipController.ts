import {
	containsEventTarget,
	getToolbarTooltipText,
	positionToolbarTooltip,
	restoreNativeTitle,
	suspendNativeTitle,
	TOOLBAR_TOOLTIP_CLASS,
	TOOLBAR_TOOLTIP_DELAY_MS,
} from './tooltipUtils';

export type TooltipTargetResolver = (target: EventTarget | null) => HTMLElement | null;

interface TooltipWindow {
	innerWidth: number;
	addEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions): void;
	removeEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions): void;
	setTimeout(handler: TimerHandler, timeout?: number, ...arguments_: unknown[]): number;
	clearTimeout(id: number | undefined): void;
}

export interface TooltipControllerOptions {
	root: HTMLElement | Document;
	getTarget: TooltipTargetResolver;
	className?: string;
	delayMs?: number;
	ownerDocument?: Document;
	view?: TooltipWindow;
	onTargetRecognized?: (target: HTMLElement) => void;
}

const TOOLTIP_ID_PREFIX = 'native-powerpoint-doc-editor-tooltip';

let tooltipIdCounter = 0;

function isDocumentRoot(root: HTMLElement | Document): root is Document {
	return root.nodeType === 9;
}

function getRootDocument(root: HTMLElement | Document): Document {
	return isDocumentRoot(root) ? root : root.ownerDocument;
}

function getRootWindow(ownerDocument: Document, view?: TooltipWindow): TooltipWindow {
	const resolvedView = view ?? ownerDocument.defaultView ?? (typeof window !== 'undefined' ? window : null);
	if (!resolvedView) {
		throw new Error('TooltipController requires a window-like view');
	}
	return resolvedView;
}

export class TooltipController {
	private readonly root: HTMLElement | Document;
	private readonly getTarget: TooltipTargetResolver;
	private readonly ownerDocument: Document;
	private readonly view: TooltipWindow;
	private readonly className: string;
	private readonly delayMs: number;
	private readonly onTargetRecognized?: (target: HTMLElement) => void;
	private readonly describedByBeforeTooltip = new WeakMap<HTMLElement, string | null>();
	private activeTarget: HTMLElement | null = null;
	private tooltipEl: HTMLDivElement | null = null;
	private tooltipTimer: number | null = null;
	private disposed = false;

	constructor(options: TooltipControllerOptions) {
		this.root = options.root;
		this.getTarget = options.getTarget;
		this.ownerDocument = options.ownerDocument ?? getRootDocument(options.root);
		this.view = getRootWindow(this.ownerDocument, options.view);
		this.className = options.className ?? TOOLBAR_TOOLTIP_CLASS;
		this.delayMs = options.delayMs ?? TOOLBAR_TOOLTIP_DELAY_MS;
		this.onTargetRecognized = options.onTargetRecognized;
		this.attach();
	}

	detach(): void {
		if (this.disposed) {
			return;
		}
		this.disposed = true;
		this.root.removeEventListener('pointerover', this.handlePointerOver, true);
		this.root.removeEventListener('pointerout', this.handlePointerOut, true);
		this.root.removeEventListener('pointerdown', this.handlePointerDown, true);
		this.root.removeEventListener('scroll', this.handleScrollOrResize, true);
		this.view.removeEventListener('resize', this.handleScrollOrResize, false);
		this.view.removeEventListener('blur', this.handleBlur, false);
		this.hideTooltip();
	}

	private attach(): void {
		this.root.addEventListener('pointerover', this.handlePointerOver, true);
		this.root.addEventListener('pointerout', this.handlePointerOut, true);
		this.root.addEventListener('pointerdown', this.handlePointerDown, true);
		this.root.addEventListener('scroll', this.handleScrollOrResize, true);
		this.view.addEventListener('resize', this.handleScrollOrResize, false);
		this.view.addEventListener('blur', this.handleBlur, false);
	}

	private clearTooltipTimer(): void {
		if (this.tooltipTimer !== null) {
			this.view.clearTimeout(this.tooltipTimer);
			this.tooltipTimer = null;
		}
	}

	private removeTooltipElement(): void {
		this.tooltipEl?.remove();
		this.tooltipEl = null;
	}

	private hideTooltip(): void {
		this.clearTooltipTimer();
		this.removeTooltipElement();
		this.restoreDescribedBy(this.activeTarget);
		restoreNativeTitle(this.activeTarget);
		this.activeTarget = null;
	}

	private scheduleTooltip(target: HTMLElement): void {
		if (target === this.activeTarget) {
			return;
		}

		this.hideTooltip();
		this.activeTarget = target;
		suspendNativeTitle(target);
		this.onTargetRecognized?.(target);
		this.tooltipTimer = this.view.setTimeout(() => {
			this.tooltipTimer = null;
			if (this.activeTarget === target) {
				this.showTooltip(target);
			}
		}, this.delayMs);
	}

	private removeExistingTooltipElements(): void {
		this.ownerDocument.querySelectorAll<HTMLElement>(`.${this.className}`).forEach((element) => {
			if (element !== this.tooltipEl) {
				element.remove();
			}
		});
	}

	private showTooltip(target: HTMLElement): void {
		const label = getToolbarTooltipText(target);
		if (!label || target.isConnected === false) {
			this.hideTooltip();
			return;
		}

		this.removeTooltipElement();
		this.removeExistingTooltipElements();

		const tooltip = this.ownerDocument.createElement('div');
		tooltip.className = this.className;
		tooltip.id = `${TOOLTIP_ID_PREFIX}-${++tooltipIdCounter}`;
		tooltip.setAttribute('role', 'tooltip');
		tooltip.textContent = label;
		this.ownerDocument.body.appendChild(tooltip);
		this.tooltipEl = tooltip;
		this.describeTarget(target, tooltip.id);
		positionToolbarTooltip(target, tooltip, { viewportWidth: this.view.innerWidth });
	}

	private describeTarget(target: HTMLElement, tooltipId: string): void {
		if (!this.describedByBeforeTooltip.has(target)) {
			this.describedByBeforeTooltip.set(target, target.getAttribute('aria-describedby'));
		}
		const current = target.getAttribute('aria-describedby');
		const ids = new Set((current ?? '').split(/\s+/).filter(Boolean));
		ids.add(tooltipId);
		target.setAttribute('aria-describedby', Array.from(ids).join(' '));
	}

	private restoreDescribedBy(target: HTMLElement | null): void {
		if (!target || !this.describedByBeforeTooltip.has(target)) {
			return;
		}
		const previous = this.describedByBeforeTooltip.get(target);
		if (previous) {
			target.setAttribute('aria-describedby', previous);
		} else {
			target.removeAttribute('aria-describedby');
		}
		this.describedByBeforeTooltip.delete(target);
	}

	private readonly handlePointerOver = (event: Event): void => {
		const target = this.getTarget(event.target);
		if (target) {
			this.scheduleTooltip(target);
		}
	};

	private readonly handlePointerOut = (event: Event): void => {
		if (!this.activeTarget || containsEventTarget(this.activeTarget, (event as PointerEvent).relatedTarget)) {
			return;
		}
		this.hideTooltip();
	};

	private readonly handlePointerDown = (event: Event): void => {
		if (this.getTarget(event.target)) {
			this.hideTooltip();
		}
	};

	private readonly handleScrollOrResize = (): void => {
		this.hideTooltip();
	};

	private readonly handleBlur = (): void => {
		this.hideTooltip();
	};
}
