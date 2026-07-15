import type { App, MarkdownPostProcessorContext, Plugin, TFile } from 'obsidian';
import { useCallback, useEffect, useRef } from 'react';
import { createRoot, Root } from 'react-dom/client';
import { DocxEditor } from '@npde/docx-editor-react';
import type { RenderedDomContext } from '@npde/docx-editor-core/plugin-api';
import type { Translations } from '@npde/docx-editor-i18n';
import { attachDocxImeTransformNeutralizer } from './docxImeTransformNeutralizer';
import { DOCX_RENDERED_PAGE_SELECTOR } from './docxEditorChromeMarkers';
import { ensureDocxDefaultStyles } from './docxStyleDefaults';
import { ensureEditorStyles } from './DocxReactView';
import { isHTMLElement } from './domGuards';
import { Component, MarkdownRenderChild } from './obsidianRuntime';
import { CoalescedTimeout } from './coalescedTimeout';

const DOCX_EMBED_SELECTOR = '.internal-embed[src], .internal-embed[data-src]';

function getEmbedLinkPath(embedEl: Element) {
	return embedEl.getAttribute('src') ?? embedEl.getAttribute('data-src') ?? '';
}

function stripSubpath(linkPath: string) {
	return linkPath.split('#')[0] ?? '';
}

function isDocxLink(linkPath: string) {
	return stripSubpath(linkPath).toLowerCase().endsWith('.docx');
}

function resolveDocxEmbed(app: App, linkPath: string, sourcePath: string) {
	const file = app.metadataCache.getFirstLinkpathDest(stripSubpath(linkPath), sourcePath);
	return file && file.extension.toLowerCase() === 'docx' ? file : null;
}

function DocxEmbedPreview({
	file,
	buffer,
	hostEl,
	i18n,
}: {
	file: TFile;
	buffer: ArrayBuffer;
	hostEl: HTMLElement;
	i18n: Translations | undefined;
}) {
	const sourceRef = useRef<HTMLDivElement>(null);
	const pagesRef = useRef<HTMLDivElement>(null);
	const renderedDomContextRef = useRef<RenderedDomContext | null>(null);
	const syncFrameRef = useRef<number | null>(null);

	const syncPages = useCallback(() => {
		const renderedDomContext = renderedDomContextRef.current;
		const pagesEl = pagesRef.current;
		if (!renderedDomContext?.pagesContainer?.isConnected || !pagesEl) {
			return;
		}

		const pages = Array.from(renderedDomContext.pagesContainer.querySelectorAll<HTMLElement>(DOCX_RENDERED_PAGE_SELECTOR));
		if (pages.length === 0) {
			return;
		}

		pagesEl.empty();
		for (const page of pages) {
			pagesEl.appendChild(page.cloneNode(true));
		}

		const firstPage = pages[0];
		if (!firstPage) {
			return;
		}

		const pageRect = firstPage.getBoundingClientRect();
		hostEl.setCssProps({
			'--native-powerpoint-doc-editor-embed-page-height': `${Math.ceil(pageRect.height)}px`,
			'--native-powerpoint-doc-editor-embed-page-width': `${Math.ceil(pageRect.width)}px`,
		});
	}, [hostEl]);

	const queueSyncPages = useCallback(() => {
		if (syncFrameRef.current !== null) {
			return;
		}

		syncFrameRef.current = window.requestAnimationFrame(() => {
			syncFrameRef.current = null;
			syncPages();
		});
	}, [syncPages]);

	const handleRenderedDomContextReady = useCallback((context: RenderedDomContext) => {
		renderedDomContextRef.current = context;
		queueSyncPages();
	}, [queueSyncPages]);

	useEffect(() => {
		const sourceEl = sourceRef.current;
		if (!sourceEl) {
			return;
		}

		let detachNeutralizer: (() => void) | undefined;
		let neutralizerRetryTimeout: number | undefined;

		const attachNeutralizer = (): boolean => {
			const editorRoot = sourceEl.querySelector<HTMLElement>('.native-powerpoint-doc-editor-embed-editor');
			if (!editorRoot) {
				return false;
			}

			detachNeutralizer?.();
			detachNeutralizer = attachDocxImeTransformNeutralizer(editorRoot);
			return true;
		};

		const observer = new MutationObserver(() => {
			queueSyncPages();
			if (!detachNeutralizer) {
				attachNeutralizer();
			}
		});
		observer.observe(sourceEl, {
			attributes: true,
			childList: true,
			subtree: true,
		});
		window.setTimeout(queueSyncPages, 0);
		window.setTimeout(queueSyncPages, 100);
		window.setTimeout(queueSyncPages, 500);

		if (!attachNeutralizer()) {
			neutralizerRetryTimeout = window.setTimeout(attachNeutralizer, 100);
		}

		return () => {
			if (syncFrameRef.current !== null) {
				window.cancelAnimationFrame(syncFrameRef.current);
				syncFrameRef.current = null;
			}
			if (neutralizerRetryTimeout !== undefined) {
				window.clearTimeout(neutralizerRetryTimeout);
			}
			detachNeutralizer?.();
			observer.disconnect();
		};
	}, [queueSyncPages]);

	return (
		<>
			<div className="native-powerpoint-doc-editor-embed-viewport">
				<div ref={pagesRef} className="native-powerpoint-doc-editor-embed-pages" />
			</div>
			<div ref={sourceRef} className="native-powerpoint-doc-editor-embed-source" aria-hidden="true">
				<DocxEditor
					key={`${file.path}-${file.stat.mtime}`}
					className="native-powerpoint-doc-editor-embed-editor"
					documentBuffer={buffer}
					disableFindReplaceShortcuts
					i18n={i18n}
					initialZoom={1}
					readOnly
					showOutlineButton={false}
					showRuler={false}
					showToolbar={false}
					showZoomControl={false}
					onFontsLoaded={queueSyncPages}
					onRenderedDomContextReady={handleRenderedDomContextReady}
				/>
			</div>
		</>
	);
}

interface EmbedInfo {
	containerEl: HTMLElement;
}

type DocxFileEmbedCreator = (info: EmbedInfo, file: TFile, subpath: string) => Component;

interface EmbedRegistry {
	registerExtension?: (extension: string, creator: DocxFileEmbedCreator) => void;
	registerExtensions?: (extensions: string[], creator: DocxFileEmbedCreator) => void;
	unregisterExtension?: (extension: string) => void;
	unregisterExtensions?: (extensions: string[]) => void;
}

function getEmbedRegistry(app: App) {
	return (app as App & { embedRegistry?: EmbedRegistry }).embedRegistry;
}

export class DocxFileEmbed extends Component {
	private root: Root | null = null;
	private unloaded = false;

	constructor(
		private info: EmbedInfo,
		private app: App,
		private file: TFile,
		private getEditorLocale: () => Translations | undefined,
		private subpath = '',
	) {
		super();
		this.info.containerEl.addClasses(['native-powerpoint-doc-editor-embed', 'native-powerpoint-doc-editor-native-embed']);
		this.registerDomEvent(this.info.containerEl, 'click', (evt) => {
			evt.stopImmediatePropagation();
		});
	}

	onload() {
		super.onload();
		void this.loadFile();
	}

	async loadFile() {
		await this.loadDocument();
	}

	private async loadDocument() {
		const { containerEl } = this.info;
		this.root?.unmount();
		this.root = null;
		containerEl.empty();
		containerEl.addClass('native-powerpoint-doc-editor-embed');
		containerEl.createDiv({ cls: 'native-powerpoint-doc-editor-embed-loading', text: `Loading ${this.file.name}...` });

		try {
			ensureEditorStyles();
			const sourceBuffer = await this.app.vault.readBinary(this.file);
			const { buffer } = await ensureDocxDefaultStyles(sourceBuffer);
			if (this.unloaded) {
				return;
			}

			containerEl.empty();
			const hostEl = containerEl.createDiv({ cls: 'native-powerpoint-doc-editor-embed-host' });
			this.root = createRoot(hostEl);
			this.root.render(<DocxEmbedPreview file={this.file} buffer={buffer} hostEl={hostEl} i18n={this.getEditorLocale()} />);
		} catch (error) {
			if (this.unloaded) {
				return;
			}

			const message = error instanceof Error ? error.message : 'Unknown error';
			containerEl.empty();
			containerEl.createDiv({
				cls: 'native-powerpoint-doc-editor-embed-error',
				text: `Could not render ${this.file.name}: ${message}`,
			});
		}
	}

	onunload() {
		this.unloaded = true;
		this.root?.unmount();
		this.root = null;
		super.onunload();
	}
}

class DocxEmbedRenderChild extends MarkdownRenderChild {
	private embed: DocxFileEmbed;

	constructor(
		containerEl: HTMLElement,
		app: App,
		file: TFile,
		subpath: string,
		getEditorLocale: () => Translations | undefined,
	) {
		super(containerEl);
		this.embed = new DocxFileEmbed({ containerEl }, app, file, getEditorLocale, subpath);
		this.addChild(this.embed);
	}
}

class DocxEmbedScanChild extends MarkdownRenderChild {
	private observer: MutationObserver | null = null;
	private readonly scanTimer: CoalescedTimeout;

	constructor(
		containerEl: HTMLElement,
		private app: App,
		private ctx: MarkdownPostProcessorContext,
		private getEditorLocale: () => Translations | undefined,
	) {
		super(containerEl);
		this.scanTimer = new CoalescedTimeout(
			containerEl.ownerDocument.defaultView ?? window,
			() => this.scan(),
		);
	}

	onload() {
		this.scan();
		this.scanTimer.schedule(0);
		this.scanTimer.schedule(100);
		this.observer = new MutationObserver(() => this.scanTimer.schedule(25));
		this.observer.observe(this.containerEl, {
			attributes: true,
			attributeFilter: ['data-src', 'src'],
			childList: true,
			subtree: true,
		});
	}

	onunload() {
		this.scanTimer.cancel();
		this.observer?.disconnect();
		this.observer = null;
		super.onunload();
	}

	private scan() {
		renderDocxEmbeds(this.app, this.containerEl, this.ctx, this.getEditorLocale);
	}
}

function collectEmbedElements(el: HTMLElement) {
	const embeds = Array.from(el.querySelectorAll(DOCX_EMBED_SELECTOR));
	if (el.matches(DOCX_EMBED_SELECTOR)) {
		embeds.unshift(el);
	}

	return embeds;
}

export function renderDocxEmbeds(
	app: App,
	el: HTMLElement,
	ctx: MarkdownPostProcessorContext,
	getEditorLocale: () => Translations | undefined,
) {
	const embeds = collectEmbedElements(el);

	for (const embedEl of embeds) {
		if (!isHTMLElement(embedEl) || embedEl.dataset.nativePowerPointDocEditorEmbed === 'true') {
			continue;
		}

		const linkPath = getEmbedLinkPath(embedEl);
		if (!isDocxLink(linkPath)) {
			continue;
		}

		const file = resolveDocxEmbed(app, linkPath, ctx.sourcePath);
		if (!file) {
			continue;
		}

		embedEl.dataset.nativePowerPointDocEditorEmbed = 'true';
		ctx.addChild(new DocxEmbedRenderChild(embedEl, app, file, linkPath.split('#')[1] ?? '', getEditorLocale));
	}
}

export function processDocxEmbeds(
	app: App,
	el: HTMLElement,
	ctx: MarkdownPostProcessorContext,
	getEditorLocale: () => Translations | undefined,
) {
	ctx.addChild(new DocxEmbedScanChild(el, app, ctx, getEditorLocale));
}

export function registerDocxFileEmbed(plugin: Plugin, getEditorLocale: () => Translations | undefined) {
	const registry = getEmbedRegistry(plugin.app);
	if (!registry) {
		return false;
	}

	const createEmbed: DocxFileEmbedCreator = (info, file, subpath) => new DocxFileEmbed(info, plugin.app, file, getEditorLocale, subpath);

	try {
		if (typeof registry.registerExtension === 'function') {
			registry.registerExtension('docx', createEmbed);
			plugin.register(() => registry.unregisterExtension?.('docx'));
			return true;
		}

		if (typeof registry.registerExtensions === 'function') {
			registry.registerExtensions(['docx'], createEmbed);
			plugin.register(() => registry.unregisterExtensions?.(['docx']));
			return true;
		}
	} catch {
		return false;
	}

	return false;
}
