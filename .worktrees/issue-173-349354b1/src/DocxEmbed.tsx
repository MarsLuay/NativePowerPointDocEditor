import type { App, MarkdownPostProcessorContext, TFile } from 'obsidian';
import { useCallback, useEffect, useRef } from 'react';
import { createRoot, Root } from 'react-dom/client';
import { DocxEditor } from './docx/runtime/bridge.mjs';
import type { RenderedDomContext, Translations } from './docx/runtime/contract';
import { attachDocxImeTransformNeutralizer } from './docxImeTransformNeutralizer';
import { DOCX_RENDERED_PAGE_SELECTOR } from './docxEditorChromeMarkers';
import { ensureDocxDefaultStyles } from './docxStyleDefaults';
import { ensureEditorStyles } from './DocxReactView';
import { isHTMLElement } from './domGuards';
import { Component, MarkdownRenderChild } from './obsidianRuntime';

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

function copyRenderedPages(pagesContainer: HTMLElement, pagesEl: HTMLElement, hostEl: HTMLElement) {
	const pages = Array.from(pagesContainer.querySelectorAll<HTMLElement>(DOCX_RENDERED_PAGE_SELECTOR));
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
}

function useDocxEmbedSync(hostEl: HTMLElement) {
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

		copyRenderedPages(renderedDomContext.pagesContainer, pagesEl, hostEl);
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

	return {
		pagesRef,
		sourceRef,
		queueSyncPages,
		handleRenderedDomContextReady,
	};
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
	const { pagesRef, sourceRef, queueSyncPages, handleRenderedDomContextReady } = useDocxEmbedSync(hostEl);

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
			ensureEditorStyles(containerEl.ownerDocument);
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
