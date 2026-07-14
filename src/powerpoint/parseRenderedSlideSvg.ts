function tagName(element: Element): string {
	return (element.tagName || element.nodeName || '').toLowerCase().replace(/^.*:/, '');
}

function walkElements(root: Element, visit: (element: Element) => void): void {
	if (!root || root.nodeType !== 1) return;
	visit(root);
	for (const child of Array.from(root.childNodes || [])) {
		walkElements(child as Element, visit);
	}
}

function matchesSelector(element: Element, selector: string): boolean {
	const match = selector.match(/^([a-z*]+|\*)\[([^\]=]+)(?:="([^"]*)")?\]$/i);
	if (!match) return false;
	const wantedTag = (match[1] ?? '*').toLowerCase();
	const attribute = match[2];
	const value = match[3];
	if (!attribute) return false;
	if (wantedTag !== '*' && tagName(element) !== wantedTag) return false;
	const actual = element.getAttribute?.(attribute);
	if (actual == null) return false;
	return value === undefined || actual === value;
}

function querySelectorAll(root: Element, selector: string): Element[] {
	const results: Element[] = [];
	walkElements(root, (element) => {
		if (matchesSelector(element, selector)) {
			results.push(element);
		}
	});
	return results;
}

function querySelector(root: Element, selector: string): Element | null {
	return querySelectorAll(root, selector)[0] ?? null;
}

function closest(element: Element, selector: string): Element | null {
	let node: Element | null = element;
	while (node && node.nodeType === 1) {
		if (matchesSelector(node, selector)) return node;
		node = node.parentNode as Element | null;
	}
	return null;
}

function augmentDomQueries(root: Element): void {
	const assignQueries = (element: Element): void => {
		const target = element as Element & {
			querySelector?: (selector: string) => Element | null;
			querySelectorAll?: (selector: string) => Element[];
			closest?: (selector: string) => Element | null;
		};
		if (typeof target.querySelector !== 'function') {
			const patch = target as unknown as Record<string, unknown>;
			patch.querySelector = (selector: string) => querySelector(element, selector);
			patch.querySelectorAll = (selector: string) => querySelectorAll(element, selector);
			patch.closest = (selector: string) => closest(element, selector);
		}
	};
	walkElements(root, assignQueries);
	assignQueries(root);
}

export function parseRenderedSlideSvg(svgString: string): SVGSVGElement {
	const doc = new DOMParser().parseFromString(svgString, 'image/svg+xml');
	const root = doc.documentElement;
	const rootTagName = root.tagName?.toLowerCase() ?? root.localName?.toLowerCase() ?? '';
	if (rootTagName !== 'svg' && !(typeof SVGSVGElement !== 'undefined' && root instanceof SVGSVGElement)) {
		throw new Error('Could not parse slide SVG.');
	}
	augmentDomQueries(root);
	return root as unknown as SVGSVGElement;
}
