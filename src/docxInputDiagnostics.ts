export type DocxInputDiagnosticEventType = 'keydown' | 'keyup' | 'beforeinput';

export interface DocxInputDiagnosticIds {
	viewId: string;
	pluginId: string;
	handlerIds: {
		keydown: string;
		keyup: string;
		beforeinput: string;
		transaction: string;
	};
}

export interface DocxInputDiagnosticEventLike {
	key?: unknown;
	code?: unknown;
	repeat?: unknown;
	timeStamp?: unknown;
	isTrusted?: unknown;
	location?: unknown;
	eventPhase?: unknown;
	defaultPrevented?: unknown;
	isComposing?: unknown;
	target?: unknown;
	currentTarget?: unknown;
	inputType?: unknown;
}

export interface DocxInputTransactionDiagnosticDetails {
	paragraphsBefore: number;
	paragraphsAfter: number;
	selectionBefore: { from: number; to: number; empty: boolean };
	selectionAfter: { from: number; to: number; empty: boolean };
	transactionCount: number;
	docChangedCount: number;
	steps: string[];
	meta: string[];
}

export interface DocxInputDiagnosticTracker {
	readonly ids: DocxInputDiagnosticIds;
	observeKeyDown(event: unknown): void;
	observeKeyUp(event: unknown): void;
	observeBeforeInput(event: unknown): void;
	beginHandler(event: unknown, source: string): DocxInputDiagnosticHandler | null;
	finishHandler(handler: DocxInputDiagnosticHandler | null, handled: boolean, defaultPreventedAfter?: boolean): void;
	getEventContext(event: unknown): Record<string, unknown>;
	recordTransaction(details: DocxInputTransactionDiagnosticDetails): void;
	mount(): void;
	unmount(): void;
}

export interface DocxInputDiagnosticHandler {
	eventId: string;
	correlationId: string;
	keySequenceId: string | null;
	handlerId: string;
	source: string;
	defaultPreventedBefore: boolean;
}

interface TrackedEvent {
	event: object | null;
	eventId: string;
	correlationId: string;
	keySequenceId: string | null;
	eventType: DocxInputDiagnosticEventType;
	key: string | null;
	code: string | null;
	inputType: string | null;
	repeat: boolean;
	repeatIndex: number;
	timeStamp: number | null;
	isTrusted: boolean | null;
	location: number | null;
	eventPhase: number | null;
	defaultPrevented: boolean;
	isComposing: boolean;
	target: string | null;
	currentTarget: string | null;
}

interface ActiveKey {
	correlationId: string;
	keySequenceId: string;
	repeatIndex: number;
}

interface DocxInputDiagnosticOptions {
	viewId: string;
	emit: (message: string, data: Record<string, unknown>) => void;
}

const MAX_ID_CHARACTERS = 96;
const MAX_SOURCE_CHARACTERS = 160;
const MAX_TARGET_CHARACTERS = 128;
const MAX_TRANSACTION_ITEMS = 32;
const MAX_TRANSACTION_ITEM_CHARACTERS = 160;
const MAX_ACTIVE_KEYS = 16;

function firstIteratorValue<T>(iterator: Iterator<T>): T | undefined {
	const result = iterator.next();
	return result.done ? undefined : result.value;
}

const RELEVANT_INPUT_TYPES = new Set([
	'insertParagraph',
	'insertLineBreak',
	'deleteContentBackward',
	'deleteContentForward',
	'deleteWordBackward',
	'deleteWordForward',
	'deleteSoftLineBackward',
	'deleteSoftLineForward',
	'deleteByCut',
	'deleteByDrag',
]);

let trackerCounter = 0;

function boundedString(value: unknown, maxCharacters: number): string | null {
	if (typeof value !== 'string' || value.length === 0) {
		return null;
	}
	return value.length <= maxCharacters ? value : `${value.slice(0, maxCharacters - 1)}…`;
}

function boundedNumber(value: unknown): number | null {
	return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function boundedBoolean(value: unknown): boolean | null {
	return typeof value === 'boolean' ? value : null;
}

function getEventRecord(event: unknown): Record<string, unknown> | null {
	return event !== null && typeof event === 'object' ? event as Record<string, unknown> : null;
}

function eventObject(event: unknown): object | null {
	return event !== null && typeof event === 'object' ? event : null;
}

function getString(event: unknown, property: keyof DocxInputDiagnosticEventLike): string | null {
	return boundedString(getEventRecord(event)?.[property], MAX_SOURCE_CHARACTERS);
}

function getTargetElement(target: unknown): Record<string, unknown> | null {
	const record = getEventRecord(target);
	if (!record) {
		return null;
	}

	if (record.nodeType === 3 && record.parentElement && typeof record.parentElement === 'object') {
		return record.parentElement as Record<string, unknown>;
	}
	return record;
}

function describeTarget(target: unknown): string | null {
	const element = getTargetElement(target);
	if (!element) {
		return null;
	}

	const tagName = boundedString(element.tagName, 32)?.toLowerCase() ?? 'target';
	const id = boundedString(element.id, 48);
	const className = boundedString(element.className, 64);
	const getAttribute = typeof element.getAttribute === 'function'
		? element.getAttribute as (name: string) => unknown
		: null;
	const role = getAttribute ? boundedString(getAttribute.call(element, 'role'), 32) : null;
	const parts = [tagName];
	if (id) {
		parts.push(`#${id}`);
	}
	if (className) {
		const classes = className.split(/\s+/).filter(Boolean).slice(0, 4).join('.');
		if (classes) {
			parts.push(`.${classes}`);
		}
	}
	if (role) {
		parts.push(`[role=${role}]`);
	}
	return boundedString(parts.join(''), MAX_TARGET_CHARACTERS);
}

function inputTypeIsRelevant(event: unknown): boolean {
	const inputType = getString(event, 'inputType');
	return inputType !== null && RELEVANT_INPUT_TYPES.has(inputType);
}

function keyIsRelevant(event: unknown): boolean {
	const key = getString(event, 'key');
	const code = getString(event, 'code');
	return key === 'Enter'
		|| key === 'Backspace'
		|| key === 'Delete'
		|| code === 'Enter'
		|| code === 'NumpadEnter'
		|| code === 'Backspace'
		|| code === 'Delete';
}

function keyToken(event: unknown): string {
	const key = getString(event, 'key') ?? '';
	const code = getString(event, 'code') ?? '';
	const location = boundedNumber(getEventRecord(event)?.location) ?? 0;
	return boundedString(`${code}|${key}|${location}`, MAX_ID_CHARACTERS) ?? 'unknown';
}

function eventData(tracked: TrackedEvent): Record<string, unknown> {
	return {
		eventId: tracked.eventId,
		correlationId: tracked.correlationId,
		keySequenceId: tracked.keySequenceId,
		eventType: tracked.eventType,
		key: tracked.key,
		code: tracked.code,
		inputType: tracked.inputType,
		repeat: tracked.repeat,
		repeatIndex: tracked.repeatIndex,
		timeStamp: tracked.timeStamp,
		isTrusted: tracked.isTrusted,
		location: tracked.location,
		eventPhase: tracked.eventPhase,
		defaultPrevented: tracked.defaultPrevented,
		isComposing: tracked.isComposing,
		target: tracked.target,
		currentTarget: tracked.currentTarget,
	};
}

function clampTransactionDetails(details: DocxInputTransactionDiagnosticDetails) {
	const clampItems = (items: string[]) => items
		.slice(0, MAX_TRANSACTION_ITEMS)
		.map((item) => boundedString(item, MAX_TRANSACTION_ITEM_CHARACTERS) ?? '');
	return {
		paragraphsBefore: details.paragraphsBefore,
		paragraphsAfter: details.paragraphsAfter,
		selectionBefore: details.selectionBefore,
		selectionAfter: details.selectionAfter,
		transactionCount: Math.min(MAX_TRANSACTION_ITEMS, Math.max(0, details.transactionCount)),
		docChangedCount: Math.min(MAX_TRANSACTION_ITEMS, Math.max(0, details.docChangedCount)),
		steps: clampItems(details.steps),
		meta: clampItems(details.meta),
	};
}

export function createDocxInputDiagnostics(options: DocxInputDiagnosticOptions): DocxInputDiagnosticTracker {
	const trackerNumber = ++trackerCounter;
	const pluginId = `docx-input-plugin-${trackerNumber}`;
	const viewId = boundedString(options.viewId, MAX_ID_CHARACTERS) ?? `docx-editor-view-${trackerNumber}`;
	const ids: DocxInputDiagnosticIds = {
		viewId,
		pluginId,
		handlerIds: {
			keydown: `${pluginId}:keydown`,
			keyup: `${pluginId}:keyup`,
			beforeinput: `${pluginId}:beforeinput`,
			transaction: `${pluginId}:transaction`,
		},
	};
	const activeKeys = new Map<string, ActiveKey>();
	const eventRecords = new WeakMap<object, TrackedEvent>();
	let eventCounter = 0;
	let correlationCounter = 0;
	let sequenceCounter = 0;
	let currentEvent: TrackedEvent | null = null;
	let mounted = false;

	const nextEventId = () => `${pluginId}:event-${++eventCounter}`;
	const nextCorrelationId = () => `${pluginId}:correlation-${++correlationCounter}`;
	const nextSequenceId = () => `${pluginId}:sequence-${++sequenceCounter}`;
	const emit = (message: string, data: Record<string, unknown>) => options.emit(message, {
		viewId,
		pluginId,
		...data,
	});
	const remember = (tracked: TrackedEvent) => {
		if (tracked.event) {
			eventRecords.set(tracked.event, tracked);
		}
		currentEvent = tracked;
		return tracked;
	};
	const deferClear = (tracked: TrackedEvent) => {
		queueMicrotask(() => {
			if (currentEvent === tracked) {
				currentEvent = null;
			}
		});
	};
	const createTracked = (
		event: unknown,
		eventType: DocxInputDiagnosticEventType,
		correlationId: string,
		keySequenceId: string | null,
		repeatIndex: number,
	): TrackedEvent => {
		const record = getEventRecord(event);
		return {
			event: eventObject(event),
			eventId: nextEventId(),
			correlationId,
			keySequenceId,
			eventType,
			key: getString(event, 'key'),
			code: getString(event, 'code'),
			inputType: getString(event, 'inputType'),
			repeat: record?.repeat === true,
			repeatIndex,
			timeStamp: boundedNumber(record?.timeStamp),
			isTrusted: boundedBoolean(record?.isTrusted),
			location: boundedNumber(record?.location),
			eventPhase: boundedNumber(record?.eventPhase),
			defaultPrevented: record?.defaultPrevented === true,
			isComposing: record?.isComposing === true,
			target: describeTarget(record?.target),
			currentTarget: describeTarget(record?.currentTarget),
		};
	};

	const observe = (tracked: TrackedEvent, origin: string, handlerId: string) => {
		remember(tracked);
		emit('DOCX input event observed', {
			...eventData(tracked),
			origin,
			source: `DocxReactView.inputDiagnosticsPlugin.${tracked.eventType}`,
			handlerId,
		});
	};

	return {
		ids,
		observeKeyDown(event) {
			if (!keyIsRelevant(event)) {
				return;
			}
			const token = keyToken(event);
			const active = activeKeys.get(token);
			const isRepeat = getEventRecord(event)?.repeat === true;
			const correlationId = active && isRepeat ? active.correlationId : nextCorrelationId();
			const keySequenceId = active && isRepeat ? active.keySequenceId : nextSequenceId();
			const repeatIndex = active && isRepeat ? active.repeatIndex + 1 : 0;
			activeKeys.delete(token);
			if (activeKeys.size >= MAX_ACTIVE_KEYS) {
				const oldest = firstIteratorValue(activeKeys.keys());
				if (typeof oldest === 'string') {
					activeKeys.delete(oldest);
				}
			}
			activeKeys.set(token, { correlationId, keySequenceId, repeatIndex });
			const tracked = createTracked(event, 'keydown', correlationId, keySequenceId, repeatIndex);
			observe(tracked, 'dom.keydown', ids.handlerIds.keydown);
		},
		observeKeyUp(event) {
			if (!keyIsRelevant(event)) {
				return;
			}
			const token = keyToken(event);
			const active = activeKeys.get(token);
			const tracked = createTracked(
				event,
				'keyup',
				active?.correlationId ?? nextCorrelationId(),
				active?.keySequenceId ?? null,
				active?.repeatIndex ?? 0,
			);
			activeKeys.delete(token);
			observe(tracked, 'dom.keyup', ids.handlerIds.keyup);
			deferClear(tracked);
		},
		observeBeforeInput(event) {
			if (!inputTypeIsRelevant(event)) {
				return;
			}
			const activeKey = firstIteratorValue(activeKeys.values());
			const active = activeKey ?? (currentEvent?.eventType === 'keydown' || currentEvent?.eventType === 'beforeinput'
				? currentEvent
				: null);
			const tracked = createTracked(
				event,
				'beforeinput',
				active?.correlationId ?? nextCorrelationId(),
				active?.keySequenceId ?? null,
				active?.repeatIndex ?? 0,
			);
			observe(tracked, 'dom.beforeinput', ids.handlerIds.beforeinput);
		},
		beginHandler(event, source) {
			const record = eventObject(event) ? eventRecords.get(eventObject(event)!) : currentEvent;
			if (!record) {
				return null;
			}
			const handlerId = record.eventType === 'keydown'
				? ids.handlerIds.keydown
				: record.eventType === 'beforeinput' ? ids.handlerIds.beforeinput : ids.handlerIds.keyup;
			return {
				eventId: record.eventId,
				correlationId: record.correlationId,
				keySequenceId: record.keySequenceId,
				handlerId,
				source: boundedString(source, MAX_SOURCE_CHARACTERS) ?? 'unknown',
				defaultPreventedBefore: record.defaultPrevented,
			};
		},
		finishHandler(handler, handled, defaultPreventedAfter) {
			if (!handler) {
				return;
			}
			emit('DOCX input handler executed', {
				eventId: handler.eventId,
				correlationId: handler.correlationId,
				keySequenceId: handler.keySequenceId,
				handlerId: handler.handlerId,
				source: handler.source,
				origin: 'editor.handler',
				handled,
				defaultPreventedBefore: handler.defaultPreventedBefore,
				defaultPreventedAfter: defaultPreventedAfter ?? handler.defaultPreventedBefore,
			});
			if (currentEvent?.eventId === handler.eventId) {
				deferClear(currentEvent);
			}
		},
		getEventContext(event) {
			const record = eventObject(event) ? eventRecords.get(eventObject(event)!) : currentEvent;
			return record ? eventData(record) : {};
		},
		recordTransaction(details) {
			const boundedDetails = clampTransactionDetails(details);
			emit('DOCX input transaction applied', {
				correlationId: currentEvent?.correlationId ?? null,
				keySequenceId: currentEvent?.keySequenceId ?? null,
				eventId: currentEvent?.eventId ?? null,
				handlerId: ids.handlerIds.transaction,
				origin: 'editor.transaction',
				source: 'DocxReactView.inputDiagnosticsPlugin.appendTransaction',
				...boundedDetails,
			});
		},
		mount() {
			if (mounted) {
				return;
			}
			mounted = true;
			emit('DOCX input diagnostics mounted', {
				handlerIds: ids.handlerIds,
				origin: 'editor.lifecycle',
				source: 'DocxReactView.inputDiagnosticsPlugin.view',
			});
		},
		unmount() {
			if (!mounted) {
				return;
			}
			mounted = false;
			activeKeys.clear();
			currentEvent = null;
			emit('DOCX input diagnostics unmounted', {
				handlerIds: ids.handlerIds,
				origin: 'editor.lifecycle',
				source: 'DocxReactView.inputDiagnosticsPlugin.view.destroy',
			});
		},
	};
}
