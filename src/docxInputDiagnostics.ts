export type DocxInputDiagnosticEventType = 'keydown' | 'keyup' | 'beforeinput' | 'input';

export interface DocxInputDiagnosticIds {
	viewId: string;
	pluginId: string;
	handlerIds: {
		keydown: string;
		keyup: string;
		beforeinput: string;
		input: string;
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

export interface DocxInputSelectionSnapshot {
	from: number;
	to: number;
	empty: boolean;
}

export interface DocxInputDiagnosticTracker {
	readonly ids: DocxInputDiagnosticIds;
	observeKeyDown(event: unknown, selection?: DocxInputSelectionSnapshot): void;
	observeKeyUp(event: unknown, selection?: DocxInputSelectionSnapshot): void;
	observeBeforeInput(event: unknown, selection?: DocxInputSelectionSnapshot): void;
	observeInput(event: unknown, selection?: DocxInputSelectionSnapshot): void;
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
	selectionBefore: DocxInputSelectionSnapshot | null;
	selectionAfter: DocxInputSelectionSnapshot | null;
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

function insertTextIsSpace(event: unknown): boolean {
	if (getString(event, 'inputType') !== 'insertText') {
		return false;
	}
	const data = getEventRecord(event)?.data;
	return data === ' ' || data === '\u00a0';
}

function inputEventIsRelevant(event: unknown): boolean {
	return inputTypeIsRelevant(event) || insertTextIsSpace(event);
}

function keyIsRelevant(event: unknown): boolean {
	const key = getString(event, 'key');
	const code = getString(event, 'code');
	return key === 'Enter'
		|| key === ' '
		|| key === 'Backspace'
		|| key === 'Delete'
		|| code === 'Enter'
		|| code === 'NumpadEnter'
		|| code === 'Space'
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
		selectionBefore: tracked.selectionBefore,
		selectionAfter: tracked.selectionAfter,
	};
}

function copySelection(selection: DocxInputSelectionSnapshot | undefined): DocxInputSelectionSnapshot | null {
	if (!selection || !Number.isFinite(selection.from) || !Number.isFinite(selection.to)) {
		return null;
	}
	return {
		from: selection.from,
		to: selection.to,
		empty: selection.empty === true,
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

interface RecentRelevantKeydown {
	key: string | null;
	timeStamp: number | null;
	repeat: boolean;
	viewId: string;
	pluginId: string;
	handlerId: string;
	attach(viewId: string, handlerId: string): void;
}

const recentRelevantKeydowns: RecentRelevantKeydown[] = [];
const MAX_RECENT_KEYDOWNS = 32;

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
			input: `${pluginId}:input`,
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
	const sequences = new Map<string, {
		key: string | null;
		keydownCount: number;
		beforeInputCount: number;
		repeatKeydowns: number;
		docChangingTransactionCount: number;
		viewIds: Set<string>;
		handlerIds: Set<string>;
		summarized: boolean;
	}>();

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
		selection?: DocxInputSelectionSnapshot,
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
			selectionBefore: copySelection(selection),
			selectionAfter: null,
		};
	};

	const sequenceFor = (correlationId: string, key: string | null) => {
		let sequence = sequences.get(correlationId);
		if (!sequence) {
			sequence = {
				key,
				keydownCount: 0,
				beforeInputCount: 0,
				repeatKeydowns: 0,
				docChangingTransactionCount: 0,
				viewIds: new Set([viewId]),
				handlerIds: new Set(),
				summarized: false,
			};
			sequences.set(correlationId, sequence);
		}
		if (key && !sequence.key) {
			sequence.key = key;
		}
		return sequence;
	};

	const emitDuplicateSummary = (correlationId: string) => {
		const sequence = sequences.get(correlationId);
		if (!sequence || sequence.summarized) {
			return;
		}
		const duplicateKeydown = sequence.keydownCount >= 2 && sequence.repeatKeydowns === 0;
		const multipleViews = sequence.viewIds.size > 1 && sequence.repeatKeydowns === 0;
		const multipleTransactions = sequence.docChangingTransactionCount >= 2;
		const keydownAndBeforeInput = sequence.keydownCount >= 1
			&& sequence.beforeInputCount >= 1
			&& multipleTransactions;
		if (!duplicateKeydown && !multipleViews && !multipleTransactions && !keydownAndBeforeInput) {
			return;
		}
		const probableBoundary = multipleViews
			? 'multiple-views'
			: duplicateKeydown
				? 'duplicate-keydown'
				: keydownAndBeforeInput
					? 'keydown-and-beforeinput'
					: 'multiple-transactions';
		sequence.summarized = true;
		emit('DOCX duplicate input candidate', {
			correlationId,
			key: sequence.key,
			keydownCount: sequence.keydownCount,
			beforeInputCount: sequence.beforeInputCount,
			handlerCount: sequence.handlerIds.size,
			docChangingTransactionCount: sequence.docChangingTransactionCount,
			viewIds: [...sequence.viewIds],
			handlerIds: [...sequence.handlerIds],
			probableBoundary,
		});
	};

	const noteSequence = (
		correlationId: string,
		key: string | null,
		kind: 'keydown' | 'beforeinput' | 'input',
		repeat: boolean,
		handlerId: string,
	) => {
		const sequence = sequenceFor(correlationId, key);
		sequence.handlerIds.add(handlerId);
		if (kind === 'keydown') {
			sequence.keydownCount += 1;
			if (repeat) {
				sequence.repeatKeydowns += 1;
			}
		} else if (kind === 'beforeinput' || kind === 'input') {
			sequence.beforeInputCount += 1;
		}
		emitDuplicateSummary(correlationId);
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
		observeKeyDown(event, selection) {
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
			const tracked = createTracked(event, 'keydown', correlationId, keySequenceId, repeatIndex, selection);
			const sequence = sequenceFor(correlationId, tracked.key);
			for (const previous of recentRelevantKeydowns) {
				const sameKeyAndTime = previous.key === tracked.key
					&& previous.timeStamp === tracked.timeStamp
					&& !previous.repeat
					&& !tracked.repeat;
				if (!sameKeyAndTime) {
					continue;
				}
				if (previous.viewId !== viewId) {
					sequence.viewIds.add(previous.viewId);
					sequence.handlerIds.add(previous.handlerId);
					previous.attach(viewId, ids.handlerIds.keydown);
				} else if (previous.pluginId === pluginId) {
					sequence.keydownCount += 1;
					sequence.handlerIds.add(previous.handlerId);
				}
			}
			recentRelevantKeydowns.push({
				key: tracked.key,
				timeStamp: tracked.timeStamp,
				repeat: tracked.repeat,
				viewId,
				pluginId,
				handlerId: ids.handlerIds.keydown,
				attach(otherViewId, otherHandlerId) {
					const owner = sequences.get(correlationId);
					if (!owner) {
						return;
					}
					owner.viewIds.add(otherViewId);
					owner.handlerIds.add(otherHandlerId);
					emitDuplicateSummary(correlationId);
				},
			});
			if (recentRelevantKeydowns.length > MAX_RECENT_KEYDOWNS) {
				recentRelevantKeydowns.shift();
			}
			observe(tracked, 'dom.keydown', ids.handlerIds.keydown);
			noteSequence(correlationId, tracked.key, 'keydown', tracked.repeat, ids.handlerIds.keydown);
		},
		observeKeyUp(event, selection) {
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
				selection,
			);
			activeKeys.delete(token);
			observe(tracked, 'dom.keyup', ids.handlerIds.keyup);
			deferClear(tracked);
		},
		observeBeforeInput(event, selection) {
			if (!inputEventIsRelevant(event)) {
				return;
			}
			const activeKey = firstIteratorValue(activeKeys.values());
			const active = activeKey ?? (currentEvent?.eventType === 'keydown' || currentEvent?.eventType === 'beforeinput' || currentEvent?.eventType === 'input'
				? currentEvent
				: null);
			const tracked = createTracked(
				event,
				'beforeinput',
				active?.correlationId ?? nextCorrelationId(),
				active?.keySequenceId ?? null,
				active?.repeatIndex ?? 0,
				selection,
			);
			observe(tracked, 'dom.beforeinput', ids.handlerIds.beforeinput);
			noteSequence(tracked.correlationId, tracked.key, 'beforeinput', false, ids.handlerIds.beforeinput);
		},
		observeInput(event, selection) {
			if (!inputEventIsRelevant(event)) {
				return;
			}
			const activeKey = firstIteratorValue(activeKeys.values());
			const active = activeKey ?? (currentEvent?.eventType === 'keydown' || currentEvent?.eventType === 'beforeinput' || currentEvent?.eventType === 'input'
				? currentEvent
				: null);
			const tracked = createTracked(
				event,
				'input',
				active?.correlationId ?? nextCorrelationId(),
				active?.keySequenceId ?? null,
				active?.repeatIndex ?? 0,
				selection,
			);
			observe(tracked, 'dom.input', ids.handlerIds.input);
			noteSequence(tracked.correlationId, tracked.key, 'input', false, ids.handlerIds.input);
		},
		beginHandler(event, source) {
			const record = eventObject(event) ? eventRecords.get(eventObject(event)!) : currentEvent;
			if (!record) {
				return null;
			}
			const handlerId = record.eventType === 'keydown'
				? ids.handlerIds.keydown
				: record.eventType === 'beforeinput'
					? ids.handlerIds.beforeinput
					: record.eventType === 'input'
						? ids.handlerIds.input
						: ids.handlerIds.keyup;
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
			const correlationId = currentEvent?.correlationId ?? null;
			emit('DOCX input transaction applied', {
				correlationId,
				keySequenceId: currentEvent?.keySequenceId ?? null,
				eventId: currentEvent?.eventId ?? null,
				handlerId: ids.handlerIds.transaction,
				origin: 'editor.transaction',
				source: 'DocxReactView.inputDiagnosticsPlugin.appendTransaction',
				...boundedDetails,
			});
			if (correlationId && boundedDetails.docChangedCount > 0) {
				const sequence = sequenceFor(correlationId, currentEvent?.key ?? null);
				sequence.docChangingTransactionCount += boundedDetails.docChangedCount;
				sequence.handlerIds.add(ids.handlerIds.transaction);
				emitDuplicateSummary(correlationId);
			}
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
