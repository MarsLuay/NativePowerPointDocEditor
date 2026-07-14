import { setIcon } from 'obsidian';

type ClassName = string | readonly string[] | undefined;

const ACTION_ROW_BUTTON_STYLE: Readonly<Record<string, string>> = {
	'-webkit-appearance': 'none',
	'align-items': 'center',
	appearance: 'none',
	background: 'var(--npde-dialog-button-bg)',
	border: '1px solid var(--npde-settings-action-border, var(--npde-dialog-border))',
	'border-radius': '5px',
	'box-shadow': 'none',
	color: 'var(--npde-dialog-text)',
	cursor: 'pointer',
	display: 'inline-flex',
	flex: '0 0 auto',
	'font-size': '12px',
	'font-weight': '500',
	height: '30px',
	'justify-content': 'center',
	padding: '4px 9px',
	'white-space': 'nowrap',
};

function classNameToString(className: ClassName): string | undefined {
	if (typeof className === 'string') {
		return className || undefined;
	}
	return className ? Array.from(className).filter(Boolean).join(' ') : undefined;
}

function classNameToArray(className: ClassName): string[] {
	if (typeof className === 'string') {
		return className ? [className] : [];
	}
	return className ? Array.from(className).filter(Boolean) : [];
}

function mergeClassNames(...classNames: ClassName[]): string[] {
	return classNames.flatMap(classNameToArray);
}

function addClassName(element: HTMLElement, className: ClassName): void {
	if (typeof className === 'string') {
		element.addClass(className);
	} else if (className) {
		element.addClasses(Array.from(className).filter(Boolean));
	}
}

function setAttributes(element: HTMLElement, attributes: Record<string, string | undefined> | undefined): void {
	if (!attributes) return;
	for (const [key, value] of Object.entries(attributes)) {
		if (value !== undefined) {
			element.setAttribute(key, value);
		}
	}
}

function setPriorityCssProps(element: HTMLElement, styles: Readonly<Record<string, string>>): void {
	for (const [property, value] of Object.entries(styles)) {
		element.style.setProperty(property, value, 'important');
	}
}

function targetNode(target: EventTarget | null): Node | null {
	return target instanceof Node ? target : null;
}

export interface PopoverShellOptions {
	className: ClassName;
	role?: string;
	attr?: Record<string, string | undefined>;
	stopPointerDown?: boolean;
}

export function createPopoverShell(parent: HTMLElement, options: PopoverShellOptions): HTMLElement {
	const className = classNameToString(options.className);
	const popover = className ? parent.createDiv({ cls: className }) : parent.createDiv();
	if (options.role) {
		popover.setAttribute('role', options.role);
	}
	setAttributes(popover, options.attr);
	if (options.stopPointerDown) {
		popover.addEventListener('pointerdown', (event) => event.stopPropagation());
	}
	return popover;
}

export function positionPopoverBelow(popover: HTMLElement, anchor: HTMLElement, offset = 4): void {
	const rect = anchor.getBoundingClientRect();
	popover.setCssProps({ left: `${rect.left}px`, top: `${rect.bottom + offset}px` });
}

export interface PopoverDismissOptions {
	popover: HTMLElement;
	anchor?: HTMLElement;
	onDismiss: () => void;
	ownerDocument?: Document;
	pointerEvent?: 'mousedown' | 'pointerdown';
	keydownTarget?: HTMLElement;
	closeOnEscape?: boolean;
	capture?: boolean;
}

export function bindPopoverDismissHandlers(options: PopoverDismissOptions): () => void {
	const ownerDocument = options.ownerDocument ?? activeDocument;
	const pointerEvent = options.pointerEvent ?? 'pointerdown';
	const keydownTarget = options.keydownTarget ?? options.popover;
	const capture = options.capture ?? true;
	const closeOnEscape = options.closeOnEscape ?? true;

	const onOutsidePointer = (event: Event): void => {
		const target = targetNode(event.target);
		if (!target) return;
		if (options.popover.contains(target)) return;
		if (options.anchor?.contains(target)) return;
		options.onDismiss();
	};
	const onKeyDown = (event: KeyboardEvent): void => {
		if (event.key !== 'Escape') return;
		event.preventDefault();
		options.onDismiss();
	};

	ownerDocument.addEventListener(pointerEvent, onOutsidePointer, capture);
	if (closeOnEscape) {
		keydownTarget.addEventListener('keydown', onKeyDown);
	}

	return () => {
		ownerDocument.removeEventListener(pointerEvent, onOutsidePointer, capture);
		if (closeOnEscape) {
			keydownTarget.removeEventListener('keydown', onKeyDown);
		}
	};
}

export interface MenuIconOptions {
	name: string;
	className?: ClassName;
}

export interface ToolbarIconButtonOptions {
	className?: ClassName;
	icon: string;
	label: string;
	tooltip?: string | false;
	iconClassName?: ClassName;
	attr?: Record<string, string | undefined>;
	onClick?: (event: MouseEvent) => void;
}

export function configureToolbarIconButton(
	button: HTMLButtonElement,
	options: ToolbarIconButtonOptions,
): HTMLButtonElement {
	button.type = 'button';
	addClassName(button, mergeClassNames('clickable-icon', options.className));
	button.setAttribute('aria-label', options.label);
	const tooltip = options.tooltip ?? options.label;
	if (tooltip === false) {
		button.removeAttribute('data-tooltip');
	} else {
		button.setAttribute('data-tooltip', tooltip);
	}
	setAttributes(button, options.attr);
	button.replaceChildren();
	if (options.iconClassName) {
		const iconEl = button.createSpan({ cls: classNameToString(options.iconClassName) });
		setIcon(iconEl, options.icon);
	} else {
		setIcon(button, options.icon);
	}
	if (options.onClick) {
		button.addEventListener('click', options.onClick);
	}
	return button;
}

export function createToolbarIconButton(
	parent: HTMLElement,
	options: ToolbarIconButtonOptions,
): HTMLButtonElement {
	const button = parent.createEl('button');
	return configureToolbarIconButton(button, options);
}

export interface MenuItemOptions {
	className: ClassName;
	text?: string;
	role?: string;
	ariaLabel?: string;
	icon?: MenuIconOptions;
	labelClassName?: ClassName;
	disabled?: boolean;
	active?: boolean;
	selected?: boolean;
	attr?: Record<string, string | undefined>;
	preventMouseDown?: boolean;
	preventDefaultOnClick?: boolean;
	stopClickPropagation?: boolean;
	onMouseEnter?: (event: MouseEvent) => void;
	onMouseLeave?: (event: MouseEvent) => void;
	onClick?: (event: MouseEvent) => void;
}

export interface ExistingMenuItemOptions extends Omit<MenuItemOptions, 'className'> {
	className?: ClassName;
}

export function configureMenuItemButton(button: HTMLButtonElement, options: ExistingMenuItemOptions): HTMLButtonElement {
	button.type = 'button';
	addClassName(button, options.className);
	setAttributes(button, options.attr);
	if (options.role) {
		button.setAttribute('role', options.role);
	}
	if (options.ariaLabel) {
		button.setAttribute('aria-label', options.ariaLabel);
	}
	if (options.icon) {
		const iconEl = button.createSpan({ cls: classNameToString(options.icon.className) });
		setIcon(iconEl, options.icon.name);
	}
	if (options.text && options.labelClassName) {
		button.createSpan({ cls: classNameToString(options.labelClassName), text: options.text });
	} else if (options.text) {
		button.setText(options.text);
	}
	if (options.disabled) {
		button.disabled = true;
	}
	if (options.active) {
		button.addClass('is-active');
	}
	if (options.selected !== undefined) {
		button.setAttribute('aria-selected', options.selected ? 'true' : 'false');
	}
	if (options.preventMouseDown) {
		button.addEventListener('mousedown', (event) => event.preventDefault());
	}
	if (options.onMouseEnter) {
		button.addEventListener('mouseenter', options.onMouseEnter);
	}
	if (options.onMouseLeave) {
		button.addEventListener('mouseleave', options.onMouseLeave);
	}
	if (options.onClick) {
		button.addEventListener('click', (event) => {
			if (options.preventDefaultOnClick) {
				event.preventDefault();
			}
			if (options.stopClickPropagation) {
				event.stopPropagation();
			}
			options.onClick?.(event);
		});
	}
	return button;
}

export function createMenuItem(parent: HTMLElement, options: MenuItemOptions): HTMLButtonElement {
	const button = parent.createEl('button', { cls: classNameToString(options.className) });
	return configureMenuItemButton(button, { ...options, className: undefined });
}

export interface InjectedMenuOptionOptions {
	onSelect: () => void;
	dismissBeforeSelect?: boolean;
	ownerDocument?: Document;
}

/**
 * Hardens a button that is injected into a third-party (Eigenpal/Radix) menu so
 * the host menu cannot swallow or reinterpret its activation. Blocks the host's
 * pointer/mouse handling, activates on click and Enter/Space, and (by default)
 * dismisses the host menu with Escape before invoking {@link onSelect}.
 */
export function hardenInjectedMenuOption(
	button: HTMLButtonElement,
	options: InjectedMenuOptionOptions,
): HTMLButtonElement {
	const ownerDocument = options.ownerDocument ?? button.ownerDocument;
	const swallow = (event: Event): void => {
		event.preventDefault();
		event.stopImmediatePropagation();
		event.stopPropagation();
	};
	const activate = (event: Event): void => {
		swallow(event);
		if (options.dismissBeforeSelect !== false) {
			ownerDocument.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		}
		options.onSelect();
	};
	button.addEventListener('pointerdown', swallow);
	button.addEventListener('mousedown', swallow);
	button.addEventListener('click', activate);
	button.addEventListener('keydown', (event) => {
		if (event.key === 'Enter' || event.key === ' ') {
			activate(event);
		}
	});
	return button;
}

export interface MenuSectionOptions {
	className: ClassName;
	text?: string;
	role?: string;
}

export function createMenuSection(parent: HTMLElement, options: MenuSectionOptions): HTMLElement {
	const section = parent.createDiv({ cls: classNameToString(options.className), text: options.text });
	if (options.role) {
		section.setAttribute('role', options.role);
	}
	return section;
}

export interface MenuRowBaseOptions {
	rowClassName: ClassName;
	rowExtraClassName?: ClassName;
	copyClassName: ClassName;
	labelClassName: ClassName;
	descriptionClassName: ClassName;
	label: string;
	description?: string;
	rowTag?: 'div' | 'label';
}

export interface MenuRowParts {
	row: HTMLElement;
	copy: HTMLElement;
}

export function createMenuRow(parent: HTMLElement, options: MenuRowBaseOptions): MenuRowParts {
	const rowClassName = classNameToString(options.rowClassName);
	const row = options.rowTag === 'label'
		? parent.createEl('label', { cls: rowClassName })
		: parent.createDiv({ cls: rowClassName });
	addClassName(row, options.rowExtraClassName);

	const copy = row.createDiv({ cls: classNameToString(options.copyClassName) });
	copy.createSpan({ cls: classNameToString(options.labelClassName), text: options.label });
	if (options.description) {
		copy.createDiv({ cls: classNameToString(options.descriptionClassName), text: options.description });
	}

	return { row, copy };
}

export interface CheckboxRowOptions extends MenuRowBaseOptions {
	controlClassName: ClassName;
	inputClassName: ClassName;
	markClassName?: ClassName;
	checked: boolean;
	onChange: (checked: boolean) => void | Promise<void>;
}

export function createCheckboxRow(parent: HTMLElement, options: CheckboxRowOptions): HTMLInputElement {
	const { row } = createMenuRow(parent, { ...options, rowTag: 'label' });
	const control = row.createSpan({ cls: classNameToString(options.controlClassName) });
	const input = control.createEl('input', {
		cls: classNameToString(options.inputClassName),
		type: 'checkbox',
	});
	input.checked = options.checked;
	if (options.markClassName) {
		control.createSpan({ cls: classNameToString(options.markClassName) });
	}
	input.addEventListener('change', () => {
		void options.onChange(input.checked);
	});
	return input;
}

export interface SelectRowOption {
	value: string;
	label: string;
}

export interface SelectRowOptions extends MenuRowBaseOptions {
	selectClassName: ClassName;
	options: readonly SelectRowOption[];
	selectedValue: string;
	onChange: (value: string) => void | Promise<void>;
}

export function createSelectRow(parent: HTMLElement, options: SelectRowOptions): HTMLSelectElement {
	const { row } = createMenuRow(parent, { ...options, rowExtraClassName: ['mod-input', ...(Array.isArray(options.rowExtraClassName) ? options.rowExtraClassName : options.rowExtraClassName ? [options.rowExtraClassName] : [])] });
	const select = row.createEl('select', { cls: classNameToString(options.selectClassName) });
	for (const item of options.options) {
		const option = select.createEl('option', { text: item.label });
		option.value = item.value;
		option.selected = item.value === options.selectedValue;
	}
	select.addEventListener('change', () => {
		void options.onChange(select.value);
	});
	return select;
}

export interface ActionRowOptions extends MenuRowBaseOptions {
	actionClassName: ClassName;
	actionLabel: string;
	onClick: () => void | Promise<void>;
}

export function createActionRow(parent: HTMLElement, options: ActionRowOptions): HTMLButtonElement {
	const { row } = createMenuRow(parent, {
		...options,
		rowExtraClassName: mergeClassNames(options.rowExtraClassName, 'mod-action'),
	});
	const button = row.createEl('button', {
		cls: classNameToString(options.actionClassName),
		text: options.actionLabel,
		type: 'button',
	});
	setPriorityCssProps(button, ACTION_ROW_BUTTON_STYLE);
	button.addEventListener('click', (event) => {
		event.preventDefault();
		event.stopPropagation();
		void options.onClick();
	});
	return button;
}
