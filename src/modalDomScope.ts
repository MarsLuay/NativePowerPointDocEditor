import { Component } from 'obsidian';

/** Returns a Component scope; register DOM events on it, then call loadModalDomScope. */
export function openModalDomScope(): Component {
	return new Component();
}

export function loadModalDomScope(scope: Component): void {
	scope.load();
}

export function closeModalDomScope(scope: Component | undefined): void {
	scope?.unload();
}
