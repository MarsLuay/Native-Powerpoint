type InstanceOfCapable = {
	instanceOf: (constructor: unknown) => boolean;
};

type DomConstructor<T> = abstract new (...args: never[]) => T;

function hasInstanceOf(value: unknown): value is InstanceOfCapable {
	return (
		(typeof value === 'object' || typeof value === 'function')
		&& value !== null
		&& typeof (value as Partial<InstanceOfCapable>).instanceOf === 'function'
	);
}

function isDomInstance<T>(value: unknown, constructor: DomConstructor<T>): value is T {
	if ((typeof value !== 'object' && typeof value !== 'function') || value === null) {
		return false;
	}

	if (hasInstanceOf(value)) {
		return value.instanceOf(constructor);
	}

	return Function.prototype[Symbol.hasInstance].call(constructor, value);
}

export function isHTMLElement(value: unknown): value is HTMLElement {
	return typeof HTMLElement !== 'undefined' && isDomInstance(value, HTMLElement);
}

export function isElement(value: unknown): value is Element {
	return typeof Element !== 'undefined' && isDomInstance(value, Element);
}

export function isHTMLButtonElement(value: unknown): value is HTMLButtonElement {
	return typeof HTMLButtonElement !== 'undefined' && isDomInstance(value, HTMLButtonElement);
}

export function isClipboardEvent(value: unknown): value is ClipboardEvent {
	return typeof ClipboardEvent !== 'undefined' && isDomInstance(value, ClipboardEvent);
}

export function isInputEvent(value: unknown): value is InputEvent {
	return typeof InputEvent !== 'undefined' && isDomInstance(value, InputEvent);
}

export function isNode(value: unknown): value is Node {
	return typeof Node !== 'undefined' && isDomInstance(value, Node);
}

export function isSVGGElement(value: unknown): value is SVGGElement {
	return typeof SVGGElement !== 'undefined' && isDomInstance(value, SVGGElement);
}

export function isSVGTextElement(value: unknown): value is SVGTextElement {
	return typeof SVGTextElement !== 'undefined' && isDomInstance(value, SVGTextElement);
}

export function isSVGTSpanElement(value: unknown): value is SVGTSpanElement {
	return typeof SVGTSpanElement !== 'undefined' && isDomInstance(value, SVGTSpanElement);
}

export function isText(value: unknown): value is Text {
	return typeof Text !== 'undefined' && isDomInstance(value, Text);
}
