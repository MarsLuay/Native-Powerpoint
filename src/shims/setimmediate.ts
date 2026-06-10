type ImmediateHandle = number;
type ImmediateCallback = (...args: unknown[]) => void;

interface ImmediateTarget {
	setImmediate?: (callback: ImmediateCallback, ...args: unknown[]) => ImmediateHandle;
	clearImmediate?: (handle: ImmediateHandle) => void;
}

const target = window as unknown as ImmediateTarget;

if (typeof target.setImmediate !== 'function') {
	target.setImmediate = (callback: ImmediateCallback, ...args: unknown[]) => {
		return window.setTimeout(() => {
			callback(...args);
		}, 0);
	};
}

if (typeof target.clearImmediate !== 'function') {
	target.clearImmediate = (handle) => window.clearTimeout(handle);
}
