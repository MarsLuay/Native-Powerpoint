type ImmediateHandle = ReturnType<typeof setTimeout>;
type ImmediateCallback = (...args: unknown[]) => void;

interface ImmediateTarget {
	setImmediate?: (callback: ImmediateCallback, ...args: unknown[]) => ImmediateHandle;
	clearImmediate?: (handle: ImmediateHandle) => void;
}

const target = globalThis as unknown as ImmediateTarget;

if (typeof target.setImmediate !== 'function') {
	target.setImmediate = (callback: ImmediateCallback, ...args: unknown[]) => {
		return setTimeout(() => {
			callback(...args);
		}, 0);
	};
}

if (typeof target.clearImmediate !== 'function') {
	target.clearImmediate = (handle) => clearTimeout(handle);
}
