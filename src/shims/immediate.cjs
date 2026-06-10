module.exports = function immediate(callback) {
	window.setTimeout(callback, 0);
};
