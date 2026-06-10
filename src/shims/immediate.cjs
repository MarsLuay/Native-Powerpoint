module.exports = function immediate(callback) {
	setTimeout(callback, 0);
};
