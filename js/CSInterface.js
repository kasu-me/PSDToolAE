
var CSInterface = function () { };

CSInterface.THEME_COLOR_CHANGED_EVENT = "com.adobe.csxs.events.ThemeColorChanged";

CSInterface.prototype.hostEnvironment = window.__adobe_cep__ ? JSON.parse(window.__adobe_cep__.getHostEnvironment()) : null;

CSInterface.prototype.getHostEnvironment = function () {
	this.hostEnvironment = window.__adobe_cep__ ? JSON.parse(window.__adobe_cep__.getHostEnvironment()) : null;
	return this.hostEnvironment || { appName: "MockApp", appSkinInfo: { panelBackgroundColor: { color: { red: 50, green: 50, blue: 50 } }, baseFontSize: 12 } };
};

CSInterface.prototype.addEventListener = function (type, listener, obj) {
	if (typeof window.__adobe_cep__ !== "undefined") {
		window.__adobe_cep__.addEventListener(type, listener, obj);
	}
};

CSInterface.prototype.evalScript = function (script, callback) {
	if (typeof window.__adobe_cep__ !== "undefined") {
		window.__adobe_cep__.evalScript(script, callback);
	} else {
		console.log("ExtendScript call (mock): " + script);
		// Mock return for JSON.parse
		if (callback) callback('{"status":"error", "message":"Mock environment (no AE connection)"}');
	}
};

CSInterface.prototype.resizeContent = function (width, height) {
	if (typeof window.__adobe_cep__ !== "undefined") {
		window.__adobe_cep__.resizeContent(width, height);
	}
};
