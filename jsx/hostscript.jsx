
// JSON Polyfill (Minimal)
if (typeof JSON !== 'object') {
	JSON = {};
}
(function () {
	'use strict';
	function f(n) { return n < 10 ? '0' + n : n; }
	if (typeof Date.prototype.toJSON !== 'function') {
		Date.prototype.toJSON = function () {
			return isFinite(this.valueOf())
				? this.getUTCFullYear() + '-' +
				f(this.getUTCMonth() + 1) + '-' +
				f(this.getUTCDate()) + 'T' +
				f(this.getUTCHours()) + ':' +
				f(this.getUTCMinutes()) + ':' +
				f(this.getUTCSeconds()) + 'Z'
				: null;
		};
		String.prototype.toJSON = Number.prototype.toJSON = Boolean.prototype.toJSON = function () { return this.valueOf(); };
	}
	var cx = /[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
		escapable = /[\\\"\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,
		meta = { '\b': '\\b', '\t': '\\t', '\n': '\\n', '\f': '\\f', '\r': '\\r', '"': '\\"', '\\': '\\\\' };
	function quote(string) {
		escapable.lastIndex = 0;
		return escapable.test(string) ? '"' + string.replace(escapable, function (a) {
			var c = meta[a];
			return typeof c === 'string' ? c : '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
		}) + '"' : '"' + string + '"';
	}
	function str(key, holder) {
		var i, k, v, length, mind = gap, partial, value = holder[key];
		if (value && typeof value === 'object' && typeof value.toJSON === 'function') { value = value.toJSON(key); }
		if (typeof rep === 'function') { value = rep.call(holder, key, value); }
		switch (typeof value) {
			case 'string': return quote(value);
			case 'number': return isFinite(value) ? String(value) : 'null';
			case 'boolean':
			case 'null': return String(value);
			case 'object':
				if (!value) return 'null';
				gap += indent;
				partial = [];
				if (Object.prototype.toString.apply(value) === '[object Array]') {
					length = value.length;
					for (i = 0; i < length; i += 1) { partial[i] = str(i, value) || 'null'; }
					v = partial.length === 0 ? '[]' : gap ? '[\n' + gap + partial.join(',\n' + gap) + '\n' + mind + ']' : '[' + partial.join(',') + ']';
					gap = mind;
					return v;
				}
				if (rep && typeof rep === 'object') {
					length = rep.length;
					for (i = 0; i < length; i += 1) { if (typeof rep[i] === 'string') { k = rep[i]; v = str(k, value); if (v) { partial.push(quote(k) + (gap ? ': ' : ':') + v); } } }
				} else {
					for (k in value) { if (Object.prototype.hasOwnProperty.call(value, k)) { v = str(k, value); if (v) { partial.push(quote(k) + (gap ? ': ' : ':') + v); } } }
				}
				v = partial.length === 0 ? '{}' : gap ? '{\n' + gap + partial.join(',\n' + gap) + '\n' + mind + '}' : '{' + partial.join(',') + '}';
				gap = mind;
				return v;
		}
	}
	if (typeof JSON.stringify !== 'function') {
		var gap, indent, rep;
		JSON.stringify = function (value, replacer, space) {
			var i; gap = ''; indent = '';
			if (typeof space === 'number') { for (i = 0; i < space; i += 1) { indent += ' '; } }
			else if (typeof space === 'string') { indent = space; }
			rep = replacer;
			if (replacer && typeof replacer !== 'function' && (typeof replacer !== 'object' || typeof replacer.length !== 'number')) { throw new Error('JSON.stringify'); }
			return str('', { '': value });
		};
	}
}());

// Main Functions

function importPSD() {
	try {
		var importOptions = new ImportOptions();
		var f = File.openDialog("PSDファイルを選択してください", "*.psd");
		if (!f) return JSON.stringify({ status: "cancelled" });

		importOptions.file = f;
		importOptions.importAs = ImportAsType.COMP;
		importOptions.sequence = false;

		var importedItem = app.project.importFile(importOptions);

		// インポートされたコンポジションを開く
		if (importedItem instanceof CompItem) {
			importedItem.openInViewer();
			return JSON.stringify({ status: "success", compId: importedItem.id, compName: importedItem.name });
		} else {
			// フォルダとしてインポートされた場合（編集可能なレイヤースタイルなど）
			// コンポジションを探す
			return JSON.stringify({ status: "success", message: "Imported but comp ID logic might need check" });
		}

	} catch (e) {
		return JSON.stringify({ status: "error", message: e.toString() });
	}
}

// ---------------------------------------------------------
// Keyframe Operations
// ---------------------------------------------------------

var g_storedKeyframes = []; // Array of { prop: Property, keyIndex: number, layerChain: Array<Layer> }

function getKeyframesFromActiveComp() {
	try {
		g_storedKeyframes = [];
		var comp = app.project.activeItem;
		if (!comp || !(comp instanceof CompItem)) {
			return "false";
		}

		// 再帰的にキーフレームを収集
		// rootCompTime, layerChain=[]
		collectKeysForMove(comp, comp.time, []);

		if (g_storedKeyframes.length > 0) {
			return "true";
		}
		return "false";
	} catch (e) {
		alert("Error in getKeyframesFromActiveComp: " + e.toString());
		return "false";
	}
}

function collectKeysForMove(comp, time, layerChain) {
	for (var i = 1; i <= comp.numLayers; i++) {
		var layer = comp.layer(i);

		// Scan properties of this layer first using current comp time
		// Pass layerChain from parent context (do not include current layer yet)
		scanPropsForMove(layer, time, layerChain);

		// If it's a pre-comp, recurse
		if (layer.source instanceof CompItem) {
			// Calculate inner time for the pre-comp
			// Simple Formula: (parentTime - startTime) * (100 / stretch)
			// Note: Does not support Time Remap for now.
			var localTime = (time - layer.startTime) * (100 / layer.stretch);

			var newChain = layerChain.concat([]); // shallow copy
			newChain.push(layer);

			collectKeysForMove(layer.source, localTime, newChain);
		}
	}
}

function scanPropsForMove(propGroup, time, layerChain) {
	// Traverse properties
	var numProps = propGroup.numProperties;
	if (!numProps) return;

	for (var i = 1; i <= numProps; i++) {
		var prop = propGroup.property(i);

		if (prop.propertyType === PropertyType.PROPERTY) {
			if (prop.numKeys > 0) {
				// Check if there is a keyframe at 'time'
				var nearestIndex = prop.nearestKeyIndex(time);

				// If no keys, nearestKeyIndex might return 0 or 1 depending on AE version/state, check bounds
				if (nearestIndex > 0 && nearestIndex <= prop.numKeys) {
					var keyTime = prop.keyTime(nearestIndex);

					// Tolerance for floating point comparison (e.g. 1/frameRate is too large, use small epsilon)
					// 0.0001 is usually enough (1/30 frame is ~0.033)
					if (Math.abs(keyTime - time) < 0.0001) {
						// Found a keyframe!
						g_storedKeyframes.push({
							property: prop,
							keyIndex: nearestIndex,
							layerChain: layerChain
						});
					}
				}
			}
		} else if (prop.propertyType === PropertyType.NAMED_GROUP || prop.propertyType === PropertyType.INDEXED_GROUP) {
			scanPropsForMove(prop, time, layerChain);
		}
	}
}

function moveKeyframesToCurrentTime() {
	try {
		if (g_storedKeyframes.length === 0) {
			return JSON.stringify({ status: "error", message: "記録中のキーフレームがありません (Memory Empty)" });
		}

		var rootComp = app.project.activeItem;
		if (!rootComp || !(rootComp instanceof CompItem)) {
			return JSON.stringify({ status: "error", message: "コンポジションがアクティブではありません" });
		}

		var rootTime = rootComp.time;
		var successCount = 0;
		var failCount = 0;
		var details = "";

		app.beginUndoGroup("Move Keyframes via Script");

		for (var i = 0; i < g_storedKeyframes.length; i++) {
			var item = g_storedKeyframes[i];
			var prop = item.property;
			var oldKeyIndex = item.keyIndex;

			// Calculate target time for this specific property
			var targetTime = rootTime;
			try {
				for (var j = 0; j < item.layerChain.length; j++) {
					var layer = item.layerChain[j];
					targetTime = (targetTime - layer.startTime) * (100 / layer.stretch);
				}
			} catch (e) {
				failCount++;
				details += "Layer Calc Error: " + e.message + "; ";
				continue;
			}

			// Perform the move
			var res = moveKeyframe(prop, oldKeyIndex, targetTime);
			if (res.success) {
				successCount++;
			} else {
				failCount++;
				details += "Prop " + prop.name + ": " + res.reason + "; ";
			}
		}

		g_storedKeyframes = []; // Clear after move
		app.endUndoGroup();

		if (successCount > 0) {
			return JSON.stringify({ status: "success", message: "Moved " + successCount + " keys (" + failCount + " failed)" });
		} else {
			return JSON.stringify({ status: "error", message: "移動失敗 (成功0, 失敗" + failCount + "): " + details });
		}

	} catch (e) {
		return JSON.stringify({ status: "error", message: "Fatal Error: " + e.toString() });
	}
}

function moveKeyframe(prop, keyIndex, newTime) {
	try {
		// Validations
		if (!prop.canVaryOverTime) return { success: false, reason: "Not time-variant" };
		if (keyIndex < 1 || keyIndex > prop.numKeys) return { success: false, reason: "Invalid Index " + keyIndex + "/" + prop.numKeys };

		// 1. Get all attributes
		var value;
		try {
			value = prop.keyValue(keyIndex);
		} catch (e) {
			return { success: false, reason: "keyValue failed" };
		}

		var inTempEase = prop.keyInTemporalEase(keyIndex);
		var outTempEase = prop.keyOutTemporalEase(keyIndex);
		var inInterp = prop.keyInInterpolationType(keyIndex);
		var outInterp = prop.keyOutInterpolationType(keyIndex);
		var tempCont = prop.keyTemporalContinuous(keyIndex);
		var tempAuto = prop.keyTemporalAutoBezier(keyIndex);

		var inSpat = null, outSpat = null, spatAuto = false, spatCont = false, roving = false;
		var isSpatial = (prop.propertyValueType === PropertyValueType.TwoD_SPATIAL || prop.propertyValueType === PropertyValueType.ThreeD_SPATIAL);

		if (isSpatial) {
			inSpat = prop.keyInSpatialTangent(keyIndex);
			outSpat = prop.keyOutSpatialTangent(keyIndex);
			spatAuto = prop.keySpatialAutoBezier(keyIndex);
			spatCont = prop.keySpatialContinuous(keyIndex);
			try { roving = prop.keyRoving(keyIndex); } catch (e) { }
		}

		// 2. Remove old key

		// If newTime is exactly same loop, do nothing but report success
		if (Math.abs(prop.keyTime(keyIndex) - newTime) < 0.0001) return { success: true };

		prop.removeKey(keyIndex);

		// 3. Add new key
		var newKeyIndex = prop.addKey(newTime);

		if (value !== null) prop.setValueAtKey(newKeyIndex, value);

		// Restore attributes
		if (inTempEase && outTempEase) prop.setTemporalEaseAtKey(newKeyIndex, inTempEase, outTempEase);
		prop.setInterpolationTypeAtKey(newKeyIndex, inInterp, outInterp);

		if (isSpatial) {
			if (inSpat && outSpat) {
				prop.setSpatialTangentsAtKey(newKeyIndex, inSpat, outSpat);
			}
			prop.setSpatialAutoBezierAtKey(newKeyIndex, spatAuto);
			prop.setSpatialContinuousAtKey(newKeyIndex, spatCont);
			try { prop.setRovingAtKey(newKeyIndex, roving); } catch (e) { }
		}

		prop.setTemporalContinuousAtKey(newKeyIndex, tempCont);

		prop.setTemporalAutoBezierAtKey(newKeyIndex, tempAuto);

		return { success: true };
	} catch (e) {
		return { success: false, reason: e.toString() };
	}
}


function removeKeyframesAtCurrentTime() {
	var undoName = "Remove Keyframes at Current Time";
	var comp = app.project.activeItem;
	if (!comp || !(comp instanceof CompItem)) {
		return JSON.stringify({ status: "error", message: "コンポジションがアクティブではありません" });
	}

	app.beginUndoGroup(undoName);

	try {
		var counter = { count: 0 };
		removeKeysRecursive(comp, comp.time, counter);

		if (counter.count > 0) {
			// Undoグループを閉じるには正常終了後に閉じる必要があるが、
			// エラー時はfinallyなどで閉じられるようにすべきか、
			// あるいはここで閉じてから結果を返す。
			app.endUndoGroup();
			return JSON.stringify({ status: "success", count: counter.count });
		} else {
			// 何も削除しなかった場合でもUndo履歴に残したくない場合はキャンセル扱いにする手もあるが、
			// ユーザーが「削除を実行したが何もなかった」とわかるようにそのまま終了する。
			app.endUndoGroup();
			return JSON.stringify({ status: "error", message: "削除対象のキーフレームが見つかりませんでした" });
		}
	} catch (e) {
		app.endUndoGroup(); // エラー時も閉じる
		return JSON.stringify({ status: "error", message: "Fatal Error: " + e.toString() });
	}
}

function removeKeysRecursive(comp, time, counter) {
	for (var i = 1; i <= comp.numLayers; i++) {
		var layer = comp.layer(i);

		// 現在のレイヤーのプロパティからキーを削除
		removeKeysFromLayer(layer, time, counter);

		// If it's a pre-comp, recurse
		if (layer.source instanceof CompItem) {
			// Calculate inner time for the pre-comp
			if (layer.timeRemapEnabled) {
				// タイムリマップ有効の場合、現在時刻でのタイムリマップ値を取得して再帰
				var trProp = layer.property("Time Remap");
				if (trProp) {
					var mappedTime = trProp.valueAtTime(time, false);
					removeKeysRecursive(layer.source, mappedTime, counter);
				}
			} else {
				// 通常の時間計算
				// comp time -> layer local time -> source comp time
				// localTime = (compTime - startTime) * stretchFactor
				// stretch: 100 means 100% speed. stretch 200 means 50% speed (2x duration).
				// AE Scripting Guide: layer.stretch is a percentage.
				// But formula is specific.
				// Correct formula: (time - layer.startTime) * (100 / layer.stretch)
				var localTime = (time - layer.startTime) * (100 / layer.stretch);
				removeKeysRecursive(layer.source, localTime, counter);
			}
		}
	}
}

function removeKeysFromLayer(propGroup, time, counter) {
	var numProps = propGroup.numProperties;
	if (!numProps) return;

	for (var i = 1; i <= numProps; i++) {
		var prop = propGroup.property(i);

		if (prop.propertyType === PropertyType.PROPERTY) {
			if (prop.numKeys > 0) {
				// nearestKeyIndex returns index of key closest to time
				var nearestIndex = prop.nearestKeyIndex(time);

				// nearestKeyIndex can return 0 if no keys, but numKeys > 0 checked above.
				// It returns a valid index 1..numKeys
				if (nearestIndex > 0 && nearestIndex <= prop.numKeys) {
					var keyTime = prop.keyTime(nearestIndex);

					// Floating point tolerance
					if (Math.abs(keyTime - time) < 0.0001) {
						prop.removeKey(nearestIndex);
						counter.count++;
						// 同じ時間に重複キーは存在しないはずなので、このプロパティでの削除は完了
					}
				}
			}
		} else if (prop.propertyType === PropertyType.NAMED_GROUP || prop.propertyType === PropertyType.INDEXED_GROUP) {
			removeKeysFromLayer(prop, time, counter);
		}
	}
}

function getHierarchy() {
	var comp = app.project.activeItem;
	if (!(comp instanceof CompItem)) {
		return JSON.stringify({ status: "error", message: "コンポジションを選択してください。" });
	}

	var projectPath = "";
	if (app.project.file) {
		projectPath = app.project.file.fsName;
	}

	var root = {
		name: comp.name,
		id: comp.id,
		projectPath: projectPath,
		type: "root",
		children: parseComp(comp)
	};
	return JSON.stringify(root);
}

function parseComp(comp) {
	var children = [];
	for (var i = 1; i <= comp.numLayers; i++) {
		var layer = comp.layer(i);
		var item = {
			index: layer.index,
			layerId: layer.id, // CSInterface / modern AE needed
			name: layer.name,
			visible: layer.enabled, // 簡易的な表示状態 (目のアイコン)
			opacity: layer.opacity.value // 不透明度
		};

		// レイヤ名の接頭辞判定
		if (item.name.indexOf("*") === 0) item.mode = "radio";
		else if (item.name.indexOf("!") === 0) item.mode = "locked";
		else item.mode = "check";

		// プリコンポジションの場合、再帰的に解析するか？
		// PSDToolの仕様上、フォルダはプリコンプになることが多い
		// ここでは、ソースがCompItemなら「フォルダ」扱いする
		if (layer.source instanceof CompItem) {
			item.type = "folder";
			item.compId = layer.source.id; // 子コンポジションのID
			// item.children = parseComp(layer.source); // ここで再帰すると巨大になる可能性があるが、構造把握には必要
			// 今回は、再帰的に取得する
			item.children = parseComp(layer.source);
		} else {
			item.type = "layer";
		}
		children.push(item);
	}
	return children;
}

// レイヤーの表示・非表示を切り替える（キーフレーム作成）
function setLayerVisibility(compId, layerIndex, visible) {
	var comp = getCompById(compId);
	if (!comp) return "error: comp not found";

	var layer = comp.layer(layerIndex);
	if (!layer) return "error: layer not found";

	// 表示状態の制御
	// PSDTool的なアプローチでは、Opacityで制御することが多い（フェードなどはしない前提）
	// あるいは「目のアイコン」 (enabled) だが、enabledはキーフレーム打てない。
	// Opacityにキーフレームを打つ。

	var opacityProp = layer.property("Opacity");
	var time = comp.time;
	var value = visible ? 100 : 0;

	// もしレイヤーが無効化(目のアイコンOFF)されていたら、表示時には有効化する
	if (visible) layer.enabled = true;

	// 現在の時間にキーフレームを追加
	var keyIndex = opacityProp.addKey(time);
	opacityProp.setValueAtKey(keyIndex, value);

	// ホールドキーフレームに設定 (ToggleHoldKeyframe 相当)
	opacityProp.setInterpolationTypeAtKey(keyIndex, KeyframeInterpolationType.HOLD, KeyframeInterpolationType.HOLD);

	return "success";
}

// 複数のレイヤーを一括設定（ラジオボタン用）
// targetLayerIndexをON、それ以外のsiblingsをOFFにする
function setRadioVisibility(compId, targetIndex, siblingIndices) {
	app.beginUndoGroup("Radio Visibility Change");
	var comp = getCompById(compId);
	if (!comp) return;

	var time = comp.time;

	// Target ON
	var targetLayer = comp.layer(targetIndex);
	if (targetLayer) targetLayer.enabled = true; // 強制的に有効化

	var tProp = targetLayer.property("Opacity");
	var tKey = tProp.addKey(time);
	tProp.setValueAtKey(tKey, 100);
	tProp.setInterpolationTypeAtKey(tKey, KeyframeInterpolationType.HOLD, KeyframeInterpolationType.HOLD);

	// Siblings OFF
	for (var i = 0; i < siblingIndices.length; i++) {
		var idx = siblingIndices[i];
		if (idx === targetIndex) continue;

		var layer = comp.layer(idx);
		// ロックされているものは無視するなど必要ならここで判定
		// レイヤ名の先頭が ! ならスキップ
		if (layer.name.indexOf("!") === 0) continue;

		var prop = layer.property("Opacity");
		var key = prop.addKey(time);
		prop.setValueAtKey(key, 0);
		prop.setInterpolationTypeAtKey(key, KeyframeInterpolationType.HOLD, KeyframeInterpolationType.HOLD);
	}
	app.endUndoGroup();
	return "success";
}

function getCompById(id) {
	for (var i = 1; i <= app.project.numItems; i++) {
		var item = app.project.item(i);
		if (item instanceof CompItem && item.id == id) {
			return item;
		}
	}
	return null;
}

function applyVisibilityToCompRaw(comp, items) {
	var time = comp.time;
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		var layer = comp.layer(item.index);
		if (!layer) continue;
		if (layer.name.indexOf("!") === 0) continue;

		if (item.visible) layer.enabled = true;
		var prop = layer.property("Opacity");
		var keyIndex = prop.addKey(time);
		var val = item.visible ? 100 : 0;
		prop.setValueAtKey(keyIndex, val);
		prop.setInterpolationTypeAtKey(keyIndex, KeyframeInterpolationType.HOLD, KeyframeInterpolationType.HOLD);
	}
}

// 全コンポジションに対して一括適用（単一のアンドゥグループ）
function applyGlobalBatchVisibility(globalDataStr) {
	app.beginUndoGroup("Apply Preset Global");
	try {
		var globalData = globalDataStr;
		if (typeof globalData === 'string') {
			try { globalData = JSON.parse(globalData); } catch (e) { }
		}

		if (!globalData || typeof globalData !== 'object') {
			return JSON.stringify({ status: "error", message: "Invalid data" });
		}

		for (var compId in globalData) {
			if (globalData.hasOwnProperty(compId)) {
				var comp = getCompById(compId);
				if (comp) {
					applyVisibilityToCompRaw(comp, globalData[compId]);
				}
			}
		}
	} catch (e) {
		return JSON.stringify({ status: "error", message: e.toString() });
	} finally {
		app.endUndoGroup();
	}
	return JSON.stringify({ status: "success" });
}

// プリセット用の一括適用
function applyBatchVisibility(compId, visibilityData) {
	app.beginUndoGroup("Apply Preset");
	var comp = getCompById(compId);
	if (!comp) {
		app.endUndoGroup();
		return JSON.stringify({ status: "error", message: "Comp not found" });
	}

	var data = visibilityData;
	if (typeof visibilityData === 'string') {
		try {
			data = JSON.parse(visibilityData);
		} catch (e) {
			app.endUndoGroup();
			return JSON.stringify({ status: "error", message: "Invalid JSON: " + e.message });
		}
	}

	if (!data || typeof data.length === 'undefined') {
		app.endUndoGroup();
		return JSON.stringify({ status: "error", message: "Data is not an array" });
	}

	applyVisibilityToCompRaw(comp, data);

	app.endUndoGroup();
	return JSON.stringify({ status: "success" });
}

var g_capturedKeyframes = [];

function captureKeyframes() {
	try {
		var comp = app.project.activeItem;
		if (!(comp instanceof CompItem)) {
			return JSON.stringify({ status: 'error', message: 'コンポジションを選択してください。' });
		}

		g_capturedKeyframes = [];
		// 再帰的に検索 (初期チェーンは空)
		scanLayersRecursive(comp, comp.time, []);

		return JSON.stringify({ status: 'success', count: g_capturedKeyframes.length });
	} catch (e) {
		return JSON.stringify({ status: 'error', message: e.toString() + ' (Line: ' + e.line + ')' });
	}
}

function scanLayersRecursive(comp, currentTime, layerChain) {
	for (var i = 1; i <= comp.numLayers; i++) {
		var layer = comp.layer(i);

		// 可視レイヤーのみ、あるいは選択レイヤーのみ？ -> 全レイヤー対象

		if (layer.source instanceof CompItem) {
			// プリコンポジションの場合 (レイヤー自体のプロパティもスキャンする)
			scanProperties(layer, currentTime, layerChain);

			// 内部時間の計算
			var childTime;
			if (layer.timeRemapEnabled) {
				// タイムリマップ有効時
				childTime = layer.property('Time Remap').valueAtTime(currentTime, false);
			} else {
				// 通常時: (親時間 - 開始時間) * (100 / ストレッチ)
				childTime = (currentTime - layer.startTime) * (100 / layer.stretch);
			}

			// 子コンポジションへ再帰 (チェーンにこのレイヤーを追加)
			var newChain = layerChain.concat(layer);
			scanLayersRecursive(layer.source, childTime, newChain);

		} else {
			// 通常レイヤー
			scanProperties(layer, currentTime, layerChain);
		}
	}
}

function scanProperties(propGroup, time, layerChain) {
	var numProps = 0;
	try { numProps = propGroup.numProperties; } catch (e) { return; }

	for (var i = 1; i <= numProps; i++) {
		var prop = propGroup.property(i);

		if (prop.propertyType === PropertyType.PROPERTY) {
			// プロパティの場合、キーフレームがあるかチェック
			if (prop.numKeys > 0) {
				// 指定時間の近くにキーフレームがあるか？
				var nearestKeyIndex = 0;
				try { nearestKeyIndex = prop.nearestKeyIndex(time); } catch (e) { continue; }

				if (nearestKeyIndex > 0) {
					var keyTime = prop.keyTime(nearestKeyIndex);
					// 許容誤差 (0.001s)
					var epsilon = 0.001;
					if (Math.abs(keyTime - time) < epsilon) {
						try {
							// ユーザー要件: "collects keyframes at the 'current playback head position' ... moves to the 'new current playback head position'"
							// ここでは、選択されているかどうかは確認せず、時間位置にあるものは全て拾う。

							g_capturedKeyframes.push({
								property: prop,
								keyIndex: nearestKeyIndex,
								originalTime: keyTime,
								layerChain: layerChain.slice() // Pathを保存
							});
						} catch (e) {
							// NOP
						}
					}
				}
			}
		} else if (prop.propertyType === PropertyType.NAMED_GROUP || prop.propertyType === PropertyType.INDEXED_GROUP) {
			// グループの場合、再帰
			scanProperties(prop, time, layerChain);
		}
	}
}


function moveKeyframes() {
	app.beginUndoGroup('Move Keyframes Hierarchy');
	try {
		var rootComp = app.project.activeItem;
		if (!(rootComp instanceof CompItem)) {
			return JSON.stringify({ status: 'error', message: 'コンポジションを選択してください。' });
		}

		var rootTime = rootComp.time;
		var movedCount = 0;

		for (var i = 0; i < g_capturedKeyframes.length; i++) {
			var item = g_capturedKeyframes[i];
			var prop = item.property;

			// プロパティ有効性チェック
			try {
				if (!prop.isValid) continue;
				if (prop.numKeys === 0) continue;
			} catch (e) { continue; }

			// ルート時間からローカル時間に変換
			var targetTime = rootTime;

			for (var j = 0; j < item.layerChain.length; j++) {
				var layer = item.layerChain[j];
				try {
					if (layer.timeRemapEnabled) {
						// タイムリマップ: 親時間におけるリマップ値が子の時間
						targetTime = layer.property('Time Remap').valueAtTime(targetTime, false);
					} else {
						targetTime = (targetTime - layer.startTime) * (100 / layer.stretch);
					}
				} catch (e) {
					// レイヤー喪失等
				}
			}

			// 移動処理
			// 1. 元のキーを探す (収集時と同じ時間にあるか？)
			var currentIndex = 0;
			// まずitem.keyIndexを確認
			if (prop.numKeys >= item.keyIndex && Math.abs(prop.keyTime(item.keyIndex) - item.originalTime) < 0.001) {
				currentIndex = item.keyIndex;
			} else {
				// インデックスずれの可能性
				currentIndex = prop.nearestKeyIndex(item.originalTime);
				if (Math.abs(prop.keyTime(currentIndex) - item.originalTime) > 0.001) {
					// キーが見つからない、移動済み、削除済み
					continue;
				}
			}

			// 元の時間と移動後の時間がほぼ同じなら何もしない
			if (Math.abs(targetTime - item.originalTime) < 0.001) continue;

			// キー情報の抽出・コピー
			var keyVal, inInterp, outInterp;
			var inTan, outTan, spatialInTan, spatialOutTan;
			var isSpatial = (prop.propertyValueType === PropertyValueType.TwoD_SPATIAL || prop.propertyValueType === PropertyValueType.ThreeD_SPATIAL);
			var roving = false;
			var tempAutoBezier, temporalContinuous;

			// コピー
			keyVal = prop.keyValue(currentIndex);
			inInterp = prop.keyInInterpolationType(currentIndex);
			outInterp = prop.keyOutInterpolationType(currentIndex);

			if (inInterp !== KeyframeInterpolationType.HOLD && inInterp !== KeyframeInterpolationType.LINEAR) {
				inTan = prop.keyInTemporalEase(currentIndex);
			}
			if (outInterp !== KeyframeInterpolationType.HOLD && outInterp !== KeyframeInterpolationType.LINEAR) {
				outTan = prop.keyOutTemporalEase(currentIndex);
			}

			if (isSpatial) {
				spatialInTan = prop.keyInSpatialTangent(currentIndex);
				spatialOutTan = prop.keyOutSpatialTangent(currentIndex);
				roving = prop.keyRoving(currentIndex);
			}

			tempAutoBezier = prop.keyTemporalAutoBezier(currentIndex);
			temporalContinuous = prop.keyTemporalContinuous(currentIndex);

			// 削除
			prop.removeKey(currentIndex);

			// 追加
			var newKeyIndex = prop.addKey(targetTime);

			// 値・補間設定
			if (keyVal !== null) prop.setValueAtKey(newKeyIndex, keyVal);

			if (inInterp && outInterp) {
				prop.setInterpolationTypeAtKey(newKeyIndex, inInterp, outInterp);
			}

			if (inTan && outTan) {
				prop.setTemporalEaseAtKey(newKeyIndex, inTan, outTan);
			}

			if (isSpatial) {
				prop.setSpatialTangentsAtKey(newKeyIndex, spatialInTan, spatialOutTan);
				prop.setRovingAtKey(newKeyIndex, roving); // Roving設定
			}

			prop.setTemporalAutoBezierAtKey(newKeyIndex, tempAutoBezier);
			prop.setTemporalContinuousAtKey(newKeyIndex, temporalContinuous);

			// 最後に選択
			prop.setSelectedAtKey(newKeyIndex, true);

			movedCount++;
		}

		return JSON.stringify({ status: 'success', count: movedCount });

	} catch (e) {
		return JSON.stringify({ status: 'error', message: e.toString() + ' (Line: ' + e.line + ')' });
	} finally {
		app.endUndoGroup();
	}
}

function moveTimeToSelectedLayerInPoint() {
	var comp = app.project.activeItem;
	if (!comp || !(comp instanceof CompItem)) {
		return "false";
	}
	var selectedLayers = comp.selectedLayers;
	if (selectedLayers.length > 0) {
		comp.time = selectedLayers[0].inPoint;
		return "true";
	}
	return "false";
}

