
var csInterface = new CSInterface();

// テーマ適用の初期実行とイベントリスナー登録
updateTheme();
csInterface.addEventListener(CSInterface.THEME_COLOR_CHANGED_EVENT, updateTheme);

// ウィンドウサイズを画面サイズの約半分に設定
try {
	var h = Math.floor(window.screen.availHeight * 0.5);
	var w = 400; // 幅は元のまま維持
	csInterface.resizeContent(w, h);
} catch (e) {
	console.log("Resize failed/ignored");
}


document.getElementById('btn-reload').addEventListener('click', loadHierarchy);
document.getElementById('btn-collapse-all').addEventListener('click', collapseAllFolders);
document.getElementById('btn-add-preset').addEventListener('click', savePreset);
document.getElementById('btn-export-preset').addEventListener('click', exportPresets);
document.getElementById('btn-import-preset').addEventListener('click', importPresets);

// Keyframe Operation Events
document.getElementById('btn-get-keys').addEventListener('click', getKeyframesFromActiveComp);
document.getElementById('btn-move-keys').addEventListener('click', moveKeyframesToCurrentTime);
document.getElementById('btn-delete-keys').addEventListener('click', removeKeyframesAtCurrentTime);

// 初期状態: moveボタンは無効化
var btnMoveKeys = document.getElementById('btn-move-keys');
if (btnMoveKeys) btnMoveKeys.disabled = true;

function getKeyframesFromActiveComp() {
	var statusDiv = document.getElementById('status-keyframe');
	statusDiv.textContent = "取得中...";
	statusDiv.style.color = "#eeeeee";
	statusDiv.style.fontWeight = "normal";

	if (btnMoveKeys) btnMoveKeys.disabled = true;

	csInterface.evalScript('getKeyframesFromActiveComp()', function (res) {
		if (res === "true") {
			statusDiv.textContent = "キーフレームを記録しました";
			statusDiv.style.color = "#4CAF50"; // 柔らかめの緑
			statusDiv.style.fontWeight = "bold";
			if (btnMoveKeys) btnMoveKeys.disabled = false;
		} else {
			statusDiv.textContent = "キーフレームが見つかりませんでした";
			statusDiv.style.color = "#ff4444"; // 鮮やかな赤
			statusDiv.style.fontWeight = "bold";
			if (btnMoveKeys) btnMoveKeys.disabled = true;
		}
	});
}

function moveKeyframesToCurrentTime() {
	var statusDiv = document.getElementById('status-keyframe');
	statusDiv.textContent = "移動中...";
	statusDiv.style.color = "#eeeeee";
	statusDiv.style.fontWeight = "normal";

	csInterface.evalScript('moveKeyframesToCurrentTime()', function (res) {
		var resultObj = { count: 0, status: "false" };
		try {
			// true/false以外にJSONが返る可能性もあるが、一旦単純化
			if (res === "true" || res === "false") {
				// 旧仕様対策
			} else {
				resultObj = JSON.parse(res);
			}
		} catch (e) { }

		if (res === "true" || resultObj.status === "success") {
			statusDiv.textContent = "移動完了 (メモリ消去済)";
			if (btnMoveKeys) btnMoveKeys.disabled = true;
			statusDiv.style.color = "#4CAF50"; // 柔らかめの緑
			statusDiv.style.fontWeight = "bold";
			setTimeout(function () {
				statusDiv.textContent = "";
				statusDiv.style.fontWeight = "normal";
			}, 3000);
		} else {
			statusDiv.textContent = resultObj.message ? resultObj.message : "記録中のキーフレームがありません";
			if (btnMoveKeys) btnMoveKeys.disabled = true;
			statusDiv.style.color = "#ff4444"; // 鮮やかな赤
			statusDiv.style.fontWeight = "bold";
		}
	});
}
function removeKeyframesAtCurrentTime() {
	if (!confirm("現在の時間にあるキーフレームを全階層から削除しますか？\nこの操作は元に戻せません。")) {
		return;
	}

	var statusDiv = document.getElementById('status-keyframe');
	statusDiv.textContent = "削除中...";
	statusDiv.style.color = "#eeeeee";
	statusDiv.style.fontWeight = "normal";

	csInterface.evalScript('removeKeyframesAtCurrentTime()', function (res) {
		var resultObj = { count: 0, status: "false" };
		try {
			if (res.startsWith("{")) {
				resultObj = JSON.parse(res);
			} else {
				// Fallback for simple boolean string responses (though hostscript creates JSON)
				if (res === "true") resultObj.status = "success";
			}
		} catch (e) { }

		if (resultObj.status === "success") {
			statusDiv.textContent = "削除完了 (" + resultObj.count + "個)";
			statusDiv.style.color = "#ff4444";
			statusDiv.style.fontWeight = "bold";
			setTimeout(function () {
				statusDiv.textContent = "";
				statusDiv.style.fontWeight = "normal";
			}, 3000);
		} else {
			statusDiv.textContent = resultObj.message ? resultObj.message : "削除に失敗、またはキーが見つかりません";
			statusDiv.style.color = "#ff4444";
			statusDiv.style.fontWeight = "bold";
		}
	});
}


// Storage Management
var modal = document.getElementById("storage-modal");
var btnManage = document.getElementById("btn-manage-storage");
var spanClose = document.getElementsByClassName("close-modal")[0];

btnManage.onclick = function () {
	renderStorageList();
	modal.style.display = "block";
}

spanClose.onclick = function () {
	modal.style.display = "none";
}

window.onclick = function (event) {
	if (event.target == modal) {
		modal.style.display = "none";
	}
}

// 起動時に自動読み込み
setTimeout(loadHierarchy, 100);

// 状態保存用グローバル変数
var g_expansionState = {}; // { compId: { "layerKey": isExpanded, ... } }
var g_lastCompId = null;
var g_currentProjectPath = null;

// Preset Logic
function getPresetKey(compId) {
	var pathPart = "";
	if (g_currentProjectPath) {
		pathPart = "_" + encodeURIComponent(g_currentProjectPath);
	}
	return 'psdtool_presets' + pathPart + '_' + compId;
}

function loadPresets(compId) {
	var list = document.getElementById('preset-list');
	list.innerHTML = '';

	var json = localStorage.getItem(getPresetKey(compId));
	if (!json) return;

	try {
		var presets = JSON.parse(json);
		presets.forEach(function (p, i) {
			var item = document.createElement('div');
			item.className = 'preset-item';

			var nameSpan = document.createElement('span');
			nameSpan.className = 'preset-name';
			nameSpan.textContent = p.name;
			nameSpan.onclick = function (e) { applyPreset(p.data, e); };
			item.appendChild(nameSpan);

			// Rename Button
			var renameBtn = document.createElement('span');
			renameBtn.className = 'preset-rename';
			renameBtn.textContent = '✎';
			renameBtn.onclick = function (e) {
				e.stopPropagation();
				renamePreset(compId, i, p.name);
			};
			item.appendChild(renameBtn);

			var delBtn = document.createElement('span');
			delBtn.className = 'preset-delete';
			delBtn.textContent = '×';
			delBtn.onclick = function (e) {
				e.stopPropagation();
				if (confirm('プリセット "' + p.name + '" を削除しますか？')) {
					deletePreset(compId, i);
				}
			};
			item.appendChild(delBtn);

			list.appendChild(item);
		});
	} catch (e) {
		console.error("Preset load error", e);
	}
}

function savePreset() {
	if (!g_lastCompId) {
		alert("コンポジションが読み込まれていません。");
		return;
	}

	var name = prompt("プリセット名を入力してください", "Preset " + (new Date().toLocaleTimeString()));
	if (!name) return;

	// 現在のチェック状態を取得
	var currentData = [];
	var lis = document.querySelectorAll('li[data-layer-key]');
	lis.forEach(function (li) {
		var key = li.dataset.layerKey;
		var input = li.querySelector('input');
		var nameSpan = li.querySelector('.layer-name');
		var layerName = nameSpan ? nameSpan.textContent : null;

		if (input && !input.disabled && (input.type === 'checkbox' || input.type === 'radio')) {
			currentData.push({
				key: key,
				layerName: layerName,
				visible: input.checked
			});
		}
	});

	var key = getPresetKey(g_lastCompId);
	var presets = [];
	var json = localStorage.getItem(key);
	if (json) {
		try { presets = JSON.parse(json); } catch (e) { }
	}

	presets.push({
		name: name,
		data: currentData
	});

	localStorage.setItem(key, JSON.stringify(presets));
	loadPresets(g_lastCompId);
}

function renamePreset(compId, index, oldName) {
	var newName = prompt("新しいプリセット名を入力してください", oldName);
	if (newName && newName !== oldName) {
		var key = getPresetKey(compId);
		var json = localStorage.getItem(key);
		if (json) {
			var presets = JSON.parse(json);
			if (presets[index]) {
				presets[index].name = newName;
				localStorage.setItem(key, JSON.stringify(presets));
				loadPresets(compId);
			}
		}
	}
}

function deletePreset(compId, index) {
	var key = getPresetKey(compId);
	var json = localStorage.getItem(key);
	if (json) {
		var presets = JSON.parse(json);
		presets.splice(index, 1);
		localStorage.setItem(key, JSON.stringify(presets));
		loadPresets(compId);
	}
}

function applyPreset(savedData, e) {
	if (!g_lastCompId) return;

	// Ctrlキーが押されている場合の特別処理
	if (e && e.ctrlKey) {
		csInterface.evalScript('moveTimeToSelectedLayerInPoint()', function (res) {
			// 移動処理が終わったら続きを実行
			// 選択がない場合は移動しないだけで通常動作
			continueApplyPreset(savedData);
		});
	} else {
		continueApplyPreset(savedData);
	}
}

function continueApplyPreset(savedData) {
	// 1. まず最新の階層構造を取得する (インデックスのズレ防止)
	csInterface.evalScript('getHierarchy()', function (res) {
		var hierarchy = JSON.parse(res);
		if (hierarchy.status === 'error') {
			alert("階層データの取得に失敗しました: " + hierarchy.message);
			return;
		}

		// 2. プリセットデータをMap化して検索しやすくする
		// Key: "ID:(id)" or "NAME:(name)" -> Value: boolean (visible)
		var presetMap = {};
		savedData.forEach(function (item) {
			if (item.key) presetMap[item.key] = item.visible;
		});

		var batchDataByComp = {}; // { compId: [ {index, visible}, ... ] }

		// 3. 最新の階層ツリーを走査して、プリセットに含まれるレイヤーを探す
		// 同時にツリーデータの visible プロパティも更新して、UI描画に備える
		function traverseAndMatch(node, currentCompId) {
			// 現在のCompIDでリスト初期化
			if (!batchDataByComp[currentCompId]) {
				batchDataByComp[currentCompId] = [];
			}

			// 子要素（レイヤー）の走査
			if (node.children) {
				node.children.forEach(function (child) {
					// もしchildがフォルダ（プリコンプ）なら、その中身はそのプリコンプIDに属する
					// currentCompIdは「childが存在するCompのID」

					// まず自分自身の処理
					// ルート以外のレイヤー/フォルダ
					if (child.index !== undefined) {
						var key = getLayerKey(child);
						if (presetMap.hasOwnProperty(key)) {
							var shouldVisible = presetMap[key];

							// AE送信データに追加 (現在のCompIDに対して)
							batchDataByComp[currentCompId].push({
								index: child.index,
								visible: shouldVisible
							});

							// UI用データ更新
							child.visible = shouldVisible;
							child.opacity = shouldVisible ? 100 : 0;
						}
					}

					// 再帰処理
					// もしこのchildがフォルダ(comp)なら、そのchildrenは child.compId に属する
					if (child.children && child.children.length > 0) {
						var nextCompId = child.compId ? child.compId : currentCompId;
						traverseAndMatch(child, nextCompId);
					}
				});
			}
		}

		// 初期呼び出し: hierarchyの直下は g_lastCompId (または hierarchy.id) に属する
		traverseAndMatch(hierarchy, hierarchy.id);

		// 4. UIの再描画 (ユーザーへの即時フィードバック)
		renderTree(hierarchy);
		if (hierarchy.id) loadPresets(hierarchy.id);

		// 5. AEへ送信 (Compごとに送信)
		// applyBatchVisibility はまだ一回のUndoGroupに対応していないので、
		// 複数回呼ぶとUndoが分かれてしまう可能性があるが、まずは機能修正優先。
		// ホストスクリプト側でまとめて処理するよう変更するのがベストだが、今回はループ呼び出しで対応。

		var compIds = Object.keys(batchDataByComp);

		// 連続呼び出し用チェーン関数
		function runBatch(index) {
			if (index >= compIds.length) return;
			var cId = compIds[index];
			var items = batchDataByComp[cId];
			if (items.length === 0) {
				runBatch(index + 1);
				return;
			}

			csInterface.evalScript('applyBatchVisibility(' + cId + ', ' + JSON.stringify(items) + ')', function (res) {
				console.log("Preset applied for comp " + cId + ": " + res);
				runBatch(index + 1);
			});
		}

		runBatch(0);
	});
}

function loadHierarchy() {
	// 現在の状態を保存してからロード
	saveExpansionState();

	csInterface.evalScript('getHierarchy()', function (res) {
		var data = JSON.parse(res);
		if (data.status === 'error') {
			renderMessage("コンポジションが開かれていません。<br>対象のコンポジションを開いてから「再読み込み」を押してください。");
		} else {
			if (data.projectPath) {
				g_currentProjectPath = data.projectPath;
			} else {
				g_currentProjectPath = null; // 未保存の場合は共用または別扱い
			}

			renderTree(data);
			if (data.id) loadPresets(data.id);
		}
	});
}

function saveExpansionState() {
	if (!g_lastCompId) return;

	var state = {};
	// ツリー内のすべてのLIをチェックし、展開状態（ulがblockかどうか）を記録
	// data-layer-key を持つ要素のみ対象
	var lis = document.querySelectorAll('li[data-layer-key]');
	lis.forEach(function (li) {
		var key = li.dataset.layerKey;
		var childUl = li.querySelector('ul');
		if (childUl && childUl.style.display !== 'none') {
			state[key] = true;
		}
	});

	g_expansionState[g_lastCompId] = state;
}

function renderMessage(msg) {
	var container = document.getElementById('tree-container');
	container.innerHTML = '<p style="color: #bbb; text-align: center; margin-top: 20px;">' + msg + '</p>';
}

function collapseAllFolders() {
	var uls = document.querySelectorAll('#tree-container ul ul'); // Root以外のUL
	uls.forEach(function (ul) {
		ul.style.display = 'none';
	});
	// アイコンも戻す
	var toggles = document.querySelectorAll('.toggle-icon');
	toggles.forEach(function (t) {
		if (t.textContent === '▼') t.textContent = '▶';
	});
}

function renderTree(data) {
	var container = document.getElementById('tree-container');
	container.innerHTML = '';

	if (data.status === 'error') {
		renderMessage(data.message);
		return;
	}

	// カレントのCompIDを記録（次回の保存のため）
	g_lastCompId = data.id;

	var rootUl = document.createElement('ul');
	rootUl.className = 'tree-root';

	// ルートの子要素（実際のレイヤー群）を描画. Depth starts at 1.
	data.children.forEach(function (child) {
		rootUl.appendChild(createNode(child, data.id, 1, true));
	});

	container.appendChild(rootUl);
}

function getLayerKey(item) {
	if (item.layerId !== undefined && item.layerId !== null) {
		return "ID:" + item.layerId;
	}
	// Fallback
	return "NAME:" + item.name;
}

function createNode(item, parentCompId, depth, parentVisible) {
	var li = document.createElement('li');

	var div = document.createElement('div');
	div.className = 'item-container';

	// 現在のアイテムの可視状態（チェックされているか）を判定
	var isSelfVisible = false;
	if (item.mode === 'locked') {
		isSelfVisible = true;
	} else {
		if (!item.visible || (item.opacity !== undefined && item.opacity <= 0)) {
			isSelfVisible = false;
		} else {
			isSelfVisible = true;
		}
	}

	// フォルダ展開用トグル
	var toggle = document.createElement('span');
	toggle.className = 'toggle-icon';

	// 展開状態の決定
	// 1. 保存された状態(g_expansionState)があればそれを優先
	// 2. なければ、デフォルトロジックに従う

	var isExpanded = false;

	// キーの生成
	var layerKey = getLayerKey(item);

	// 復元ロジック
	var savedState = g_expansionState[g_lastCompId];
	var hasSavedState = (savedState !== undefined);

	if (hasSavedState) {
		// 保存されていた場合、そのKeyがTrueなら展開
		if (savedState[layerKey] === true) {
			isExpanded = true;
		} else {
			isExpanded = false;
		}
	} else {
		// デフォルトロジック
		if (depth < 2) {
			// 第一階層: 自身が表示状態なら展開
			isExpanded = isSelfVisible;
		} else {
			isExpanded = false;
		}
	}

	if (item.children && item.children.length > 0) {
		// 現在の状態に合わせてアイコン設定
		toggle.textContent = isExpanded ? '▼' : '▶';

		toggle.onclick = function (e) {
			var childUl = li.querySelector('ul');
			if (childUl) {
				var isHidden = childUl.style.display === 'none';
				var targetDisplay = isHidden ? 'block' : 'none';
				var targetIcon = isHidden ? '▼' : '▶';

				if (e.altKey) {
					// Altキーが押されていたら、配下すべてのフォルダに適用
					var allUls = li.querySelectorAll('ul');
					allUls.forEach(function (u) { u.style.display = targetDisplay; });

					var allToggles = li.querySelectorAll('.toggle-icon');
					allToggles.forEach(function (t) {
						if (t.textContent.trim() !== '') {
							t.textContent = targetIcon;
						}
					});
				} else {
					// 通常動作
					childUl.style.display = targetDisplay;
					toggle.textContent = targetIcon;
				}
			}
		};
	} else {
		toggle.innerHTML = '&nbsp;';
	}
	div.appendChild(toggle);

	// 入力コントロール
	var input = document.createElement('input');

	if (item.mode === 'locked') {
		input.type = 'checkbox';
		input.checked = true;
		input.disabled = true;
	} else if (item.mode === 'radio') {
		input.type = 'radio';
		input.name = 'radio_group_' + parentCompId;
	} else {
		input.type = 'checkbox';
	}

	// 初期状態反映
	if (item.mode !== 'locked') {
		if (!isSelfVisible) {
			input.checked = false;
		} else {
			input.checked = true;
		}
	}

	// ロジック実行関数
	var executeLogic = function () {
		var currentCompId = parentCompId;

		if (item.mode === 'radio') {
			var siblingsIndices = [];
			var parentUl = li.parentElement;
			var childLis = parentUl.children;
			for (var i = 0; i < childLis.length; i++) {
				var cLi = childLis[i];
				var cInput = cLi.querySelector('input');
				if (cInput && cInput.type === 'radio') {
					siblingsIndices.push(parseInt(cLi.dataset.index));
				}
			}

			csInterface.evalScript('setRadioVisibility(' + currentCompId + ', ' + item.index + ', ' + JSON.stringify(siblingsIndices) + ')', function (res) {
				console.log(res);
			});

		} else if (item.mode === 'check' && !input.disabled) {
			var visible = input.checked;
			csInterface.evalScript('setLayerVisibility(' + currentCompId + ', ' + item.index + ', ' + visible + ')', function (res) {
				console.log(res);
			});
		}
	};

	input.onclick = function (e) {
		executeLogic();
	};

	div.appendChild(input);

	var labelClick = function (e) {
		if (input.disabled) return;
		input.click();
	};

	var icon = document.createElement('span');
	icon.className = item.type === 'folder' ? 'layer-icon folder-icon' : 'layer-icon';
	icon.onclick = labelClick;
	div.appendChild(icon);

	var label = document.createElement('span');
	label.className = 'layer-name ' + (item.mode === 'locked' ? 'locked' : '');
	label.textContent = item.name;
	label.onclick = labelClick;
	div.appendChild(label);

	li.appendChild(div);
	li.dataset.index = item.index;
	li.dataset.layerKey = layerKey; // キー保存

	if (item.children && item.children.length > 0) {
		var childUl = document.createElement('ul');
		childUl.style.display = isExpanded ? 'block' : 'none';

		var nextCompId = item.compId ? item.compId : parentCompId;

		item.children.forEach(function (child) {
			// 子階層へも isSelfVisibleなどの状態を渡す（現状ではロジックに使っていないが、構造維持）
			childUl.appendChild(createNode(child, nextCompId, depth + 1, isSelfVisible));
		});
		li.appendChild(childUl);
	}

	return li;
}

function updateTheme() {
	var hostEnv = csInterface.getHostEnvironment();
	if (!hostEnv || !hostEnv.appSkinInfo) return;

	var skin = hostEnv.appSkinInfo;
	var panelBg = skin.panelBackgroundColor.color;

	var bgStr = toHex(panelBg);
	document.body.style.backgroundColor = bgStr;

	var brightness = (panelBg.red * 299 + panelBg.green * 587 + panelBg.blue * 114) / 1000;

	if (brightness > 128) {
		// Light Theme
		document.body.style.color = "#333333";
		setCSSVar('--btn-bg', "#e0e0e0");
		setCSSVar('--btn-fg', "#333333");
		setCSSVar('--btn-border', "#aaaaaa");
		setCSSVar('--btn-hover', "#d0d0d0");
		setCSSVar('--item-hover', "#e8e8e8");
	} else {
		// Dark Theme
		document.body.style.color = "#cccccc";
		setCSSVar('--btn-bg', "#383838");
		setCSSVar('--btn-fg', "#f0f0f0");
		setCSSVar('--btn-border', "#505050");
		setCSSVar('--btn-hover', "#484848");
		setCSSVar('--item-hover', "#2a2a2a");
	}

	if (skin.baseFontSize) {
		document.body.style.fontSize = skin.baseFontSize + "px";
	}
}

function toHex(color) {
	function hex(n) { return (n < 16 ? "0" : "") + Math.round(n).toString(16); }
	return "#" + hex(color.red) + hex(color.green) + hex(color.blue);
}

function setCSSVar(name, value) {
	document.documentElement.style.setProperty(name, value);
}

// ------ Import / Export Logic ------

function exportPresets() {
	if (!g_lastCompId) {
		alert("コンポジションが読み込まれていません。");
		return;
	}

	var key = getPresetKey(g_lastCompId);
	var json = localStorage.getItem(key);

	if (!json || JSON.parse(json).length === 0) {
		alert("保存されているプリセットがありません。");
		return;
	}

	// 階層情報を取得してIDから名前へのマッピングを行う（別プロジェクトへの移行用）
	csInterface.evalScript('getHierarchy()', function (res) {
		var hierarchy = JSON.parse(res);
		if (hierarchy.status === 'error') {
			alert("コンポジション情報の取得に失敗しました。");
			return;
		}

		// ID -> Name Map
		var idToName = {};
		function traverse(node) {
			if (node.layerId !== undefined) {
				idToName["ID:" + node.layerId] = node.name;
			}
			if (node.children) node.children.forEach(traverse);
		}
		traverse(hierarchy);

		var savedPresets = JSON.parse(json);
		var exportData = [];

		savedPresets.forEach(function (p) {
			var newData = [];
			if (p.data) {
				p.data.forEach(function (item) {
					var k = item.key;
					var v = item.visible;
					if (k && k.indexOf("ID:") === 0) {
						// IDキーなら名前に変換してエクスポート
						// マップにあれば変換。なければスキップ（存在しないレイヤーはエクスポートしない方が安全）
						if (idToName[k]) {
							newData.push({ key: "NAME:" + idToName[k], visible: v });
						}
					} else if (k && k.indexOf("NAME:") === 0) {
						// 既にNAMEキーならそのまま
						newData.push({ key: k, visible: v });
					}
				});
			}
			if (newData.length > 0) {
				exportData.push({ name: p.name, data: newData });
			}
		});

		if (exportData.length === 0) {
			alert("エクスポート可能なデータが見つかりませんでした。");
			return;
		}

		var result = window.cep.fs.showSaveDialogEx(
			"プリセットのエクスポート",
			g_lastCompId + "_presets.json",
			["json"],
			"psdtool_presets.json"
		);

		if (result.data) {
			var err = window.cep.fs.writeFile(result.data, JSON.stringify(exportData, null, 2));
			if (err.err === 0) {
				alert("エクスポートしました。");
			} else {
				alert("保存に失敗しました: " + err.err);
			}
		}
	});
}

function importPresets() {
	if (!g_lastCompId) {
		alert("コンポジションが読み込まれていません。");
		return;
	}

	var result = window.cep.fs.showOpenDialogEx(
		true, false, "プリセットのインポート", null, ["json"]
	);

	if (result.data && result.data.length > 0) {
		var filePath = result.data[0];
		var readRes = window.cep.fs.readFile(filePath);

		if (readRes.err === 0) {
			try {
				var importedPresets = JSON.parse(readRes.data);
				if (!Array.isArray(importedPresets)) {
					alert("不正なファイル形式です。");
					return;
				}

				// バリデーションのために現在の階層構造を取得
				csInterface.evalScript('getHierarchy()', function (res) {
					var hierarchy = JSON.parse(res);
					if (hierarchy.status === 'error') {
						alert("コンポジション情報の取得に失敗したため、インポートを中止します。");
						return;
					}

					// 名前 -> ID Map の作成 (現在のコンポジション環境に合わせるため)
					var nameToId = {};
					function collectKeys(node) {
						if (node.layerId !== undefined) {
							// "NAME:レイヤー名" -> "ID:レイヤーID" の対応を作る
							nameToId["NAME:" + node.name] = "ID:" + node.layerId;
						}
						if (node.children) node.children.forEach(collectKeys);
					}
					if (hierarchy.children) hierarchy.children.forEach(collectKeys);

					// インポートデータの検証
					var errorMessages = [];
					var validPresets = [];

					importedPresets.forEach(function (preset) {
						if (!preset.data || !Array.isArray(preset.data)) return;

						var convertedData = [];
						var missingCount = 0;

						preset.data.forEach(function (item) {
							var k = item.key;
							if (!k) return;

							// エクスポートデータは "NAME:xxx" になっているはず
							if (k.indexOf("NAME:") === 0) {
								if (nameToId[k]) {
									// 現在のCompに対応するレイヤーIDが見つかった -> IDキーに変換して取り込む
									convertedData.push({ key: nameToId[k], visible: item.visible });
								} else {
									missingCount++;
								}
							} else {
								// もしIDキーのまま渡されたら、別Compでは使えないのでNGとする
								missingCount++;
							}
						});

						if (missingCount > 0) {
							errorMessages.push("プリセット '" + preset.name + "' に含まれる " + missingCount + " 個のレイヤーが現在のコンポジションで見つかりません（名前不一致）。");
						} else {
							if (convertedData.length > 0) {
								validPresets.push({ name: preset.name, data: convertedData });
							}
						}
					});

					if (errorMessages.length > 0) {
						alert("インポートエラー:\n" + errorMessages.join("\n") + "\n\nすべてのレイヤー名が一致する必要があります。");
						return;
					}

					if (validPresets.length === 0) {
						alert("インポートできる有効なデータがありませんでした。");
						return;
					}

					// 検証OKなら追加
					var key = getPresetKey(g_lastCompId);
					var existingJson = localStorage.getItem(key);
					var existingPresets = existingJson ? JSON.parse(existingJson) : [];

					// 重複チェックはせず追記する（名前が同じでも内容は違う可能性があるため）
					validPresets.forEach(function (p) {
						existingPresets.push(p);
					});

					localStorage.setItem(key, JSON.stringify(existingPresets));
					loadPresets(g_lastCompId);
					alert("インポートが完了しました。");

				});

			} catch (e) {
				alert("JSONパースエラー: " + e);
			}
		} else {
			alert("ファイル読み込みエラー: " + readRes.err);
		}
	}
}

// ------ Storage Management ------

function renderStorageList() {
	var listContainer = document.getElementById("storage-list-container");
	listContainer.innerHTML = "";

	var items = [];
	for (var i = 0; i < localStorage.length; i++) {
		var key = localStorage.key(i);
		if (key.indexOf("psdtool_presets") === 0) {
			var data = localStorage.getItem(key);
			var presetCount = 0;
			try {
				var json = JSON.parse(data);
				if (Array.isArray(json)) presetCount = json.length;
			} catch (e) { }

			// Parse Key
			var lastUnder = key.lastIndexOf("_");
			var compId = key.substring(lastUnder + 1);
			var prefix = key.substring(0, lastUnder);
			var path = "(未保存プロジェクト)";

			if (prefix === "psdtool_presets") {
				// No path part
			} else if (prefix.indexOf("psdtool_presets_") === 0) {
				var encoded = prefix.substring("psdtool_presets_".length);
				try {
					path = decodeURIComponent(encoded);
				} catch (e) {
					path = encoded;
				}
			}

			items.push({
				key: key,
				path: path,
				compId: compId,
				count: presetCount
			});
		}
	}

	if (items.length === 0) {
		listContainer.innerHTML = "<p style='color:#888; text-align:center; padding-top:20px;'>保存されたプリセットはありません。</p>";
		return;
	}

	// Sort by path
	items.sort(function (a, b) {
		return a.path.localeCompare(b.path) || a.compId.localeCompare(b.compId);
	});

	items.forEach(function (item) {
		var div = document.createElement("div");
		div.className = "storage-item";

		var infoDiv = document.createElement("div");
		infoDiv.className = "storage-info";

		// ファイル名を抽出して大きく表示
		var fileName = item.path;
		if (fileName.indexOf("\\") > 0) fileName = fileName.split("\\").pop();
		if (fileName.indexOf("/") > 0) fileName = fileName.split("/").pop();

		var nameSpan = document.createElement("span");
		nameSpan.className = "storage-name";
		nameSpan.textContent = fileName;
		nameSpan.title = "Project: " + fileName;

		var pathSpan = document.createElement("span");
		pathSpan.className = "storage-path";
		pathSpan.title = item.path;
		pathSpan.textContent = item.path;

		var detailsSpan = document.createElement("span");
		detailsSpan.className = "storage-comp";
		detailsSpan.textContent = "CompID: " + item.compId + " / Presets: " + item.count;

		infoDiv.appendChild(nameSpan);
		infoDiv.appendChild(pathSpan);
		infoDiv.appendChild(detailsSpan);

		var actionsDiv = document.createElement("div");
		actionsDiv.className = "storage-actions";

		var importBtn = document.createElement("button");
		importBtn.className = "storage-import";
		importBtn.textContent = "インポート";
		importBtn.onclick = function () {
			importPresetsFromData(item.key);
		};

		var delBtn = document.createElement("button");
		delBtn.className = "storage-delete";
		delBtn.textContent = "削除";
		delBtn.onclick = function () {
			if (confirm("以下のデータを削除しますか？\n\nパス: " + item.path + "\nCompID: " + item.compId)) {
				localStorage.removeItem(item.key);
				renderStorageList(); // Refresh list

				// Make sure to refresh main UI if we deleted the currently active presets
				if (g_lastCompId) {
					var currentKey = getPresetKey(g_lastCompId);
					if (currentKey === item.key) {
						loadPresets(g_lastCompId);
					}
				}
			}
		};

		actionsDiv.appendChild(importBtn);
		actionsDiv.appendChild(delBtn);

		div.appendChild(infoDiv);
		div.appendChild(actionsDiv);
		listContainer.appendChild(div);
	});
}

function importPresetsFromData(storageKey) {
	if (!g_lastCompId) {
		alert("コンポジションが読み込まれていません。");
		return;
	}

	var json = localStorage.getItem(storageKey);
	if (!json) {
		alert("データが見つかりませんでした。");
		return;
	}

	try {
		var importedPresets = JSON.parse(json);
		if (!Array.isArray(importedPresets)) {
			alert("データ形式が不正です。");
			return;
		}

		// バリデーションのために現在の階層構造を取得
		csInterface.evalScript('getHierarchy()', function (res) {
			var hierarchy = JSON.parse(res);
			if (hierarchy.status === 'error') {
				alert("コンポジション情報の取得に失敗したため、インポートを中止します。");
				return;
			}

			// 名前 -> ID Map と 有効ID Map の作成
			var nameToId = {};
			var validIds = {};
			function collectKeys(node) {
				if (node.layerId !== undefined) {
					nameToId["NAME:" + node.name] = "ID:" + node.layerId;
					validIds["ID:" + node.layerId] = true;
				}
				if (node.children) node.children.forEach(collectKeys);
			}
			if (hierarchy.children) hierarchy.children.forEach(collectKeys);

			var validPresets = [];
			var errorMessages = [];

			importedPresets.forEach(function (preset) {
				if (!preset.data || !Array.isArray(preset.data)) return;

				var convertedData = [];
				var missingCount = 0;

				preset.data.forEach(function (item) {
					var k = item.key;
					if (!k) return;

					if (k.indexOf("NAME:") === 0) {
						if (nameToId[k]) {
							convertedData.push({ key: nameToId[k], visible: item.visible });
						} else {
							missingCount++;
						}
					} else if (k.indexOf("ID:") === 0) {
						if (validIds[k]) {
							convertedData.push({ key: k, visible: item.visible });
						} else if (item.layerName && nameToId["NAME:" + item.layerName]) {
							// IDは一致しないが、バックアップされた名前で一致するレイヤーがある場合（プロジェクト移行時など）
							// 現在のプロジェクトのIDに変換して取り込む
							convertedData.push({ key: nameToId["NAME:" + item.layerName], visible: item.visible });
						} else {
							missingCount++;
						}
					}
				});

				if (missingCount > 0) {
					errorMessages.push("プリセット '" + preset.name + "' に含まれる " + missingCount + " 個のレイヤーが現在のコンポジションで見つかりません。");
				} else {
					if (convertedData.length > 0) {
						validPresets.push({ name: preset.name, data: convertedData });
					}
				}
			});

			if (errorMessages.length > 0) {
				alert("インポートエラー:\n" + errorMessages.join("\n") + "\n\nすべてのレイヤーが現在のコンポジションに存在する必要があります。");
				return;
			}

			if (validPresets.length === 0) {
				alert("インポートできる有効なデータがありませんでした。");
				return;
			}

			// 追加
			var activeKey = getPresetKey(g_lastCompId);
			var existingJson = localStorage.getItem(activeKey);
			var existingPresets = existingJson ? JSON.parse(existingJson) : [];

			validPresets.forEach(function (p) {
				existingPresets.push(p);
			});

			localStorage.setItem(activeKey, JSON.stringify(existingPresets));
			loadPresets(g_lastCompId);

			alert("インポートが完了しました。");

			// モーダルを閉じる
			document.getElementById("storage-modal").style.display = "none";
		});

	} catch (e) {
		alert("データの読み込みに失敗しました: " + e);
	}
}
