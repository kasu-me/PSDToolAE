(function (thisObj) {
	// UIの構築
	function buildUI(thisObj) {
		var panel = (thisObj instanceof Panel) ? thisObj : new Window("palette", "字幕自動反映", undefined, { resizeable: true });
		var btn = panel.add("button", undefined, "テキスト反映");

		btn.onClick = function () {
			app.beginUndoGroup("テキストを反映");
			processSelectedLayers();
			app.endUndoGroup();
		};

		panel.layout.layout(true);
		return panel;
	}

	// 実際の処理
	function processSelectedLayers() {
		var comp = app.project.activeItem;
		if (!comp || !(comp instanceof CompItem)) {
			alert("コンポジションを選択してから実行してください。");
			return;
		}

		var selectedLayers = comp.selectedLayers;
		if (selectedLayers.length === 0) {
			alert("対象の音声レイヤーを選択してください。");
			return;
		}

		for (var i = 0; i < selectedLayers.length; i++) {
			var layer = selectedLayers[i];

			// ソースファイル（実体のファイル）を持っているか確認
			if (!layer.source || !layer.source.mainSource || !layer.source.mainSource.file) {
				continue;
			}

			var audioFile = layer.source.mainSource.file;
			var fileName = decodeURI(audioFile.name); // 例: 000-AAAA-BBBB-CCCC.wav

			// ----------------------------------------------------
			// 1. テキストファイルの読み込み
			// ----------------------------------------------------
			// 拡張子を除いたファイルパスを取得して.txtに変更
			var docIndex = audioFile.fsName.lastIndexOf(".");
			var txtFilePath = audioFile.fsName.substring(0, docIndex > -1 ? docIndex : audioFile.fsName.length) + ".txt";
			var txtFile = new File(txtFilePath);

			// テキストファイルが存在しない場合はスキップ
			if (!txtFile.exists) {
				continue;
			}

			txtFile.encoding = "UTF-8"; // 文字化けを防ぐためUTF-8指定
			txtFile.open("r");
			var textContent = txtFile.read();
			txtFile.close();

			// ----------------------------------------------------
			// 2. ファイル名から「BBBB」を抽出
			// ----------------------------------------------------
			// 拡張子を除いたファイル名で分割する場合も考慮
			var fileNameWithoutExt = fileName.substring(0, fileName.lastIndexOf("."));
			var nameParts = fileNameWithoutExt.split("-");

			// フォーマット通りでない場合はスキップ
			if (nameParts.length < 3) {
				continue;
			}

			var targetPrefix = nameParts[2]; // 例: BBBB

			// ----------------------------------------------------
			// 3. 対象のテキストレイヤーを検索してキーフレームを追加
			// ----------------------------------------------------
			var targetTextLayer = null;

			for (var j = 1; j <= comp.numLayers; j++) {
				var l = comp.layer(j);
				// テキストレイヤーであり、かつ名前がBBBBから始まるか(前方一致)確認
				if (l instanceof TextLayer && l.name.indexOf(targetPrefix) === 0) {
					targetTextLayer = l;
					break;
				}
			}

			// 見つからなかった場合はスキップ
			if (!targetTextLayer) {
				continue;
			}

			// ソーステキストプロパティを取得 (言語設定に依存しないADBE名を使用)
			var sourceTextProp = targetTextLayer.property("ADBE Text Properties").property("ADBE Text Document");
			var opacityProp = targetTextLayer.property("ADBE Transform Group").property("ADBE Opacity");

			if (sourceTextProp) {
				// 現在のテキストドキュメントのスタイルを保持したままテキストだけ書き換える
				var textDoc = sourceTextProp.value;
				textDoc.text = textContent;

				// インポイントの時間をフレームぴったりの位置に合わせる（スナップ）
				var snappedTime = Math.round(layer.inPoint / comp.frameDuration) * comp.frameDuration;

				// 音声レイヤーのインポイント(開始時間)にキーフレームを追加
				sourceTextProp.setValueAtTime(snappedTime, textDoc);

				// コンポジション内のすべてのテキストレイヤーの不透明度を設定
				for (var k = 1; k <= comp.numLayers; k++) {
					var l = comp.layer(k);
					if (l instanceof TextLayer) {
						var otherOpacityProp = l.property("ADBE Transform Group").property("ADBE Opacity");
						if (otherOpacityProp) {
							if (l === targetTextLayer) {
								// 対象のテキストレイヤーは不透明度を100%に
								otherOpacityProp.setValueAtTime(snappedTime, 100);
							} else {
								// 対象以外のテキストレイヤーは不透明度を0%に
								otherOpacityProp.setValueAtTime(snappedTime, 0);
							}
						}
					}
				}
			}
		}
	}

	var myPanel = buildUI(thisObj);
	if (myPanel instanceof Window) {
		myPanel.center();
		myPanel.show();
	}
})(this);
