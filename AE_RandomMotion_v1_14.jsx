// RandomMotion - After Effects ScriptUI Panel
// AE 2024対応（ヘルプ機能：ツールチップ＆ステータスバー付き）

(function(thisObj) {
    "use strict";
    
    // ========================================
    // グローバル設定
    // ========================================
    var SCRIPT_NAME = "RandomMotion";
    var VERSION = "1.0.0";
    var PRESET_FILE = "RandomMotion_Presets.json"
    var UIREF = null; // UIコントロール参照を保持

    
    // デフォルト値
    // ★DEFAULT_VALUES を以下に置き換え
    var DEFAULT_VALUES = {
        xMin: 0, xMax: 0,
        yMin: 0, yMax: 0,
        zMin: 0, zMax: 0,
        rMin: 0, rMax: 0,
        sMin: 0, sMax: 0,
        tMin: 0, tMax: 0,
        frameOffset: 20,

        // ▼ 旧: signMode は互換のため残さず、個別符号モードに移行
        // signMode: 0,
        // ▼ 新規：プロパティ別の符号モード（0..4 = +/- , + , - , +-+- , -+-+）
        signXMode: 0,
        signYMode: 0,
        signRMode: 0,
        signSMode: 0,
        signTMode: 0, // UIは未表示（内部用、必要なら今後UI化可能）

        posMode: 0,   // 0:拡散, 1:座標, 2:X/Y/Z, 3:↑→↓←(2D), 4:X→Y→Z(3D)
        scaleXY: 0,   // 0:X&Y, 1:X, 2:Y, 3:X/Y(Random), 4:XYXY, 5:YXYX
        autoStartIn: false,
        autoStartOut: false
    };

    var settings = {};
    for (var key in DEFAULT_VALUES) {
        settings[key] = DEFAULT_VALUES[key];
    }

    var presets = [];

    function resetToDefaults() {
        // DEFAULT_VALUES を深いコピーで settings に反映
        function deepClone(v) {
            if (v && typeof v === "object") {
                if (v instanceof Array) {
                    var a = [];
                    for (var i = 0; i < v.length; i++) a[i] = deepClone(v[i]);
                    return a;
                } else {
                    var o = {};
                    for (var k in v) if (v.hasOwnProperty(k)) o[k] = deepClone(v[k]);
                    return o;
                }
            }
            return v;
        }
        for (var key in DEFAULT_VALUES) {
            if (DEFAULT_VALUES.hasOwnProperty(key)) {
                settings[key] = deepClone(DEFAULT_VALUES[key]);
            }
        }
    }
  

  
    
    // --- JSONセーフラッパー（ExtendScriptの古環境対策） ---
    function _jsonStringify(obj) {
        try {
            if (typeof JSON !== "undefined" && JSON && JSON.stringify) {
                return JSON.stringify(obj, null, 2);
            }
        } catch (e) {}
        function esc(s){return String(s).replace(/\\/g,"\\\\").replace(/"/g,'\\"').replace(/\r?\n/g,"\\n");}
        function ser(x){
            if (x === null) return "null";
            var t = typeof x;
            if (t === "string")  return '"' + esc(x) + '"';
            if (t === "number")  return isFinite(x) ? String(x) : "null";
            if (t === "boolean") return x ? "true" : "false";
            if (x instanceof Array) { var a=[]; for (var i=0;i<x.length;i++) a.push(ser(x[i])); return "["+a.join(",")+"]"; }
            if (t === "object")  { var kv=[]; for (var k in x) if (x.hasOwnProperty(k)) {
                                    var v=x[k]; if (typeof v!=="function" && typeof v!=="undefined") kv.push('"'+esc(k)+'":'+ser(v));
                                } return "{"+kv.join(",")+"}"; }
            return "null";
        }
        return ser(obj);
    }
    function _jsonParse(text) {
        try { if (typeof JSON!=="undefined" && JSON && JSON.parse) return JSON.parse(text); } catch(e){}
        try { return eval("(" + text + ")"); } catch(e2){ return []; }
    }

    function _jsonParse(text) {
        try {
            if (typeof JSON !== "undefined" && JSON && JSON.parse) return JSON.parse(text);
        } catch (e) {}
        // 自分で書き出した JSON しか読まない前提で eval 代替（外部入力は使用しない想定）
        try { return eval("(" + text + ")"); } catch (e2) { return []; }
    }

    // ========================================
    // プリセット管理
    // ========================================
    // Folder.userData は Folderオブジェクトなので文字列化して使う
    function getPresetFilePath() {
        // 1) 実行中スクリプトと同じフォルダを最優先
        try {
            var scriptFile = new File($.fileName); // 実行中の .jsx / .jsxbin
            if (scriptFile && scriptFile.parent && scriptFile.parent.exists) {
                var scriptDir = scriptFile.parent.fullName;
                return scriptDir + "/" + PRESET_FILE; // 例: AE_RandomMotion_v1_04.jsx と同じ場所
            }
        } catch (e) {
            // 無視してフォールバックへ
        }

        // 2) フォールバック: 従来の userData パス
        var basePath = (Folder.userData && Folder.userData.fsName)
            ? Folder.userData.fsName
            : Folder.userData.fullName;

        var targetPath = basePath + "/Adobe/After Effects/RandomMotion";

        // 必要な中間ディレクトリも作成
        var folder = new Folder(targetPath);
        if (!folder.exists) {
            var parts = targetPath.split("/");
            var acc = parts[0];
            for (var i = 1; i < parts.length; i++) {
                acc += "/" + parts[i];
                var f = new Folder(acc);
                if (!f.exists) f.create();
            }
        }
        return folder.fullName + "/" + PRESET_FILE;
    }

    function loadPresets() {
        // 候補: スクリプトと同じ場所 → userData の順で探す
        var candidates = [];

        // スクリプト隣
        try {
            var scriptFile = new File($.fileName);
            if (scriptFile && scriptFile.parent && scriptFile.parent.exists) {
                candidates.push(scriptFile.parent.fullName + "/" + PRESET_FILE);
            }
        } catch (e) {}

        // userData 側
        var basePath = (Folder.userData && Folder.userData.fsName)
            ? Folder.userData.fsName
            : Folder.userData.fullName;
        var targetPath = basePath + "/Adobe/After Effects/RandomMotion";
        candidates.push(targetPath + "/" + PRESET_FILE);

        for (var i = 0; i < candidates.length; i++) {
            var f = new File(candidates[i]);
            if (f.exists) {
                if (f.open("r")) {
                    var content = f.read();
                    f.close();
                    presets = _jsonParse(content) || [];
                    return;
                }
            }
        }
        presets = []; // 見つからない場合
    }

    
    function savePresets() {
        // 先に文字列化（ここで失敗したらファイルを触らない）
        var text = _jsonStringify(presets);
        if (typeof text !== "string" || text === "") text = "[]";

        // 優先：スクリプトと同じ場所
        var primaryPath = getPresetFilePath(); // 既存（スクリプト横→無理なら userData を返すでもOK）
        var tried = [];
        if (primaryPath) {
            tried.push(primaryPath);
            try {
                var pf = new File(primaryPath);
                var pdir = pf.parent; if (!pdir.exists) pdir.create();
                if (pf.open("w")) {
                    pf.encoding = "UTF-8";
                    pf.lineFeed = "Unix";
                    pf.write(text);
                    pf.close();
                    return;
                }
            } catch (e) {}
        }

        // フォールバック：userData 側
        var basePath = (Folder.userData && Folder.userData.fsName) ? Folder.userData.fsName : Folder.userData.fullName;
        var targetDir = basePath + "/Adobe/After Effects/RandomMotion";
        var folder = new Folder(targetDir);
        if (!folder.exists) {
            var parts = targetDir.split("/"); var acc = parts[0];
            for (var i=1;i<parts.length;i++){ acc+="/"+parts[i]; var d=new Folder(acc); if(!d.exists) d.create(); }
        }
        var fallbackPath = folder.fullName + "/" + PRESET_FILE;
        tried.push(fallbackPath);
        try {
            var ff = new File(fallbackPath);
            if (ff.open("w")) {
                ff.encoding = "UTF-8";
                ff.lineFeed = "Unix";
                ff.write(text);
                ff.close();
                return;
            }
        } catch (e2) {}

        alert("プリセットを保存できませんでした。\n試行した場所:\n- " + tried.join("\n- "));
    }


    
    function addPreset(name, values) {
        // 名前の「,」を空白に置換
        name = name.replace(/,/g, " ");
        if (name === "") {
            name = "Preset_" + (presets.length + 1);
        }
        presets.push({name: name, values: values});
        savePresets();
    }
    
    function removePreset(index) {
        presets.splice(index, 1);
        savePresets();
    }
    
    function movePreset(index, direction) {
        if (direction === -1 && index > 0) {
            var temp = presets[index];
            presets[index] = presets[index - 1];
            presets[index - 1] = temp;
            savePresets();
            return index - 1;
        } else if (direction === 1 && index < presets.length - 1) {
            var temp = presets[index];
            presets[index] = presets[index + 1];
            presets[index + 1] = temp;
            savePresets();
            return index + 1;
        }
        return index;
    }
    
    function exportPresets() {
        var file = File.saveDialog("プリセットをエクスポート", "JSON:*.json");
        if (!file) return;

        // 先に安全に文字列化（ここで失敗したらファイルを触らない）
        var text = _jsonStringify(presets);
        if (typeof text !== "string" || text === "") text = "[]";

        // 拡張子 .json を付与
        if (!/\.json$/i.test(file.name)) file = new File(file.fsName + ".json");

        // 書き込み
        if (file.open("w")) {
            file.encoding = "UTF-8";
            file.lineFeed = "Unix";
            file.write(text);
            file.close();
            alert("プリセットをエクスポートしました");
        } else {
            alert("ファイルを開けませんでした：\n" + file.fsName);
        }
    }

    function importPresets() {
        var file = File.openDialog("プリセットをインポート", "JSON:*.json");
        if (!file) return;

        if (file.open("r")) {
            var content = file.read();
            file.close();

            var arr = _jsonParse(content);
            if (!(arr instanceof Array)) {
                alert("プリセットファイルの形式が不正です。");
                return;
            }
            presets = arr;      // 置換（必要ならマージに変更可）
            savePresets();      // 内部保存も更新
            alert("プリセットをインポートしました（既存のプリセットは置換されました）");
        } else {
            alert("ファイルを開けませんでした：\n" + file.fsName);
        }
    }

    // ========================================
    // ユーティリティ関数
    // ========================================
    function clamp(val, min, max) {
        return Math.max(min, Math.min(max, val));
    }
    
    function randomRange(min, max) {
        return min + Math.random() * (max - min);
    }
    
    function validateMinMax() {
        // min と max が逆になっていたら自動で入れ替える（数値化もここで保証）
        function normalizePair(minKey, maxKey) {
            var minVal = Number(settings[minKey]);
            var maxVal = Number(settings[maxKey]);

            // 数値でない場合の保険
            if (!isFinite(minVal)) minVal = 0;
            if (!isFinite(maxVal)) maxVal = 0;

            // 逆なら入れ替え
            if (minVal > maxVal) {
                var tmp = minVal;
                minVal = maxVal;
                maxVal = tmp;
            }

            // settings に戻す（数値として揃える）
            settings[minKey] = minVal;
            settings[maxKey] = maxVal;
        }

        normalizePair("xMin", "xMax");
        normalizePair("yMin", "yMax");

        // ▼追加：3D用 Z
        normalizePair("zMin", "zMax");

        normalizePair("rMin", "rMax");
        normalizePair("sMin", "sMax");
        normalizePair("tMin", "tMax");

        // UIがあれば表示も更新
        if (typeof syncSettingsToUI === "function" && UIREF) {
            syncSettingsToUI();
        }

        // ここでは常に true（エラー扱いにしない）
        return true;
    }


    
    function allValuesZero() {
        return settings.xMin === 0 && settings.xMax === 0 &&
               settings.yMin === 0 && settings.yMax === 0 &&
               settings.zMin === 0 && settings.zMax === 0 &&
               settings.rMin === 0 && settings.rMax === 0 &&
               settings.sMin === 0 && settings.sMax === 0 &&
               settings.tMin === 0 && settings.tMax === 0;
    }
    
    // ★getSign を以下に置き換え（modeを受け取る形）
    function getSign(index, mode) {
        switch (mode|0) {
            case 0: return Math.random() < 0.5 ? 1 : -1;         // +/-（ランダム）
            case 1: return 1;                                     // +
            case 2: return -1;                                    // -
            case 3: return (index % 2 === 0) ? 1 : -1;            // +-+-
            case 4: return (index % 2 === 0) ? -1 : 1;            // -+-+
            default: return 1;
        }
    }

    
    // ========================================
    // キーフレーム生成
    // ========================================
    function applyRandomMotion() {
        var comp = app.project.activeItem;
        if (!comp || !(comp instanceof CompItem)) {
            alert("コンポジションを開いてください");
            return;
        }
        
        var layers = comp.selectedLayers;
        if (layers.length === 0) {
            alert("レイヤーを選択してください");
            return;
        }
        
        if (!validateMinMax()) {
            alert("エラー: 最小値が最大値より大きい項目があります");
            return;
        }
        
        if (allValuesZero()) {
            alert("すべての値が0です。キーフレームは生成されません");
            return;
        }
        
        app.beginUndoGroup(SCRIPT_NAME + " - Apply");
        
        try {
            for (var i = 0; i < layers.length; i++) {
                applyToLayer(layers[i], comp, i);
            }
        } catch(e) {
            alert("エラー: " + e.toString());
        }
        
        app.endUndoGroup();
    }
        
    function applyToLayer(layer, comp, layerIndex) {
        var currentTime   = comp.time;
        var frameOffset   = settings.frameOffset;
        var frameDuration = comp.frameDuration;

        // 始点時間の決定
        var startTime;
        if (settings.autoStartIn && settings.autoStartOut) {
            // 両方ON: イン点とアウト点に2キーを固定（既存仕様のまま）
            applyTwoKeys(layer, layer.inPoint, layer.outPoint, layerIndex);
            return;
        } else if (settings.autoStartIn) {
            startTime = layer.inPoint;
        } else if (settings.autoStartOut) {
            startTime = layer.outPoint;
        } else {
            startTime = currentTime;
        }

        // オフセット時間（フレーム → 秒）
        var offsetTime = frameOffset * frameDuration;

        // 修正ポイント：
        // 正負にかかわらず、
        //   time1 = startTime + offsetTime  (← ランダム適用後の値を打つ時刻)
        //   time2 = startTime               (← 元の値を保持する時刻)
        //
        // これにより frameOffset = -10 の場合は
        //   time1 = 現在 - 10f
        //   time2 = 現在
        // となり、「現在から10フレーム前に“ランダム適用したキー”が打たれる」挙動になります。
        var time1 = startTime + offsetTime; // ランダム適用後
        var time2 = startTime;              // 元の値

        // 各プロパティに適用（既存関数は time1=適用／time2=元 を前提に実装済み）
        applyPosition(layer, time1, time2, layerIndex);
        applyRotation(layer, time1, time2, layerIndex);
        applyScale(layer, time1, time2, layerIndex);
        applyOpacity(layer, time1, time2, layerIndex);
    }

    
    function applyTwoKeys(layer, time1, time2, layerIndex) {
        // イン点とアウト点に2キーを配置（自動始点両方ON時）
        applyPosition(layer, time1, time2, layerIndex);
        applyRotation(layer, time1, time2, layerIndex);
        applyScale(layer, time1, time2, layerIndex);
        applyOpacity(layer, time1, time2, layerIndex);
    }
        
function applyPosition(layer, time1, time2, layerIndex) {
    // X/Y/Z が全部 0 なら何もしない
    if (settings.xMin === 0 && settings.xMax === 0 &&
        settings.yMin === 0 && settings.yMax === 0 &&
        settings.zMin === 0 && settings.zMax === 0) return;

    var tr = layer.property("ADBE Transform Group");
    var pos = tr.property("ADBE Position");

    // 3D判定（レイヤーが3Dの時だけZを動かす）
    var is3DLayer = !!layer.threeDLayer;

    // 符号モード
    var signX = getSign(layerIndex, settings.signXMode);
    var signY = getSign(layerIndex, settings.signYMode);
    var signZ = getSign(layerIndex, settings.signZMode);

    // Separate Dimensions
    var isSeparate = pos.dimensionsSeparated;

    if (isSeparate) {
        var xProp = tr.property("ADBE Position_0");
        var yProp = tr.property("ADBE Position_1");

        // 3Dなら Z も（無ければ null）
        var zProp = null;
        if (is3DLayer) zProp = tr.property("ADBE Position_2");

        var originalX = xProp.valueAtTime(time2, false);
        var originalY = yProp.valueAtTime(time2, false);
        var originalZ = (zProp) ? zProp.valueAtTime(time2, false) : 0;

        var deltaX = 0, deltaY = 0, deltaZ = 0;
        calculatePositionDelta(layerIndex, is3DLayer, function (dx, dy, dz) {
            deltaX = (dx || 0) * signX;
            deltaY = (dy || 0) * signY;
            deltaZ = (dz || 0) * signZ;
        });

        xProp.setValueAtTime(time1, originalX + deltaX);
        xProp.setValueAtTime(time2, originalX);

        yProp.setValueAtTime(time1, originalY + deltaY);
        yProp.setValueAtTime(time2, originalY);

        if (zProp) {
            zProp.setValueAtTime(time1, originalZ + deltaZ);
            zProp.setValueAtTime(time2, originalZ);
        }

    } else {
        var original = pos.valueAtTime(time2, false);
        var hasZ = is3DLayer && (original instanceof Array) && (original.length >= 3);

        var deltaX2 = 0, deltaY2 = 0, deltaZ2 = 0;
        calculatePositionDelta(layerIndex, is3DLayer, function (dx, dy, dz) {
            deltaX2 = (dx || 0) * signX;
            deltaY2 = (dy || 0) * signY;
            deltaZ2 = (dz || 0) * signZ;
        });

        var newVal;
        if (hasZ) {
            newVal = [original[0] + deltaX2, original[1] + deltaY2, original[2] + deltaZ2];
        } else {
            // 2DはZ無視
            newVal = [original[0] + deltaX2, original[1] + deltaY2];
        }

        pos.setValueAtTime(time1, newVal);
        pos.setValueAtTime(time2, original);
    }
}

    // ★applyRotation を以下に置き換え
    function applyRotation(layer, time1, time2, layerIndex) {
        if (settings.rMin === 0 && settings.rMax === 0) return;

        var rot = layer.property("ADBE Transform Group").property("ADBE Rotate Z");
        var original = rot.valueAtTime(time2, false);
        var sign = getSign(layerIndex, settings.signRMode);
        var delta = randomRange(settings.rMin, settings.rMax) * sign;

        rot.setValueAtTime(time1, original + delta);
        rot.setValueAtTime(time2, original);
    }

    // ★applyScale を以下に置き換え（S-Modeは scaleXY。符号は signSMode）
    function applyScale(layer, time1, time2, layerIndex) {
        if (settings.sMin === 0 && settings.sMax === 0) return;

        var scale = layer.property("ADBE Transform Group").property("ADBE Scale");
        var original = scale.valueAtTime(time2, false);

        var hasZ = (original instanceof Array) && (original.length >= 3);

        var sign = getSign(layerIndex, settings.signSMode);
        var delta = randomRange(settings.sMin, settings.sMax) * sign;

        var newX = original[0];
        var newY = original[1];
        var newZ = hasZ ? original[2] : 0;

        switch (settings.scaleXY | 0) {
            case 0: // X&Y
                newX = original[0] + delta;
                newY = original[1] + delta;
                break;
            case 1: // X
                newX = original[0] + delta;
                break;
            case 2: // Y
                newY = original[1] + delta;
                break;
            case 3: // X/Y(Random)
                if (Math.random() < 0.5) newX = original[0] + delta;
                else                     newY = original[1] + delta;
                break;
            case 4: // XYXY
                if (layerIndex % 2 === 0) newX = original[0] + delta;
                else                      newY = original[1] + delta;
                break;
            case 5: // YXYX
                if (layerIndex % 2 === 0) newY = original[1] + delta;
                else                      newX = original[0] + delta;
                break;

            // ▼追加（末尾追加で互換維持）
            case 6: // Z
                if (hasZ) newZ = original[2] + delta;
                break;

            case 7: // X/Y/Z(Random) ※2DならX/Yのみ
                if (hasZ) {
                    var pick = Math.floor(Math.random() * 3);
                    if (pick === 0) newX = original[0] + delta;
                    else if (pick === 1) newY = original[1] + delta;
                    else newZ = original[2] + delta;
                } else {
                    if (Math.random() < 0.5) newX = original[0] + delta;
                    else                     newY = original[1] + delta;
                }
                break;
        }

        // 0跨ぎ防止
        if ((original[0] > 0 && newX < 0) || (original[0] < 0 && newX > 0)) newX = 0;
        if ((original[1] > 0 && newY < 0) || (original[1] < 0 && newY > 0)) newY = 0;
        if (hasZ) {
            if ((original[2] > 0 && newZ < 0) || (original[2] < 0 && newZ > 0)) newZ = 0;
        }

        if (hasZ) {
            scale.setValueAtTime(time1, [newX, newY, newZ]);
        } else {
            scale.setValueAtTime(time1, [newX, newY]);
        }
        scale.setValueAtTime(time2, original);
    }

    // ★applyOpacity を以下に置き換え（内部 signTMode を使用）
    function applyOpacity(layer, time1, time2, layerIndex) {
        if (settings.tMin === 0 && settings.tMax === 0) return;

        var opacity = layer.property("ADBE Transform Group").property("ADBE Opacity");
        var original = opacity.valueAtTime(time2, false);
        var sign = getSign(layerIndex, settings.signTMode);
        var delta = randomRange(settings.tMin, settings.tMax) * sign;

        var newValue = clamp(original - delta, 0, 100);
        opacity.setValueAtTime(time1, newValue);
        opacity.setValueAtTime(time2, original);
    }

        
    // 方向（符号）はここでは決めない版
    function calculatePositionDelta(layerIndex, is3DLayer, callback) {
        // 入力レンジを絶対値レンジに正規化（負の設定でも大きさとして解釈）
        var axMin = Math.min(Math.abs(settings.xMin || 0), Math.abs(settings.xMax || 0));
        var axMax = Math.max(Math.abs(settings.xMin || 0), Math.abs(settings.xMax || 0));
        var ayMin = Math.min(Math.abs(settings.yMin || 0), Math.abs(settings.yMax || 0));
        var ayMax = Math.max(Math.abs(settings.yMin || 0), Math.abs(settings.yMax || 0));

        // ▼追加：Z
        var azMin = Math.min(Math.abs(settings.zMin || 0), Math.abs(settings.zMax || 0));
        var azMax = Math.max(Math.abs(settings.zMin || 0), Math.abs(settings.zMax || 0));

        // 3DかつZレンジが有効なときだけ使う
        var canZ = !!is3DLayer && ((azMin !== 0) || (azMax !== 0));

        function randAbs(minV, maxV) {
            var v = randomRange(minV, maxV);
            return Math.abs(v);
        }

        var dx = 0, dy = 0, dz = 0;

        switch ((settings.posMode | 0)) {

            case 2: // X/Y/Z（どれか一方）※2DならX/Yのみ
                if (canZ) {
                    var pick3 = Math.floor(Math.random() * 3);
                    if (pick3 === 0) dx = randAbs(axMin, axMax);
                    else if (pick3 === 1) dy = randAbs(ayMin, ayMax);
                    else dz = randAbs(azMin, azMax);
                } else {
                    if (Math.random() < 0.5) dx = randAbs(axMin, axMax);
                    else                     dy = randAbs(ayMin, ayMax);
                }
                break;

            case 3: // ↑→↓←（2D想定：XかYのみ。Zは動かさない）
                if ((layerIndex % 2) === 0) dx = randAbs(axMin, axMax);
                else                        dy = randAbs(ayMin, ayMax);
                break;

            case 4: // ▼追加：X→Y→Z（3Dの分布用。2DならX→Y）
                if (canZ) {
                    var pickSeq = (layerIndex % 3);
                    if (pickSeq === 0) dx = randAbs(axMin, axMax);
                    else if (pickSeq === 1) dy = randAbs(ayMin, ayMax);
                    else dz = randAbs(azMin, azMax);
                } else {
                    if ((layerIndex % 2) === 0) dx = randAbs(axMin, axMax);
                    else                        dy = randAbs(ayMin, ayMax);
                }
                break;

            // 0:拡散, 1:座標（ここは従来どおり「複数軸」扱い。3DならZも）
            default:
                dx = randAbs(axMin, axMax);
                dy = randAbs(ayMin, ayMax);
                if (canZ) dz = randAbs(azMin, azMax);
                break;
        }

        if (typeof callback === "function") callback(dx, dy, dz);
    }

    // ========================================
    // ヘルプ（ツールチップ＆ステータスバー）
    // ========================================
    function attachHelpTips(ui) {
        // ボタン系
        ui.btnApply.helpTip    = "設定に基づいてキーフレームを生成します。\n全プロパティが0なら何もしません。";
        ui.btnGotoIn.helpTip   = "選択レイヤーのイン点へ移動（I相当）。\n右クリックで“始点=イン点固定”トグル。";
        ui.btnGotoOut.helpTip  = "選択レイヤーのアウト点へ移動（O相当）。\n右クリックで“始点=アウト点固定”トグル。";
        ui.btnPreset.helpTip   = "プリセットウィンドウを開きます（パネル右クリックでもカーソル位置に開く）。";
        ui.btnSignFlip.helpTip = "フレームオフセットの±を一括反転します。";

        // フレーム
        ui.frameSlider.helpTip = "現在時間からのずらし量（フレーム）。\n正:「動いた後→元の値」/ 負:「元→動いた後」。";
        ui.frameText.helpTip   = "フレームの整数入力。±どちらも可。";

        // 値設定
        ui.xMin.helpTip = "X移動の最小値。";  ui.xMax.helpTip = "X移動の最大値。";
        ui.yMin.helpTip = "Y移動の最小値。";  ui.yMax.helpTip = "Y移動の最大値。";

        // ▼追加
        if (ui.zMin) ui.zMin.helpTip = "Z移動の最小値（3Dレイヤーのみ有効）。";
        if (ui.zMax) ui.zMax.helpTip = "Z移動の最大値（3Dレイヤーのみ有効）。";

        ui.rMin.helpTip = "回転の最小値（度）。";  ui.rMax.helpTip = "回転の最大値（度）。";
        ui.sMin.helpTip = "スケールの最小値。正負を跨ぐ場合は0で停止。";  ui.sMax.helpTip = "スケールの最大値。";
        ui.tMin.helpTip = "不透明度の最小差分（現在値から差し引く）。";     ui.tMax.helpTip = "不透明度の最大差分（現在値から差し引く）。";

        // スケールXY指定（▼Z系追記）
        ui.scaleMode.helpTip = "スケール適用の軸・順序：X&Y / X / Y / XまたはY / XYXY / YXYX / Z(3D) / XまたはYまたはZ(3D)";

        // 位置モード（▼追加モード込み）
        var posHelps = [
            "拡散：X/Y（3DならZも）を同時に動かします。",
            "座標：X/Y（3DならZも）を同時に動かします（現実装は拡散と同等）。",
            "X/Y/Z：1軸だけ動かします（2DはX/Y）。",
            "↑→↓←：2D想定でXかYのみ（Zは動かしません）。",
            "X→Y→Z：レイヤー順で軸を分配（2DはX→Y）。"
        ];
        if (ui.rbPos && ui.rbPos.length) {
            for (var j = 0; j < ui.rbPos.length && j < posHelps.length; j++) {
                ui.rbPos[j].helpTip = posHelps[j];
            }
        }
    }


    // ステータスバーに hover 説明を反映
    function bindHoverHelp(ctrl, text, helpBar) {
        if (!ctrl) return;
        var firstLine = (text && text.split) ? text.split("\n")[0] : ""; // ★未定義でも安全
        try {
            if (ctrl.addEventListener) {
                ctrl.addEventListener("mouseover", function(){ helpBar.text = "ヒント：" + firstLine; });
                ctrl.addEventListener("mouseout",  function(){ helpBar.text = "ヒント："; });
            }
        } catch(e) {}
    }
    
    // ========================================
    // UI構築
    // ========================================
    // ★buildUI を丸ごと置き換え（符号DDのラベルをUnicodeマイナスに修正）
    function buildUI(thisObj) {
        var win = (thisObj instanceof Panel) ? thisObj : new Window("palette", SCRIPT_NAME, undefined, {resizeable: true});

        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.spacing = 5;
        win.margins = 10;

        // --- パネル: モーション設定 ---
        var motionGroup = win.add("panel", undefined, "モーション設定");
        motionGroup.orientation = "column";
        motionGroup.alignChildren = ["fill", "top"];
        motionGroup.spacing = 5;
        motionGroup.margins = 10;

        // 値設定：行1/行2（▼Z追加のため 3+3 配置に）
        var row1 = motionGroup.add("group"); row1.alignment = ["fill", "top"]; row1.spacing = 20;
        var row2 = motionGroup.add("group"); row2.alignment = ["fill", "top"]; row2.spacing = 20;

        // X / Y / Z
        var s1 = createValueControl(row1, "X", "xMin", "xMax", -2000, 2000);
        var s2 = createValueControl(row1, "Y", "yMin", "yMax", -2000, 2000);
        var s6 = createValueControl(row1, "Z", "zMin", "zMax", -2000, 2000);

        // R / S / T
        var s3 = createValueControl(row2, "R", "rMin", "rMax", -180, 180);
        var s4 = createValueControl(row2, "S", "sMin", "sMax", -400, 400);
        var s5 = createValueControl(row2, "T", "tMin", "tMax", 0, 100);

        // ▼ 符号ドロップダウン（ラベルをUnicodeマイナスに）
        var MINUS = "\u2212"; // "−"
        var SIGN_LABELS = ["+/-", "+", MINUS, "+\u2212+\u2212", "\u2212+\u2212+"]; // ["+/-", "+", "−", "+−+−", "−+−+"]

        function attachSignDropdown(valueGroup, labelText, settingsKey) {
            var g = valueGroup.add("group");
            g.alignment = ["left","top"];
            g.spacing = 5;

            g.add("statictext", undefined, "符号:");
            var dd = g.add("dropdownlist", undefined, SIGN_LABELS);
            dd.selection = (settings[settingsKey] | 0) || 0;
            dd.onChange = function(){ settings[settingsKey] = dd.selection.index; };
            dd.helpTip = labelText + " の符号モード（+/-, +, " + MINUS + ", +\u2212+\u2212, \u2212+\u2212+）";
            return dd;
        }

        var ddSignX = attachSignDropdown(s1, "X", "signXMode");
        var ddSignY = attachSignDropdown(s2, "Y", "signYMode");

        // ▼追加：Z符号
        var ddSignZ = attachSignDropdown(s6, "Z", "signZMode");

        var ddSignR = attachSignDropdown(s3, "R", "signRMode");
        var ddSignS = attachSignDropdown(s4, "S", "signSMode");

        // ▼ S-Mode（末尾にZ系追加して互換維持）
        var sModeRow = s4.add("group");
        sModeRow.alignment = ["left", "top"];
        sModeRow.add("statictext", undefined, "S-Mode");
        sModeRow.add("statictext", undefined, ":");

        var SCALE_MODE_ITEMS = ["X&Y", "X", "Y", "X/Y", "XYXY", "YXYX", "Z", "X/Y/Z"];
        var scaleXYDD = sModeRow.add("dropdownlist", undefined, SCALE_MODE_ITEMS);

        // selection は items から安全に
        if (scaleXYDD.items && scaleXYDD.items.length) {
            var idx = (settings.scaleXY | 0);
            if (idx < 0 || idx >= scaleXYDD.items.length) idx = 0;
            scaleXYDD.selection = scaleXYDD.items[idx];
        }
        scaleXYDD.onChange = function(){ settings.scaleXY = scaleXYDD.selection.index; };
        scaleXYDD.helpTip = "スケール適用の軸・順序：X&Y / X / Y / XまたはY / XYXY / YXYX / Z(3D) / XまたはYまたはZ(3D)";

        // フレーム指定
        var frameGroup = motionGroup.add("group");
        frameGroup.alignment = ["fill", "top"];
        frameGroup.add("statictext", undefined, "フレーム:");
        var frameSlider = frameGroup.add("slider", undefined, settings.frameOffset || 10, -100, 100);
        frameSlider.preferredSize.width = 150;
        var frameText = frameGroup.add("edittext", undefined, String(settings.frameOffset || 10));
        frameText.characters = 5;
        var btnSignFlip = frameGroup.add("button", undefined, "+-");
        btnSignFlip.preferredSize.width = 30;

        frameSlider.onChanging = function(){
            settings.frameOffset = Math.round(frameSlider.value);
            frameText.text = settings.frameOffset;
        };
        frameText.onChange = function(){
            settings.frameOffset = parseInt(frameText.text, 10) || 0;
            frameSlider.value = settings.frameOffset;
        };
        btnSignFlip.onClick = function () {
            var v = parseInt(frameText.text, 10);
            if (isNaN(v)) v = settings.frameOffset || 0;
            settings.frameOffset = -v;
            frameText.text = String(settings.frameOffset);
            frameSlider.value = settings.frameOffset;
        };

        // ▼ 位置モードラジオ（分布：3D用を追加）
        var posGroup = motionGroup.add("group");
        posGroup.add("statictext", undefined, "分布:");
        var posRadios = [];

        // 既存 0..3 を維持し、4 を追加
        var posLabels = ["拡散", "座標", "X/Y/Z", "↑→↓←", "X→Y→Z"];
        for (var j=0; j<posLabels.length; j++){
            var rb2 = posGroup.add("radiobutton", undefined, posLabels[j]);
            posRadios.push(rb2);
            rb2.value = (j === (settings.posMode|0));
            (function(idx){ rb2.onClick = function(){ settings.posMode = idx; }; })(j);
        }

        // --- ボタン行 ---
        var btnGroup = motionGroup.add("group");
        btnGroup.alignment = ["fill", "top"];

        var btnGotoIn  = btnGroup.add("button", undefined, "<");  btnGotoIn.preferredSize.width = 40;
        var btnGotoOut = btnGroup.add("button", undefined, ">");  btnGotoOut.preferredSize.width = 40;

        btnGroup.add("statictext", undefined, ""); // spacer

        var btnPreset  = btnGroup.add("button", undefined, "プリセット");  btnPreset.preferredSize.width = 70;

        var btnReset   = btnGroup.add("button", undefined, "Reset");
        btnReset.helpTip = "既定値に戻す（DEFAULT_VALUES を反映）";

        var btnApply   = btnGroup.add("button", undefined, "実行");

        // ▼イベント結線
        btnApply.onClick = function(){ applyRandomMotion(); };
        btnGotoIn.onClick  = function(){ gotoInPoint();  };
        btnGotoOut.onClick = function(){ gotoOutPoint(); };

        // < / > 右クリックで自動始点トグル
        if (btnGotoIn.addEventListener){
            btnGotoIn.addEventListener("mousedown", function(e){
                if (e.button === 2){ settings.autoStartIn = !settings.autoStartIn; if (typeof syncSettingsToUI==="function") syncSettingsToUI(); }
            });
        }
        if (btnGotoOut.addEventListener){
            btnGotoOut.addEventListener("mousedown", function(e){
                if (e.button === 2){ settings.autoStartOut = !settings.autoStartOut; if (typeof syncSettingsToUI==="function") syncSettingsToUI(); }
            });
        }

        // プリセット
        btnPreset.onClick = function(){ showPresetWindow(false); };
        if (btnPreset.addEventListener){
            btnPreset.addEventListener("mousedown", function(e){
                if (e.button === 2) showPresetWindow(true);
            });
        }

        // Reset
        btnReset.onClick = function(){
            resetToDefaults();
            if (typeof syncSettingsToUI === "function") syncSettingsToUI();
        };

        // ヘルプバー
        var helpBar = win.add("statictext", undefined, "Ready.");
        helpBar.alignment = ["fill", "bottom"];

        // UI参照束ね
        var ui = {
            xMin: s1.children[1].children[1], xMax: s1.children[2].children[1],
            yMin: s2.children[1].children[1], yMax: s2.children[2].children[1],

            // ▼追加
            zMin: s6.children[1].children[1], zMax: s6.children[2].children[1],

            rMin: s3.children[1].children[1], rMax: s3.children[2].children[1],
            sMin: s4.children[1].children[1], sMax: s4.children[2].children[1],
            tMin: s5.children[1].children[1], tMax: s5.children[2].children[1],

            xMinSlider: s1.children[1].children[2], xMaxSlider: s1.children[2].children[2],
            yMinSlider: s2.children[1].children[2], yMaxSlider: s2.children[2].children[2],

            // ▼追加
            zMinSlider: s6.children[1].children[2], zMaxSlider: s6.children[2].children[2],

            rMinSlider: s3.children[1].children[2], rMaxSlider: s3.children[2].children[2],
            sMinSlider: s4.children[1].children[2], sMaxSlider: s4.children[2].children[2],
            tMinSlider: s5.children[1].children[2], tMaxSlider: s5.children[2].children[2],

            row1_X: s1, row1_Y: s2, row1_Z: s6,
            row2_R: s3, row2_S: s4, row2_T: s5,

            // 符号DD
            signXDD: ddSignX,
            signYDD: ddSignY,
            signZDD: ddSignZ,
            signRDD: ddSignR,
            signSDD: ddSignS,

            scaleMode: scaleXYDD,

            rbPos:  posRadios,
            frameSlider: frameSlider,
            frameText: frameText,
            btnSignFlip: btnSignFlip,
            btnGotoIn: btnGotoIn,
            btnGotoOut: btnGotoOut,
            btnPreset: btnPreset,
            btnReset: btnReset,
            btnApply: btnApply
        };

        attachHelpTips(ui);
        UIREF = ui;
        if (typeof syncSettingsToUI === "function") syncSettingsToUI();

        // ホバー時のヘルプ
        function linkHover(ctrl, text){ bindHoverHelp(ctrl, text, helpBar); }
        linkHover(ui.btnApply,   ui.btnApply.helpTip);
        linkHover(ui.btnGotoIn,  ui.btnGotoIn.helpTip);
        linkHover(ui.btnGotoOut, ui.btnGotoOut.helpTip);
        linkHover(ui.btnPreset,  ui.btnPreset.helpTip);
        linkHover(ui.btnReset,   ui.btnReset.helpTip);
        linkHover(ui.btnSignFlip,ui.btnSignFlip.helpTip);
        linkHover(ui.frameSlider,ui.frameSlider.helpTip);
        linkHover(ui.frameText,  ui.frameText.helpTip);
        linkHover(ui.scaleMode,  ui.scaleMode.helpTip);
        if (ui.signXDD) linkHover(ui.signXDD, ui.signXDD.helpTip);
        if (ui.signYDD) linkHover(ui.signYDD, ui.signYDD.helpTip);
        if (ui.signZDD) linkHover(ui.signZDD, ui.signZDD.helpTip);
        if (ui.signRDD) linkHover(ui.signRDD, ui.signRDD.helpTip);
        if (ui.signSDD) linkHover(ui.signSDD, ui.signSDD.helpTip);

        win.onResizing = win.onResize = function(){ this.layout.resize(); };
        if (win instanceof Window) { win.center(); win.show(); } else { win.layout.layout(true); }

        return win;
    }


    function syncSettingsToUI() {
        if (!UIREF) return;

        function setMinMax(group, minKey, maxKey) {
            var minText   = group.children[1].children[1];
            var minSlider = group.children[1].children[2];
            var maxText   = group.children[2].children[1];
            var maxSlider = group.children[2].children[2];

            minText.text    = String(settings[minKey]);
            minSlider.value = settings[minKey];
            maxText.text    = String(settings[maxKey]);
            maxSlider.value = settings[maxKey];
        }

        setMinMax(UIREF.row1_X, "xMin", "xMax");
        setMinMax(UIREF.row1_Y, "yMin", "yMax");

        // ▼追加
        if (UIREF.row1_Z) setMinMax(UIREF.row1_Z, "zMin", "zMax");

        setMinMax(UIREF.row2_R, "rMin", "rMax");
        setMinMax(UIREF.row2_S, "sMin", "sMax");
        setMinMax(UIREF.row2_T, "tMin", "tMax");

        // スケールXYモード（範囲外でも落ちない）
        if (UIREF.scaleMode && UIREF.scaleMode.items && UIREF.scaleMode.items.length) {
            var idx = (settings.scaleXY | 0);
            if (idx < 0 || idx >= UIREF.scaleMode.items.length) idx = 0;
            UIREF.scaleMode.selection = UIREF.scaleMode.items[idx];
        }

        // フレームオフセット
        if (UIREF.frameSlider && UIREF.frameText) {
            UIREF.frameSlider.value = settings.frameOffset;
            UIREF.frameText.text    = String(settings.frameOffset);
        }

        // 位置モードラジオ（追加分も含めて）
        if (UIREF.rbPos && UIREF.rbPos.length) {
            for (var j = 0; j < UIREF.rbPos.length; j++) {
                UIREF.rbPos[j].value = (j === (settings.posMode|0));
            }
        }

        // 自動始点トグル表示（ボタンのラベルに * を付ける仕様）
        if (UIREF.btnGotoIn)  UIREF.btnGotoIn.text  = settings.autoStartIn  ? "<*" : "<";
        if (UIREF.btnGotoOut) UIREF.btnGotoOut.text = settings.autoStartOut ? ">*" : ">";
    }


    function createValueControl(parent, label, minKey, maxKey, sliderMin, sliderMax) {
        var group = parent.add("group");
        group.orientation = "column";
        group.alignChildren = ["left", "top"];
        
        group.add("statictext", undefined, label);
        
        var minGroup = group.add("group");
        minGroup.add("statictext", undefined, "min:");
        var minText = minGroup.add("edittext", undefined, "0");
        minText.characters = 6;
        var minSlider = minGroup.add("slider", undefined, 0, sliderMin, sliderMax);
        minSlider.preferredSize.width = 80;
        
        var maxGroup = group.add("group");
        maxGroup.add("statictext", undefined, "max:");
        var maxText = maxGroup.add("edittext", undefined, "0");
        maxText.characters = 6;
        var maxSlider = maxGroup.add("slider", undefined, 0, sliderMin, sliderMax);
        maxSlider.preferredSize.width = 80;
        
        minSlider.onChanging = function() {
            settings[minKey] = Math.round(minSlider.value);
            minText.text = settings[minKey];
        };
        
        minText.onChange = function() {
            var val = parseFloat(minText.text) || 0;
            settings[minKey] = clamp(val, sliderMin, sliderMax);
            minSlider.value = settings[minKey];
        };
        
        maxSlider.onChanging = function() {
            settings[maxKey] = Math.round(maxSlider.value);
            maxText.text = settings[maxKey];
        };
        
        maxText.onChange = function() {
            var val = parseFloat(maxText.text) || 0;
            settings[maxKey] = clamp(val, sliderMin, sliderMax);
            maxSlider.value = settings[maxKey];
        };
        
        return group;
    }
    
    function gotoInPoint() {
        var comp = app.project.activeItem;
        if (!comp || !(comp instanceof CompItem)) return;
        var layers = comp.selectedLayers;
        if (layers.length === 0) return;
        comp.time = layers[0].inPoint;
    }
    
    function gotoOutPoint() {
        var comp = app.project.activeItem;
        if (!comp || !(comp instanceof CompItem)) return;
        var layers = comp.selectedLayers;
        if (layers.length === 0) return;
        comp.time = layers[0].outPoint;
    }
    
    // ========================================
    // プリセットウィンドウ
    // ========================================
    function showPresetWindow(atCursor) {
        var win = new Window("palette", "プリセット", undefined);
        win.orientation = "column";
        win.alignChildren = ["fill", "fill"];
        win.spacing = 10;
        win.margins = 10;

        // リストボックス
        var list = win.add("listbox", undefined, [], { multiselect: false });
        list.preferredSize = [300, 200];
        updatePresetList(list);

        // ダブルクリックで適用（※閉じない）
        list.onDoubleClick = function () {
            if (list.selection) {
                applyPreset(list.selection.index); // ← win.close() しない
            }
        };

        // キー操作：Delete=削除、Enter=適用（※閉じない）
        list.addEventListener("keydown", function (e) {
            if (e.keyName === "Delete" && list.selection) {
                removePreset(list.selection.index);
                updatePresetList(list);
            } else if (e.keyName === "Enter" && list.selection) {
                applyPreset(list.selection.index); // ← win.close() しない
            }
        });

        // 右クリックで削除
        list.addEventListener("mousedown", function (e) {
            if (e.button === 2 && list.selection) {
                removePreset(list.selection.index);
                updatePresetList(list);
            }
        });

        // ボタン行
        var btnGroup = win.add("group");
        btnGroup.alignment = ["fill", "top"];

        var btnUp = btnGroup.add("button", undefined, "↑");
        btnUp.preferredSize.width = 30;
        var btnDown = btnGroup.add("button", undefined, "↓");
        btnDown.preferredSize.width = 30;

        btnGroup.add("statictext", undefined, ""); // spacer

        var nameText = btnGroup.add("edittext", undefined, "");
        nameText.characters = 15;
        var btnSave = btnGroup.add("button", undefined, "Save");
        var btnImport = btnGroup.add("button", undefined, "Import");
        var btnExport = btnGroup.add("button", undefined, "Export");

        // 並べ替え
        btnUp.onClick = function () {
            if (list.selection) {
                var newIndex = movePreset(list.selection.index, -1);
                updatePresetList(list);
                list.selection = newIndex;
            }
        };
        btnDown.onClick = function () {
            if (list.selection) {
                var newIndex = movePreset(list.selection.index, 1);
                updatePresetList(list);
                list.selection = newIndex;
            }
        };

        // 保存
        btnSave.onClick = function () {
            var name = nameText.text;
            var values = {};
            for (var key in settings) values[key] = settings[key];
            addPreset(name, values);
            updatePresetList(list);
            nameText.text = "";
        };

        // インポート／エクスポート
        btnImport.onClick = function () {
            importPresets();
            updatePresetList(list);
        };
        btnExport.onClick = function () {
            exportPresets();
        };

        // ウィンドウ位置
        if (atCursor) {
            // カーソル近辺（厳密取得不可なので近似）
            win.location = [$.screens[0].left + 100, $.screens[0].top + 100];
        } else {
            win.center();
        }

        win.show();
    }

    
    function updatePresetList(list) {
        list.removeAll();
        for (var i = 0; i < presets.length; i++) {
            list.add("item", presets[i].name);
        }
    }
    
    function applyPreset(index) {
        if (index < 0 || index >= presets.length) return;
        var preset = presets[index];
        for (var key in preset.values) {
            settings[key] = preset.values[key];
        }
        syncSettingsToUI();
        // ← ここにあった alert(...) を削除
    }

    
    // ========================================
    // 初期化
    // ========================================
    loadPresets();
    buildUI(thisObj);
    
})(this);
