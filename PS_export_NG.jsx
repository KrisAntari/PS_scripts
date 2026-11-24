#target photoshop

(function() {
    // ==== НАСТРОЙКИ ====
    var BASE_LAYER_NAME = "";             // например, "Маска" — точное имя верхнеуровневого слоя
    var BASE_LAYER_INDEX_FROM_TOP = 0;    // 0 = верхний слой; используется, если BASE_LAYER_NAME пусто
    // ===================

    var doc, layers, outFolder, vis = [];
    try {
        if (app.documents.length === 0) { alert("Откройте PSD."); return; }
        doc = app.activeDocument;
        layers = doc.layers;
        if (layers.length < 2) { alert("Меньше двух слоёв."); return; }

        // Определяем опорный слой
        var baseIndex = -1;
        if (BASE_LAYER_NAME) {
            for (var k = 0; k < layers.length; k++)
                if (layers[k].name === BASE_LAYER_NAME) { baseIndex = k; break; }
            if (baseIndex < 0) {
                var list = [];
                for (var t = 0; t < layers.length; t++) list.push(t + ": " + layers[t].name);
                alert("Слой '" + BASE_LAYER_NAME + "' не найден.\n\nВерхнеуровневые слои:\n" + list.join("\n"));
                return;
            }
        } else {
            baseIndex = BASE_LAYER_INDEX_FROM_TOP;
            if (baseIndex < 0 || baseIndex >= layers.length) {
                alert("BASE_LAYER_INDEX_FROM_TOP вне диапазона."); return;
            }
        }

        // Проверка на группу
        if (layers[baseIndex].typename === "LayerSet") {
            var anyVisible = false;
            for (var c = 0; c < layers[baseIndex].layers.length; c++) {
                if (layers[baseIndex].layers[c].visible) { anyVisible = true; break; }
            }
            if (!anyVisible) {
                alert("Опорный слой — группа '" + layers[baseIndex].name +
                      "', и все её дочерние слои скрыты.\nВыберите другой опорный слой " +
                      "(или сделайте дочерние видимыми вручную).");
                return;
            }
        }

        if (!layers[baseIndex].visible) {
            if (!confirm("Опорный слой '" + layers[baseIndex].name + "' сейчас скрыт. Продолжить?")) return;
        }

        outFolder = Folder.selectDialog(
            "Опорный слой: '" + layers[baseIndex].name + "' (индекс " + baseIndex + ")\n" +
            "Режим: " + doc.mode + ", бит: " + doc.bitsPerChannel + "\n" +
            "Слоёв: " + layers.length + "\n\nПапка для JPG:"
        );
        if (!outFolder) return;

        var testFile = new File(outFolder.fsName + "/_test_write.txt");
        if (!testFile.open("w")) { alert("Нет прав на запись в:\n" + outFolder.fsName); return; }
        testFile.write("ok"); testFile.close(); testFile.remove();

        if (doc.mode !== DocumentMode.RGB) doc.changeMode(ChangeMode.RGB);
        if (doc.bitsPerChannel !== BitsPerChannelType.EIGHT) doc.bitsPerChannel = BitsPerChannelType.EIGHT;

        for (var i = 0; i < layers.length; i++) vis.push(layers[i].visible);

        var opts = new JPEGSaveOptions();
        opts.quality = 12;
        opts.embedColorProfile = true;

        app.displayDialogs = DialogModes.NO;

        var saved = 0, errors = [];
        for (var i = 0; i < layers.length; i++) {
            if (i === baseIndex) continue;

            for (var j = 0; j < layers.length; j++)
                if (layers[j].visible) { try { layers[j].visible = false; } catch (e) {} }

            try {
                layers[baseIndex].visible = true;
                layers[i].visible = true;
            } catch (e) {
                errors.push("Переключение видимости: " + e.message);
                continue;
            }

            app.refresh();  // даём Photoshop применить состояние

            if (!layers[baseIndex].visible || !layers[i].visible) {
                errors.push("Видимость не установилась при i=" + i + " (" + layers[i].name + ")");
                continue;
            }

            var base = String(layers[i].name)
                .replace(/[\\\/:*?"<>|]/g, "_")   // запрещённые символы
                .replace(/\.(jpe?g|png|tiff?|psd|gif|bmp|webp)$/i, "");  // срезаем известное расширение
            var fname = base + ".jpg";
            try {
                doc.saveAs(new File(outFolder.fsName + "/" + fname), opts, true, Extension.LOWERCASE);
                saved++;
            } catch (e) {
                errors.push('"' + layers[i].name + '": ' + e.message);
            }
        }

        app.displayDialogs = DialogModes.ALL;

        var msg = "Опорный слой: '" + layers[baseIndex].name + "'\n" +
                  "Сохранено: " + saved + " из " + (layers.length - 1);
        if (errors.length) msg += "\n\nОшибки:\n" + errors.slice(0, 5).join("\n");
        alert(msg);

    } catch (e) {
        app.displayDialogs = DialogModes.ALL;
        alert("ОШИБКА:\n" + e.message + "\nСтрока " + e.line);
    } finally {
        app.displayDialogs = DialogModes.ALL;
        try {
            if (layers && vis.length === layers.length)
                for (var k = 0; k < layers.length; k++) layers[k].visible = vis[k];
        } catch (e) {}
    }
})();