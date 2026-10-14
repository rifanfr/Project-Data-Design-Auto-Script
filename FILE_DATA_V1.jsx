#target illustrator

(function () {
    if (app.documents.length === 0) {
        alert("Tidak ada dokumen terbuka.");
        return;
    }

    var doc = app.activeDocument;

    var folder = null;
    try { folder = doc.path; } catch (e0) {}
    if (!folder || !folder.fsName) {
        alert("Simpan file terlebih dahulu.");
        return;
    }

    try {
        var isi = "USER - " + folder.fsName;

        // Isi teks tetap bisa diubah walau layer atau objek terkunci; kuncinya dikembalikan setelahnya
        function setContents(tf, text) {
            try { tf.contents = text; return true; } catch (e2) {}
            var layer = null, layerLocked = false, itemLocked = false;
            try { layer = tf.layer; layerLocked = layer.locked; layer.locked = false; } catch (e3) {}
            try { itemLocked = tf.locked; tf.locked = false; } catch (e4) {}
            var ok = false;
            try { tf.contents = text; ok = true; } catch (e5) {}
            try { tf.locked = itemLocked; } catch (e6) {}
            try { if (layer) layer.locked = layerLocked; } catch (e7) {}
            return ok;
        }

        var masterCount = 0, copyCount = 0, failed = 0;

        for (var i = 0; i < doc.textFrames.length; i++) {
            var tf = doc.textFrames[i];
            var nm = tf.name;
            if (nm !== "FILE_DATA_MASTER" && nm !== "FILE_DATA_COPY") continue;
            if (setContents(tf, isi)) {
                if (nm === "FILE_DATA_MASTER") masterCount++; else copyCount++;
            } else {
                failed++;
            }
        }

        if (masterCount === 0) {
            alert("FILE_DATA_MASTER tidak ditemukan" + (failed ? " atau tidak bisa diubah." : "."));
            return;
        }

        alert(
            "FILE DATA berhasil diupdate.\n\n" +
            "MASTER: " + masterCount + "\n" +
            "COPY: " + copyCount +
            (failed ? "\nGagal diubah: " + failed : "") +
            "\n\n" + isi
        );
        return true;
    } catch (err) {
        alert("Error: " + err.message + (err.line ? "\nBaris: " + err.line : ""));
        return false;
    }
})();
