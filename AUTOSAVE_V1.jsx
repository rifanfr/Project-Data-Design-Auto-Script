#target illustrator

(function () {
    if (app.documents.length === 0) { alert("Buka TEMPLATE_AUTOMATION.ai dulu."); return; }
    var doc = app.activeDocument;

    function str(v) { return (v === undefined || v === null) ? "" : String(v); }
    function trimText(v) { return str(v).replace(/^\s+|\s+$/g, ""); }
    // Illustrator memakai \r sebagai pemisah baris pada teks
    function toText(v) { return trimText(v).replace(/\r\n/g, "\r").replace(/\n/g, "\r"); }
    function upper(v) { return trimText(v).toUpperCase(); }
    function formatFlute(v) {
        var x = upper(v);
        if (x === "B") return "B/F";
        if (x === "C") return "C/F";
        if (x === "BC") return "BC/F";
        return x;
    }

    // Cari root OneDrive kerja dari environment Windows terlebih dahulu.
    // OneDrive for Business biasanya tersedia sebagai OneDriveCommercial.
    var targetFolderName = "DESIGN DATA PWR AT Sharepoint";
    var dataFolder = null;
    var rootsToTry = [];

    function pushRoot(pathText) {
        if (!pathText) return;
        var f = new Folder(pathText);
        if (f.exists) rootsToTry.push(f);
    }

    try { pushRoot($.getenv("OneDriveCommercial")); } catch (e0) {}
    try { pushRoot($.getenv("OneDrive")); } catch (e1) {}
    try { pushRoot($.getenv("UserProfile")); } catch (e2) {}
    pushRoot(new Folder("~").fsName);

    // Cari target folder pada root kandidat dan satu level di bawahnya.
    for (var r = 0; r < rootsToTry.length && !dataFolder; r++) {
        var direct = new Folder(rootsToTry[r].fsName + "/" + targetFolderName);
        if (direct.exists) {
            dataFolder = direct;
            break;
        }

        var children = [];
        try { children = rootsToTry[r].getFiles(function (f) { return f instanceof Folder; }); } catch (e3) {}
        for (var c = 0; c < children.length; c++) {
            var nested = new Folder(children[c].fsName + "/" + targetFolderName);
            if (nested.exists) {
                dataFolder = nested;
                break;
            }
        }
    }

    // Fallback aman: kalau Windows/Illustrator memakai path OneDrive yang tidak standar,
    // user cukup pilih folder DESIGN DATA PWR AT Sharepoint sekali pada run ini.
    if (!dataFolder) {
        alert("Folder OneDrive belum terdeteksi otomatis.\n\nPilih folder:\n" + targetFolderName + "\n\ndi jendela berikutnya.");
        var picked = Folder.selectDialog("Pilih folder " + targetFolderName);
        if (!picked) return;
        if (picked.name !== targetFolderName) {
            var child = new Folder(picked.fsName + "/" + targetFolderName);
            if (child.exists) picked = child;
        }
        dataFolder = picked;
    }

    var files = dataFolder.getFiles(function (f) {
        return (f instanceof File) && /^design_data_.*\.json$/i.test(f.name);
    });
    if (!files || files.length === 0) {
        alert("Tidak ada design_data_*.json di folder:\n" + dataFolder.fsName + "\n\nPastikan OneDrive sudah selesai sync.");
        return;
    }

    // Pilih JSON terbaru berdasarkan waktu modified.
    var dataFile = files[0];
    for (var i = 1; i < files.length; i++) {
        if (files[i].modified > dataFile.modified) dataFile = files[i];
    }

    dataFile.encoding = "UTF-8";
    if (!dataFile.open("r")) { alert("JSON terbaru tidak bisa dibuka:\n" + dataFile.fsName); return; }
    var raw = dataFile.read();
    dataFile.close();
    raw = raw.replace(/^\uFEFF/, "");

    var data;
    try { data = (typeof JSON !== "undefined" && JSON.parse) ? JSON.parse(raw) : eval("(" + raw + ")"); }
    catch (eJson) { alert("Format JSON tidak valid:\n" + eJson + "\n\nFile: " + dataFile.fsName); return; }
    if (!data || typeof data !== "object") { alert("Isi JSON tidak berisi data yang valid:\n" + dataFile.name); return; }

    try {

    var fluteValue = formatFlute(data.flute);
    var jointValue = upper(data.joint_type);
    var salesValue = upper(data.sales);
    var sizeValue = trimText(data.length) + "x" + trimText(data.width) + "x" + trimText(data.height) + " MM";
    var qtyValue = trimText(data.qty_per_ikat) + " /IKAT";

    var textData = {
        "CUSTOMER_01": toText(data.customer), "CUSTOMER_02": toText(data.customer),
        "ITEM_01": toText(data.item), "ITEM_02": toText(data.item),
        "SIZE_01": sizeValue, "SIZE_02": sizeValue,
        "SUBSTANCE_01": toText(data.substance), "SUBSTANCE_02": toText(data.substance),
        "FLUTE_01": fluteValue, "FLUTE_02": fluteValue,
        "NO_MG_01": toText(data.no_mg), "NO_MG_02": toText(data.no_mg),
        "SALES_02": salesValue,
        "BOX_TYPE_01": toText(data.box_type),
        "JOINT_TYPE_01": jointValue, "JOINT_TYPE_02": jointValue,
        "QTY_IKAT_01": qtyValue,
        "TANGGAL_01": toText(data.generated_date),
        "DESIGN_DATE_01": toText(data.generated_month)
    };

    var score = str(data.jenis_score).toLowerCase();
    var ikatan = str(data.jenis_ikatan_box).toLowerCase();
    var visibilityData = {
        "SCORE_1T": score.indexOf("1") >= 0,
        "SCORE_2T": score.indexOf("2") >= 0,
        "SCORE_3T": score.indexOf("3") >= 0,
        "IKATAN_LAWAN_ARAH_FLUTE": ikatan.indexOf("berlawanan") >= 0,
        "IKATAN_SEARAH_FLUTE": ikatan.indexOf("searah") >= 0 && ikatan.indexOf("berlawanan") < 0,
        "IKATAN_MENYILANG": ikatan.indexOf("menyilang") >= 0
    };

    function getVariable(name) {
        for (var i = 0; i < doc.variables.length; i++) if (doc.variables[i].name === name) return doc.variables[i];
        return null;
    }

    var items = [];
    function collect(container) {
        var list;
        try { list = container.pageItems; } catch (e) { return; }
        for (var i = 0; i < list.length; i++) {
            var it = list[i]; items.push(it);
            if (it.typename === "GroupItem") collect(it);
        }
    }
    collect(doc);

    var textUpdated = 0, visibilityUpdated = 0, missing = [];
    var textTotal = 0, visibilityTotal = 0;
    for (var tk in textData) if (textData.hasOwnProperty(tk)) textTotal++;
    for (var vk in visibilityData) if (visibilityData.hasOwnProperty(vk)) visibilityTotal++;

    for (var key in textData) {
        if (!textData.hasOwnProperty(key)) continue;
        var v = getVariable(key);
        if (!v) { missing.push(key + " (variable tidak ada)"); continue; }
        var target = null;
        try { target = v.pageItem; } catch (e1) {}
        if (target && target.typename === "TextFrame") {
            try { target.contents = textData[key]; textUpdated++; continue; } catch (e2) {}
        }
        var found = false;
        for (var ti = 0; ti < items.length; ti++) {
            var it = items[ti]; if (it.typename !== "TextFrame") continue;
            try {
                if (it.contentVariable && it.contentVariable.name === key) {
                    it.contents = textData[key]; textUpdated++; found = true; break;
                }
            } catch (e3) {}
        }
        if (!found) missing.push(key + " (text binding tidak ditemukan)");
    }

    for (var key2 in visibilityData) {
        if (!visibilityData.hasOwnProperty(key2)) continue;
        var vv = getVariable(key2);
        if (!vv) { missing.push(key2 + " (variable tidak ada)"); continue; }
        var vt = null;
        try { vt = vv.pageItem; } catch (e4) {}
        if (vt) {
            try { vt.hidden = !visibilityData[key2]; visibilityUpdated++; continue; } catch (e5) {}
        }
        var vf = false;
        for (var vi = 0; vi < items.length; vi++) {
            var pit = items[vi];
            try {
                if (pit.visibilityVariable && pit.visibilityVariable.name === key2) {
                    pit.hidden = !visibilityData[key2]; visibilityUpdated++; vf = true; break;
                }
            } catch (e6) {}
        }
        if (!vf) missing.push(key2 + " (visibility binding tidak ditemukan)");
    }

    app.redraw();

    var msg = "DESIGN DATA selesai.\n\n" +
              "Sumber OneDrive: " + dataFolder.fsName + "\n" +
              "JSON terbaru: " + dataFile.name + "\n" +
              "No. MG: " + str(data.no_mg) + "\n" +
              "Text di-update: " + textUpdated + "\n" +
              "Visibility di-update: " + visibilityUpdated + "\n\n" +
              "Format output:\nFLUTE: " + fluteValue + "\nJOINT: " + jointValue + "\nSALES: " + salesValue;

    if (missing.length) msg += "\n\nBelum match:\n- " + missing.join("\n- ");
    else msg += "\n\nSEMUA MAPPING BERHASIL.";

    // Auto Save As, seperti V3.
    if (missing.length === 0 && textUpdated === textTotal && visibilityUpdated === visibilityTotal) {
        var badChars = ["\\", "/", ":", "*", "?", String.fromCharCode(34), "<", ">", "|", "\r", "\n"];
        var safeName = function (v) {
            var s = str(v);
            for (var bc = 0; bc < badChars.length; bc++) s = s.split(badChars[bc]).join("_");
            return trimText(s).replace(/[. ]+$/, "");
        };
        var safeItem = safeName(data.item); if (!safeItem) safeItem = "DESIGN";
        var outputName = safeItem + "_MG" + safeName(data.no_mg) + ".ai";
        var outputFolder = Folder.selectDialog("Pilih folder penyimpanan file AI hasil");
        if (outputFolder) {
            var outputFile = new File(outputFolder.fsName + "/" + outputName);
            if (outputFile.exists && !confirm("File sudah ada:\n" + outputName + "\n\nTimpa file tersebut?")) {
                alert(msg + "\n\nSave dibatalkan karena file sudah ada."); return;
            }
            try {

    var opt = new IllustratorSaveOptions();

    doc.saveAs(outputFile, opt);

    // AUTO FILE DATA
    var isi = "USER - " + doc.path.fsName;
    saveOk = false;

    for (var tf = 0; tf < doc.textFrames.length; tf++) {

        if (
            doc.textFrames[tf].name === "FILE_DATA_MASTER" ||
            doc.textFrames[tf].name === "FILE_DATA_COPY"
        ) {

            doc.textFrames[tf].contents = isi;
        }
    }

    doc.save();
    saveOk = true;

    msg += "\n\nSAVE AS BERHASIL:\n" + outputFile.fsName;

} catch (se) {

    msg += "\n\nSAVE AS GAGAL:\n" + se;
}
        } else msg += "\n\nSave As dibatalkan.";
    } else msg += "\n\nSave As tidak dijalankan karena mapping belum " + textTotal + " text + " + visibilityTotal + " visibility.";

    alert(msg);
    return saveOk === true;
    } catch (err) {
        alert("Error: " + err.message + (err.line ? "\nBaris: " + err.line : ""));
        return false;
    }
})();
