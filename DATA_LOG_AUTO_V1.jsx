#target illustrator
(function () {
    if (app.documents.length === 0) {
        alert("Buka file AI tujuan dulu.");
        return;
    }
    var doc = app.activeDocument;
    var targetFolderName = "DESIGN DATA PWR AT Sharepoint";
    function str(v) { return (v === undefined || v === null) ? "" : String(v); }
    function trimText(v) { return str(v).replace(/^\s+|\s+$/g, ""); }
    function display(v) { var x = trimText(v); return x === "" ? "-" : x; }
    function upper(v) { return trimText(v).toUpperCase(); }
    function formatFlute(v) { var x = upper(v); if (x === "B") return "B/F"; return x; }
    function lineCount(v) {
        var x = str(v).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
        var count = x.split("\n").length;
        return count < 1 ? 1 : count;
    }

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

    for (var r = 0; r < rootsToTry.length && !dataFolder; r++) {
        var direct = new Folder(rootsToTry[r].fsName + "/" + targetFolderName);
        if (direct.exists) { dataFolder = direct; break; }
        var children = [];
        try { children = rootsToTry[r].getFiles(function (f) { return f instanceof Folder; }); } catch (e3) {}
        for (var c = 0; c < children.length; c++) {
            var nested = new Folder(children[c].fsName + "/" + targetFolderName);
            if (nested.exists) { dataFolder = nested; break; }
        }
    }
    if (!dataFolder) {
        alert("Folder OneDrive belum terdeteksi otomatis.\n\nPilih folder:\n" + targetFolderName);
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
        alert("Tidak ada design_data_*.json di:\n" + dataFolder.fsName);
        return;
    }
    var dataFile = files[0];
    for (var i = 1; i < files.length; i++) {
        if (files[i].modified > dataFile.modified) dataFile = files[i];
    }
    dataFile.encoding = "UTF-8";
    if (!dataFile.open("r")) {
        alert("JSON terbaru tidak bisa dibuka:\n" + dataFile.fsName);
        return;
    }
    var raw = dataFile.read();
    dataFile.close();
    raw = raw.replace(/^\uFEFF/, "");
    var data;
    try { data = eval("(" + raw + ")"); }
    catch (jsonErr) { alert("Format JSON tidak valid:\n" + jsonErr); return; }
    if (!data || typeof data !== "object") {
        alert("Isi JSON tidak berisi data yang valid:\n" + dataFile.name);
        return;
    }

    var rows = [
        ["NO. MG", display(data.no_mg)],
        ["CUSTOMER", display(data.customer)],
        ["ITEM", display(data.item)],
        ["", ""],
        ["PIC SALES SUPPORT", upper(data.pic_sales_support)],
        ["PIC SALES", upper(data.sales)],
        ["", ""],
        ["SUBSTANCE", display(data.substance)],
        ["FLUTE", formatFlute(data.flute)],
        ["BOX TYPE", upper(data.box_type)],
        ["SHEET LENGTH", display(data.sheet_length)],
        ["SHEET WIDTH", display(data.sheet_width)],
        ["", ""],
        ["INSIDE BOX LENGTH", display(data.length) + " mm"],
        ["INSIDE BOX WIDTH", display(data.width) + " mm"],
        ["INSIDE BOX HEIGHT", display(data.height) + " mm"],
        ["", ""],
        ["POSISI KUPING", display(data.posisi_kuping)],
        ["JENIS SCORE", display(data.jenis_score)],
        ["QTY PER IKAT", display(data.qty_per_ikat)],
        ["JOINT TYPE", upper(data.joint_type)],
        ["LOGO KODE LOT", display(data.logo_kode_lot)],
        ["POSISI KODE LOT", display(data.posisi_kode_lot)],
        ["JENIS IKATAN BOX", display(data.jenis_ikatan_box)],
        ["FIRST PRODUCTION", display(data.first_production)],
        ["TOOLING", display(data.tooling)],
        ["", ""],
        ["REMARK", display(data.remark)],
        ["", ""],
        ["SOURCE", dataFile.name],
        ["GENERATED ON", display(data.generated_datetime)]
    ];

    try {
    var layer = null;
    for (var li = 0; li < doc.layers.length; li++) {
        if (doc.layers[li].name.toLowerCase() === "lampiran") { layer = doc.layers[li]; break; }
    }
    if (!layer) { layer = doc.layers.add(); layer.name = "LAMPIRAN"; }
    layer.visible = true;
    layer.locked = false;

    var abIndex = doc.artboards.getActiveArtboardIndex();
    var ab = doc.artboards[abIndex].artboardRect;
    var startX = ab[2] + 36;
    var startY = ab[1];
    var labelX = startX;
    var colonX = startX + 132;
    var valueX = startX + 146;
    var lineH = 16;
    var y = startY;

    function makeText(contents, x, yy, size, isBold) {
        var tf = layer.textFrames.add();
        // Illustrator memakai \r sebagai pemisah baris pada teks
        tf.contents = str(contents).replace(/\r\n/g, "\r").replace(/\n/g, "\r");
        tf.position = [x, yy];
        tf.textRange.characterAttributes.size = size;
        if (isBold) {
            try { tf.textRange.characterAttributes.textFont = app.textFonts.getByName("Arial-BoldMT"); } catch (ee) {}
        }
        return tf;
    }

    makeText("DESIGN DATA LOG", startX, y, 12, true);
    y -= 18;
    makeText("____________________________________________________________", startX, y, 8, false);
    y -= 18;

    for (var ri = 0; ri < rows.length; ri++) {
        var label = rows[ri][0];
        var value = rows[ri][1];
        if (label === "" && value === "") { y -= 8; continue; }
        makeText(label, labelX, y, 10, false);
        makeText(":", colonX, y, 10, false);
        makeText(value, valueX, y, 10, false);

        // V4: posisi baris berikut otomatis mengikuti jumlah baris pada value.
        y -= lineH * lineCount(value);
    }

    makeText("____________________________________________________________", startX, y + 7, 8, false);
    app.redraw();
    alert(
        "DATA LOG AUTO V4 BERHASIL.\n\n" +
        "JSON: " + dataFile.name + "\n" +
        "NO. MG: " + display(data.no_mg) + "\n" +
        "GENERATED ON: " + display(data.generated_datetime) + "\n" +
        "Layer: LAMPIRAN"
    );
    return true;
    } catch (err) {
        alert("Error: " + err.message + (err.line ? "\nBaris: " + err.line : ""));
        return false;
    }
})();
