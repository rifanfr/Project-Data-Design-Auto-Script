#target illustrator

/*
 * LAUNCHER.V1
 * Copyright (c) 2026 RFR. All rights reserved.
 */

(function () {
    var dlg = new Window("dialog", "LAUNCHER.V1");
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 9;
    dlg.margins = 16;

    var title = dlg.add("statictext", undefined, "MENU");
    title.alignment = "center";

    var allBtn = dlg.add("button", undefined, "CREATE ALL");
    var sep1 = dlg.add("panel"); sep1.preferredSize.height = 2;
    var masterBtn = dlg.add("button", undefined, "DATA MC");
    var logAutoBtn = dlg.add("button", undefined, "DATA LOG");
    var dielineManualBtn = dlg.add("button", undefined, "DIE LINE MANUAL");
    var sep2 = dlg.add("panel"); sep2.preferredSize.height = 2;
    var closeBtn = dlg.add("button", undefined, "CLOSE", {name:"cancel"});
    var copyright = dlg.add("statictext", undefined, "\u00A9 2026 RFR. All rights reserved.");
    copyright.alignment = "center";
    try { copyright.graphics.font = ScriptUI.newFont(copyright.graphics.font.name, "REGULAR", 9); } catch (eFont) {}

    var LAUNCHER_FOLDER = new File($.fileName).parent;
    function baseFolder() { return LAUNCHER_FOLDER; }

    // Saat CREATE ALL, semua notifikasi dikumpulkan lalu ditampilkan sekali di akhir.
    var realAlert = alert;
    var collected = null;
    function notify(msg) {
        if (collected) collected.push(String(msg));
        else realAlert(msg);
    }
    function flushNotifications() {
        var list = collected;
        collected = null;
        if (list && list.length) realAlert(list.join("\n\n------------------------------\n\n"));
    }

    function findScript(names) {
        if (!(names instanceof Array)) names = [names];
        var bases = [baseFolder(), new Folder(baseFolder().fsName + "/Automation"), baseFolder().parent];
        for (var i = 0; i < names.length; i++) {
            for (var b = 0; b < bases.length; b++) {
                if (!bases[b]) continue;
                var f = new File(bases[b].fsName + "/" + names[i]);
                if (f.exists) return f;
            }
        }
        return null;
    }

    function evalScriptFile(fileObj, label, strict) {
        // Membayangi alert agar skrip anak (eval) ikut dikumpulkan saat CREATE ALL.
        var alert = notify;
        if (!fileObj) {
            alert(label + " tidak ditemukan di folder Launcher.\n\nFolder yang dicek:\n" + LAUNCHER_FOLDER.fsName);
            return false;
        }
        fileObj.encoding = "UTF-8";
        if (!fileObj.open("r")) {
            alert("Tidak bisa membuka " + label + ":\n" + fileObj.fsName);
            return false;
        }
        var code = fileObj.read();
        fileObj.close();
        code = code.replace(/^\uFEFF/, "").replace(/^\s*#target\s+illustrator\s*/i, "");
        try {
            var result = eval(code);
            // Skrip anak mengembalikan true jika sukses; strict dipakai di CREATE ALL.
            return strict ? (result === true) : true;
        }
        catch (e) {
            alert("Gagal menjalankan " + label + ":\n" + e + (e && e.line ? "\nBaris: " + e.line : ""));
            return false;
        }
    }

    function findMaster() { return findScript("AUTOSAVE_V1.jsx"); }
    function findDielineAuto() { return findScript("DIELINE_AUTO_V1.jsx"); }
    function findLogAuto() { return findScript("DATA_LOG_AUTO_V1.jsx"); }
    function findDielineManual() { return findScript("V1.jsx"); }

    // Final checkbox enforcement. Runs only during CREATE ALL, immediately before final save.
    function applyFinalCheckboxPatch() {
        var alert = notify;
        if (app.documents.length === 0) {
            alert("Checkbox patch gagal: dokumen aktif tidak tersedia.");
            return false;
        }
        var doc = app.activeDocument;
        var targetFolderName = "DESIGN DATA PWR AT Sharepoint";

        function str(v) { return (v === undefined || v === null) ? "" : String(v); }
        function trimText(v) { return str(v).replace(/^\s+|\s+$/g, ""); }

        var dataFolder = null;
        var roots = [];
        function pushRoot(p) {
            if (!p) return;
            var f = new Folder(p);
            if (f.exists) roots.push(f);
        }
        try { pushRoot($.getenv("OneDriveCommercial")); } catch (e0) {}
        try { pushRoot($.getenv("OneDrive")); } catch (e1) {}
        try { pushRoot($.getenv("UserProfile")); } catch (e2) {}
        pushRoot(new Folder("~").fsName);

        for (var r = 0; r < roots.length && !dataFolder; r++) {
            var direct = new Folder(roots[r].fsName + "/" + targetFolderName);
            if (direct.exists) { dataFolder = direct; break; }
            var children = [];
            try { children = roots[r].getFiles(function (f) { return f instanceof Folder; }); } catch (e3) {}
            for (var c = 0; c < children.length; c++) {
                var nested = new Folder(children[c].fsName + "/" + targetFolderName);
                if (nested.exists) { dataFolder = nested; break; }
            }
        }
        if (!dataFolder) {
            alert("Checkbox patch gagal: folder JSON OneDrive tidak ditemukan.");
            return false;
        }

        var files = dataFolder.getFiles(function (f) {
            return (f instanceof File) && /^design_data_.*\.json$/i.test(f.name);
        });
        if (!files || files.length === 0) {
            alert("Checkbox patch gagal: tidak ada design_data_*.json.");
            return false;
        }

        var dataFile = files[0];
        for (var fi = 1; fi < files.length; fi++) {
            if (files[fi].modified > dataFile.modified) dataFile = files[fi];
        }
        dataFile.encoding = "UTF-8";
        if (!dataFile.open("r")) {
            alert("Checkbox patch gagal: JSON tidak bisa dibuka.");
            return false;
        }
        var raw = dataFile.read(); dataFile.close();
        raw = raw.replace(/^\uFEFF/, "");
        var data;
        try { data = eval("(" + raw + ")"); }
        catch (e4) { alert("Checkbox patch gagal: JSON tidak valid.\n" + e4); return false; }
        if (!data || typeof data !== "object") { alert("Checkbox patch gagal: isi JSON bukan objek."); return false; }

        var score = trimText(data.jenis_score).toLowerCase();
        var ikatan = trimText(data.jenis_ikatan_box).toLowerCase();
        var state = {
            "SCORE_1T": score.indexOf("1") >= 0,
            "SCORE_2T": score.indexOf("2") >= 0,
            "SCORE_3T": score.indexOf("3") >= 0,
            "IKATAN_LAWAN_ARAH_FLUTE": ikatan.indexOf("berlawanan") >= 0,
            "IKATAN_SEARAH_FLUTE": ikatan.indexOf("searah") >= 0 && ikatan.indexOf("berlawanan") < 0,
            "IKATAN_MENYILANG": ikatan.indexOf("menyilang") >= 0
        };

        function getVar(name) {
            for (var vi = 0; vi < doc.variables.length; vi++) {
                if (doc.variables[vi].name === name) return doc.variables[vi];
            }
            return null;
        }

        var items = [];
        function collect(container) {
            var list;
            try { list = container.pageItems; } catch (e5) { return; }
            for (var ci = 0; ci < list.length; ci++) {
                items.push(list[ci]);
                if (list[ci].typename === "GroupItem") collect(list[ci]);
            }
        }
        collect(doc);

        var updated = 0;
        var totalKeys = 0;
        for (var k0 in state) { if (state.hasOwnProperty(k0)) totalKeys++; }
        var missing = [];
        for (var key in state) {
            if (!state.hasOwnProperty(key)) continue;
            var v = getVar(key);
            if (!v) { missing.push(key); continue; }

            var target = null;
            try { target = v.pageItem; } catch (e6) {}
            if (target) {
                try { target.hidden = !state[key]; updated++; continue; } catch (e7) {}
            }

            var found = false;
            for (var ii = 0; ii < items.length; ii++) {
                try {
                    if (items[ii].visibilityVariable && items[ii].visibilityVariable.name === key) {
                        items[ii].hidden = !state[key];
                        updated++;
                        found = true;
                        break;
                    }
                } catch (e8) {}
            }
            if (!found) missing.push(key);
        }

        app.redraw();
        if (missing.length || updated !== totalKeys) {
            alert("CREATE ALL berhenti di checkbox patch.\nUpdated: " + updated + "/" + totalKeys +
                  (missing.length ? "\nMissing:\n- " + missing.join("\n- ") : ""));
            return false;
        }
        return true;
    }

    function runAll() {
        var alert = notify;
        var master = findMaster(), dieline = findDielineAuto(), logAuto = findLogAuto();
        var missing = [];
        if (!master) missing.push("Master Card JSON OneDrive");
        if (!dieline) missing.push("Dieline JSON OneDrive V3");
        if (!logAuto) missing.push("DATA LOG AUTO");
        if (missing.length) {
            alert("CREATE ALL dibatalkan. Script tidak ditemukan:\n- " + missing.join("\n- "));
            return;
        }

        var lines = [];
        // Notifikasi sukses dari skrip anak dibuang; detail hanya dipertahankan jika langkah gagal.
        function step(name, fn) {
            var mark = collected.length;
            var ok = false;
            try { ok = fn() === true; } catch (es) { alert(name + ": " + es.message); }
            if (ok) {
                collected.length = mark;
                lines.push("[OK] " + name);
                return true;
            }
            lines.push("[GAGAL] " + name);
            var details = collected.slice(mark).join("\n\n");
            collected.length = mark;
            if (details) lines.push("\n" + details);
            return false;
        }
        function hasDoc() {
            if (app.documents.length > 0) return true;
            alert("Dokumen aktif tidak tersedia.");
            return false;
        }

        var allOk =
            step("Data MC", function () { return evalScriptFile(master, "Master Card JSON OneDrive", true); }) &&
            step("Die Line", function () { return hasDoc() && evalScriptFile(dieline, "Die Line JSON OneDrive", true); }) &&
            step("Data Log", function () { return hasDoc() && evalScriptFile(logAuto, "DATA LOG JSON OneDrive", true); }) &&
            step("Checkbox", function () { return hasDoc() && applyFinalCheckboxPatch(); }) &&
            step("Save", function () { app.activeDocument.save(); return true; });

        collected.length = 0;
        lines.unshift(allOk ? "CREATE ALL SELESAI\n" : "CREATE ALL BERHENTI\n");
        collected.push(lines.join("\n"));
    }

    allBtn.onClick = function () {
        dlg.close();
        collected = [];
        try { runAll(); }
        catch (eAll) { notify("Error CREATE ALL: " + eAll.message + (eAll.line ? "\nBaris: " + eAll.line : "")); }
        flushNotifications();
    };

    masterBtn.onClick = function () { dlg.close(); evalScriptFile(findMaster(), "Master Card JSON OneDrive"); };
    logAutoBtn.onClick = function () { dlg.close(); evalScriptFile(findLogAuto(), "DATA LOG JSON OneDrive"); };
    dielineManualBtn.onClick = function () { dlg.close(); evalScriptFile(findDielineManual(), "Die Line Manual V1"); };

    dlg.center();
    dlg.show();
})();
