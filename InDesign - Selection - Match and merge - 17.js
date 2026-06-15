/**
 * Selection Inspector
 *
 * TEXT MODE:  Shows paragraph style, character style(s), font size, leading,
 *             tracking. Match Properties with tolerance; step-through navigation.
 *
 * FRAME MODE: Shows frame type, object style, width, height.
 *             Match Properties with tolerance; step-through navigation.
 *
 * CELL MODE:  Shows cell style, height, width, insets, and parent table info.
 *             Match against any combination; tags cells from the same table.
 *
 * TABLE MODE: Shows table style, row count, column count.
 *             Match against style, rows, and/or columns.
 *
 * Navigation and Merge powered by app.doScript() for reliable engine access.
 */

(function () {

    if (!app.documents.length) { alert("No document open."); return; }

    var doc = app.activeDocument;
    var sel = app.selection;

    if (!sel || !sel.length) { alert("Nothing selected."); return; }

    var s0 = sel[0];
    if (!s0 || !s0.isValid) { alert("Selection is not valid."); return; }

    // ── Helpers ──────────────────────────────────────────────────────────────

    function safeStr(fn) {
        try { var v = fn(); return (v !== null && v !== undefined) ? String(v) : "\u2014"; }
        catch (e) { return "\u2014"; }
    }

    function roundTo(n, dec) {
        var f = Math.pow(10, dec);
        return Math.round(n * f) / f;
    }

    function getCharStylesInPara(para) {
        var seen = {}, names = [];
        try {
            var chars = para.characters;
            for (var i = 0; i < chars.length; i++) {
                try {
                    var n = chars[i].appliedCharacterStyle.name;
                    if (!seen[n]) { seen[n] = true; names.push(n); }
                } catch (e) {}
            }
        } catch (e) {}
        return names.length ? names.join(", ") : "[None]";
    }

    function getParaPage(para) {
        try {
            var frame = para.parentTextFrames[0];
            if (frame && frame.isValid) {
                var page = frame.parentPage;
                if (page && page.isValid) return page.name;
            }
        } catch (e) {}
        return "?";
    }

    function getItemPage(item) {
        try {
            var page = item.parentPage;
            if (page && page.isValid) return page.name;
        } catch (e) {}
        return "?";
    }

    // Convert FirstBaseline enum value to a human-readable display string
    function baselineOffsetName(val) {
        try { if (val === FirstBaseline.ASCENT_OFFSET) return "Ascent"; }      catch (e) {}
        try { if (val === FirstBaseline.CAP_HEIGHT)    return "Cap Height"; }   catch (e) {}
        try { if (val === FirstBaseline.LEADING)       return "Leading"; }      catch (e) {}
        try { if (val === FirstBaseline.EMBOX_HEIGHT)  return "Em Box Height"; }catch (e) {}
        try { if (val === FirstBaseline.FIXED_HEIGHT)  return "Fixed"; }        catch (e) {}
        try { if (val === FirstBaseline.X_HEIGHT)      return "x Height"; }     catch (e) {}
        return String(val);
    }

    function getTablePage(tbl) {
        try {
            var fr = tbl.parent;
            while (fr && fr.constructor && fr.constructor.name !== "TextFrame") {
                fr = fr.parent;
            }
            if (fr && fr.parentPage) return fr.parentPage.name;
        } catch (e) {}
        return "?";
    }

    function trunc(str, len) {
        try {
            var s = str.replace(/\r|\n/g, " ");
            return s.length > len ? s.substring(0, len) + "\u2026" : s;
        } catch (e) { return ""; }
    }

    function fmtDev(delta) {
        var r = roundTo(delta, 2);
        return (r >= 0 ? "+" : "") + r;
    }

    function buildDevFlag(dev) {
        if (!dev) return "";
        var parts = [];
        if (dev.fontSize        !== undefined && dev.fontSize        !== 0)  parts.push("size "    + fmtDev(dev.fontSize)   + " pt");
        if (dev.leading         !== undefined && dev.leading         !== 0)  parts.push("leading " + fmtDev(dev.leading)    + " pt");
        if (dev.tracking        !== undefined && dev.tracking        !== 0)  parts.push("tracking "+ fmtDev(dev.tracking));
        if (dev.width           !== undefined && dev.width           !== 0)  parts.push("w "       + fmtDev(dev.width)      + " pt");
        if (dev.height          !== undefined && dev.height          !== 0)  parts.push("h "       + fmtDev(dev.height)     + " pt");
        if (dev.cellWidth       !== undefined && dev.cellWidth       !== 0)  parts.push("cell-w "  + fmtDev(dev.cellWidth)  + " pt");
        if (dev.cellHeight      !== undefined && dev.cellHeight      !== 0)  parts.push("cell-h "  + fmtDev(dev.cellHeight) + " pt");
        if (dev.colWidths !== undefined) {
            var cwHasNonZero = false;
            for (var cwbi = 0; cwbi < dev.colWidths.length; cwbi++) { if (dev.colWidths[cwbi] !== 0) { cwHasNonZero = true; break; } }
            if (cwHasNonZero) {
                var cwFmtParts = [];
                for (var cwfi = 0; cwfi < dev.colWidths.length; cwfi++) cwFmtParts.push(fmtDev(dev.colWidths[cwfi]));
                parts.push("col-w [" + cwFmtParts.join(", ") + "] pt");
            }
        }
        // String-valued passive style deviation — rendered with → rather than ±
        if (dev.styleDeviation  !== undefined && dev.styleDeviation  !== "") parts.push("style \u2192 \"" + dev.styleDeviation + "\"");
        return parts.length ? "  [\u03b4 " + parts.join(", ") + "]" : "";
    }

    // Returns true if any property in dev is non-zero (i.e. the match is fuzzy)
    function hasDev(dev) {
        if (!dev) return false;
        for (var k in dev) {
            if (!dev.hasOwnProperty(k)) continue;
            var v = dev[k];
            if (typeof v === "string"  && v !== "") return true;
            if (typeof v === "number"  && v !== 0)  return true;
            if (typeof v === "object"  && v !== null && typeof v.length === "number") {
                for (var ai = 0; ai < v.length; ai++) { if (v[ai] !== 0) return true; }
            }
        }
        return false;
    }

    // ── Navigation functions ─────────────────────────────────────────────────

    function navToPara(para) {
        try {
            var frame = para.parentTextFrames[0];
            if (!frame || !frame.isValid) return;
            var page = frame.parentPage;
            if (!page || !page.isValid) return;
            app.activeWindow.activePage = page;
            para.select();
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    function navToItem(item) {
        try {
            var page = item.parentPage;
            if (!page || !page.isValid) return;
            app.activeWindow.activePage = page;
            app.select(item);
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    function navToCell(cell) {
        try {
            var fr = cell;
            while (fr && fr.constructor && fr.constructor.name !== "TextFrame") fr = fr.parent;
            if (fr && fr.parentPage) app.activeWindow.activePage = fr.parentPage;
            app.select(cell);
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    function navToTable(tbl) {
        try {
            var fr = tbl.parent;
            while (fr && fr.constructor && fr.constructor.name !== "TextFrame") fr = fr.parent;
            if (fr && fr.parentPage) app.activeWindow.activePage = fr.parentPage;
            app.select(tbl);
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    // ── Keyword find/change helper ───────────────────────────────────────────

    function applyKeywordReplace(obj, findStr, replaceStr) {
        try {
            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
            app.findTextPreferences.findWhat = findStr;
            app.findTextPreferences.caseSensitive = false;
            app.changeTextPreferences.changeTo = replaceStr;
            try { obj.changeText(); } catch (e2) {}
        } catch (e) {}
        try {
            app.findTextPreferences = NothingEnum.nothing;
            app.changeTextPreferences = NothingEnum.nothing;
        } catch (e3) {}
    }

    // ── Listbox helpers ──────────────────────────────────────────────────────

    function getSelectedIndices(lb) {
        var indices = [];
        if (!lb.selection) return indices;
        var sel = lb.selection;
        if (typeof sel.length === "number") {
            for (var i = 0; i < sel.length; i++) indices.push(sel[i].index);
        } else {
            indices.push(sel.index);
        }
        indices.sort(function (a, b) { return a - b; });
        return indices;
    }

    function restoreSelection(lb, indices) {
        for (var i = 0; i < lb.items.length; i++) lb.items[i].selected = false;
        for (var j = 0; j < indices.length; j++) {
            if (indices[j] >= 0 && indices[j] < lb.items.length) {
                lb.items[indices[j]].selected = true;
            }
        }
    }

    // ── Adjust Merge Targets dialog ──────────────────────────────────────────
    //
    // fields[] entries: { applyKey, valueKey, label, unit, type, devKey,
    //   styleCollection?, minValue?, displayStr?, enabled }
    //   type: "string" | "number" | "enum" | "colWidths"
    //
    function buildAdjustDialog(mergeTarget, fields, deviantMatches) {
        // Collect which dev keys have non-zero deviations across all deviants
        var deviantKeys = {};
        for (var dmi = 0; dmi < deviantMatches.length; dmi++) {
            var dmdev = deviantMatches[dmi].dev;
            if (!dmdev) continue;
            for (var ddk in dmdev) {
                if (!dmdev.hasOwnProperty(ddk)) continue;
                var ddv = dmdev[ddk];
                if (typeof ddv === "string" && ddv !== "")  { deviantKeys[ddk] = true; }
                else if (typeof ddv === "number" && ddv !== 0)   { deviantKeys[ddk] = true; }
                else if (typeof ddv === "object" && ddv !== null) {
                    for (var dai = 0; dai < ddv.length; dai++) { if (ddv[dai] !== 0) { deviantKeys[ddk] = true; break; } }
                }
            }
        }

        var adjDlg = new Window("dialog", "Adjust Merge Targets");
        adjDlg.alignChildren = ["fill", "top"];
        adjDlg.margins = 18;
        adjDlg.spacing = 10;
        adjDlg.preferredSize.width = 420;

        var adjPropSec = addSection(adjDlg, "Properties to apply");
        adjPropSec.spacing = 6;

        var allAdjCbs    = [];
        var adjFldCtrls  = [];

        for (var fi = 0; fi < fields.length; fi++) {
            var f = fields[fi];
            if (!f.enabled) continue;

            if (f.type === "countWithDefault") {
                var cdrGrp = adjPropSec.add("group");
                cdrGrp.alignChildren = ["left", "center"];
                cdrGrp.spacing = 8;
                var cdrCb = cdrGrp.add("checkbox", undefined, "");
                cdrCb.value = false;
                allAdjCbs.push(cdrCb);
                var cdrLbl = cdrGrp.add("statictext", undefined, f.label + ":");
                cdrLbl.preferredSize.width = 150;
                var cdrMain = cdrGrp.add("edittext", undefined, String(mergeTarget[f.valueKey]));
                cdrMain.preferredSize.width = 60;
                if (f.unit) cdrGrp.add("statictext", undefined, f.unit);

                var cdrSubRow = adjPropSec.add("group");
                cdrSubRow.alignChildren = ["left", "center"];
                cdrSubRow.spacing = 8;
                cdrSubRow.margins = [24, 0, 0, 0];
                var cdrSubLbl = cdrSubRow.add("statictext", undefined, f.sub.label + ":");
                cdrSubLbl.preferredSize.width = 150;
                var cdrSubInp = cdrSubRow.add("edittext", undefined, String(mergeTarget[f.sub.valueKey]));
                cdrSubInp.preferredSize.width = 60;
                cdrSubRow.add("statictext", undefined, "pt");
                cdrSubRow.visible = false;

                (function (thisCb, thisSubRow) {
                    thisCb.onClick = function () {
                        thisSubRow.visible = thisCb.value;
                        adjDlg.layout.layout(true);
                        adjDlg.layout.resize();
                        updateAdjMergeBtn();
                    };
                })(cdrCb, cdrSubRow);

                adjFldCtrls.push({ cb: cdrCb, inputs: [cdrMain], subInputs: [cdrSubInp], field: f });
                continue;
            }

            if (f.type === "keywordReplace") {
                var kwrGrp = adjPropSec.add("group");
                kwrGrp.alignChildren = ["left", "center"];
                kwrGrp.spacing = 8;
                var kwrCb = kwrGrp.add("checkbox", undefined, "");
                kwrCb.value = false;
                allAdjCbs.push(kwrCb);
                var kwrLbl = kwrGrp.add("statictext", undefined, "Replace \"" + f.keyword + "\" with:");
                kwrLbl.preferredSize.width = 200;
                var kwrInp = kwrGrp.add("edittext", undefined, "");
                kwrInp.preferredSize.width = 130;
                adjFldCtrls.push({ cb: kwrCb, inputs: [kwrInp], subInputs: [], field: f });
                continue;
            }

            if (f.type === "colWidths") {
                var cwHdrGrp = adjPropSec.add("group");
                cwHdrGrp.alignChildren = ["left", "center"];
                cwHdrGrp.spacing = 8;
                var cwCb = cwHdrGrp.add("checkbox", undefined, f.label);
                cwCb.value = !!(f.devKey && deviantKeys[f.devKey]);
                allAdjCbs.push(cwCb);

                var cwBodyGrp = adjPropSec.add("group");
                cwBodyGrp.orientation = "column";
                cwBodyGrp.alignChildren = ["left", "top"];
                cwBodyGrp.margins = [20, 0, 0, 0];
                cwBodyGrp.spacing = 4;

                var cwInputs = [];
                var cwVals = mergeTarget[f.valueKey];
                for (var cwi = 0; cwi < cwVals.length; cwi++) {
                    var cwRow = cwBodyGrp.add("group");
                    cwRow.alignChildren = ["left", "center"];
                    cwRow.spacing = 6;
                    var cwLbl = cwRow.add("statictext", undefined, "Col " + (cwi + 1) + ":");
                    cwLbl.preferredSize.width = 50;
                    var cwFld = cwRow.add("edittext", undefined, String(cwVals[cwi]));
                    cwFld.preferredSize.width = 70;
                    cwRow.add("statictext", undefined, "pt");
                    cwInputs.push(cwFld);
                }
                adjFldCtrls.push({ cb: cwCb, inputs: cwInputs, field: f });
                continue;
            }

            var adjRow = adjPropSec.add("group");
            adjRow.alignChildren = ["left", "center"];
            adjRow.spacing = 8;

            var adjCb = adjRow.add("checkbox", undefined, "");
            adjCb.value = !!(f.devKey && deviantKeys[f.devKey]);
            allAdjCbs.push(adjCb);

            var adjLbl = adjRow.add("statictext", undefined, f.label + ":");
            adjLbl.preferredSize.width = 150;

            var adjInp;
            if (f.type === "enum") {
                adjInp = adjRow.add("statictext", undefined, f.displayStr || "—");
                adjInp.preferredSize.width = 130;
            } else {
                var adjInitVal = mergeTarget[f.valueKey];
                adjInp = adjRow.add("edittext", undefined,
                    (adjInitVal !== null && adjInitVal !== undefined) ? String(adjInitVal) : "");
                adjInp.preferredSize.width = 130;
            }

            if (f.unit) adjRow.add("statictext", undefined, f.unit);

            adjFldCtrls.push({ cb: adjCb, inputs: [adjInp], field: f });
        }

        var adjBtnGrp = adjDlg.add("group");
        adjBtnGrp.alignment = "right";
        adjBtnGrp.spacing = 8;
        var adjMergeBtn  = adjBtnGrp.add("button", undefined, "Merge");
        var adjCancelBtn = adjBtnGrp.add("button", undefined, "Cancel");

        function updateAdjMergeBtn() {
            var any = false;
            for (var aci = 0; aci < allAdjCbs.length; aci++) { if (allAdjCbs[aci].value) { any = true; break; } }
            adjMergeBtn.enabled = any;
        }
        for (var adjCbi = 0; adjCbi < allAdjCbs.length; adjCbi++) {
            (function (thisCb) {
                // Skip checkboxes that already have a custom onClick (e.g. countWithDefault)
                if (!thisCb.onClick) {
                    thisCb.onClick = function () { updateAdjMergeBtn(); };
                }
            })(allAdjCbs[adjCbi]);
        }
        updateAdjMergeBtn();

        var adjConfirmed = false;

        adjMergeBtn.onClick = function () {
            for (var vi = 0; vi < adjFldCtrls.length; vi++) {
                var vfc = adjFldCtrls[vi];
                if (!vfc.cb.value) continue;
                if (vfc.field.type === "number") {
                    var nv = parseFloat(vfc.inputs[0].text);
                    if (isNaN(nv)) { alert("\"" + vfc.field.label + "\" requires a valid number."); return; }
                    if (vfc.field.minValue !== null && vfc.field.minValue !== undefined && nv <= vfc.field.minValue) {
                        alert("\"" + vfc.field.label + "\" must be greater than " + vfc.field.minValue + "."); return;
                    }
                }
                if (vfc.field.type === "colWidths") {
                    for (var cwvi = 0; cwvi < vfc.inputs.length; cwvi++) {
                        var cwv = parseFloat(vfc.inputs[cwvi].text);
                        if (isNaN(cwv) || cwv <= 0) { alert("Column " + (cwvi + 1) + " width must be a positive number."); return; }
                    }
                }
                if (vfc.field.type === "countWithDefault") {
                    var cdv = parseInt(vfc.inputs[0].text);
                    if (isNaN(cdv) || cdv <= 0) { alert("\"" + vfc.field.label + "\" must be a positive whole number."); return; }
                    var sdv = parseFloat(vfc.subInputs[0].text);
                    if (isNaN(sdv) || sdv <= 0) { alert("\"" + vfc.field.sub.label + "\" must be a positive number."); return; }
                }
                if (vfc.field.styleCollection) {
                    var sname = vfc.inputs[0].text;
                    var scoll = vfc.field.styleCollection();
                    var sfound = false;
                    for (var sci2 = 0; sci2 < scoll.length; sci2++) { if (scoll[sci2].name === sname) { sfound = true; break; } }
                    if (!sfound) { alert("Style not found: \"" + sname + "\".\nCheck the name and try again."); return; }
                }
            }
            for (var wi = 0; wi < adjFldCtrls.length; wi++) {
                var wfc = adjFldCtrls[wi];
                mergeTarget[wfc.field.applyKey] = wfc.cb.value;
                if (wfc.field.type === "number" || wfc.field.type === "string") {
                    mergeTarget[wfc.field.valueKey] = wfc.inputs[0].text;
                }
                if (wfc.field.type === "colWidths") {
                    for (var cwwi = 0; cwwi < wfc.inputs.length; cwwi++) {
                        mergeTarget[wfc.field.valueKey][cwwi] = parseFloat(wfc.inputs[cwwi].text);
                    }
                }
                if (wfc.field.type === "countWithDefault") {
                    mergeTarget[wfc.field.valueKey]      = wfc.inputs[0].text;
                    mergeTarget[wfc.field.sub.valueKey]  = wfc.subInputs[0].text;
                }
                if (wfc.field.type === "keywordReplace") {
                    mergeTarget[wfc.field.valueKey] = wfc.inputs[0].text;
                }
            }
            adjConfirmed = true;
            adjDlg.close();
        };

        adjCancelBtn.onClick = function () { adjDlg.close(); };

        adjDlg.show();
        return adjConfirmed;
    }

    // ── Core navigator — runs inside app.doScript ────────────────────────────
    //
    // Each match object: { ref, baseEntry, entry, dev, props }
    //   ref       — live InDesign object
    //   baseEntry — display string without deviation flag
    //   entry     — baseEntry + buildDevFlag(dev), updated after a Merge
    //   dev       — { fontSize: δ, leading: δ, … } — zero means exact
    //   props     — "Matched on" summary lines
    //
    // mergeFn(match) — mode-specific function that writes target values back
    //   to the InDesign object and zeroes the relevant dev keys.
    //   Pass null for modes with no numeric deviations (Table mode).
    //
    function runStepThrough(matches, navFn, label, mergeFn, adjustAndConfirmFn) {
        if (!matches || matches.length === 0) { alert("No matching " + label + " found."); return; }

        var total        = matches.length;
        var subset       = [];
        var subsetCursor = -1;
        var navStarted   = false;

        var entryStrings = [];
        for (var ei = 0; ei < matches.length; ei++) entryStrings.push(matches[ei].entry);

        var listDlg = new Window("dialog", "Match Results");
        listDlg.alignChildren = ["fill", "top"];
        listDlg.margins = 18;
        listDlg.spacing = 10;
        listDlg.preferredSize.width = 460;

            // ── Matched-on summary ───────────────────────────────────────────
            var hdrSec = listDlg.add("panel", undefined, "Matched on");
            hdrSec.alignChildren = ["left", "top"];
            hdrSec.margins = [10, 15, 10, 10];
            hdrSec.alignment = ["fill", "top"];
            var propsArr = matches[0].props;
            for (var pi = 0; pi < propsArr.length; pi++) {
                hdrSec.add("statictext", undefined, "\u2022  " + propsArr[pi]);
            }

            // ── Scrollable result list (multi-select) ────────────────────────
            var listSec = listDlg.add("panel", undefined, "Results  (" + total + " " + label + ")");
            listSec.alignChildren = ["fill", "top"];
            listSec.margins = [10, 15, 10, 10];
            listSec.alignment = ["fill", "top"];

            var lb = listSec.add("listbox", undefined, entryStrings, { multiselect: true });
            lb.preferredSize = [420, 220];

        var navLabel = listSec.add("statictext", undefined, "");
        navLabel.alignment = ["right", "top"];

        // ── Buttons ──────────────────────────────────────────────────────
            var btnGrp = listDlg.add("group");
            btnGrp.alignment = "right";
            btnGrp.spacing = 8;
            var mergeBtn = btnGrp.add("button", undefined, "Merge");
            var prevBtn  = btnGrp.add("button", undefined, "\u2190 Previous");
            var nextBtn  = btnGrp.add("button", undefined, "Next \u2192");
            var closeBtn = btnGrp.add("button", undefined, "Close");

        mergeBtn.enabled = false;
        prevBtn.enabled  = false;
        nextBtn.enabled  = false;

        lb.onChange = function () {
                var newSel = getSelectedIndices(lb);
                var hasNew = (newSel.length > 0);
                nextBtn.enabled  = hasNew;
                prevBtn.enabled  = hasNew;
                mergeBtn.enabled = (mergeFn !== null && hasNew);
                if (newSel.join(",") !== subset.join(",")) navStarted = false;
            };

        nextBtn.onClick = function () {
                var current = getSelectedIndices(lb);
                if (current.length === 0) return;
                if (!navStarted || current.join(",") !== subset.join(",")) {
                    subset = current; subsetCursor = 0; navStarted = true;
                } else {
                    subsetCursor = Math.min(subsetCursor + 1, subset.length - 1);
                }
                navFn(matches[subset[subsetCursor]].ref);
                navLabel.text = "Showing " + (subsetCursor + 1) + " of " + subset.length;
            };

            prevBtn.onClick = function () {
                var current = getSelectedIndices(lb);
                if (current.length === 0) return;
                if (!navStarted || current.join(",") !== subset.join(",")) {
                    subset = current; subsetCursor = subset.length - 1; navStarted = true;
                } else {
                    subsetCursor = Math.max(subsetCursor - 1, 0);
                }
                navFn(matches[subset[subsetCursor]].ref);
                navLabel.text = "Showing " + (subsetCursor + 1) + " of " + subset.length;
            };

            mergeBtn.onClick = function () {
                var current = getSelectedIndices(lb);
                if (current.length === 0) return;

                // Stage 4: open Adjust Merge Targets dialog for all selected items
                var selectedMatchObjs = [];
                for (var si2 = 0; si2 < current.length; si2++) selectedMatchObjs.push(matches[current[si2]]);
                if (!adjustAndConfirmFn(selectedMatchObjs)) return;
                for (var mi = 0; mi < current.length; mi++) mergeFn(matches[current[mi]]);
                subset = current;
                lb.removeAll();
                for (var ri = 0; ri < matches.length; ri++) lb.add("item", matches[ri].entry);
                restoreSelection(lb, subset);
                mergeBtn.enabled = false;
            };

        closeBtn.onClick = function () {
            listDlg.close();
        };

        listDlg.show();
    }

    // ── Mode detection ───────────────────────────────────────────────────────

    var TEXT_TYPES = {
        "InsertionPoint": true, "Text": true, "Character": true,
        "Word": true, "Line": true, "Paragraph": true, "TextStyleRange": true
    };
    var typeName   = s0.constructor ? s0.constructor.name : "";
    var inTextMode = !!TEXT_TYPES[typeName];

    // ── ScriptUI helpers ─────────────────────────────────────────────────────

    function addSection(parent, title) {
        var p = parent.add("panel", undefined, title);
        p.alignChildren = ["left", "top"];
        p.margins = [10, 15, 10, 10];
        p.spacing = 8;
        p.alignment = ["fill", "top"];
        return p;
    }

    function addRow(parent, label, value) {
        var g = parent.add("group");
        g.alignChildren = ["left", "top"];
        g.spacing = 8;
        var lbl = g.add("statictext", undefined, label, { multiline: true });
        lbl.preferredSize.width = 130;
        lbl.justify = "right";
        var val = g.add("statictext", undefined, String(value), { multiline: true });
        val.preferredSize.width = 220;
    }

    function addCheckRow(parent, label) {
        var g = parent.add("group");
        g.alignChildren = ["left", "center"];
        g.spacing = 8;
        var cb = g.add("checkbox", undefined, label);
        cb.value = false;
        return cb;
    }

    function addToleranceRow(parent, label, unit) {
        var g = parent.add("group");
        g.alignChildren = ["left", "center"];
        g.spacing = 8;
        var cb = g.add("checkbox", undefined, label);
        cb.preferredSize.width = 240;
        g.add("statictext", undefined, "\u00b1");
        var tol = g.add("edittext", undefined, "0");
        tol.preferredSize.width = 46;
        g.add("statictext", undefined, unit);
        return { cb: cb, tol: tol };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // TEXT MODE
    // ══════════════════════════════════════════════════════════════════════════

    if (inTextMode) {

        var para;
        try { para = (typeName === "Paragraph") ? s0 : s0.paragraphs[0]; } catch (e) {}
        if (!para || !para.isValid) { alert("Could not resolve paragraph."); return; }

        var paraStyleRaw = safeStr(function () { return para.appliedParagraphStyle.name; });
        var fontSizeRaw  = safeStr(function () { return roundTo(para.pointSize, 2); });
        var leadingRaw   = safeStr(function () { return roundTo(para.leading,   2); });
        var trackingRaw  = safeStr(function () { return para.tracking; });
        var charStyles   = getCharStylesInPara(para);

        var mergeTargetTx = {
            applyStyle: false, style:    paraStyleRaw,
            applySize:  false, fontSize: fontSizeRaw,
            applyLeading:  false, leading:  leadingRaw,
            applyTracking: false, tracking: trackingRaw
        };

        var dlg = new Window("dialog", "Selection Inspector");
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = 18;
        dlg.spacing = 12;
        dlg.preferredSize.width = 420;

        var infoSec = addSection(dlg, "Paragraph");
        addRow(infoSec, "Paragraph Style:",    paraStyleRaw);
        addRow(infoSec, "Character Style(s):", charStyles);
        addRow(infoSec, "Font Size:",          fontSizeRaw + " pt");
        addRow(infoSec, "Leading:",            leadingRaw  + " pt");
        addRow(infoSec, "Tracking:",           trackingRaw);

        var propSec = addSection(dlg, "Match Properties");
        var cbParaStyle = addCheckRow(propSec,     "Paragraph Style  (" + paraStyleRaw + ")");
        var rFontSize   = addToleranceRow(propSec, "Font Size  ("  + fontSizeRaw + " pt)", "pt");
        var rLeading    = addToleranceRow(propSec, "Leading  ("    + leadingRaw  + " pt)", "pt");
        var rTracking   = addToleranceRow(propSec, "Tracking  ("   + trackingRaw + ")",    "u");
        var cbFontSize  = rFontSize.cb; var tolFontSize = rFontSize.tol;
        var cbLeading   = rLeading.cb;  var tolLeading  = rLeading.tol;
        var cbTracking  = rTracking.cb; var tolTracking = rTracking.tol;
        var kwRowTx = propSec.add("group");
        kwRowTx.alignChildren = ["left", "center"];
        kwRowTx.spacing = 8;
        var cbKeyword = kwRowTx.add("checkbox", undefined, "Contains:");
        cbKeyword.preferredSize.width = 90;
        var kwInput = kwRowTx.add("edittext", undefined, "");
        kwInput.preferredSize.width = 200;

        var btnGrp = dlg.add("group");
        btnGrp.alignment = "right";
        btnGrp.spacing = 8;
        var matchBtn = btnGrp.add("button", undefined, "Match");
        btnGrp.add("button", undefined, "Close", { name: "cancel" });

        var matchRequested = false;
        matchBtn.onClick = function () {
            if (!cbParaStyle.value && !cbFontSize.value && !cbLeading.value && !cbTracking.value &&
                    !(cbKeyword.value && kwInput.text !== "")) {
                alert("Please select at least one property to match against.");
                return;
            }
            matchRequested = true;
            dlg.close(1);
        };

        dlg.show();
        if (!matchRequested) return;

        var tfTolV = parseFloat(tolFontSize.text) || 0;
        var ldTolV = parseFloat(tolLeading.text)  || 0;
        var trTolV = parseFloat(tolTracking.text) || 0;
        var kwEnabled = cbKeyword.value && kwInput.text !== "";
        var kwStr = kwEnabled ? kwInput.text : "";

        mergeTargetTx.applyKeyword = false;
        mergeTargetTx.keyword = kwStr;
        mergeTargetTx.keywordReplacement = "";

        var activeProps = [];
        if (cbParaStyle.value) activeProps.push("Style: " + paraStyleRaw);
        if (cbFontSize.value)  activeProps.push(tfTolV > 0
            ? "Size: " + fontSizeRaw + " pt \u00b1 " + tfTolV + " pt \u2192 " + roundTo(parseFloat(fontSizeRaw) - tfTolV, 2) + "\u2013" + roundTo(parseFloat(fontSizeRaw) + tfTolV, 2) + " pt"
            : "Size: " + fontSizeRaw + " pt");
        if (cbLeading.value)   activeProps.push(ldTolV > 0
            ? "Leading: " + leadingRaw + " pt \u00b1 " + ldTolV + " pt \u2192 " + roundTo(parseFloat(leadingRaw) - ldTolV, 2) + "\u2013" + roundTo(parseFloat(leadingRaw) + ldTolV, 2) + " pt"
            : "Leading: " + leadingRaw + " pt");
        if (cbTracking.value)  activeProps.push(trTolV > 0
            ? "Tracking: " + trackingRaw + " \u00b1 " + trTolV + " \u2192 " + roundTo(parseFloat(trackingRaw) - trTolV, 2) + "\u2013" + roundTo(parseFloat(trackingRaw) + trTolV, 2)
            : "Tracking: " + trackingRaw);
        if (kwEnabled) activeProps.push("Keyword: \"" + kwStr + "\"");

        function paraMatchesFilters(tp) {
            try {
                if (!tp || !tp.isValid) return null;
                if (cbParaStyle.value && tp.appliedParagraphStyle.name !== paraStyleRaw) return null;
                var dev = {};
                if (cbFontSize.value) {
                    var d = roundTo(tp.pointSize, 2) - parseFloat(fontSizeRaw);
                    if (Math.abs(d) > tfTolV) return null;
                    dev.fontSize = roundTo(d, 2);
                }
                if (cbLeading.value) {
                    var d2 = roundTo(tp.leading, 2) - parseFloat(leadingRaw);
                    if (Math.abs(d2) > ldTolV) return null;
                    dev.leading = roundTo(d2, 2);
                }
                if (cbTracking.value) {
                    var d3 = tp.tracking - parseFloat(trackingRaw);
                    if (Math.abs(d3) > trTolV) return null;
                    dev.tracking = roundTo(d3, 2);
                }
                // Passive style observation — flag if style differs but wasn't a filter
                if (!cbParaStyle.value) {
                    try {
                        var candStyle = tp.appliedParagraphStyle.name;
                        if (candStyle !== paraStyleRaw) dev.styleDeviation = candStyle;
                    } catch (e) {}
                }
                return dev;
            } catch (e) { return null; }
        }

        function textMergeFn(match) {
            try {
                var tp = match.ref;
                if (!tp || !tp.isValid) return;
                if (mergeTargetTx.applySize)     { tp.pointSize = parseFloat(mergeTargetTx.fontSize); match.dev.fontSize  = 0; }
                if (mergeTargetTx.applyLeading)  { tp.leading   = parseFloat(mergeTargetTx.leading);  match.dev.leading   = 0; }
                if (mergeTargetTx.applyTracking) { tp.tracking  = parseFloat(mergeTargetTx.tracking); match.dev.tracking  = 0; }
                if (mergeTargetTx.applyStyle) {
                    var targetStyle = null;
                    var allPS = doc.allParagraphStyles;
                    for (var si = 0; si < allPS.length; si++) {
                        if (allPS[si].name === mergeTargetTx.style) { targetStyle = allPS[si]; break; }
                    }
                    if (targetStyle && targetStyle.isValid) {
                        tp.applyParagraphStyle(targetStyle, false);
                        match.dev.styleDeviation = "";
                    }
                }
                if (mergeTargetTx.applyKeyword && mergeTargetTx.keyword) {
                    applyKeywordReplace(tp, mergeTargetTx.keyword, mergeTargetTx.keywordReplacement);
                }
                match.entry = match.baseEntry + buildDevFlag(match.dev);
            } catch (e) {}
        }

        var matches = [];
        try {
            var stories = doc.stories;
            for (var si = 0; si < stories.length; si++) {
                try {
                    var paras = stories[si].paragraphs;
                    for (var pi = 0; pi < paras.length; pi++) {
                        try {
                            var tp = paras[pi];
                            var dev = paraMatchesFilters(tp);
                            if (!dev) continue;
                            if (kwEnabled) {
                                try { if (tp.contents.toLowerCase().indexOf(kwStr.toLowerCase()) === -1) continue; } catch (e2) { continue; }
                            }
                            var baseE = "p." + getParaPage(tp) + "  \u2014  \"" + trunc(tp.contents, 40) + "\"";
                            matches.push({
                                ref: tp, baseEntry: baseE,
                                entry: baseE + buildDevFlag(dev),
                                dev: dev, props: activeProps
                            });
                        } catch (e) {}
                    }
                    var tables = stories[si].tables;
                    for (var ti = 0; ti < tables.length; ti++) {
                        try {
                            var rows = tables[ti].rows;
                            for (var ri = 0; ri < rows.length; ri++) {
                                try {
                                    var cells = rows[ri].cells;
                                    for (var ci = 0; ci < cells.length; ci++) {
                                        try {
                                            var cellParas = cells[ci].paragraphs;
                                            for (var cpi = 0; cpi < cellParas.length; cpi++) {
                                                try {
                                                    var cp   = cellParas[cpi];
                                                    var devC = paraMatchesFilters(cp);
                                                    if (!devC) continue;
                                                    if (kwEnabled) {
                                                        try { if (cp.contents.toLowerCase().indexOf(kwStr.toLowerCase()) === -1) continue; } catch (e3) { continue; }
                                                    }
                                                    var baseCE = "p." + getParaPage(cp) + "  \u2014  [Table]  \"" + trunc(cp.contents, 35) + "\"";
                                                    matches.push({
                                                        ref: cp, baseEntry: baseCE,
                                                        entry: baseCE + buildDevFlag(devC),
                                                        dev: devC, props: activeProps
                                                    });
                                                } catch (e) {}
                                            }
                                        } catch (e) {}
                                    }
                                } catch (e) {}
                            }
                        } catch (e) {}
                    }
                } catch (e) {}
            }
        } catch (e) {}

        if (matches.length === 0) { alert("No matching paragraphs found."); return; }

        var adjFieldsTx = [
            { applyKey: "applyStyle",    valueKey: "style",    label: "Paragraph Style", unit: "",   type: "string", devKey: "styleDeviation", styleCollection: function () { return doc.allParagraphStyles; }, enabled: true },
            { applyKey: "applySize",     valueKey: "fontSize", label: "Font Size",        unit: "pt", type: "number", devKey: "fontSize",        minValue: 0,     enabled: true },
            { applyKey: "applyLeading",  valueKey: "leading",  label: "Leading",          unit: "pt", type: "number", devKey: "leading",         minValue: null,  enabled: true },
            { applyKey: "applyTracking", valueKey: "tracking", label: "Tracking",         unit: "u",  type: "number", devKey: "tracking",        minValue: null,  enabled: true }
        ];
        if (kwEnabled) adjFieldsTx.push({ applyKey: "applyKeyword", valueKey: "keywordReplacement", label: "Keyword replacement", type: "keywordReplace", keyword: kwStr, enabled: true });
        function textAdjustFn(devMatches) { return buildAdjustDialog(mergeTargetTx, adjFieldsTx, devMatches); }

        app.doScript(
            function () { runStepThrough(matches, navToPara, "paragraphs", textMergeFn, textAdjustFn); },
            ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Selection Inspector"
        );
        return;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CELL MODE
    // ══════════════════════════════════════════════════════════════════════════

    if (typeName === "Cell") {

        var cell = s0;

        var cellStyleRaw  = safeStr(function () { return cell.appliedCellStyle.name; });
        var cellHRaw      = safeStr(function () { return roundTo(cell.height, 3); });
        var cellWRaw      = safeStr(function () { return roundTo(cell.width,  3); });
        var cellTopRaw    = safeStr(function () { return roundTo(cell.topInset,    3); });
        var cellLeftRaw   = safeStr(function () { return roundTo(cell.leftInset,   3); });
        var cellRightRaw  = safeStr(function () { return roundTo(cell.rightInset,  3); });
        var cellBotRaw    = safeStr(function () { return roundTo(cell.bottomInset, 3); });
        var cellInsetsStr = "Top: " + cellTopRaw + "  Left: " + cellLeftRaw +
                            "  Right: " + cellRightRaw + "  Bottom: " + cellBotRaw;

        var mergeTargetC = {
            applyStyle:  false, style:  cellStyleRaw,
            applyHeight: false, height: cellHRaw,
            applyWidth:  false, width:  cellWRaw
        };

        var parentTableId  = -1;
        var parentTblStyle = "\u2014";
        var parentTblRows  = "\u2014";
        var parentTblCols  = "\u2014";
        try {
            var pt = cell.parent;
            parentTableId  = pt.id;
            parentTblStyle = pt.appliedTableStyle.name;
            parentTblRows  = String(pt.rows.length);
            parentTblCols  = String(pt.columns.length);
        } catch (e) {}

        var dlgC = new Window("dialog", "Selection Inspector");
        dlgC.alignChildren = ["fill", "top"];
        dlgC.margins = 18;
        dlgC.spacing = 12;
        dlgC.preferredSize.width = 420;

        var cellInfoSec = addSection(dlgC, "Cell");
        addRow(cellInfoSec, "Cell Style:", cellStyleRaw);
        addRow(cellInfoSec, "Height:",     cellHRaw + " pt");
        addRow(cellInfoSec, "Width:",      cellWRaw + " pt");
        addRow(cellInfoSec, "Insets:",     cellInsetsStr);

        var tblInfoSec = addSection(dlgC, "Parent Table");
        addRow(tblInfoSec, "Table Style:", parentTblStyle);
        addRow(tblInfoSec, "Rows:",        parentTblRows);
        addRow(tblInfoSec, "Columns:",     parentTblCols);

        var cellPropSec  = addSection(dlgC, "Match Properties");
        var cbCellStyle  = addCheckRow(cellPropSec,     "Cell Style  (" + cellStyleRaw + ")");
        var rCellH       = addToleranceRow(cellPropSec, "Height  (" + cellHRaw + " pt)", "pt");
        var rCellW       = addToleranceRow(cellPropSec, "Width  ("  + cellWRaw + " pt)", "pt");
        var cbCellInsets = addCheckRow(cellPropSec,     "Insets  [" + cellInsetsStr + "]");
        var cbCellH = rCellH.cb; var tolCellH = rCellH.tol;
        var cbCellW = rCellW.cb; var tolCellW = rCellW.tol;
        var kwRowC = cellPropSec.add("group");
        kwRowC.alignChildren = ["left", "center"];
        kwRowC.spacing = 8;
        var cbKeywordC = kwRowC.add("checkbox", undefined, "Contains:");
        cbKeywordC.preferredSize.width = 90;
        var kwInputC = kwRowC.add("edittext", undefined, "");
        kwInputC.preferredSize.width = 200;

        var btnGrpC = dlgC.add("group");
        btnGrpC.alignment = "right";
        btnGrpC.spacing = 8;
        var matchBtnC = btnGrpC.add("button", undefined, "Match");
        btnGrpC.add("button", undefined, "Close", { name: "cancel" });

        var matchRequestedC = false;
        matchBtnC.onClick = function () {
            if (!cbCellStyle.value && !cbCellH.value && !cbCellW.value && !cbCellInsets.value &&
                    !(cbKeywordC.value && kwInputC.text !== "")) {
                alert("Please select at least one property to match against.");
                return;
            }
            matchRequestedC = true;
            dlgC.close(1);
        };

        dlgC.show();
        if (!matchRequestedC) return;

        var chTolV = parseFloat(tolCellH.text) || 0;
        var cwTolV = parseFloat(tolCellW.text) || 0;
        var kwEnabledC = cbKeywordC.value && kwInputC.text !== "";
        var kwStrC = kwEnabledC ? kwInputC.text : "";
        mergeTargetC.applyKeyword = false;
        mergeTargetC.keyword = kwStrC;
        mergeTargetC.keywordReplacement = "";

        var activePropsC = [];
        if (cbCellStyle.value)  activePropsC.push("Cell Style: " + cellStyleRaw);
        if (cbCellH.value)      activePropsC.push(chTolV > 0
            ? "Height: " + cellHRaw + " pt \u00b1 " + chTolV + " pt \u2192 " + roundTo(parseFloat(cellHRaw) - chTolV, 2) + "\u2013" + roundTo(parseFloat(cellHRaw) + chTolV, 2) + " pt"
            : "Height: " + cellHRaw + " pt");
        if (cbCellW.value)      activePropsC.push(cwTolV > 0
            ? "Width: " + cellWRaw + " pt \u00b1 " + cwTolV + " pt \u2192 " + roundTo(parseFloat(cellWRaw) - cwTolV, 2) + "\u2013" + roundTo(parseFloat(cellWRaw) + cwTolV, 2) + " pt"
            : "Width: " + cellWRaw + " pt");
        if (cbCellInsets.value) activePropsC.push("Insets: [" + cellInsetsStr + "]");
        if (kwEnabledC) activePropsC.push("Keyword: \"" + kwStrC + "\"");

        function cellMergeFn(match) {
            try {
                var cand = match.ref;
                if (!cand || !cand.isValid) return;
                if (mergeTargetC.applyHeight) { cand.height = parseFloat(mergeTargetC.height); match.dev.cellHeight = 0; }
                if (mergeTargetC.applyWidth)  { cand.width  = parseFloat(mergeTargetC.width);  match.dev.cellWidth  = 0; }
                if (mergeTargetC.applyStyle) {
                    var targetCS = null;
                    var allCS = doc.allCellStyles;
                    for (var csi = 0; csi < allCS.length; csi++) {
                        if (allCS[csi].name === mergeTargetC.style) { targetCS = allCS[csi]; break; }
                    }
                    if (targetCS && targetCS.isValid) {
                        cand.appliedCellStyle = targetCS;
                        match.dev.styleDeviation = "";
                    }
                }
                if (mergeTargetC.applyKeyword && mergeTargetC.keyword) {
                    applyKeywordReplace(cand, mergeTargetC.keyword, mergeTargetC.keywordReplacement);
                }
                match.entry = match.baseEntry + buildDevFlag(match.dev);
            } catch (e) {}
        }

        var matchesC = [];
        try {
            var storiesC = doc.stories;
            for (var sci = 0; sci < storiesC.length; sci++) {
                try {
                    var tablesC = storiesC[sci].tables;
                    for (var tci = 0; tci < tablesC.length; tci++) {
                        try {
                            var tblC   = tablesC[tci];
                            var isSame = (tblC.id === parentTableId);
                            var rowsC  = tblC.rows;
                            for (var rci = 0; rci < rowsC.length; rci++) {
                                try {
                                    var cellsC = rowsC[rci].cells;
                                    for (var cci = 0; cci < cellsC.length; cci++) {
                                        try {
                                            var cand = cellsC[cci];
                                            if (!cand.isValid) continue;
                                            if (cbCellStyle.value && cand.appliedCellStyle.name !== cellStyleRaw) continue;
                                            var devCH = 0, devCW = 0;
                                            if (cbCellH.value) {
                                                devCH = roundTo(cand.height - parseFloat(cellHRaw), 2);
                                                if (Math.abs(devCH) > chTolV) continue;
                                            }
                                            if (cbCellW.value) {
                                                devCW = roundTo(cand.width - parseFloat(cellWRaw), 2);
                                                if (Math.abs(devCW) > cwTolV) continue;
                                            }
                                            if (cbCellInsets.value) {
                                                if (roundTo(cand.topInset,    3) !== parseFloat(cellTopRaw))   continue;
                                                if (roundTo(cand.leftInset,   3) !== parseFloat(cellLeftRaw))  continue;
                                                if (roundTo(cand.rightInset,  3) !== parseFloat(cellRightRaw)) continue;
                                                if (roundTo(cand.bottomInset, 3) !== parseFloat(cellBotRaw))   continue;
                                            }
                                            if (kwEnabledC) {
                                                try { if (cand.contents.toLowerCase().indexOf(kwStrC.toLowerCase()) === -1) continue; } catch (e2) { continue; }
                                            }
                                            var devCObj = {};
                                            if (cbCellH.value) devCObj.cellHeight = devCH;
                                            if (cbCellW.value) devCObj.cellWidth  = devCW;
                                            // Passive cell style observation
                                            if (!cbCellStyle.value) {
                                                try {
                                                    var candCStyle = cand.appliedCellStyle.name;
                                                    if (candCStyle !== cellStyleRaw) devCObj.styleDeviation = candCStyle;
                                                } catch (e) {}
                                            }
                                            var sameTag   = isSame ? "  [Same table]" : "";
                                            var pageNameC = getTablePage(tblC);
                                            var baseCE    = "p." + pageNameC + sameTag + "  \u2014  row " + (rci + 1) + ", col " + (cci + 1);
                                            matchesC.push({
                                                ref: cand, baseEntry: baseCE,
                                                entry: baseCE + buildDevFlag(devCObj),
                                                dev: devCObj, props: activePropsC
                                            });
                                        } catch (e) {}
                                    }
                                } catch (e) {}
                            }
                        } catch (e) {}
                    }
                } catch (e) {}
            }
        } catch (e) {}

        if (matchesC.length === 0) { alert("No matching cells found."); return; }

        var adjFieldsC = [
            { applyKey: "applyStyle",  valueKey: "style",  label: "Cell Style", unit: "",   type: "string", devKey: "styleDeviation", styleCollection: function () { return doc.allCellStyles; }, enabled: true },
            { applyKey: "applyHeight", valueKey: "height", label: "Height",     unit: "pt", type: "number", devKey: "cellHeight",      minValue: 0,     enabled: true },
            { applyKey: "applyWidth",  valueKey: "width",  label: "Width",      unit: "pt", type: "number", devKey: "cellWidth",       minValue: 0,     enabled: true }
        ];
        if (kwEnabledC) adjFieldsC.push({ applyKey: "applyKeyword", valueKey: "keywordReplacement", label: "Keyword replacement", type: "keywordReplace", keyword: kwStrC, enabled: true });
        function cellAdjustFn(devMatches) { return buildAdjustDialog(mergeTargetC, adjFieldsC, devMatches); }

        app.doScript(
            function () { runStepThrough(matchesC, navToCell, "cells", cellMergeFn, cellAdjustFn); },
            ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Selection Inspector"
        );
        return;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // TABLE MODE
    // ══════════════════════════════════════════════════════════════════════════

    if (typeName === "Table") {

        var table = s0;

        var tblStyleRaw = safeStr(function () { return table.appliedTableStyle.name; });
        var tblRowsRaw  = safeStr(function () { return table.rows.length; });
        var tblColsRaw  = safeStr(function () { return table.columns.length; });

        var tblColWidths = [];
        try {
            var refCols = table.columns;
            for (var rwi = 0; rwi < refCols.length; rwi++)
                tblColWidths.push(roundTo(refCols[rwi].width, 3));
        } catch (e) {}
        var tblColWidthsStr = tblColWidths.length ? tblColWidths.join(", ") + " pt" : "—";

        var tblDefaultRowH = "14.173";
        var tblDefaultColW = "56.693";
        try { tblDefaultRowH = String(roundTo(table.rows[0].height, 3)); } catch (e) {}
        try { tblDefaultColW = String(roundTo(table.columns[0].width, 3)); } catch (e) {}

        var mergeTargetTb = {
            applyStyle:       false, style:           tblStyleRaw,
            applyColWidths:   false, colWidths:        tblColWidths.slice(),
            applyRowCount:    false, targetRows:       tblRowsRaw,  defaultRowHeight: tblDefaultRowH,
            applyColCount:    false, targetCols:       tblColsRaw,  defaultColWidth:  tblDefaultColW
        };

        var dlgT = new Window("dialog", "Selection Inspector");
        dlgT.alignChildren = ["fill", "top"];
        dlgT.margins = 18;
        dlgT.spacing = 12;
        dlgT.preferredSize.width = 420;

        var tblInfoSecT = addSection(dlgT, "Table");
        addRow(tblInfoSecT, "Table Style:",    tblStyleRaw);
        addRow(tblInfoSecT, "Rows:",           tblRowsRaw);
        addRow(tblInfoSecT, "Columns:",        tblColsRaw);
        addRow(tblInfoSecT, "Column Widths:",  tblColWidthsStr);

        var tblPropSec = addSection(dlgT, "Match Properties");
        var cbTblStyle  = addCheckRow(tblPropSec,     "Table Style  (" + tblStyleRaw + ")");
        var cbTblRows   = addCheckRow(tblPropSec,     "Rows  ("  + tblRowsRaw + ")");
        var cbTblCols   = addCheckRow(tblPropSec,     "Columns  (" + tblColsRaw + ")");
        var rColWidths  = addToleranceRow(tblPropSec, "Column Widths  [" + tblColWidthsStr + "]", "pt");
        var cbColWidths = rColWidths.cb; var tolColWidths = rColWidths.tol;
        cbColWidths.enabled = (tblColWidths.length > 0);
        var kwRowT = tblPropSec.add("group");
        kwRowT.alignChildren = ["left", "center"];
        kwRowT.spacing = 8;
        var cbKeywordT = kwRowT.add("checkbox", undefined, "Contains:");
        cbKeywordT.preferredSize.width = 90;
        var kwInputT = kwRowT.add("edittext", undefined, "");
        kwInputT.preferredSize.width = 200;

        var btnGrpT = dlgT.add("group");
        btnGrpT.alignment = "right";
        btnGrpT.spacing = 8;
        var matchBtnT = btnGrpT.add("button", undefined, "Match");
        btnGrpT.add("button", undefined, "Close", { name: "cancel" });

        var matchRequestedT = false;
        matchBtnT.onClick = function () {
            if (!cbTblStyle.value && !cbTblRows.value && !cbTblCols.value && !cbColWidths.value &&
                    !(cbKeywordT.value && kwInputT.text !== "")) {
                alert("Please select at least one property to match against.");
                return;
            }
            matchRequestedT = true;
            dlgT.close(1);
        };

        dlgT.show();
        if (!matchRequestedT) return;

        var cwTolV = parseFloat(tolColWidths.text) || 0;
        var kwEnabledT = cbKeywordT.value && kwInputT.text !== "";
        var kwStrT = kwEnabledT ? kwInputT.text : "";
        mergeTargetTb.applyKeyword = false;
        mergeTargetTb.keyword = kwStrT;
        mergeTargetTb.keywordReplacement = "";

        var activePropsT = [];
        if (cbTblStyle.value)  activePropsT.push("Table Style: " + tblStyleRaw);
        if (cbTblRows.value)   activePropsT.push("Rows: "        + tblRowsRaw);
        if (cbTblCols.value)   activePropsT.push("Columns: "     + tblColsRaw);
        if (cbColWidths.value && tblColWidths.length > 0) activePropsT.push(cwTolV > 0
            ? "Column Widths: [" + tblColWidthsStr + "] ± " + cwTolV + " pt"
            : "Column Widths: [" + tblColWidthsStr + "]");
        if (kwEnabledT) activePropsT.push("Keyword: \"" + kwStrT + "\"");

        var matchesT = [];
        try {
            var storiesT = doc.stories;
            for (var sti = 0; sti < storiesT.length; sti++) {
                try {
                    var tablesT = storiesT[sti].tables;
                    for (var tti = 0; tti < tablesT.length; tti++) {
                        try {
                            var tblT = tablesT[tti];
                            if (!tblT.isValid) continue;
                            if (cbTblStyle.value && tblT.appliedTableStyle.name !== tblStyleRaw) continue;
                            if (cbTblRows.value  && String(tblT.rows.length)    !== tblRowsRaw)  continue;
                            if (cbTblCols.value  && String(tblT.columns.length) !== tblColsRaw)  continue;
                            if (cbColWidths.value && tblColWidths.length > 0) {
                                if (tblT.columns.length !== tblColWidths.length) continue;
                                var cwOk = true;
                                for (var cwfi2 = 0; cwfi2 < tblColWidths.length; cwfi2++) {
                                    if (Math.abs(roundTo(tblT.columns[cwfi2].width, 3) - tblColWidths[cwfi2]) > cwTolV) {
                                        cwOk = false; break;
                                    }
                                }
                                if (!cwOk) continue;
                            }
                            if (kwEnabledT) {
                                var kwFoundT = false;
                                var kwLowT = kwStrT.toLowerCase();
                                try {
                                    var kwCellsT = tblT.cells;
                                    for (var kwciT = 0; kwciT < kwCellsT.length; kwciT++) {
                                        try {
                                            var kwTxtT = kwCellsT[kwciT].contents;
                                            if (typeof kwTxtT === "string" && kwTxtT.toLowerCase().indexOf(kwLowT) !== -1) { kwFoundT = true; break; }
                                        } catch (e2) {}
                                    }
                                } catch (e3) {}
                                if (!kwFoundT) continue;
                            }
                            var baseTbE = "p." + getTablePage(tblT) + "  \u2014  " + tblT.rows.length + "\u00d7" + tblT.columns.length + "  [" + tblT.appliedTableStyle.name + "]";
                            var devTbl = {};
                            if (cbColWidths.value && tblColWidths.length > 0 && tblT.columns.length === tblColWidths.length) {
                                var colDevArr = [];
                                for (var cwdi = 0; cwdi < tblColWidths.length; cwdi++)
                                    colDevArr.push(roundTo(tblT.columns[cwdi].width - tblColWidths[cwdi], 2));
                                devTbl.colWidths = colDevArr;
                            }
                            // Passive table style observation
                            if (!cbTblStyle.value) {
                                try {
                                    var candTStyle = tblT.appliedTableStyle.name;
                                    if (candTStyle !== tblStyleRaw) devTbl.styleDeviation = candTStyle;
                                } catch (e) {}
                            }
                            matchesT.push({
                                ref: tblT, baseEntry: baseTbE,
                                entry: baseTbE + buildDevFlag(devTbl),
                                dev: devTbl, props: activePropsT
                            });
                        } catch (e) {}
                    }
                } catch (e) {}
            }
        } catch (e) {}

        if (matchesT.length === 0) { alert("No matching tables found."); return; }

        function tableMergeFn(match) {
            try {
                var tblT = match.ref;
                if (!tblT || !tblT.isValid) return;

                // Structural changes first (row/col count) so column indices are stable
                if (mergeTargetTb.applyRowCount) {
                    try {
                        var targetR = parseInt(mergeTargetTb.targetRows);
                        var defH    = parseFloat(mergeTargetTb.defaultRowHeight);
                        while (tblT.rows.length < targetR) {
                            var nr = tblT.rows.add(LocationOptions.AT_END);
                            try { nr.autoGrow = false; } catch (e2) {}
                            try { nr.height = defH; } catch (e2) {}
                        }
                        while (tblT.rows.length > targetR) {
                            tblT.rows[tblT.rows.length - 1].remove();
                        }
                    } catch (e3) {
                        alert("Could not adjust rows on p." + getTablePage(tblT) + ":\n" + e3.message);
                    }
                }

                if (mergeTargetTb.applyColCount) {
                    try {
                        var targetC = parseInt(mergeTargetTb.targetCols);
                        var defW    = parseFloat(mergeTargetTb.defaultColWidth);
                        while (tblT.columns.length < targetC) {
                            var nc = tblT.columns.add(LocationOptions.AT_END);
                            try { nc.width = defW; } catch (e4) {}
                        }
                        while (tblT.columns.length > targetC) {
                            tblT.columns[tblT.columns.length - 1].remove();
                        }
                    } catch (e5) {
                        alert("Could not adjust columns on p." + getTablePage(tblT) + ":\n" + e5.message);
                    }
                }

                if (mergeTargetTb.applyColWidths) {
                    var cols = tblT.columns;
                    for (var cwmi = 0; cwmi < mergeTargetTb.colWidths.length && cwmi < cols.length; cwmi++) {
                        cols[cwmi].width = parseFloat(mergeTargetTb.colWidths[cwmi]);
                    }
                }

                if (mergeTargetTb.applyStyle) {
                    var targetTS = null;
                    var allTS = doc.allTableStyles;
                    for (var tsi = 0; tsi < allTS.length; tsi++) {
                        if (allTS[tsi].name === mergeTargetTb.style) { targetTS = allTS[tsi]; break; }
                    }
                    if (targetTS && targetTS.isValid) {
                        tblT.appliedTableStyle = targetTS;
                    }
                }

                if (mergeTargetTb.applyKeyword && mergeTargetTb.keyword) {
                    applyKeywordReplace(tblT, mergeTargetTb.keyword, mergeTargetTb.keywordReplacement);
                }

                // Rebuild entry from live values — structural changes invalidate old dev arithmetic
                match.baseEntry = "p." + getTablePage(tblT) + "  —  " +
                    tblT.rows.length + "×" + tblT.columns.length +
                    "  [" + safeStr(function () { return tblT.appliedTableStyle.name; }) + "]";
                match.dev   = {};
                match.entry = match.baseEntry;
            } catch (e) {}
        }

        var adjFieldsTb = [
            { applyKey: "applyStyle",     valueKey: "style",      label: "Table Style",         unit: "",     type: "string",           devKey: "styleDeviation", styleCollection: function () { return doc.allTableStyles; }, enabled: true },
            { applyKey: "applyColWidths", valueKey: "colWidths",  label: "Column Widths",        unit: "pt",   type: "colWidths",        devKey: "colWidths",      enabled: tblColWidths.length > 0 },
            { applyKey: "applyRowCount",  valueKey: "targetRows", label: "Target Row Count",     unit: "rows", type: "countWithDefault", devKey: null,             enabled: true,
              sub: { valueKey: "defaultRowHeight", label: "Default row height" } },
            { applyKey: "applyColCount",  valueKey: "targetCols", label: "Target Column Count",  unit: "cols", type: "countWithDefault", devKey: null,             enabled: true,
              sub: { valueKey: "defaultColWidth",  label: "Default column width" } }
        ];
        if (kwEnabledT) adjFieldsTb.push({ applyKey: "applyKeyword", valueKey: "keywordReplacement", label: "Keyword replacement", type: "keywordReplace", keyword: kwStrT, enabled: true });
        function tableAdjustFn(selMatches) {
            if (!buildAdjustDialog(mergeTargetTb, adjFieldsTb, selMatches)) return false;

            // Pre-flight: collect tables that would lose rows or columns
            var destructive = [];
            for (var pfi = 0; pfi < selMatches.length; pfi++) {
                var pt = selMatches[pfi].ref;
                if (!pt || !pt.isValid) continue;
                var pfMsgs = [];
                if (mergeTargetTb.applyRowCount) {
                    var tr = parseInt(mergeTargetTb.targetRows);
                    if (!isNaN(tr) && pt.rows.length > tr)
                        pfMsgs.push("remove " + (pt.rows.length - tr) + " row(s)");
                }
                if (mergeTargetTb.applyColCount) {
                    var tc = parseInt(mergeTargetTb.targetCols);
                    if (!isNaN(tc) && pt.columns.length > tc)
                        pfMsgs.push("remove " + (pt.columns.length - tc) + " col(s)");
                }
                if (pfMsgs.length > 0)
                    destructive.push("p." + getTablePage(pt) + ": " + pfMsgs.join(", "));
            }

            if (destructive.length === 0) return true;

            // Custom Yes/No warning dialog
            var wDlg = new Window("dialog", "Confirm Destructive Changes");
            wDlg.alignChildren = ["fill", "top"];
            wDlg.margins = 18;
            wDlg.spacing = 10;
            wDlg.preferredSize.width = 400;
            wDlg.add("statictext", undefined, "The following tables will have content permanently deleted:");
            var wList = wDlg.add("listbox", undefined, destructive);
            wList.preferredSize = [380, Math.min(destructive.length * 22 + 10, 160)];
            var wNote = wDlg.add("statictext", undefined, "This cannot be undone within this merge operation.");
            wNote.graphics.foregroundColor = wNote.graphics.newPen(wNote.graphics.PenType.SOLID_COLOR, [0.7, 0, 0, 1], 1);
            var wBtns = wDlg.add("group");
            wBtns.alignment = "right";
            wBtns.spacing = 8;
            var wProceed = false;
            var wYes = wBtns.add("button", undefined, "Proceed");
            var wNo  = wBtns.add("button", undefined, "Cancel");
            wYes.onClick = function () { wProceed = true;  wDlg.close(); };
            wNo.onClick  = function () { wProceed = false; wDlg.close(); };
            wDlg.show();
            return wProceed;
        }

        app.doScript(
            function () { runStepThrough(matchesT, navToTable, "tables", tableMergeFn, tableAdjustFn); },
            ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Selection Inspector"
        );
        return;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // FRAME MODE
    // ══════════════════════════════════════════════════════════════════════════

    var frame     = s0;
    var frameType = safeStr(function () { return frame.constructor.name; });
    var objStyle  = safeStr(function () { return frame.appliedObjectStyle.name; });

    var rawW = null, rawH = null;
    var w = "\u2014", h = "\u2014";
    try {
        var b = frame.geometricBounds;
        rawW = roundTo(b[3] - b[1], 3);
        rawH = roundTo(b[2] - b[0], 3);
        w = rawW + " pt";
        h = rawH + " pt";
    } catch (e) {}

    var rawBaselineOffset    = null;
    var rawBaselineOffsetStr = "\u2014";
    var rawBaselineMin       = null;
    var rawBaselineMinStr    = "\u2014";
    try {
        var tfp = frame.textFramePreferences;
        rawBaselineOffset    = tfp.firstBaselineOffset;
        rawBaselineOffsetStr = baselineOffsetName(rawBaselineOffset);
        rawBaselineMin       = roundTo(tfp.minimumFirstBaselineOffset, 3);
        rawBaselineMinStr    = rawBaselineMin + " pt";
    } catch (e) {}

    var mergeTargetF = {
        applyStyle:       false, style:       objStyle,
        applyWidth:       false, width:       rawW !== null ? String(rawW) : null,
        applyHeight:      false, height:      rawH !== null ? String(rawH) : null,
        applyBaseline:    false, baseline:    rawBaselineOffset,
        applyBaselineMin: false, baselineMin: rawBaselineMin !== null ? String(rawBaselineMin) : null
    };

    var dlg2 = new Window("dialog", "Selection Inspector");
    dlg2.alignChildren = ["fill", "top"];
    dlg2.margins = 18;
    dlg2.spacing = 12;
    dlg2.preferredSize.width = 420;

    var infoSec2 = addSection(dlg2, "Frame");
    addRow(infoSec2, "Frame Type:",            frameType);
    addRow(infoSec2, "Object Style:",          objStyle);
    addRow(infoSec2, "Width:",                 w);
    addRow(infoSec2, "Height:",                h);
    addRow(infoSec2, "Baseline Offset:",       rawBaselineOffsetStr);
    addRow(infoSec2, "Baseline Minimum:",      rawBaselineMinStr);

    var propSec2      = addSection(dlg2, "Match Properties");
    var cbObjStyle    = addCheckRow(propSec2,     "Object Style  (" + objStyle + ")");
    var rWidth        = addToleranceRow(propSec2, "Width  ("  + w + ")", "pt");
    var rHeight       = addToleranceRow(propSec2, "Height  (" + h + ")", "pt");
    var cbBaselineOff = addCheckRow(propSec2,     "Baseline Offset  (" + rawBaselineOffsetStr + ")");
    var rBaselineMin  = addToleranceRow(propSec2, "Baseline Minimum  (" + rawBaselineMinStr + ")", "pt");
    var cbWidth    = rWidth.cb;  var tolWidth  = rWidth.tol;
    var cbHeight   = rHeight.cb; var tolHeight = rHeight.tol;
    var cbBaselineMin = rBaselineMin.cb; var tolBaselineMin = rBaselineMin.tol;

    // Disable baseline checkboxes if values were not accessible on this frame
    cbBaselineOff.enabled = (rawBaselineOffset !== null);
    cbBaselineMin.enabled = (rawBaselineMin    !== null);
    var kwRowF = propSec2.add("group");
    kwRowF.alignChildren = ["left", "center"];
    kwRowF.spacing = 8;
    var cbKeywordF = kwRowF.add("checkbox", undefined, "Contains:");
    cbKeywordF.preferredSize.width = 90;
    var kwInputF = kwRowF.add("edittext", undefined, "");
    kwInputF.preferredSize.width = 200;

    var btnGrp2 = dlg2.add("group");
    btnGrp2.alignment = "right";
    btnGrp2.spacing = 8;
    var matchBtn2 = btnGrp2.add("button", undefined, "Match");
    btnGrp2.add("button", undefined, "Close", { name: "cancel" });

    var matchRequested2 = false;
    matchBtn2.onClick = function () {
        if (!cbObjStyle.value && !cbWidth.value && !cbHeight.value &&
            !cbBaselineOff.value && !cbBaselineMin.value &&
            !(cbKeywordF.value && kwInputF.text !== "")) {
            alert("Please select at least one property to match against.");
            return;
        }
        matchRequested2 = true;
        dlg2.close(1);
    };

    dlg2.show();
    if (!matchRequested2) return;

    var wTolV      = parseFloat(tolWidth.text)       || 0;
    var hTolV      = parseFloat(tolHeight.text)      || 0;
    var bmTolV     = parseFloat(tolBaselineMin.text) || 0;
    var kwEnabledF = cbKeywordF.value && kwInputF.text !== "";
    var kwStrF = kwEnabledF ? kwInputF.text : "";
    mergeTargetF.applyKeyword = false;
    mergeTargetF.keyword = kwStrF;
    mergeTargetF.keywordReplacement = "";

    var activeProps2 = [];
    if (cbObjStyle.value)    activeProps2.push("Object Style: " + objStyle);
    if (cbWidth.value)       activeProps2.push(wTolV > 0
        ? "Width: " + w + " \u00b1 " + wTolV + " pt \u2192 " + roundTo(rawW - wTolV, 2) + "\u2013" + roundTo(rawW + wTolV, 2) + " pt"
        : "Width: " + w);
    if (cbHeight.value)      activeProps2.push(hTolV > 0
        ? "Height: " + h + " \u00b1 " + hTolV + " pt \u2192 " + roundTo(rawH - hTolV, 2) + "\u2013" + roundTo(rawH + hTolV, 2) + " pt"
        : "Height: " + h);
    if (cbBaselineOff.value) activeProps2.push("Baseline Offset: " + rawBaselineOffsetStr);
    if (cbBaselineMin.value) activeProps2.push(bmTolV > 0
        ? "Baseline Minimum: " + rawBaselineMinStr + " \u00b1 " + bmTolV + " pt \u2192 " + roundTo(rawBaselineMin - bmTolV, 2) + "\u2013" + roundTo(rawBaselineMin + bmTolV, 2) + " pt"
        : "Baseline Minimum: " + rawBaselineMinStr);
    if (kwEnabledF) activeProps2.push("Keyword: \"" + kwStrF + "\"");

    function frameMergeFn(match) {
        try {
            var item = match.ref;
            if (!item || !item.isValid) return;
            var bnds;
            if (mergeTargetF.applyWidth && mergeTargetF.width !== null) {
                bnds = item.geometricBounds;
                item.geometricBounds = [bnds[0], bnds[1], bnds[2], bnds[1] + parseFloat(mergeTargetF.width)];
                match.dev.width = 0;
            }
            if (mergeTargetF.applyHeight && mergeTargetF.height !== null) {
                bnds = item.geometricBounds;
                item.geometricBounds = [bnds[0], bnds[1], bnds[0] + parseFloat(mergeTargetF.height), bnds[3]];
                match.dev.height = 0;
            }
            if (mergeTargetF.applyBaseline && mergeTargetF.baseline !== null) {
                try { item.textFramePreferences.firstBaselineOffset = mergeTargetF.baseline; } catch (e2) {}
            }
            if (mergeTargetF.applyBaselineMin && mergeTargetF.baselineMin !== null) {
                try { item.textFramePreferences.minimumFirstBaselineOffset = parseFloat(mergeTargetF.baselineMin); } catch (e3) {}
            }
            if (mergeTargetF.applyStyle) {
                var targetOS = null;
                var allOS = doc.allObjectStyles;
                for (var osi = 0; osi < allOS.length; osi++) {
                    if (allOS[osi].name === mergeTargetF.style) { targetOS = allOS[osi]; break; }
                }
                if (targetOS && targetOS.isValid) {
                    item.appliedObjectStyle = targetOS;
                    match.dev.styleDeviation = "";
                }
            }
            if (mergeTargetF.applyKeyword && mergeTargetF.keyword) {
                applyKeywordReplace(item, mergeTargetF.keyword, mergeTargetF.keywordReplacement);
            }
            match.entry = match.baseEntry + buildDevFlag(match.dev);
        } catch (e) {}
    }

    var matches2 = [];
    try {
        var allItems = doc.allPageItems;
        for (var ii = 0; ii < allItems.length; ii++) {
            try {
                var item = allItems[ii];
                if (!item.isValid) continue;
                if (cbObjStyle.value && item.appliedObjectStyle.name !== objStyle) continue;
                var iw = 0, ih = 0;
                if (cbWidth.value || cbHeight.value) {
                    var ib = item.geometricBounds;
                    iw = roundTo(ib[3] - ib[1], 3);
                    ih = roundTo(ib[2] - ib[0], 3);
                    if (cbWidth.value  && Math.abs(iw - rawW) > wTolV) continue;
                    if (cbHeight.value && Math.abs(ih - rawH) > hTolV) continue;
                }
                // Baseline exact-match filters
                if (cbBaselineOff.value && rawBaselineOffset !== null) {
                    try { if (item.textFramePreferences.firstBaselineOffset !== rawBaselineOffset) continue; }
                    catch (e) { continue; }
                }
                if (cbBaselineMin.value && rawBaselineMin !== null) {
                    try { if (Math.abs(roundTo(item.textFramePreferences.minimumFirstBaselineOffset, 3) - rawBaselineMin) > bmTolV) continue; }
                    catch (e) { continue; }
                }
                if (kwEnabledF) {
                    var kwFoundF = false;
                    try {
                        var kwStoryF = item.parentStory;
                        if (kwStoryF && kwStoryF.isValid) {
                            var kwTxtF = kwStoryF.contents;
                            if (typeof kwTxtF === "string" && kwTxtF.toLowerCase().indexOf(kwStrF.toLowerCase()) !== -1) kwFoundF = true;
                        }
                    } catch (e2) {}
                    if (!kwFoundF) continue;
                }
                var devF = {};
                if (cbWidth.value)  devF.width  = roundTo(iw - rawW, 2);
                if (cbHeight.value) devF.height = roundTo(ih - rawH, 2);
                // Passive object style observation
                if (!cbObjStyle.value) {
                    try {
                        var candOStyle = item.appliedObjectStyle.name;
                        if (candOStyle !== objStyle) devF.styleDeviation = candOStyle;
                    } catch (e) {}
                }
                var itemType  = item.constructor ? item.constructor.name : "Frame";
                var itemStyle = safeStr(function () { return item.appliedObjectStyle.name; });
                var baseFE    = "p." + getItemPage(item) + "  \u2014  " + itemType + " [" + itemStyle + "]";
                matches2.push({
                    ref: item, baseEntry: baseFE,
                    entry: baseFE + buildDevFlag(devF),
                    dev: devF, props: activeProps2
                });
            } catch (e) {}
        }
    } catch (e) {}

    if (matches2.length === 0) { alert("No matching frames found."); return; }

    var adjFieldsF = [
        { applyKey: "applyStyle",       valueKey: "style",       label: "Object Style",     unit: "",   type: "string", devKey: "styleDeviation", styleCollection: function () { return doc.allObjectStyles; }, enabled: true },
        { applyKey: "applyWidth",       valueKey: "width",       label: "Width",            unit: "pt", type: "number", devKey: "width",          minValue: 0,    enabled: rawW !== null },
        { applyKey: "applyHeight",      valueKey: "height",      label: "Height",           unit: "pt", type: "number", devKey: "height",         minValue: 0,    enabled: rawH !== null },
        { applyKey: "applyBaseline",    valueKey: "baseline",    label: "Baseline Offset",  unit: "",   type: "enum",   devKey: null,             displayStr: rawBaselineOffsetStr, enabled: rawBaselineOffset !== null },
        { applyKey: "applyBaselineMin", valueKey: "baselineMin", label: "Baseline Minimum", unit: "pt", type: "number", devKey: null,             minValue: null, enabled: rawBaselineMin !== null }
    ];
    if (kwEnabledF) adjFieldsF.push({ applyKey: "applyKeyword", valueKey: "keywordReplacement", label: "Keyword replacement", type: "keywordReplace", keyword: kwStrF, enabled: true });
    function frameAdjustFn(devMatches) { return buildAdjustDialog(mergeTargetF, adjFieldsF, devMatches); }

    app.doScript(
        function () { runStepThrough(matches2, navToItem, "frames", frameMergeFn, frameAdjustFn); },
        ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Selection Inspector"
    );

})();
