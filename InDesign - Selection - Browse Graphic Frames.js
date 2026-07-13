(function () {
    var doc = app.activeDocument;

    // Every placed graphic (image, EPS, PDF, etc.) anywhere in the
    // document, regardless of whether its containing frame is a
    // standalone page item or an anchored inline object. Each graphic's
    // .parent is the frame that holds it -- that's what gets selected
    // and navigated to, not the graphic itself.
    var allGraphics = doc.allGraphics;
    var frames = [];
    for (var g = 0; g < allGraphics.length; g++) {
        try {
            var frame = allGraphics[g].parent;
            if (frame && frame.isValid) frames.push(frame);
        } catch (e) {}
    }

    if (frames.length === 0) {
        alert("No image/graphic frames found in this document.");
        return;
    }

    // Selects the frame and moves the viewport to it. Adapted verbatim
    // from navToItem() in "InDesign - Selection - Match and merge - 19.js".
    function navToItem(item) {
        try {
            var page = item.parentPage;
            if (!page || !page.isValid) return;
            app.activeWindow.activePage = page;
            app.select(item);
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    var idx = 0;

    var dlg = new Window("dialog", "Browse Graphic Frames");
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 18;
    dlg.spacing = 12;

    var statusLabel = dlg.add("statictext", undefined, "");
    statusLabel.alignment = "center";

    function updateStatus() {
        statusLabel.text = "Showing " + (idx + 1) + " of " + frames.length;
    }

    var navGrp = dlg.add("group");
    navGrp.alignment = "center";
    navGrp.spacing = 8;
    var prevBtn = navGrp.add("button", undefined, "← Previous");
    var nextBtn = navGrp.add("button", undefined, "Next →");

    var doneGrp = dlg.add("group");
    doneGrp.alignment = "right";
    var doneBtn = doneGrp.add("button", undefined, "Done");

    // Wraps at both ends -- a deliberate departure from the clamp-at-the-
    // ends behavior used in Match-and-merge's step-through dialog, since
    // this is a pure "browse everything" tool rather than a narrowed
    // candidate list.
    prevBtn.onClick = function () {
        idx = (idx - 1 + frames.length) % frames.length;
        navToItem(frames[idx]);
        updateStatus();
    };

    nextBtn.onClick = function () {
        idx = (idx + 1) % frames.length;
        navToItem(frames[idx]);
        updateStatus();
    };

    doneBtn.onClick = function () {
        dlg.close();
    };

    updateStatus();
    navToItem(frames[0]);

    dlg.show();
})();
