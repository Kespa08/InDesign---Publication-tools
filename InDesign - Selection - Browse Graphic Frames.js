// Q: Why is the whole script wrapped in (function () { ... })(); ?
// A: This is a common pattern called an "immediately invoked function
//    expression." Writing "(function () { ... })()" defines a function
//    and calls it in the same breath. Everything declared inside stays
//    private to this script instead of leaking into InDesign's shared
//    scripting environment, where it could clash with variables from
//    some other script. You'll see this same shape at the top of every
//    script in this repo.
(function () {

    // Q: What are "app" and "doc"?
    // A: "app" is provided automatically by InDesign -- it represents the
    //    running application itself. "doc" is just a shorthand variable we
    //    create, pointing at app.activeDocument (whichever document is
    //    currently open and in front). Every line below reads "doc"
    //    instead of retyping "app.activeDocument" each time.
    var doc = app.activeDocument;

    // Q: Why loop through doc.allGraphics instead of looking at pages
    //    directly?
    // A: doc.allGraphics is a ready-made list of every placed picture
    //    (image, EPS, PDF, etc.) anywhere in the whole document, no
    //    matter which page it's on or whether it's a normal placed image
    //    or one anchored inline inside a paragraph of text. That saves us
    //    from having to separately walk every page AND every story to
    //    make sure nothing is missed.
    //
    // Q: What does ".parent" mean on line 20?
    // A: A placed picture always sits INSIDE a frame -- think of the
    //    frame as the box, and the picture as the photo inside the box.
    //    When you click an image on an InDesign page, you're actually
    //    selecting the box (the frame), not the raw picture data. So
    //    ".parent" on a graphic gives us that containing box, which is
    //    the thing this script actually needs to select and jump to.
    //
    // Q: What is try { ... } catch (e) { } doing here, and everywhere else
    //    in this script?
    // A: "try" means "attempt to run this code." If anything inside goes
    //    wrong (for example, a frame that's somehow already been deleted),
    //    "catch" quietly steps in instead of crashing the entire script.
    //    The "(e)" is just a name for whatever error was caught -- we
    //    don't use it here, we just want the failure contained to this
    //    one item so the loop can carry on to the next one.
    var allGraphics = doc.allGraphics;
    var frames = [];
    for (var g = 0; g < allGraphics.length; g++) {
        try {
            var frame = allGraphics[g].parent;
            if (frame && frame.isValid) frames.push(frame);
        } catch (e) {}
    }

    // Q: Why check frames.length === 0 before doing anything else?
    // A: If the document has no images at all, there's nothing to browse.
    //    Popping up a dialog with no content to show would just be
    //    confusing, so we alert the user and "return" -- which means
    //    "stop running this function right now" -- before we ever try to
    //    build the dialog.
    if (frames.length === 0) {
        alert("No image/graphic frames found in this document.");
        return;
    }

    // Q: What does this function actually do, in plain terms?
    // A: Three steps, one per line: (1) turn to the page the frame lives
    //    on, (2) select the frame, just like clicking it with the mouse,
    //    (3) zoom the view so the whole page is visible and the selected
    //    frame is easy to see. It's wrapped in its own try/catch (and a
    //    second, inner one just around the zoom step) because zooming can
    //    occasionally fail in ways that shouldn't stop the selection from
    //    having worked.
    //
    // Q: Why is this copied from another script instead of written fresh?
    // A: "InDesign - Selection - Match and merge - 19.js" already has this
    //    exact "select it and move the view to it" logic, tested and
    //    working. Reusing it instead of writing a new version means one
    //    less thing that could have a new, different bug.
    function navToItem(item) {
        try {
            var page = item.parentPage;
            if (!page || !page.isValid) return;
            app.activeWindow.activePage = page;
            app.select(item);
            try { app.activeWindow.zoom(ZoomOptions.fitPage); } catch (e2) {}
        } catch (e) {}
    }

    // Q: What is "idx"?
    // A: Short for "index" -- a plain number tracking which frame in our
    //    list we're currently looking at. It starts at 0 because, like
    //    almost everywhere in programming, the first item in a list is
    //    counted as position 0, not position 1.
    var idx = 0;

    // Q: What is "Window", and what's a "dialog"?
    // A: ScriptUI is InDesign's built-in toolkit for building small pop-up
    //    windows with text and buttons, without needing to know any real
    //    UI programming. "new Window('dialog', title)" creates a modal
    //    pop-up box -- "modal" means it takes over until it's closed, so
    //    the rest of the script pauses and waits for you to interact with
    //    it. Every script in this repo that shows a pop-up uses this same
    //    starting line.
    var dlg = new Window("dialog", "Browse Graphic Frames");
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 18;
    dlg.spacing = 12;

    // Q: What is "statictext"?
    // A: The ScriptUI term for a plain, non-editable line of text inside a
    //    dialog -- just a label. We'll update its wording every time you
    //    click Previous/Next.
    var statusLabel = dlg.add("statictext", undefined, "");
    statusLabel.alignment = "center";

    // Q: Why is this its own separate function instead of just being
    //    written inline wherever we need it?
    // A: We need to set this exact same label text in three different
    //    places below (Previous, Next, and once before the dialog even
    //    opens). Writing it once as a function and calling
    //    "updateStatus()" each time avoids repeating the same line three
    //    times -- if we ever want to change the wording, there's only one
    //    place to edit.
    function updateStatus() {
        statusLabel.text = "Showing " + (idx + 1) + " of " + frames.length;
    }

    // Q: What is a "group"?
    // A: A simple container for laying other controls (like buttons) out
    //    next to each other in a row, instead of stacked vertically.
    var navGrp = dlg.add("group");
    navGrp.alignment = "center";
    navGrp.spacing = 8;
    var prevBtn = navGrp.add("button", undefined, "← Previous");
    var nextBtn = navGrp.add("button", undefined, "Next →");

    var doneGrp = dlg.add("group");
    doneGrp.alignment = "right";
    var doneBtn = doneGrp.add("button", undefined, "Done");

    // Q: What does the % symbol do in "(idx + 1) % frames.length"?
    // A: % is the "modulo" operator -- it gives you the remainder left
    //    over after division. For example, if there are 5 frames
    //    (frames.length is 5) and idx is currently 4 (the last frame,
    //    since counting starts at 0), then idx + 1 is 5, and
    //    5 % 5 is 0 -- the remainder of 5 divided by 5 is 0. So instead of
    //    counting past the end of the list, it wraps back around to 0,
    //    the first frame. The "+ frames.length" in the Previous handler
    //    below does the same trick in the other direction, so subtracting
    //    1 from position 0 wraps around to the last frame instead of
    //    going negative.
    //
    // Q: What is ".onClick = function () { ... }" ?
    // A: This assigns a function to run whenever that specific button is
    //    clicked -- InDesign calls it an "event handler." Nothing inside
    //    it runs immediately when the script starts; it only runs later,
    //    at the moment you actually click that button.
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

    // Q: Why doesn't Done do anything else yet?
    // A: This script is deliberately a first, simple step -- just
    //    browsing, no edits. "dlg.close()" just dismisses the pop-up;
    //    later, this is the exact spot where any future action (applying
    //    a change to whichever frame you last viewed, for example) would
    //    be added.
    doneBtn.onClick = function () {
        dlg.close();
    };

    // Q: Why call updateStatus() and navToItem() here, before dlg.show()?
    // A: So the dialog doesn't open on a blank state -- by the time it
    //    appears, the viewport has already jumped to the first frame and
    //    the label already reads "Showing 1 of N," instead of waiting for
    //    you to click Next once just to see anything happen.
    updateStatus();
    navToItem(frames[0]);

    // Q: Why is dlg.show() the very last line?
    // A: This is the line that actually displays the dialog on screen.
    //    Because a dialog is "modal" (see the note above), this line also
    //    pauses the script right here -- everything after it won't run
    //    until you close the dialog (either with Done or the window's own
    //    close button). There's nothing after it in this script, since
    //    once you're done browsing, the script's job is finished.
    dlg.show();
})();
