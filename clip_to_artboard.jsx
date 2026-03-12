// clip_to_artboard.jsx
// Clips all content on the active layer to the active artboard bounds.
// Run via Illustrator: File > Scripts > Other Script

(function () {
    var doc = app.activeDocument;
    var layer = doc.activeLayer;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    // ab: [left, top, right, bottom]

    var r = layer.pathItems.rectangle(ab[1], ab[0], ab[2] - ab[0], ab[1] - ab[3]);
    r.filled = false;
    r.stroked = false;
    r.zOrder(ZOrderMethod.BRINGTOFRONT);

    app.executeMenuCommand('selectall');
    app.executeMenuCommand('makeMask');
})();
