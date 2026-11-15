/*

	Add Labels v4 (2026)

	(c) Dan Ichimescu 2023-2026 (constantindan@gmail.com)
	very modified version of:

	
	Adds or clears labels used by Batch processing automated resizing.jsx.
	Released under MIT License:
	https://choosealicense.com/licenses/mit/

	27 oct 2026 - v3:
	- all labels from Redimensionari_multiMasters_v22.jsx (labelActions) are now available:
	  page edges, visible area, bleed (all sides, 14, 23, 42, 123, 234, smart), HW / DTI / line,
	  Vodafone logo left/right, key visual, 20%, headline center, vertical center -HW, align to key object
	- new interface: buttons are placed like on a page in InDesign
	    red   = bleed (expand object to bleed)
	    white = page edges (trim)
	    green = visible area (margins)
	    grey  = Health Warning / DTI strip at the bottom
	- "Combine with existing label" on = add the label with "+", off = replace the label
	- the same label is not added twice
	- Health_Warning and DeTeIi always replace the label (the resizing script checks them as exact label)
	- "HW logo height %" writes "hw<number>" on the child rectangle of HW (e.g. hw60 = 60% of the HW band)
	- the label of the selection is shown live, can be edited and set by hand
	- "Remove last" removes only the last label from a combined label
	- palette stays open (own targetengine), InDesign preferences are restored when the palette is closed
	- fix: removed arrow characters from buttons (InDesign showed them garbled), file is plain ASCII
	- fix: zone titles are drawn by hand in black/white (statictext was white on light background)
	- "Top+Right+Bottom" renamed "Except left" (text was cut)
	- bleed buttons renamed "Expand to <sides>" with numbers 1 = top, 2 = right, 3 = bottom, 4 = left
	  (same numbers as the expandToBleed labels), legend in the BLEED title

	13 nov 2026:
	- new button "Expand to 4+1+2" (label "expandToBleed412"), top-middle of the BLEED area:
	  left, top and right edges expanded to bleed, bottom edge stays in place

*/

#targetengine "addLabelsForResizing"

// types processed by aliniereElemente in the resizing script
var LABEL_TYPES = { "Rectangle": 1, "Oval": 1, "Polygon": 1, "TextFrame": 1, "Group": 1, "Button": 1, "GraphicLine": 1 };
// labels checked with === in the resizing script, cannot be combined
var EXCLUSIVE_LABELS = { "Health_Warning": 1, "DeTeIi": 1 };

var BW = 108; // button width
var BH = 24;  // button height
var BBW = 130; // bleed button width

if (typeof addLabelsUI !== "undefined" && addLabelsUI instanceof Window) {
	// already open: bring it to front instead of opening a second palette
	addLabelsUI.show();
	addLabelsUI.active = true;
} else {
	var oldURL = app.generalPreferences.ungroupRemembersLayers;
	var oldPRL = app.clipboardPreferences.pasteRemembersLayers;
	app.generalPreferences.ungroupRemembersLayers = true;
	app.clipboardPreferences.pasteRemembersLayers = true;

	var addLabelsUI = CreateUI();
	addLabelsUI.show();
	updateStatus();
}


function CreateUI() {
	var ui = new Window("palette", "Add labels for automated resizing", undefined, { resizeable: false });
	ui.orientation = "column";
	ui.alignChildren = ["fill", "top"];
	ui.spacing = 6;
	ui.margins = 10;

	ui.add("statictext", undefined, "Select object(s) in the master, then click the place where the object must go in the resized file.");

	var gMain = ui.add("group");
	gMain.orientation = "row";
	gMain.alignChildren = ["left", "top"];

	// ================= BLEED (red) =================
	var gBleed = area(gMain, [0.86, 0.55, 0.55], "BLEED  -  expand object to bleed:   1 = top,  2 = right,  3 = bottom,  4 = left", [0, 0, 0]);

	var bRow1 = row(gBleed);
	btn(bRow1, "Expand to 1+4", "expandToBleed14", "Expand to bleed on top and left side.\nRight and bottom edges stay in place.", BBW);
	spacer(bRow1, BW);
	btn(bRow1, "Expand to 4+1+2", "expandToBleed412", "Expand to bleed on left, top and right side.\nBottom edge stays in place.", BBW);
	spacer(bRow1, BW);
	btn(bRow1, "Expand to 1+2+3+4", "expandToBleed", "Expand to bleed on all 4 sides (whole page + bleed).", BBW);

	var bRow2 = row(gBleed);
	btn(bRow2, "Expand to 4+2", "expandToBleed42", "Expand to bleed on left and right side.\nTop and bottom edges stay in place.", BBW);

	// ================= PAGE (white) =================
	var gPage = area(bRow2, [0.96, 0.96, 0.96], "PAGE (trim)  -  align to page edges", [0, 0, 0]);

	btn(bRow2, "Expand to 1+2+3", "expandToBleed123", "Expand to bleed on top, right and bottom.\nLeft edge stays in place.", BBW);

	var pRow1 = row(gPage);
	spacer(pRow1, BW);
	btn(pRow1, "Page top", "alignToPageTop", "Align top edge of object to top edge of the page.");
	spacer(pRow1, BW);

	var pRow2 = row(gPage);
	btn(pRow2, "Page left", "alignToPageLeft", "Align left edge of object to left edge of the page.");

	// ================= VISIBLE (green) =================
	var gVis = area(pRow2, [0.74, 0.88, 0.74], "VISIBLE AREA (margins)", [0, 0, 0]);

	btn(pRow2, "Page right", "alignToPageRight", "Align right edge of object to right edge of the page.");

	var vRow1 = row(gVis);
	btn(vRow1, "Top left", "alignTo_Visible_Corner_leftTop", "Align to top-left corner of the visible area.");
	btn(vRow1, "Top", "alignTo_Visible_Top", "Align to top of the visible area.");
	btn(vRow1, "Top right", "alignTo_Visible_Corner_RightTop", "Align to top-right corner of the visible area.");

	var vRow2 = row(gVis);
	btn(vRow2, "Left", "alignTo_Visible_Left", "Align to left of the visible area.");
	var vCenter = vRow2.add("group");
	vCenter.orientation = "column";
	vCenter.spacing = 4;
	btn(vCenter, "Center H", "alignHorizontalCenter", "Center horizontally in the visible area.");
	btn(vCenter, "Fill visible", "expandToVizibil", "Object takes exactly the size of the visible area.");
	btn(vRow2, "Right", "alignTo_Visible_Right", "Align to right of the visible area.");

	var vRow3 = row(gVis);
	btn(vRow3, "Bottom left", "alignTo_Visible_Corner_leftBottom-HW", "Align to bottom-left corner of the visible area,\nabove Health Warning / DTI (if they exist).");
	btn(vRow3, "Bottom", "alignTo_Visible_Bottom-HW", "Align to bottom of the visible area,\nabove Health Warning / DTI (if they exist).");
	btn(vRow3, "Bottom right", "alignTo_Visible_Corner_rightBottom-HW", "Align to bottom-right corner of the visible area,\nabove Health Warning / DTI (if they exist).");

	var vRow4 = row(gVis);
	btn(vRow4, "VDF logo left", "vdf_stanga", "Vodafone logo bottom-left of visible area.\nHeight = 4% of visible diagonal, safe space = 50% of logo height.\nAdds safe area guides.");
	spacer(vRow4, BW);
	btn(vRow4, "VDF logo right", "vdf_dreapta", "Vodafone logo bottom-right of visible area.\nHeight = 4% of visible diagonal, safe space = 50% of logo height.\nAdds safe area guides.");

	// ================= HW / DTI strip (grey) =================
	var gHw = area(gVis, [0.38, 0.38, 0.38], "HEALTH WARNING / DTI  (bottom of visible, to bleed left/right)", [1, 1, 1]);
	var hRow1 = row(gHw);
	btn(hRow1, "DTI 5%", "DeTeIi", "DTI band = 5% of visible height (above HW).\nObject must be a rectangle with the linked file in a child rectangle.\nReplaces the existing label.");
	btn(hRow1, "Line", "linia", "Line between DTI and HW, from bleed to bleed.");
	btn(hRow1, "HW 10%", "Health_Warning", "Health Warning band = 10% of visible height.\nObject must be a rectangle with the linked file in a child rectangle.\nReplaces the existing label.");

	var hRow2 = row(gHw);
	caption(hRow2, "HW logo height % (select the child rectangle):", [1, 1, 1]);
	var etHwPct = hRow2.add("edittext", undefined, "60");
	etHwPct.characters = 4;
	var btnHwPct = hRow2.add("button", undefined, "Set");
	btnHwPct.preferredSize = [50, BH];
	btnHwPct.helpTip = "Writes label hw<number> on the selected child rectangle of HW\n(e.g. hw60 = child is 60% of the HW band height).";

	var pRow3 = row(gPage);
	spacer(pRow3, BW);
	btn(pRow3, "Page bottom", "alignToPageBottom", "Align bottom edge of object to bottom edge of the page.");
	spacer(pRow3, BW);

	var bRow3 = row(gBleed);
	btn(bRow3, "Expand to 2+3+4", "expandToBleed234", "Expand to bleed on right, bottom and left.\nTop edge stays in place.", BBW);
	spacer(bRow3, BW);
	btn(bRow3, "Smart bleed", "smartBleed", "Rectangles only: every side closer than 5% to the page edge\n(or already outside) is expanded to bleed.", BBW);
	spacer(bRow3, BW);
	btn(bRow3, "Expand to 2+3", "expandToBleed23", "Expand to bleed on right and bottom side.\nTop and left edges stay in place.", BBW);

	// ================= SPECIAL =================
	var pSpecial = gMain.add("panel", undefined, "Special");
	pSpecial.orientation = "column";
	pSpecial.alignChildren = ["fill", "top"];
	pSpecial.margins = [10, 15, 10, 10];
	btn(pSpecial, "Key visual", "cheia", "Scale and center the key visual in the visible area above HW:\nportrait = 80% of visible width, landscape = 85% of visible height.\nAdds a guide in the middle of it.", 150);
	btn(pSpecial, "20% at bottom", "douzecilasuta", "Scale object to 20% of visible height and align to bottom of visible.", 150);
	btn(pSpecial, "Vertical center -HW", "alignVerticalCentersMinusHW", "Center vertically on the page height minus the Health Warning.", 150);
	btn(pSpecial, "Headline center", "hl_centerToLeft", "Keep the left edge, extend the frame to the right margin\nand center the content between them.", 150);
	btn(pSpecial, "Align to key object", "alignedTo_key_first", "Align left edge to the rectangle named \"firstreferenceobject\"\ninside the group named \"myGroup\". Runs after all other labels.", 150);

	// ================= label of selection + actions =================
	var pStatus = ui.add("panel", undefined, "Label of selection");
	pStatus.orientation = "column";
	pStatus.alignChildren = ["fill", "top"];
	pStatus.margins = [10, 15, 10, 10];

	var gStatus = pStatus.add("group");
	gStatus.alignChildren = ["left", "center"];
	ui.labelField = gStatus.add("edittext", undefined, "");
	ui.labelField.characters = 80;
	ui.labelField.helpTip = "Label of the selected object(s). You can edit it and click \"Set label\".";
	var btnSet = gStatus.add("button", undefined, "Set label");
	btnSet.helpTip = "Write the text from the field as label on the selected object(s) (replaces).";

	var gActions = pStatus.add("group");
	gActions.alignChildren = ["left", "center"];
	ui.combine = gActions.add("checkbox", undefined, "Combine with existing label (+)");
	ui.combine.value = true;
	ui.combine.helpTip = "On = add the new label to the existing one with \"+\" (e.g. alignTo_Visible_Top+smartBleed).\nOff = replace the existing label.";
	var btnRemoveLast = gActions.add("button", undefined, "Remove last");
	btnRemoveLast.helpTip = "Remove only the last label from the selected object(s).";
	var btnClear = gActions.add("button", undefined, "Clear label/s");
	btnClear.helpTip = "Clear labels of the selection.\nWith nothing selected: clear ALL labels from the document.";
	var btnClose = gActions.add("button", undefined, "Close");

	// ================= behaviour =================
	btnHwPct.onClick = function () {
		var n = Number(String(etHwPct.text).replace(/,/g, "."));
		if (isNaN(n) || n <= 0 || n > 100) { alert("HW logo height must be a number between 1 and 100."); return; }
		setLabels(function () { return "hw" + n; }, "Label: hw" + n);
	};

	btnSet.onClick = function () {
		var txt = String(ui.labelField.text).replace(/^\s+|\s+$/g, "");
		if (txt.charAt(0) === "(") return; // status message, not a label
		setLabels(function () { return txt; }, "Set label");
	};

	btnRemoveLast.onClick = function () {
		setLabels(function (old) {
			var parts = splitLabel(old);
			parts.pop();
			return parts.join("+");
		}, "Remove last label");
	};

	btnClear.onClick = function () {
		if (!app.documents.length) return;
		var doc = app.activeDocument;
		if (doc.selection.length > 0) {
			setLabels(function () { return ""; }, "Clear labels");
			return;
		}
		if (!confirm("Nothing is selected.\nAll labels from the document will be deleted. Are you sure?")) return;
		app.doScript(function () {
			var items = doc.allPageItems;
			for (var i = 0; i < items.length; i++) items[i].label = "";
		}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Clear all labels");
		updateStatus();
	};

	btnClose.onClick = function () { ui.close(); };

	ui.listeners = [
		app.addEventListener("afterSelectionChanged", updateStatus),
		app.addEventListener("afterContextChanged", updateStatus)
	];

	ui.onClose = function () {
		for (var i = 0; i < ui.listeners.length; i++) {
			try { ui.listeners[i].remove(); } catch (e) { }
		}
		app.generalPreferences.ungroupRemembersLayers = oldURL;
		app.clipboardPreferences.pasteRemembersLayers = oldPRL;
		addLabelsUI = null;
	};

	return ui;
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------

// colored group that looks like a zone of the page
function area(parent, rgb, title, textRgb) {
	var g = parent.add("group");
	g.orientation = "column";
	g.alignChildren = ["center", "top"];
	g.margins = 8;
	g.spacing = 6;
	g.graphics.backgroundColor = g.graphics.newBrush(g.graphics.BrushType.SOLID_COLOR, rgb);
	caption(g, title, textRgb);
	return g;
}

function row(parent) {
	var g = parent.add("group");
	g.orientation = "row";
	g.alignChildren = ["center", "center"];
	g.spacing = 6;
	return g;
}

// InDesign ignores foregroundColor on statictext (dark UI = white text everywhere),
// so the caption is drawn by hand in the wanted color
function caption(parent, text, textRgb) {
	var g = parent.add("group");
	var size = g.graphics.measureString(text);
	g.preferredSize = [Math.ceil(size[0]) + 4, Math.ceil(size[1]) + 2];
	g.onDraw = function () {
		var gr = this.graphics;
		gr.drawString(text, gr.newPen(gr.PenType.SOLID_COLOR, textRgb.concat(1), 1), 2, 1);
	};
	return g;
}

function spacer(parent, w) {
	var g = parent.add("group");
	g.preferredSize = [w, BH];
	return g;
}

function btn(parent, text, label, tip, w) {
	var b = parent.add("button", undefined, text);
	b.preferredSize = [w || BW, BH];
	b.helpTip = tip + "\n\nLabel: " + label;
	b.onClick = function () { addLabel(label); };
	return b;
}

// ---------------------------------------------------------------------------
// labels
// ---------------------------------------------------------------------------

function splitLabel(lbl) {
	var out = [];
	var parts = String(lbl || "").split("+");
	for (var i = 0; i < parts.length; i++) {
		var p = parts[i].replace(/^\s+|\s+$/g, "");
		if (p) out.push(p);
	}
	return out;
}

function addLabel(label) {
	setLabels(function (old) {
		if (EXCLUSIVE_LABELS[label] || !addLabelsUI.combine.value) return label;
		var parts = splitLabel(old);
		for (var i = 0; i < parts.length; i++) if (parts[i] === label) return old; // already there
		// HW / DTI labels can't be combined: replace them
		if (parts.length === 1 && EXCLUSIVE_LABELS[parts[0]]) parts = [];
		parts.push(label);
		return parts.join("+");
	}, "Label: " + label);
}

// newLabelFn(oldLabel) -> new label, applied on every selected object (one undo step)
function setLabels(newLabelFn, undoName) {
	if (!app.documents.length) { alert("Please open a master file."); return; }
	var sel = app.activeDocument.selection;
	if (sel.length === 0) { alert("Select one or more objects first."); return; }

	var skipped = 0;
	app.doScript(function () {
		for (var i = 0; i < sel.length; i++) {
			if (!LABEL_TYPES[sel[i].constructor.name]) { skipped++; continue; }
			sel[i].label = newLabelFn(sel[i].label);
		}
	}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, undoName);

	updateStatus();
	if (skipped > 0) alert(skipped + " selected item(s) were skipped (text or image content).\nSelect the frame / group with the Selection tool.");
}

function updateStatus() {
	try {
		if (!addLabelsUI) return;
		var st = addLabelsUI.labelField;
		if (!app.documents.length) { st.text = "(no document open)"; return; }
		var sel = app.activeDocument.selection;
		if (sel.length === 0) { st.text = "(nothing selected)"; return; }
		var first = null, same = true;
		for (var i = 0; i < sel.length; i++) {
			if (!LABEL_TYPES[sel[i].constructor.name]) continue;
			if (first === null) first = sel[i].label;
			else if (sel[i].label !== first) { same = false; break; }
		}
		if (first === null) st.text = "(selection can't have a label - select the frame)";
		else if (!same) st.text = "(" + sel.length + " objects with different labels)";
		else st.text = first || "(no label)";
	} catch (e) { }
}
