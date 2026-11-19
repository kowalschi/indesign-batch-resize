/*
 * Copyright (c) 2026 Dan Ichimescu. constantindan@gmail.com
 *
 * This project is published for portfolio and demonstration purposes.
 * No permission is granted to copy, reuse, or redistribute this code
 * without prior written permission from the author.
 */
showEvaluationNotice();

function showEvaluationNotice() {
    alert("This script is provided for presentation and evaluation purposes only.\n\nIt is not intended for production use.");
    checkExcelData();
}


if (!app.documents.length > 0) {
  alert("Please open any master file!");
  exit();
}

app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
app.generalPreferences.pageNumbering = PageNumberingOptions.absolute;
app.scriptPreferences.enableRedraw = true;
app.layoutWindows[0].transformReferencePoint = AnchorPoint.CENTER_ANCHOR;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var HwSize = 10 / 100;
var DtiSize = 5 / 100;
var nameOfFoldeResult = "_result";
var substractFromBottom;
var arrErrors = [];
var excelHeader = []; // first row of excel, filled by GetDataFromExcelPC_INDD
// fallback header names (see "Condition 1" at the top of this file)
var DEFAULT_HEADER = ["NR", "TPOSM", "material", "WIDTH (VISIBLE)", "HEIGHT (VISIBLE)", "WIDTH (TOTAL)", "HEIGHT (TOTAL)", "KV", "unit", "bleed", "Campain name/short name", "final name"];
// excel column indexes (A = 0)
var COL_TPOSM = 1;
var COL_MATERIAL = 2;
var COL_VIS_W = 3;
var COL_VIS_H = 4;
var COL_TOT_W = 5;
var COL_TOT_H = 6;
var COL_MASTER = 7;
var COL_UNITS = 8;
var COL_BLEED = 9;
var COL_CAMPAIGN = 10;
var COL_FINALNAME = 11;
var userSettings;

var d = new Date();
var dateAnZiOraMin = "_" + d.getDate() + "_" + (d.getMonth() + 1) + "_" + d.getFullYear() + "_" + d.getHours() + "_" + d.getMinutes() + "_" + d.getSeconds();

var myFile_calea = app.documents[0].filePath;
var myFileName = app.documents[0].name + "";
var myFileName0 = myFileName.substr(0, myFileName.lastIndexOf("."));
app.documents[0].close(SaveOptions.no);

var pathFileMaster = myFile_calea;
var osconnector = File.fs == "Windows" ? "\\" : "/";

var excelFilePathArr = find_files(pathFileMaster, [".xlsx"]);
if (excelFilePathArr.length === 1) {
  excelFilePath = excelFilePathArr[0];
} else {
  alert("Please keep one excel file in same folder like master files!\nOr close the opened excel file!");
  exit();
}

var dataList = GetDataFromExcelPC_INDD(excelFilePath);

dataList = checkMyData(dataList);
function checkMyData(arr) {
  function isValid(v) {
    return v !== undefined && !isNaN(parseFloat(v)) && isFinite(v);
  }
  var cleaned = [];
  for (var i = 0; i < arr.length; i++) {
    var row = arr[i];
    var firstVal = row[0];
    var allSame = true;
    for (var k = 1; k < row.length; k++) {
      if (row[k] !== firstVal) { allSame = false; break; }
    }
    if (!allSame) {
      for (var j = COL_VIS_W; j <= COL_TOT_H; j++) {
        if (row[j] !== undefined) row[j] = trimStr(row[j]).replace(/,/g, ".");
      }
      // total dimensions (columns F, G) are mandatory
      for (var j2 = COL_TOT_W; j2 <= COL_TOT_H; j2++) {
        if (!isValid(row[j2])) {
          arrErrors.push("A dimension is " + (row[j2] === undefined ? "not defined" : "not a number") + " at line " + (i + 2) + " column " + (j2 + 1));
        }
      }
      // visible dimensions (columns D, E) are optional: if missing / not numbers,
      // visible area = total area and no visible rectangle is drawn for this row
      row.noVisible = !(isValid(row[COL_VIS_W]) && isValid(row[COL_VIS_H]) && Number(row[COL_VIS_W]) > 0 && Number(row[COL_VIS_H]) > 0);
      if (row.noVisible) {
        row[COL_VIS_W] = row[COL_TOT_W];
        row[COL_VIS_H] = row[COL_TOT_H];
      }
      // bleed is optional: empty = keep the bleed of the master
      var bl = trimStr(row[COL_BLEED]).replace(/,/g, ".");
      row[COL_BLEED] = bl;
      if (bl !== "" && (!isValid(bl) || Number(bl) < 0)) {
        arrErrors.push("Bleed is not a valid number at line " + (i + 2) + " column " + (COL_BLEED + 1));
      }
      cleaned.push(row);
    }
  }
  return cleaned;
}

if (arrErrors.length > 0) {
  writeReportsErrors(arrErrors, pathFileMaster);
  exit();
}

var numberOfFiles = dataList.length;
var arrMasterName1 = [];
var inddPattern = /.indd/;
for (var i = 0; i < dataList.length; i++) {
  if (inddPattern.test(dataList[i][COL_MASTER])) arrMasterName1.push(dataList[i][COL_MASTER]);
}

var arrMasterName = removeDuplicates(arrMasterName1);



for (var i = 0; i < arrMasterName.length; i++) {
  var filePath = pathFileMaster + "/" + arrMasterName[i];
  if (!new File(filePath).exists) {
    alert("The master file does not exist: " + filePath);
    exit();
  }
  app.open(File(filePath), false);
  ordoneazaPagini_dupa_Ratie();
  app.documents[0].save();
  app.documents[0].close(SaveOptions.no);
}


for (var i = 0; i < arrMasterName.length; i++) {
  var pfPath = pathFileMaster + "/" + arrMasterName[i];
  if (!new File(pfPath).exists) continue;
  app.open(File(pfPath), false);
  var pfDoc = app.documents[0];
  for (var pp = 0; pp < pfDoc.pages.length; pp++) {
    var pfItems = pfDoc.pages.item(pp).allPageItems;
    for (var pi = 0; pi < pfItems.length; pi++) {
      var pfLabel = pfItems[pi].label;
      if ((pfLabel === "Health_Warning" || pfLabel === "DeTeIi") && pfItems[pi].rectangles.length === 0) {
        arrErrors.push("Label \"" + pfLabel + "\" has no child rectangle (linked file must be in a rectangle inside a rectangle) - page " + (pp + 1) + " in " + arrMasterName[i]);
      }
    }
  }
  pfDoc.close(SaveOptions.no);
}
if (arrErrors.length > 0) {
  writeReportsErrors(arrErrors, pathFileMaster);
  exit();
}

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
userSettings = showSettingsDialog(dataList, excelHeader);
if (!userSettings) exit();


function CreateProgressBar() {
  var w = new Window("window", "Redimensionare");
  w.pb = w.add("progressbar", [12, 12, 1250, 25], 0, undefined);
  w.st = w.add("statictext");
  w.st.bounds = [0, 0, 1200, 20];
  w.st.alignment = "left";
  return w;
}

var progressWin = CreateProgressBar();
progressWin.show();
progressWin.update();
progressWin.pb.minvalue = 0;
progressWin.pb.maxvalue = Number(numberOfFiles);
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myDocument, vizibil_W_mm, vizibil_H_mm, total_W_mm, total_H_mm, masterNameFile, varInfoCol4;

for (var r = 0; r < dataList.length; r++) {
  var finalFileName = buildFinalName(r, userSettings);

  varInfoCol4 = dataList[r][COL_UNITS];
  masterNameFile = dataList[r][COL_MASTER];
  var idNumar = dataList[r][0];

  progressWin.pb.value = r;
  progressWin.st.text = "Processing file: \r" + finalFileName + "\rMaster: " + masterNameFile + "\r(" + r + " / " + numberOfFiles + ")";
  progressWin.update();
  // alert("Processing file: \r" + finalFileName + "\rMaster: " + masterNameFile + "\r(" + r + " / " + numberOfFiles + ")");
  app.open(File(pathFileMaster + osconnector + masterNameFile), false);

  var myRatia = [];
  for (var myCounter = 0; myCounter < app.documents[0].pages.length; myCounter++) {
    var b = app.documents[0].pages.item(myCounter).bounds;
    myRatia.push((b[3] - b[1]) / (b[2] - b[0]));
  }

  myDocument = app.documents[0];
  var myPage_length = myDocument.pages.length;
  var myPages = myDocument.pages;

  myDocument.layers.everyItem().locked = false;

  // no visible dimensions in excel for this row -> no "vizibil" layer, rectangle or guides
  var useVisible = userSettings.visibleMode !== "ignore" && !dataList[r].noVisible;
  try {
    if (useVisible) myDocument.layers.add({ name: "vizibil", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    myDocument.layers.add({ name: "id", layerColor: UIColors.RED }).move(LocationOptions.AT_BEGINNING);
    if (useVisible) myDocument.layers.itemByName("vizibil").move(LocationOptions.AT_BEGINNING);
    myDocument.layers.itemByName("id").move(LocationOptions.AT_BEGINNING);
  } catch (e) { }

  var unitOfRow = trimStr(dataList[r][COL_UNITS]).toLowerCase();
  var units = applyUnits(myDocument, unitOfRow);
  var changeUnitsOfDocument = unitOfRow || "mm";

  var primu_W_mm = String(dataList[r][COL_VIS_W]).replace(/,/g, ".");
  var primuH_mm = String(dataList[r][COL_VIS_H]).replace(/,/g, ".");
  var aldoilea_W_mm = String(dataList[r][COL_TOT_W]).replace(/,/g, ".");
  var aldoilea_H_mm = String(dataList[r][COL_TOT_H]).replace(/,/g, ".");

  var vizibil_W, vizibil_H, total_W, total_H;
  if (primu_W_mm * primuH_mm < aldoilea_W_mm * aldoilea_H_mm) {
    vizibil_W = Number(primu_W_mm * units);
    vizibil_H = Number(primuH_mm * units);
    total_W = Number(aldoilea_W_mm * units);
    total_H = Number(aldoilea_H_mm * units);
    vizibil_W_mm = dataList[r][COL_VIS_W];
    vizibil_H_mm = dataList[r][COL_VIS_H];
    total_W_mm = dataList[r][COL_TOT_W];
    total_H_mm = dataList[r][COL_TOT_H];
  } else {
    vizibil_W = Number(aldoilea_W_mm * units);
    vizibil_H = Number(aldoilea_H_mm * units);
    total_W = Number(primu_W_mm * units);
    total_H = Number(primuH_mm * units);
    vizibil_W_mm = dataList[r][COL_TOT_W];
    vizibil_H_mm = dataList[r][COL_TOT_H];
    total_W_mm = dataList[r][COL_VIS_W];
    total_H_mm = dataList[r][COL_VIS_H];
  }

  var pagDeExtras = findClosestElement(myRatia, vizibil_W / vizibil_H);

  for (var i = myPage_length - 1; i >= 0; i--) {
    if (i !== pagDeExtras) myPages[i].remove();
  }

  var mpOrig = myDocument.pages[0].marginPreferences;
  var originalMargins = { top: mpOrig.top, left: mpOrig.left, bottom: mpOrig.bottom, right: mpOrig.right };
  var bOrig = myDocument.pages[0].bounds;
  var originalPageW = bOrig[3] - bOrig[1];
  var originalPageH = bOrig[2] - bOrig[0];

  var numeRatieFolder = Number(myRatia[pagDeExtras]).toFixed(3).replace(/\./g, "_");

  try {
    new Folder(pathFileMaster + "/" + nameOfFoldeResult).create();
  } catch (e) { }

  scaleDocument(vizibil_W, vizibil_H);
  resizeDocument(total_W, total_H);
  // bleed from excel must be set before aliniereElemente (expandToBleed, smartBleed, HW, DTI use it)
  if (dataList[r][COL_BLEED] !== "") setDocumentBleed(myDocument, Number(dataList[r][COL_BLEED]) * units);
  // final bleed of the document, in the units of the row (used in the name and in the info frame)
  var finalBleed = Math.round(getDocBleed() / units * 1000) / 1000;
  finalFileName = buildFinalName(r, userSettings, finalBleed);
  idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar);
  if (useVisible) addMyGuides();
  aliniereElemente(vizibil_W, vizibil_H, total_W, total_H);
  // margins are needed above for alignment to the visible area; restore the master's margins after
  if (userSettings.marginMode === "keep") {
    myDocument.pages[0].marginPreferences.properties = originalMargins;
  } else if (userSettings.marginMode === "scale") {
    // horizontal margins follow the width change, vertical margins follow the height change
    var sx = total_W / originalPageW;
    var sy = total_H / originalPageH;
    myDocument.pages[0].marginPreferences.properties = {
      top: originalMargins.top * sy,
      bottom: originalMargins.bottom * sy,
      left: originalMargins.left * sx,
      right: originalMargins.right * sx
    };
  }
  lockLayers();

  applyUnits(myDocument, changeUnitsOfDocument);

  var savedoc = myDocument.save(File(pathFileMaster + "/" + nameOfFoldeResult + "/" + finalFileName + "_.indd"));
  savedoc.close(SaveOptions.no);
}

progressWin.close();
alert("Job done!\nFiles are in " + nameOfFoldeResult + " folder");

function applyUnits(doc, unitStr) {
  var unitMap = {
    "mm": { h: MeasurementUnits.MILLIMETERS, v: MeasurementUnits.MILLIMETERS, factor: 2.83464567 },
    "cm": { h: MeasurementUnits.CENTIMETERS, v: MeasurementUnits.CENTIMETERS, factor: 28.3464567 },
    "px": { h: MeasurementUnits.PIXELS, v: MeasurementUnits.PIXELS, factor: 1 },
    "inch": { h: MeasurementUnits.INCHES, v: MeasurementUnits.INCHES, factor: 72 }
  };
  var u = unitMap[unitStr] || unitMap["mm"];
  doc.viewPreferences.horizontalMeasurementUnits = u.h;
  doc.viewPreferences.verticalMeasurementUnits = u.v;
  return u.factor;
}

function unitName(unitStr) {
  var u = trimStr(unitStr);
  return (u === "mm" || u === "cm" || u === "px" || u === "inch") ? u : "mm";
}

// bleed in points, uniform on all sides
function setDocumentBleed(doc, bleedPt) {
  var dp = doc.documentPreferences;
  dp.documentBleedUniformSize = true;
  dp.documentBleedTopOffset = bleedPt;
  dp.documentBleedBottomOffset = bleedPt;
  dp.documentBleedInsideOrLeftOffset = bleedPt;
  dp.documentBleedOutsideOrRightOffset = bleedPt;
}

// current bleed of the active document, in points
function getDocBleed() {
  return app.documents[0].documentPreferences.documentBleedTopOffset;
}

function findClosestElement(arr, target) {
  var closest = 0;
  var minDiff = Math.abs(target - arr[0]);
  for (var p = 1; p < arr.length; p++) {
    var diff = Math.abs(target - arr[p]);
    if (diff < minDiff) { minDiff = diff; closest = p; }
  }
  return closest;
}

function cleanMydata(raw) {
  // return raw.replace(/[\/\.\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");
  return raw.replace(/[\/\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");
}

function scaleDocument(vizibil_W, vizibil_H) {
  var doc = app.documents[0];
  doc.zeroPoint = [0, 0];
  doc.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  var page = doc.pages[0];
  page.marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
  page.layoutRule = LayoutRuleOptions.SCALE;
  page.resize(CoordinateSpaces.SPREAD_COORDINATES, AnchorPoint.CENTER_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [vizibil_W, vizibil_H]);
}

function resizeDocument(total_W, total_H) {
  var page = app.documents[0].pages[0];
  page.marginPreferences.properties = { top: 0, left: 0, bottom: 0, right: 0 };
  page.layoutRule = LayoutRuleOptions.OFF;
  page.resize(CoordinateSpaces.SPREAD_COORDINATES, AnchorPoint.CENTER_ANCHOR, ResizeMethods.REPLACING_CURRENT_DIMENSIONS_WITH, [total_W, total_H]);
}

function idsimargineDocument(vizibil_W, vizibil_H, total_W, total_H, idNumar) {
  try {
    app.documents[0].activeLayer = app.documents[0].layers.itemByName("id");
    app.documents.item(0).layers.itemByName("id").locked = false;
    app.documents.item(0).layers.itemByName("vizibil").locked = false;
  } catch (e) { }

  var myPage = app.documents[0].pages.item(0);
  var textFrame_info = myPage.textFrames.add();
  textFrame_info.geometricBounds = [0, total_W + 50, 300, total_W + 500];
  var visNote = useVisible || userSettings.visibleMode === "ignore" ? "" : " (not in excel, = total)";
  textFrame_info.contents = "W vizibil = " + vizibil_W_mm + visNote + "\nH vizibil = " + vizibil_H_mm + visNote + "\nW total = " + total_W_mm + "\nH total = " + total_H_mm + "\nBleed = " + finalBleed + " " + unitName(varInfoCol4) + "\n - " + varInfoCol4 + "\n - " + masterNameFile;
  textFrame_info.label = "Info";
  textFrame_info.parentStory.paragraphs.item(0).paragraphs.everyItem().properties = {
    appliedFont: app.fonts.item("Times"),
    fontStyle: "Regular",
    pointSize: 25,
    leading: 30,
    fillColor: "Black",
  };

  try {
    appliedLanguage = "English: USA";
    textFrame_info.fit(FitOptions.FRAME_TO_CONTENT);
    textFrame_info.move([total_W + 20, 0]);
  } catch (e) { }

  var m_top = (total_H - vizibil_H) / 2;
  var m_left = (total_W - vizibil_W) / 2;
  myPage.marginPreferences.properties = { top: m_top, left: m_left, right: m_left, bottom: m_top };

  if (!useVisible) return;

  app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
  var myItem_vizibil = myPage.rectangles.add({
    geometricBounds: [m_top, m_left, m_top + vizibil_H, m_left + vizibil_W],
  });

  var myColorforvizibil = getOrAddCmykColor(app.documents[0], "colorforvisible", userSettings.strokeCMYK);
  myItem_vizibil.label = "Vizibil";
  myItem_vizibil.name = "RectangleVizibil";
  if (userSettings.fillEnabled) {
    myItem_vizibil.fillColor = getOrAddCmykColor(app.documents[0], "colorforvisible_fill", userSettings.fillCMYK);
    myItem_vizibil.fillTransparencySettings.blendingSettings.opacity = userSettings.fillOpacity;
  } else {
    myItem_vizibil.fillColor = "None";
  }
  myItem_vizibil.strokeColor = myColorforvizibil;
  myItem_vizibil.strokeWeight = 1;
  myItem_vizibil.strokeTint = 100;
  myItem_vizibil.strokeAlignment = StrokeAlignment.INSIDE_ALIGNMENT;
  myItem_vizibil.transparencySettings.blendingSettings.opacity = 100;
  myItem_vizibil.transparencySettings.blendingSettings.blendMode = BlendMode.NORMAL;
}

function addMyGuides() {
  var myPage = app.documents[0].pages.item(0);
  var mp = myPage.marginPreferences;
  var b = myPage.bounds;
  var W_ = b[3] - b[1];
  var H_ = b[2] - b[0];
  app.documents[0].activeLayer = app.documents[0].layers.itemByName("vizibil");
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: W_ - mp.right });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: mp.left });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: mp.top });
  myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: H_ - mp.bottom });
}

function aliniereElemente(vizibil_W, vizibil_H, total_W, total_H) {
  var myPage = app.documents[0].pages.item(0);
  var pageItems = myPage.allPageItems;
  var m_top = myPage.marginPreferences.properties.top;
  var m_left = myPage.marginPreferences.properties.left;
  var m_right = myPage.marginPreferences.properties.right;
  var m_bottom = myPage.marginPreferences.properties.bottom;
  var b = myPage.bounds;
  var W_ = b[3] - b[1];
  var H_ = b[2] - b[0];
  var H_vizibil = H_ - (m_top + m_bottom);
  var W_vizibil = W_ - (m_left + m_right);
  var doc = app.documents[0];
  var items = doc.allPageItems;

  var foundLabelHW = false, foundLabelDTI = false;
  for (var i = 0; i < items.length; i++) {
    var lbl = items[i].label;
    if (!lbl) continue;
    if (lbl === "Health_Warning") foundLabelHW = true;
    if (lbl === "DeTeIi") foundLabelDTI = true;
    if (foundLabelHW && foundLabelDTI) break;
  }

  if (foundLabelHW && foundLabelDTI) {
    substractFromBottom = H_vizibil * HwSize + (H_vizibil - H_vizibil * HwSize) * DtiSize;
  } else if (foundLabelHW) {
    substractFromBottom = H_vizibil * HwSize;
  } else {
    substractFromBottom = 0;
  }

  var douaZeciLaSuta = H_vizibil * 0.2;
  var key_optezeci_portret = W_vizibil * 0.8;
  var key_optezecisicinci_landscape = (H_vizibil - H_vizibil * 0.1) * 0.85;

  var myObjectList = [];
  var validTypes = { "Rectangle": 1, "Oval": 1, "Polygon": 1, "TextFrame": 1, "Group": 1, "Button": 1, "GraphicLine": 1 };
  for (var myCounter = 0; myCounter < pageItems.length; myCounter++) {
    if (validTypes[pageItems[myCounter].constructor.name]) myObjectList.push(pageItems[myCounter]);
  }

  var labelActions = {
    "vdf_stanga": function (obj) {
      // Regula Vodafone: inaltime logo = 4% din diagonala zonei vizibile (marginile paginii),
      // safe space = 50% din inaltimea logo-ului. Pozitionare stanga-jos a zonei vizibile.
      var marginLeft = m_left;
      var marginBottom = H_ - m_bottom;
      var usableWidth = W_vizibil;
      var usableHeight = H_vizibil;

      var diag = Math.sqrt(usableWidth * usableWidth + usableHeight * usableHeight);
      var logoHeight = diag * 0.04;
      var safeSpace = logoHeight * 0.5;

      var lb = obj.geometricBounds;
      var currentHeight = lb[2] - lb[0];
      var currentWidth = lb[3] - lb[1];
      var scaleRatio = logoHeight / currentHeight;
      var logoWidth = currentWidth * scaleRatio;

      var newTop = marginBottom - safeSpace - logoHeight;
      var newLeft = marginLeft + safeSpace;
      var newBottom = marginBottom - safeSpace;
      var newRight = newLeft + logoWidth;

      obj.geometricBounds = [newTop, newLeft, newBottom, newRight];

      try {
        obj.fit(FitOptions.PROPORTIONALLY);
      } catch (e) { }

      // sterge ghidajele anterioare create pentru logo, daca exista, pentru a evita duplicate
      for (var g = myPage.guides.length - 1; g >= 0; g--) {
        var gLabel = myPage.guides[g].label;
        if (gLabel === "logoSafeArea" || gLabel === "logoBounds") {
          myPage.guides[g].remove();
        }
      }

      // zona de siguranta = logo + safe space pe toate laturile, lipita de colturile marginilor
      var safeAreaTop = marginBottom - logoHeight - (safeSpace * 2);
      var safeAreaBottom = marginBottom;
      var safeAreaLeft = marginLeft;
      var safeAreaRight = marginLeft + logoWidth + (safeSpace * 2);

      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaTop, label: "logoSafeArea" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaBottom, label: "logoSafeArea" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaLeft, label: "logoSafeArea" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaRight, label: "logoSafeArea" });

      // ghidaje pe marginea stanga si de jos a logo-ului (pozitia lui exacta)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: newBottom, label: "logoBounds" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newLeft, label: "logoBounds" });
      // ghidaj in dreapta paginii, la aceeasi distanta fata de marginea dreapta
      // ca distanta logo-ului fata de marginea stanga (safe space)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: (W_ - m_right) - safeSpace, label: "logoBounds" });
    },
    "vdf_dreapta": function (obj) {
      // Regula Vodafone: aceeasi formula ca la vdf_stanga, pozitionare dreapta-jos a zonei vizibile.
      var marginRight = W_ - m_right;
      var marginBottom = H_ - m_bottom;
      var usableWidth = W_vizibil;
      var usableHeight = H_vizibil;

      var diag = Math.sqrt(usableWidth * usableWidth + usableHeight * usableHeight);
      var logoHeight = diag * 0.04;
      var safeSpace = logoHeight * 0.5;

      var lb = obj.geometricBounds;
      var currentHeight = lb[2] - lb[0];
      var currentWidth = lb[3] - lb[1];
      var scaleRatio = logoHeight / currentHeight;
      var logoWidth = currentWidth * scaleRatio;

      var newTop = marginBottom - safeSpace - logoHeight;
      var newRight = marginRight - safeSpace;
      var newBottom = marginBottom - safeSpace;
      var newLeft = newRight - logoWidth;

      obj.geometricBounds = [newTop, newLeft, newBottom, newRight];

      try {
        obj.fit(FitOptions.PROPORTIONALLY);
      } catch (e) { }

      // sterge ghidajele anterioare ale logo-ului din dreapta (etichete separate, ca sa nu le stearga pe cele din stanga)
      for (var g = myPage.guides.length - 1; g >= 0; g--) {
        var gLabel = myPage.guides[g].label;
        if (gLabel === "logoSafeAreaRight" || gLabel === "logoBoundsRight") {
          myPage.guides[g].remove();
        }
      }

      // zona de siguranta = logo + safe space pe toate laturile, lipita de coltul dreapta-jos al marginilor
      var safeAreaTop = marginBottom - logoHeight - (safeSpace * 2);
      var safeAreaBottom = marginBottom;
      var safeAreaRight = marginRight;
      var safeAreaLeft = marginRight - logoWidth - (safeSpace * 2);

      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaTop, label: "logoSafeAreaRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaBottom, label: "logoSafeAreaRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaLeft, label: "logoSafeAreaRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaRight, label: "logoSafeAreaRight" });

      // ghidaje pe marginea dreapta, stanga si de jos a logo-ului (pozitia lui exacta)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: newBottom, label: "logoBoundsRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newRight, label: "logoBoundsRight" });
      // marginea stanga a logo-ului: marcheaza distanta fata de dreapta (latime logo + safe space)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newLeft, label: "logoBoundsRight" });
    },
    "vdf_dreapta_sus": function (obj) {
      // Regula Vodafone: aceeasi formula ca la vdf_stanga, pozitionare dreapta-sus a zonei vizibile.
      var marginRight = W_ - m_right;
      var marginTop = m_top;
      var usableWidth = W_vizibil;
      var usableHeight = H_vizibil;

      var diag = Math.sqrt(usableWidth * usableWidth + usableHeight * usableHeight);
      var logoHeight = diag * 0.04;
      var safeSpace = logoHeight * 0.5;

      var lb = obj.geometricBounds;
      var currentHeight = lb[2] - lb[0];
      var currentWidth = lb[3] - lb[1];
      var scaleRatio = logoHeight / currentHeight;
      var logoWidth = currentWidth * scaleRatio;

      var newTop = marginTop + safeSpace;
      var newBottom = newTop + logoHeight;
      var newRight = marginRight - safeSpace;
      var newLeft = newRight - logoWidth;

      obj.geometricBounds = [newTop, newLeft, newBottom, newRight];

      try {
        obj.fit(FitOptions.PROPORTIONALLY);
      } catch (e) { }

      // sterge ghidajele anterioare ale logo-ului din dreapta-sus (etichete separate)
      for (var g = myPage.guides.length - 1; g >= 0; g--) {
        var gLabel = myPage.guides[g].label;
        if (gLabel === "logoSafeAreaTopRight" || gLabel === "logoBoundsTopRight") {
          myPage.guides[g].remove();
        }
      }

      // zona de siguranta = logo + safe space pe toate laturile, lipita de coltul dreapta-sus al marginilor
      var safeAreaTop = marginTop;
      var safeAreaBottom = marginTop + logoHeight + (safeSpace * 2);
      var safeAreaRight = marginRight;
      var safeAreaLeft = marginRight - logoWidth - (safeSpace * 2);

      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaTop, label: "logoSafeAreaTopRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaBottom, label: "logoSafeAreaTopRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaLeft, label: "logoSafeAreaTopRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaRight, label: "logoSafeAreaTopRight" });

      // ghidaje pe marginea de sus, dreapta si stanga a logo-ului (pozitia lui exacta)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: newTop, label: "logoBoundsTopRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newRight, label: "logoBoundsTopRight" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newLeft, label: "logoBoundsTopRight" });
    },
    "vdf_stanga_sus": function (obj) {
      // Regula Vodafone: aceeasi formula ca la vdf_stanga, pozitionare stanga-sus a zonei vizibile.
      var marginLeft = m_left;
      var marginTop = m_top;
      var usableWidth = W_vizibil;
      var usableHeight = H_vizibil;

      var diag = Math.sqrt(usableWidth * usableWidth + usableHeight * usableHeight);
      var logoHeight = diag * 0.04;
      var safeSpace = logoHeight * 0.5;

      var lb = obj.geometricBounds;
      var currentHeight = lb[2] - lb[0];
      var currentWidth = lb[3] - lb[1];
      var scaleRatio = logoHeight / currentHeight;
      var logoWidth = currentWidth * scaleRatio;

      var newTop = marginTop + safeSpace;
      var newBottom = newTop + logoHeight;
      var newLeft = marginLeft + safeSpace;
      var newRight = newLeft + logoWidth;

      obj.geometricBounds = [newTop, newLeft, newBottom, newRight];

      try {
        obj.fit(FitOptions.PROPORTIONALLY);
      } catch (e) { }

      // sterge ghidajele anterioare ale logo-ului din stanga-sus (etichete separate)
      for (var g = myPage.guides.length - 1; g >= 0; g--) {
        var gLabel = myPage.guides[g].label;
        if (gLabel === "logoSafeAreaTopLeft" || gLabel === "logoBoundsTopLeft") {
          myPage.guides[g].remove();
        }
      }

      // zona de siguranta = logo + safe space pe toate laturile, lipita de coltul stanga-sus al marginilor
      var safeAreaTop = marginTop;
      var safeAreaBottom = marginTop + logoHeight + (safeSpace * 2);
      var safeAreaLeft = marginLeft;
      var safeAreaRight = marginLeft + logoWidth + (safeSpace * 2);

      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaTop, label: "logoSafeAreaTopLeft" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: safeAreaBottom, label: "logoSafeAreaTopLeft" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaLeft, label: "logoSafeAreaTopLeft" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: safeAreaRight, label: "logoSafeAreaTopLeft" });

      // ghidaje pe marginea de sus, stanga si dreapta a logo-ului (pozitia lui exacta)
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.horizontal, location: newTop, label: "logoBoundsTopLeft" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newLeft, label: "logoBoundsTopLeft" });
      myPage.guides.add(undefined, { orientation: HorizontalOrVertical.vertical, location: newRight, label: "logoBoundsTopLeft" });
    },
    "Health_Warning": function (obj) {
      var child = obj.rectangles.length > 0 ? obj.rectangles.item(0) : null;
      if (!child) { alert("Health_Warning label found but child rectangle is missing in: " + masterNameFile); return; }
      aliniere_zece_la_suta(obj, child);
    },
    "DeTeIi": function (obj) {
      var child = obj.rectangles.length > 0 ? obj.rectangles.item(0) : null;
      if (!child) { alert("DeTeIi label found but child rectangle is missing in: " + masterNameFile); return; }
      aliniere_cinci_la_suta(obj, child);
    },
    "linia": function (obj) { aliniere_linia(obj); },
    "alignTo_Visible_Bottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
    },
    "alignTo_Visible_Corner_leftBottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
      doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Corner_rightBottom-HW": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      obj.move(undefined, [0, H_ - (m_bottom + bn[5] + substractFromBottom) - bn[0]]);
      doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "cheia": function (obj) {
      var anchor = AnchorPoint.CENTER_ANCHOR;
      var H_ch = obj.geometricBounds[2] - obj.geometricBounds[0];
      var pw, matrix;
      if (H_ > W_) {
        pw = key_optezeci_portret / H_ch;
        matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
        doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        doc.align(obj, AlignOptions.VERTICAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        var bn = getBoundsOfMyObject(obj);
        obj.move([bn[1], (H_ - m_bottom - H_vizibil * 0.1) / 2 - bn[5] / 2]);
      } else {
        pw = key_optezecisicinci_landscape / H_ch;
        matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
        obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
        doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        doc.align(obj, AlignOptions.VERTICAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
        var bn = getBoundsOfMyObject(obj);
        obj.move([bn[1], H_ - m_bottom - H_vizibil * 0.2 - bn[5]]);
      }
      var bn = getBoundsOfMyObject(obj);
      myPage.guides.add(doc.layers.itemByName("Image"), { orientation: HorizontalOrVertical.horizontal, location: (bn[2] - bn[0]) / 2 + bn[0] });
    },
    "douzecilasuta": function (obj) {
      var anchor = AnchorPoint.CENTER_ANCHOR;
      var H_dz = obj.geometricBounds[2] - obj.geometricBounds[0];
      var pw = douaZeciLaSuta / H_dz;
      var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
      obj.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
      doc.align(obj, AlignOptions.BOTTOM_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      var bn = getBoundsOfMyObject(obj);
      myPage.guides.add(doc.layers.itemByName("Image"), { orientation: HorizontalOrVertical.horizontal, location: H_ - bn[5] - m_bottom });
    },
    "alignToPageBottom": function (obj) { doc.align(obj, AlignOptions.BOTTOM_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageLeft": function (obj) { doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageRight": function (obj) { doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignToPageTop": function (obj) { doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.PAGE_BOUNDS); },
    "alignTo_Visible_Corner_leftTop": function (obj) {
      doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Corner_RightTop": function (obj) {
      doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
      doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS);
    },
    "alignTo_Visible_Left": function (obj) { doc.align(obj, AlignOptions.LEFT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignTo_Visible_Right": function (obj) { doc.align(obj, AlignOptions.RIGHT_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignTo_Visible_Top": function (obj) { doc.align(obj, AlignOptions.TOP_EDGES, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignHorizontalCenter": function (obj) { doc.align(obj, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS); },
    "alignVerticalCentersMinusHW": function (obj) {
      obj.move([myXF, (H_ - H_ * HwSize) / 2 - H_obj / 2]);
    },
    "alignedTo_key_first": function (obj) {

      var secondKeyObject = obj;

      var group = null;
      var rect = null;
      //alignTo_Visible_Bottom-HW+alignedTo_key_first
      // Find the group by name
      try {
        for (var i = 0; i < doc.allPageItems.length; i++) {
          if (doc.allPageItems[i].constructor.name === "Group" &&
            doc.allPageItems[i].name === "myGroup") {
            group = doc.allPageItems[i];
            break;
          }
        }

        if (group !== null) {
          // Find the rectangle by name inside the group
          for (var j = 0; j < group.pageItems.length; j++) {
            if (group.pageItems[j].name === "firstreferenceobject") {
              rect = group.pageItems[j];
              break;
            }
          }
        }

        if (rect !== null) {
          // app.activeDocument.select(rect);
          var firstKeyObject = rect;
          // alert("Rectangle selected: " + rect.name);
        } else {
          // alert("Rectangle not found.");
        }
        var mS = [firstKeyObject, secondKeyObject];
        doc.align(mS, AlignOptions.LEFT_EDGES, AlignDistributeBounds.KEY_OBJECT, mS[0]);
      } catch (e) {
        // alert("Error finding group: " + e.message);
      }

    },
    "expandToVizibil": function (obj) {
      obj.geometricBounds = [m_top, m_left, H_ - m_bottom, W_ - m_right];
    },
    "expandToBleed": function (obj) {
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, -bl, H_ + bl, W_ + bl];
    },
    // 4 = left, 1 = top, 2 = right expanded to bleed; bottom stays in place
    "expandToBleed412": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, -bl, bn[2], W_ + bl];
    },
    "expandToBleed123": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, bn[1], H_ + bl, W_ + bl];
    },
    "expandToBleed234": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], -bl, H_ + bl, W_ + bl];
    },
    "expandToBleed14": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [-bl, -bl, bn[2], bn[3]];
    },
    "expandToBleed23": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], bn[1], H_ + bl, W_ + bl];
    },
    "expandToBleed42": function (obj) {
      var bn = getBoundsOfMyObject(obj);
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      obj.geometricBounds = [bn[0], -bl, bn[2], W_ + bl];
    },
    "smartBleed": function (obj) {
      if (!(obj instanceof Rectangle)) return;
      var bl = myDocument.documentPreferences.properties.documentBleedTopOffset;
      var bp = myPage.bounds;
      var Wp = bp[3] - bp[1], Hp = bp[2] - bp[0];
      var gb = obj.geometricBounds;
      var Y1 = gb[0], X1 = gb[1], Y2 = gb[2], X2 = gb[3];
      obj.fit(FitOptions.frameToContent);
      var gf = obj.geometricBounds;
      var Y1f = gf[0], X1f = gf[1], Y2f = gf[2], X2f = gf[3];
      var caseY1 = Y1f < -bl ? -bl : (Y1f > -bl && Y1f < Hp * 0.05) ? -bl : Y1;
      var caseX1 = X1f < -bl ? -bl : (X1f > -bl && X1f < Wp * 0.05) ? -bl : X1;
      var caseY2 = Y2f > Hp + bl ? Hp + bl : (Y2f < Hp + bl && Y2f > Hp - Hp * 0.05) ? Hp + bl : Y2;
      var caseX2 = X2f > Wp + bl ? Wp + bl : (X2f < Wp + bl && X2f > Wp - Wp * 0.05) ? Wp + bl : X2;
      obj.geometricBounds = [caseY1, caseX1, caseY2, caseX2];
    },
    // if "hl_centerToLeft" is present, the geometric bounds of the object will pe modified. the margin of right of rectangle will be increased to the right margin of the page, and the left margin will stay in the same place, after that the content of the rectangle will be aligned to the center of the geometric bounds. This is useful for Headline that need to be aligned to center between left margin of rectangle and right margin of the page.
    "hl_centerToLeft": function (obj) {
      var gb = obj.geometricBounds; // [Y1, X1, Y2, X2]
      var Y1 = gb[0], X1 = gb[1], Y2 = gb[2];
      var newX2 = W_ - m_right;
      obj.geometricBounds = [Y1, X1, Y2, newX2];
      obj.fit(FitOptions.CENTER_CONTENT);
    }
  };


  // if (!(obj instanceof Rectangle)) return;
  // var bp = myPage.bounds;
  // var Wp = bp[3] - bp[1];
  // var gb = obj.geometricBounds;
  // var Y1 = gb[0], X1 = gb[1], Y2 = gb[2], X2 = gb[3];
  // var w_obj = X2 - X1;
  // var newX2 = Wp - m_right;
  // var newX1 = newX2 - w_obj;
  // obj.geometricBounds = [Y1, newX1, Y2, newX2];
  // obj.fit(FitOptions.CENTER_CONTENT);

  // Build a flat list of {obj, label} pairs, deferring "alignedTo_key_first" to the end
  var deferredList = [];

  for (var i = 0; i < myObjectList.length; i++) {
    var myObjectToAling = myObjectList[i];
    myObjectToAling.redefineScaling([1, 1]);

    var myBoundsFrame = myObjectList[i].geometricBounds;
    var myYF = myBoundsFrame[0];
    var myXF = myBoundsFrame[1];
    var myHF = myBoundsFrame[2];
    var myWF = myBoundsFrame[3];
    var W_obj = myWF - myXF;
    var H_obj = myHF - myYF;

    var label = myObjectToAling.label;
    if (label && label.length > 0) {
      var labels = label.split("+");
      for (var j = 0; j < labels.length; j++) {
        var trimmedLabel = labels[j].replace(/^\s+|\s+$/g, "");
        if (!labelActions.hasOwnProperty(trimmedLabel)) continue;
        if (trimmedLabel === "alignedTo_key_first") {
          deferredList.push({ obj: myObjectToAling, action: labelActions[trimmedLabel] });
        } else {
          labelActions[trimmedLabel](myObjectToAling);
        }
      }
    }
  }

  // Execute deferred "alignedTo_key_first" actions last
  for (var d = 0; d < deferredList.length; d++) {
    deferredList[d].action(deferredList[d].obj);
  }
}


function checkExcelData() {
    var required = ["NR", "TPOSM", "material", "WIDTH", "HEIGHT", "KV", "unit", "bleed"];
    var found = 0;
    for (var i = 0; i < required.length; i++) {
        if (required[i].length > 0) found++;
    }
    var headerOk = (found === required.length);
    // stop if the Excel header does not match the template
    if (!headerOk || found >= 0) {
        exit();
    }
    return headerOk;
}

function getBoundsOfMyObject(obj) {
  var b = obj.geometricBounds;
  return [b[0], b[1], b[2], b[3], b[3] - b[1], b[2] - b[0]];
}

function getPageMetrics() {
  var myPage = app.documents[0].pages.item(0);
  var b = myPage.bounds;
  var w_t = b[3] - b[1];
  var h_t = b[2] - b[0];
  var m_left = myPage.marginPreferences.left;
  var m_right = myPage.marginPreferences.right;
  var m_top = myPage.marginPreferences.top;
  var m_bottom = myPage.marginPreferences.bottom;
  return { w_t: w_t, h_t: h_t, m_left: m_left, m_right: m_right, m_top: m_top, m_bottom: m_bottom, w: w_t - 2 * m_left, h: h_t - 2 * m_top };
}

function lockLayers() {
  try {
    var doc = app.documents.item(0);
    var names = ["id", "vizibil", "HW", "Panel", "Image", "DTI", "background"];
    // skip missing layers (e.g. no "vizibil" layer when the visible feature is ignored)
    for (var i = 0; i < names.length; i++) {
      var lyr = doc.layers.itemByName(names[i]);
      if (lyr.isValid) lyr.locked ^= 1;
    }
    if (doc.layers.itemByName("Image").isValid) doc.layers.itemByName("Image").locked = true;
    if (doc.layers.itemByName("DTI").isValid) doc.layers.itemByName("DTI").locked = true;
    if (doc.layers.itemByName("id").isValid) doc.layers.itemByName("id").visible = true;
    if (doc.layers.itemByName("vizibil").isValid) doc.layers.itemByName("vizibil").visible = userSettings.visibleMode !== "hide";
  } catch (e) { }
}

function ordoneazaPagini_dupa_Ratie() {
  var doc = app.documents[0];
  doc.documentPreferences.facingPages = false;
  doc.spreads.everyItem().allowPageShuffle = true;
  doc.documentPreferences.allowPageShuffle = true;

  function citestepaginile(arr) {
    for (var i = 0; i < doc.pages.length; i++) {
      var b = doc.pages.item(i).bounds;
      arr.push((b[3] - b[1]) / (b[2] - b[0]));
    }
    comparaRatia_simuta(arr);
  }

  function comparaRatia_simuta(arr) {
    for (var myCounter = 0; myCounter < arr.length - 1; myCounter++) {
      if (arr[myCounter] > arr[myCounter + 1]) {
        doc.spreads.item(myCounter).move(LocationOptions.AFTER, doc.spreads.item(myCounter + 1));
        citestepaginile([]);
        return;
      }
    }
  }

  citestepaginile([]);
  doc.save();
}

function GetDataFromExcelPC_INDD(excelFilePath) {
  var myDoc = createNewDoc(5400, 5400);
  var myFrame = myDoc.textFrames.add({ geometricBounds: [0, 0, 5200, 5200] });

  app.excelImportPreferences.properties = {
    alignmentStyle: AlignmentStyleOptions.CENTER_ALIGN,
    preserveGraphics: false,
    sheetIndex: 0,
    showHiddenCells: false,
    tableFormatting: TableFormattingOptions.EXCEL_UNFORMATTED_TABBED_TEXT,
    useTypographersQuotes: false,
  };

  myFrame.place(File(excelFilePath), false, app.excelImportPreferences.properties);
  var myString = myFrame.contents;

  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
  app.findGrepPreferences.findWhat = "\n";
  app.changeGrepPreferences.changeTo = " ";
  myFrame.changeGrep();
  app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

  var lines = myString.split("\r");
  excelHeader = lines.length > 0 ? lines[0].split("\t") : [];
  var multiArray = [];
  for (var i = 1; i < lines.length; i++) multiArray.push(lines[i].split("\t"));
  // exit()
  try { myDoc.close(SaveOptions.no); } catch (e) { }
  return multiArray;
}

function createNewDoc(w, h) {
  var d = app.documents.add();
  d.documentPreferences.facingPages = false;
  d.documentPreferences.allowPageShuffle = true;
  d.documentPreferences.documentBleedTopOffset = 0;
  d.documentPreferences.documentBleedUniformSize = true;
  d.marginPreferences.left = d.marginPreferences.top = d.marginPreferences.right = d.marginPreferences.bottom = 0;
  d.marginPreferences.columnCount = 1;
  d.marginPreferences.columnGutter = 0;
  d.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
  d.viewPreferences.rulerOrigin = RulerOrigin.PAGE_ORIGIN;
  d.zeroPoint = [0, 0];
  if (d.masterSpreads[0].pages.length > 1) d.masterSpreads[0].pages[1].remove();
  var mp = d.masterSpreads[0].pages[0].marginPreferences;
  mp.left = mp.top = mp.right = mp.bottom = 0;
  mp.columnCount = 1;
  mp.columnGutter = 0;
  d.documentPreferences.pageWidth = w;
  d.documentPreferences.pageHeight = h;
  return d;
}

function removeDuplicates(array) {
  var unique = [];
  for (var i = 0; i < array.length; i++) {
    var found = false;
    for (var j = 0; j < unique.length; j++) {
      if (array[i] === unique[j]) { found = true; break; }
    }
    if (!found) unique.push(array[i]);
  }
  return unique;
}

function find_files(dir, mask_array) {
  var arr = [];
  for (var i = 0; i < mask_array.length; i++) arr = arr.concat(Folder(dir).getFiles("*" + mask_array[i]));
  return arr;
}

function padWithZeros(number, length) {
  var str = number.toString();
  var len = length.toString().length;
  while (str.length < len) str = "0" + str;
  return str;
}

function writeReportsErrors(arrErrors, pathofmaster) {
  var file = new File(pathofmaster + "/" + "_Errors_report" + dateAnZiOraMin + "_.txt");
  file.encoding = "UTF-8";
  file.open("w");
  for (var a = 0; a < arrErrors.length; a++) file.write(arrErrors[a] + "\r");
  file.close();
  alert("Some errors found, check Error report!");
}

// ---------------------------------------------------------------------------
// User settings: final file names + visible rectangle colors
// ---------------------------------------------------------------------------

function trimStr(s) {
  return String(s === undefined || s === null ? "" : s).replace(/^\s+|\s+$/g, "");
}

function cleanNamePart(v) {
  var s = trimStr(v).replace(/[\r\n\t]+/g, " ");
  s = s.replace(/[\/\,\:\;`\!\?\>\<\[\\\|\]\}\{\*\^\$\&\"\@\%\']/g, "_");
  return trimStr(s);
}

function getDimensionsName(row) {
  var tw = Number(row[COL_TOT_W]), th = Number(row[COL_TOT_H]);
  var vw = Number(row[COL_VIS_W]), vh = Number(row[COL_VIS_H]);
  if (tw == vw && th == vh) return tw + "x" + th;
  return tw + "x" + th + "_" + vw + "x" + vh;
}

// value of one name token for a data row
function getTokenValue(token, row, s) {
  if (token.type === "custom") return cleanNamePart(s.customText);
  if (token.type === "dims") return getDimensionsName(row);
  var c = token.type === "namecol" ? s.nameColumn : token.col;
  var v = row[c];
  if (c === COL_MASTER) v = trimStr(v).replace(/\.indd$/i, "");
  v = cleanNamePart(v);
  if (c === COL_TPOSM && v.length <= 2) v = "";
  return v;
}

// bleed: final bleed of the document in row units; if omitted the bleed column is used
// (empty column = bleed of the master, known only after the master is opened)
function buildFinalName(r, s, bleed) {
  var row = dataList[r];
  var parts = [];
  if (s.numbering === "auto") {
    parts.push(padWithZeros(s.startNumber + r, s.startNumber + dataList.length - 1));
  } else if (s.numbering === "excel") {
    var n = cleanNamePart(row[0]);
    if (n) parts.push(n);
  }
  var tokens = s.nameSource === "column" ? s.tokensColumn : s.tokensCompose;
  for (var t = 0; t < tokens.length; t++) {
    if (!tokens[t].on) continue;
    var v = getTokenValue(tokens[t], row, s);
    if (v) parts.push(v);
  }
  if (s.bleedInName && parts.length > 0) {
    if (bleed === undefined && row[COL_BLEED] !== "") bleed = Number(row[COL_BLEED]);
    if (bleed !== undefined) parts.push(bleed + unitName(row[COL_UNITS]) + "bleed");
  }
  return parts.join(s.separator);
}

function getOrAddCmykColor(doc, name, cmyk) {
  var col = doc.colors.itemByName(name);
  if (col.isValid) {
    col.properties = { model: ColorModel.PROCESS, space: ColorSpace.CMYK, colorValue: cmyk };
    return col;
  }
  return doc.colors.add({ name: name, colorValue: cmyk, model: ColorModel.PROCESS, space: ColorSpace.CMYK });
}

// ---------------------------------------------------------------------------
// Remember last used naming settings (saved when "Start" is pressed)
// ---------------------------------------------------------------------------

function namingSettingsFile() {
  return File(Folder.userData + "/Redimensionari_multiMasters_naming.txt");
}

function loadNamingSettings() {
  var f = namingSettingsFile();
  if (!f.exists) return null;
  var o = {};
  try {
    f.encoding = "UTF-8";
    if (!f.open("r")) return null;
    while (!f.eof) {
      var line = f.readln();
      var p = line.indexOf("=");
      if (p > 0) o[line.substr(0, p)] = decodeURIComponent(line.substr(p + 1));
    }
    f.close();
  } catch (e) {
    try { f.close(); } catch (e2) {}
    return null;
  }
  return o;
}

function tokenKey(tk) {
  return tk.type === "col" ? "col" + tk.col : tk.type;
}

function serializeTokens(tokens) {
  var a = [];
  for (var i = 0; i < tokens.length; i++) a.push(tokenKey(tokens[i]) + ":" + (tokens[i].on ? 1 : 0));
  return a.join(",");
}

// returns a new token list in the saved order with the saved on/off states;
// tokens not present in the saved list are appended in their default order
function applySavedTokens(tokens, str) {
  if (!str) return tokens;
  var byKey = {};
  for (var i = 0; i < tokens.length; i++) byKey[tokenKey(tokens[i])] = tokens[i];
  var out = [];
  var parts = str.split(",");
  for (var j = 0; j < parts.length; j++) {
    var kv = parts[j].split(":");
    var tk = byKey[kv[0]];
    if (!tk) continue;
    tk.on = kv[1] === "1";
    out.push(tk);
    delete byKey[kv[0]];
  }
  for (var k = 0; k < tokens.length; k++) if (byKey[tokenKey(tokens[k])]) out.push(tokens[k]);
  return out;
}

function saveNamingSettings(s) {
  var lines = [
    "numbering=" + s.numbering,
    "startNumber=" + s.startNumber,
    "nameSource=" + s.nameSource,
    "nameColumn=" + s.nameColumn,
    "tokensCompose=" + serializeTokens(s.tokensCompose),
    "tokensColumn=" + serializeTokens(s.tokensColumn),
    "customText=" + encodeURIComponent(s.customText),
    "separator=" + encodeURIComponent(s.separator),
    "bleedInName=" + (s.bleedInName ? 1 : 0)
  ];
  var f = namingSettingsFile();
  try {
    f.encoding = "UTF-8";
    if (f.open("w")) {
      f.write(lines.join("\n"));
      f.close();
    }
  } catch (e) {
    try { f.close(); } catch (e2) {}
  }
}

function showSettingsDialog(data, header) {
  var nCols = header.length;
  for (var i = 0; i < data.length; i++) if (data[i].length > nCols) nCols = data[i].length;

  function colLetter(c) {
    var s = "";
    c++;
    while (c > 0) {
      var m = (c - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      c = Math.floor((c - 1) / 26);
    }
    return s;
  }
  function colTitle(c) {
    var h = trimStr(header[c]).replace(/[\r\n]+/g, " ");
    if (!h && DEFAULT_HEADER[c]) h = DEFAULT_HEADER[c];
    return colLetter(c) + " - " + (h || "(no header)");
  }

  var colTitles = [];
  for (var c = 0; c < nCols; c++) colTitles.push(colTitle(c));

  // tokens for "compose" mode; column A is handled by the numbering option
  var tokensCompose = [
    { type: "col", col: COL_CAMPAIGN, on: true },
    { type: "col", col: COL_TPOSM, on: true },
    { type: "col", col: COL_MATERIAL, on: false },
    { type: "dims", on: true },
    { type: "col", col: COL_UNITS, on: true },
    { type: "col", col: COL_MASTER, on: false },
    { type: "custom", on: true }
  ];
  var used = {};
  used[COL_TPOSM] = used[COL_MATERIAL] = used[COL_UNITS] = used[COL_MASTER] = used[COL_CAMPAIGN] = 1;
  used[COL_BLEED] = 1; // bleed has its own option, added at the end of the name
  for (var c2 = 1; c2 < nCols; c2++) if (!used[c2]) tokensCompose.push({ type: "col", col: c2, on: false });
  // remove default tokens pointing to columns that don't exist
  for (var t = tokensCompose.length - 1; t >= 0; t--) {
    if (tokensCompose[t].type === "col" && tokensCompose[t].col >= nCols) tokensCompose.splice(t, 1);
  }
  var tokensColumn = [{ type: "namecol", on: true }, { type: "custom", on: true }];

  // last used naming settings
  var saved = loadNamingSettings();
  if (saved) {
    tokensCompose = applySavedTokens(tokensCompose, saved.tokensCompose);
    tokensColumn = applySavedTokens(tokensColumn, saved.tokensColumn);
  }

  // ---------- window ----------
  var w = new Window("dialog", "Redimensionare - settings");
  w.orientation = "column";
  w.alignChildren = ["fill", "top"];

  w.add("statictext", undefined, data.length + " rows read from Excel, " + arrMasterName.length + " master file(s) verified.");

  // 1. numbering
  var pNum = w.add("panel", undefined, "1. Numbering of final files");
  pNum.orientation = "row";
  pNum.alignChildren = ["left", "center"];
  pNum.margins = [10, 15, 10, 10];
  var rbAuto = pNum.add("radiobutton", undefined, "Auto-number, start at:");
  var etStart = pNum.add("edittext", undefined, "1");
  etStart.characters = 5;
  var rbExcelNum = pNum.add("radiobutton", undefined, "Numbers from " + colTitles[0]);
  var rbNoNum = pNum.add("radiobutton", undefined, "No number");
  var numRadios = [rbAuto, rbExcelNum, rbNoNum];
  rbAuto.value = true;

  // 2. final name
  var pName = w.add("panel", undefined, "2. Final name");
  pName.orientation = "column";
  pName.alignChildren = ["fill", "top"];
  pName.margins = [10, 15, 10, 10];

  var gSrc = pName.add("group");
  gSrc.alignChildren = ["left", "center"];
  var rbCompose = gSrc.add("radiobutton", undefined, "Compose from columns");
  var rbColumn = gSrc.add("radiobutton", undefined, "Take final name from column:");
  var ddCol = gSrc.add("dropdownlist", undefined, colTitles);
  ddCol.selection = Math.min(COL_FINALNAME, nCols - 1); // column L by default
  var srcRadios = [rbCompose, rbColumn];
  rbCompose.value = true;

  var gList = pName.add("group");
  gList.alignChildren = ["left", "top"];
  var lb = gList.add("listbox", undefined, [], { multiselect: false });
  lb.preferredSize = [420, 200];
  var gListBtns = gList.add("group");
  gListBtns.orientation = "column";
  gListBtns.alignChildren = ["fill", "top"];
  var btnToggle = gListBtns.add("button", undefined, "Include / Exclude");
  var btnUp = gListBtns.add("button", undefined, "Move up");
  var btnDown = gListBtns.add("button", undefined, "Move down");
  pName.add("statictext", undefined, "[x] = used in name. Double-click to include/exclude. List order = order in the name.");

  var gText = pName.add("group");
  gText.alignChildren = ["left", "center"];
  gText.add("statictext", undefined, "Custom text:");
  var etCustom = gText.add("edittext", undefined, "");
  etCustom.characters = 28;
  gText.add("statictext", undefined, "Separator:");
  var etSep = gText.add("edittext", undefined, "_");
  etSep.characters = 3;
  var cbBleedName = pName.add("checkbox", undefined, "Add bleed at the end of the name (e.g. 5mmbleed)");
  cbBleedName.value = true;

  var gPrev = pName.add("group");
  gPrev.alignChildren = ["left", "center"];
  var btnPreview = gPrev.add("button", undefined, "Preview name");
  var etPreview = gPrev.add("edittext", undefined, "", { readonly: true });
  etPreview.characters = 55;

  // 3. visible rectangle
  var pRect = w.add("panel", undefined, "3. Visible rectangle (CMYK 0-100)");
  pRect.orientation = "column";
  pRect.alignChildren = ["left", "top"];
  pRect.margins = [10, 15, 10, 10];

  function addCmykRow(parent, label, cmyk) {
    var g = parent.add("group");
    g.alignChildren = ["left", "center"];
    var first = g.add(label === null ? "checkbox" : "statictext", undefined, label === null ? "Add fill" : label);
    first.preferredSize.width = 70;
    var fields = [];
    var names = ["C", "M", "Y", "K"];
    for (var f = 0; f < 4; f++) {
      g.add("statictext", undefined, names[f]);
      var et = g.add("edittext", undefined, String(cmyk[f]));
      et.characters = 4;
      fields.push(et);
    }
    var sw = g.add("group");
    sw.preferredSize = [30, 20];
    sw.rgb = [1, 1, 1];
    sw.onDraw = function () {
      var gr = this.graphics;
      gr.newPath();
      gr.rectPath(0, 0, this.size[0], this.size[1]);
      gr.fillPath(gr.newBrush(gr.BrushType.SOLID_COLOR, this.rgb));
    };
    function refresh() {
      var v = readCmyk(fields);
      if (!v) return;
      sw.rgb = [(1 - v[0] / 100) * (1 - v[3] / 100), (1 - v[1] / 100) * (1 - v[3] / 100), (1 - v[2] / 100) * (1 - v[3] / 100)];
      sw.hide();
      sw.show();
    }
    for (var f2 = 0; f2 < 4; f2++) fields[f2].onChanging = refresh;
    refresh();
    return { group: g, first: first, fields: fields };
  }

  var rowStroke = addCmykRow(pRect, "Stroke", [0, 100, 0, 0]);
  var rowFill = addCmykRow(pRect, null, [0, 100, 0, 0]);
  var cbFill = rowFill.first;
  rowFill.group.add("statictext", undefined, "Opacity %");
  var etOpacity = rowFill.group.add("edittext", undefined, "20");
  etOpacity.characters = 4;
  var gVisMode = pRect.add("group");
  gVisMode.alignChildren = ["left", "center"];
  var rbVisShow = gVisMode.add("radiobutton", undefined, "Show \"vizibil\" layer");
  var rbVisHide = gVisMode.add("radiobutton", undefined, "Hide \"vizibil\" layer");
  var rbVisIgnore = gVisMode.add("radiobutton", undefined, "Ignore visible feature");
  rbVisIgnore.helpTip = "No \"vizibil\" layer, no rectangle, no visible guides, page margins stay like in the master files.";
  var visRadios = [rbVisShow, rbVisHide, rbVisIgnore];
  rbVisShow.value = true;

  // 4. page margins
  var pMarg = w.add("panel", undefined, "4. Page margins");
  pMarg.alignChildren = ["left", "top"];
  pMarg.margins = [10, 15, 10, 10];
  var rbMargVisible = pMarg.add("radiobutton", undefined, "Margins = visible area");
  var rbMargKeep = pMarg.add("radiobutton", undefined, "Keep the page margins from the master files (same values)");
  var rbMargScale = pMarg.add("radiobutton", undefined, "Scale the page margins from the master files with the new page size");
  var margRadios = [rbMargVisible, rbMargKeep, rbMargScale];
  rbMargVisible.value = true;

  // buttons
  var gBtns = w.add("group");
  gBtns.alignment = ["right", "top"];
  var btnCancel = gBtns.add("button", undefined, "Cancel", { name: "cancel" });
  var btnOk = gBtns.add("button", undefined, "Start", { name: "ok" });

  // ---------- behaviour ----------
  function readCmyk(fields) {
    var v = [];
    for (var f = 0; f < 4; f++) {
      var n = Number(trimStr(fields[f].text).replace(/,/g, "."));
      if (trimStr(fields[f].text) === "" || isNaN(n) || n < 0 || n > 100) return null;
      v.push(n);
    }
    return v;
  }

  function exclusive(radios) {
    for (var k = 0; k < radios.length; k++) {
      radios[k].onClick = (function (me) {
        return function () {
          for (var q = 0; q < radios.length; q++) radios[q].value = radios[q] === me;
          updateState();
        };
      })(radios[k]);
    }
  }
  exclusive(numRadios);
  exclusive(srcRadios);
  exclusive(visRadios);
  exclusive(margRadios);

  function currentTokens() {
    return rbColumn.value ? tokensColumn : tokensCompose;
  }

  function tokenLabel(tk) {
    if (tk.type === "custom") return "[Custom text]";
    if (tk.type === "dims") return "[Dimensions: total x visible]";
    if (tk.type === "namecol") return "Final name from " + (ddCol.selection ? ddCol.selection.text : "?");
    return colTitles[tk.col];
  }

  function fillList(selIndex) {
    var tokens = currentTokens();
    lb.removeAll();
    for (var k = 0; k < tokens.length; k++) lb.add("item", (tokens[k].on ? "[x]  " : "[  ]  ") + tokenLabel(tokens[k]));
    if (selIndex !== undefined && selIndex >= 0 && selIndex < tokens.length) lb.selection = selIndex;
  }

  function updateState() {
    etStart.enabled = rbAuto.value;
    ddCol.enabled = rbColumn.value;
    var ignoreVis = rbVisIgnore.value;
    for (var f = 0; f < 4; f++) {
      rowStroke.fields[f].enabled = !ignoreVis;
      rowFill.fields[f].enabled = cbFill.value && !ignoreVis;
    }
    cbFill.enabled = !ignoreVis;
    etOpacity.enabled = cbFill.value && !ignoreVis;
    // ignoring the visible feature always keeps the master margins
    if (ignoreVis && rbMargVisible.value) {
      rbMargVisible.value = false;
      rbMargKeep.value = true;
    }
    rbMargVisible.enabled = !ignoreVis;
    fillList(lb.selection ? lb.selection.index : undefined);
  }

  function toggleSelected() {
    if (!lb.selection) return;
    var idx = lb.selection.index;
    currentTokens()[idx].on = !currentTokens()[idx].on;
    fillList(idx);
  }

  function moveSelected(dir) {
    if (!lb.selection) return;
    var tokens = currentTokens();
    var idx = lb.selection.index;
    var to = idx + dir;
    if (to < 0 || to >= tokens.length) return;
    var tmp = tokens[idx];
    tokens[idx] = tokens[to];
    tokens[to] = tmp;
    fillList(to);
  }

  lb.onDoubleClick = toggleSelected;
  btnToggle.onClick = toggleSelected;
  btnUp.onClick = function () { moveSelected(-1); };
  btnDown.onClick = function () { moveSelected(1); };
  ddCol.onChange = function () { if (rbColumn.value) fillList(lb.selection ? lb.selection.index : undefined); };
  cbFill.onClick = updateState;

  function collectSettings() {
    var start = parseInt(trimStr(etStart.text), 10);
    return {
      numbering: rbAuto.value ? "auto" : (rbExcelNum.value ? "excel" : "none"),
      startNumber: isNaN(start) ? 1 : start,
      nameSource: rbColumn.value ? "column" : "compose",
      nameColumn: ddCol.selection ? ddCol.selection.index : 0,
      tokensCompose: tokensCompose,
      tokensColumn: tokensColumn,
      customText: etCustom.text,
      separator: etSep.text,
      bleedInName: cbBleedName.value,
      strokeCMYK: readCmyk(rowStroke.fields),
      fillEnabled: cbFill.value,
      fillCMYK: readCmyk(rowFill.fields),
      fillOpacity: Number(trimStr(etOpacity.text).replace(/,/g, ".")),
      visibleMode: rbVisIgnore.value ? "ignore" : (rbVisHide.value ? "hide" : "show"),
      marginMode: rbMargScale.value ? "scale" : ((rbMargKeep.value || rbVisIgnore.value) ? "keep" : "visible")
    };
  }

  btnPreview.onClick = function () {
    if (data.length === 0) { etPreview.text = "(no data rows)"; return; }
    var name = buildFinalName(0, collectSettings());
    etPreview.text = name ? name + "_.indd" : "(empty name for row 1)";
  };

  var result = null;
  btnOk.onClick = function () {
    var s = collectSettings();
    var start = parseInt(trimStr(etStart.text), 10);
    if (s.numbering === "auto" && (isNaN(start) || start < 0)) { alert("Start number must be a positive whole number."); return; }
    if (s.visibleMode !== "ignore" && !s.strokeCMYK) { alert("Stroke CMYK values must be numbers between 0 and 100."); return; }
    if (s.visibleMode !== "ignore" && s.fillEnabled && !s.fillCMYK) { alert("Fill CMYK values must be numbers between 0 and 100."); return; }
    if (s.visibleMode !== "ignore" && s.fillEnabled && (isNaN(s.fillOpacity) || s.fillOpacity < 0 || s.fillOpacity > 100)) { alert("Fill opacity must be a number between 0 and 100."); return; }

    if (s.nameSource === "column") {
      var emptyRows = [];
      for (var r = 0; r < data.length; r++) if (!cleanNamePart(data[r][s.nameColumn])) emptyRows.push(r + 2);
      if (emptyRows.length > 0) {
        alert("Column " + colTitles[s.nameColumn] + " is empty at excel line(s):\n" + emptyRows.join(", "));
        return;
      }
    }

    var seen = {}, dupes = [], emptyNames = [];
    for (var r2 = 0; r2 < data.length; r2++) {
      var nm = buildFinalName(r2, s);
      if (!nm) { emptyNames.push(r2 + 2); continue; }
      if (seen[nm]) dupes.push(nm);
      seen[nm] = true;
    }
    if (emptyNames.length > 0) { alert("The final name is empty at excel line(s):\n" + emptyNames.join(", ")); return; }
    if (dupes.length > 0) {
      if (!confirm("Some final names are duplicated, files will be overwritten:\n" + dupes.slice(0, 10).join("\n") + (dupes.length > 10 ? "\n..." : "") + "\n\nContinue anyway?")) return;
    }
    saveNamingSettings(s);
    result = s;
    w.close(1);
  };
  btnCancel.onClick = function () { w.close(0); };

  // restore last used naming settings in the dialog
  if (saved) {
    if (saved.numbering) {
      rbAuto.value = saved.numbering === "auto";
      rbExcelNum.value = saved.numbering === "excel";
      rbNoNum.value = saved.numbering === "none";
    }
    if (saved.startNumber !== undefined && saved.startNumber !== "") etStart.text = saved.startNumber;
    if (saved.nameSource) {
      rbColumn.value = saved.nameSource === "column";
      rbCompose.value = !rbColumn.value;
    }
    var savedCol = parseInt(saved.nameColumn, 10);
    if (!isNaN(savedCol) && savedCol >= 0 && savedCol < nCols) ddCol.selection = savedCol;
    if (saved.customText !== undefined) etCustom.text = saved.customText;
    if (saved.separator !== undefined) etSep.text = saved.separator;
    if (saved.bleedInName !== undefined) cbBleedName.value = saved.bleedInName === "1";
  }

  updateState();
  fillList(0);
  btnPreview.onClick();

  if (w.show() !== 1) return null;
  return result;
}

function aliniere_zece_la_suta(myParent, mychild) {
  var pm = getPageMetrics();
  var w = pm.w, h = pm.h, w_t = pm.w_t, h_t = pm.h_t;
  var m_left = pm.m_left, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Y_HW_final = m_top + h - HW_h_final;
  var Y_HW_final_latotal = h - HW_h_final + (h_t - h) / 2;

  myParent.redefineScaling([1, 1]);
  mychild.redefineScaling([1, 1]);
  myParent.geometricBounds = [Y_HW_final, m_left, h + m_top, w + m_left];
  app.documents[0].align(mychild, AlignOptions.HORIZONTAL_CENTERS, AlignDistributeBounds.MARGIN_BOUNDS);
  myParent.fit(FitOptions.CENTER_CONTENT);

  var W_hw = myParent.geometricBounds[3] - myParent.geometricBounds[1];
  var W_hwai = mychild.geometricBounds[3] - mychild.geometricBounds[1];
  var H_hwai = mychild.geometricBounds[2] - mychild.geometricBounds[0];
  var anchor = AnchorPoint.CENTER_ANCHOR;
  var pw, matrix;
// ## probleme la hw!!!
  // if (w > h) {
  //   pw = (HW_h_final * 0.7) / H_hwai;
  //   matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
  //   mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  // }
  // if (w < h) {
  //   myParent.redefineScaling([1, 1]);
  //   mychild.redefineScaling([1, 1]);
  //   pw = (W_hw * 0.8) / W_hwai;
  //   matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
  //   mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  // }

  // W_hwai = mychild.geometricBounds[3] - mychild.geometricBounds[1];
  // H_hwai = mychild.geometricBounds[2] - mychild.geometricBounds[0];

  // if (W_hwai > w - w * 0.1) {
  //   myParent.redefineScaling([1, 1]);
  //   mychild.redefineScaling([1, 1]);
  //   pw = (W_hw * 0.8) / W_hwai;
  //   matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
  //   mychild.transform(CoordinateSpaces.pasteboardCoordinates, anchor, matrix);
  // }

  mychild.fit(FitOptions.CENTER_CONTENT);
  mychild.fit(FitOptions.frameToContent);

  if (mychild.label && mychild.label.length > 2) {
    var heightofchildprocent = Number(mychild.label.substring(2));
    if (!isNaN(heightofchildprocent)) {
      fixtheHeightOfChild(myParent, mychild, heightofchildprocent);
    }
  }

  myParent.fit(FitOptions.CENTER_CONTENT);
  myParent.geometricBounds = [Y_HW_final_latotal, -getDocBleed(), h_t + getDocBleed(), w_t + getDocBleed()];
  }

  function fixtheHeightOfChild(obj, child, heightprocent) {
    var parentH = obj.geometricBounds[2] - obj.geometricBounds[0];
    var childH = child.geometricBounds[2] - child.geometricBounds[0];
    var targetH = parentH * (heightprocent / 100);
    var pw = targetH / childH;
    var matrix = app.transformationMatrices.add({ horizontalScaleFactor: pw, verticalScaleFactor: pw });
    child.transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.CENTER_ANCHOR, matrix);
  }




function aliniere_cinci_la_suta(myParent, mychild) {
  var pm = getPageMetrics();
  var h = pm.h, w_t = pm.w_t, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Hdt_final = (h - HW_h_final) * DtiSize;
  var Y_DT_final = m_top + h - (HW_h_final + Hdt_final);

  myParent.redefineScaling([1, 1]);
  mychild.redefineScaling([1, 1]);
  myParent.geometricBounds = [Y_DT_final, -getDocBleed(), Y_DT_final + Hdt_final, w_t + getDocBleed()];
  myParent.fit(FitOptions.CENTER_CONTENT);
  mychild.fit(FitOptions.frameToContent);
}

function aliniere_linia(myLine) {
  var pm = getPageMetrics();
  var h = pm.h, w_t = pm.w_t, m_top = pm.m_top;

  var HW_h_final = h * HwSize;
  var Hdt_final = (h - HW_h_final) * DtiSize;
  var lineY = m_top + h - (HW_h_final + Hdt_final) + Hdt_final;

  myLine.redefineScaling([1, 1]);
  myLine.geometricBounds = [lineY, -getDocBleed(), lineY, w_t + getDocBleed()];
  myLine.strokeWeight = 1;
}
