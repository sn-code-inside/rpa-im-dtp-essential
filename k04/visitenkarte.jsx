/**
 * RPA im DTP essential | Beispiel-Skript: Visitenkarte
 *
 * @summary Adobe InDesign-Skript zur Generierung einer beispielhaften Visitenkarte
 * @author Ennis Gundogan <e@gndgn.dev>
 */

app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;

var doc = app.documents.add();
doc.documentPreferences.pageHeight = 55;
doc.documentPreferences.pageWidth = 85;
doc.pageItemDefaults.properties = { 
	fillColor: "None", strokeColor: "None", strokeWidth: 0 
};
	
doc.pages[0].marginPreferences.properties = { 
	top: 6, left: 6, right: 6, bottom: 6 
};

var tf1 = doc.textFrames.add({geometricBounds: [30, 6, 55, 40]});
var tf2 = doc.textFrames.add({geometricBounds: [30, 45, 55, 79]});
var tfs = doc.textFrames;

for(i = 0; i < tfs.length; i++) {
	var txt = tfs[i].texts[0];
	txt.appliedFont = app.fonts.itemByName("Myriad Pro");
	txt.appliedLanguage = "Deutsch: 2006 Rechtschreibreform";
	txt.fontStyle = "Regular";
	txt.justification = Justification.LEFT_ALIGN;
	txt.leading = "9.5pt";
	txt.pointSize = "7.5pt";
}

tf1.contents = "Vorname Nachname\rPosition";
tf1.paragraphs[0].fontStyle = "Bold";
tf1.paragraphs[0].pointSize = "10pt";

tf2.contents = "Unternehmen GmbH\rStrasse 123\r45678 Ort\r";
tf2.contents += "Tel.: 0123 456 78-9\rMobil: 0123 45 678 90\r";
tf2.contents += "E-Mail: info@xyz.org";
tf2.paragraphs[0].fontStyle = "Bold";
tf2.texts[0].baselineShift = "-2pt";

for(i = 0; i < tfs.length; i++) {
	tfs[i].textFramePreferences.properties = {
		autoSizingReferencePoint: AutoSizingReferenceEnum.TOP_LEFT_POINT,
		autoSizingType: AutoSizingTypeEnum.HEIGHT_AND_WIDTH,
		useNoLineBreaksForAutoSizing: true
	};
}

var img = doc.pages.firstItem().rectangles.add({
	geometricBounds: [5, 40, 17, 81]
});

var imgfile = new File("~/Documents/RPA-im-DTP_Beispiele/K-4-2-2_Visitenkarte/Assets/RPA-im-DTP_essential_K-4-2-2_Logo-Dummy.pdf");

if(imgfile.exists) {
	img.place(imgfile);
	img.fit(FitOptions.PROPORTIONALLY);
} else {
	alert("Datei existiert nicht:\n" + imgfile);
}