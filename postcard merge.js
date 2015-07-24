var hummus = require('hummus');
var fs = require('fs');
var excel = require('excel');
var applescript = require('applescript');

var sourceDB, baseSourceFolder, baseDestinationFolder, destinationFile, i, j, k, thisRow;

sourceDB = "/Volumes/HOM/PRODUCTS/NOTE CARD CAFE/POSTCARDS/DBs/Build Postcard Print Files.xlsx";
baseSourceFolder = "/Volumes/HOM/PRODUCTS/NOTE CARD CAFE/POSTCARDS/printfiles/";
baseDestinationFolder = "/Volumes/HOM/PRODUCTS/NOTE CARD CAFE/POSTCARDS/merged/";

// parse excel file to get data
excel(sourceDB, function(err, data) {
	if(err) throw err;

	// big loop for each set
	for (i = 1; i < data.length-5; i=i+6) {

		thisRow = data[i];
		console.log(thisRow)
		var setName = thisRow[0];
		var packageInsert = baseSourceFolder + thisRow[1];
		var outputFile = baseDestinationFolder + setName + ".122page.pdf"

		var pdfWriter =	hummus.createWriter(outputFile);
		pdfWriter.appendPDFPagesFromPDF(packageInsert);

		// loop for going through all the singles of the set
		for (j = i; j <  i + 6; j++) {
			thisRow = data[j];
			var cardFront = baseSourceFolder + thisRow[2];
			var cardBack = baseSourceFolder + thisRow[3];
			console.log(cardFront)
			console.log(cardBack);
			// loop for adding each single 10 times
			for (k = 0; k < 10; k++) {
				pdfWriter.appendPDFPagesFromPDF(cardFront);
				pdfWriter.appendPDFPagesFromPDF(cardBack)
			};


		};

		pdfWriter.end();
	};

});
