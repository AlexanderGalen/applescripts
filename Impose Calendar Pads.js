var modulesPath = "/Volumes/RESOURCE/node_modules/"

var fs = require('fs');
var hummus = require(modulesPath + 'hummus');
var mv = require(modulesPath + 'mv');
var path = require('path');

// script expects two parameters to be passed to it: the path to the file to be imposed, and the shape of the file, either CACC or CACH
var fileToImpose = process.argv[2];
var shape = process.argv[3];


// function to impose a file. imposes slightly differently depending on the shape passed to it
// shape can either be "CACC" or "CACH".
function impose(shape) {

	// builds the finished pdf name based on the path to the file to impose
	var parentFolder = path.dirname(fileToImpose);
	var jobNumber = path.basename(parentFolder)
	var finishedPDFName = parentFolder + "/" + jobNumber + "." + shape + ".pdf"

	//creates the file for the imposed PDF, a new page, and a content context for writing to.
	var pdfWriter = hummus.createWriter(finishedPDFName);
	var page = pdfWriter.createPage(0,0,pageWidth,pageHeight);
	var cxt = pdfWriter.startPageContentContext(page);

	// make form IDs for xobject of image to impose
	var formIDs = pdfWriter.createFormXObjectsFromPDF(fileToImpose,hummus.ePDFPageBoxMediaBox);

	// defining some variables for the imposition loop
	var xPos, yPos, startXPos, bigXIncrement, smallXIncrement, yIncrement, columnLimit, rowLimit;

	if(shape == "CACC") {
		xPos = 113.76;
		yPos = 9;
		startXPos = 113.76;
		bigXIncrement = 211.5;
		smallXIncrement = 162;
		yIncrement = 266.4;
		columnLimit = 5;
		rowLimit = 6;
	}
	else {
		xPos = 77.748;
		yPos = 9;
		startXPos = 77.748;
		bigXIncrement = 211.5;
		smallXIncrement = 162;
		yIncrement = 319.464;
		columnLimit = 5;
		rowLimit = 5;
	}

	// the imposition loop. inside loop creates a row, outside loop increments that up and makes all the rows
	for(i = 0; i < rowLimit; i++) {
		for(j = 0; j < columnLimit; j++) {
			// add one to evenOddCounter if pad is house shaped, because those start rotated 270 instead of 90, or essentially, one more column further
			// this was just a quick/easy/effecient way of accounting for the different rotations in each shape, rather than creating a new function or something
			var evenOddCounter = j;
			if(shape == "CACH") evenOddCounter++;
			// if its on an even column, draw it rotated 90º and increment the small increment
			if(evenOddCounter%2 == 0) {
				console.log(3, "rotating 90º" )
				//cxt.drawImage(xPos,yPos,fileToImpose,rotate270);
				// rotate current context, then draw the xobject for the image to impose, then reset the context.
				cxt.q()
				.cm(1,0,0,1,xPos,yPos) // this translates state to current position
				.cm(0,-1,1,0,0,padWidth) // this rotates 270º
				.doXObject(page.getResourcesDictionary().addFormXObjectMapping(formIDs[0]));
				cxt.Q();
				xPos = xPos + smallXIncrement;
			}
			// if its on an odd column, draw it rotated 270º and increment the big increment
			else {
				console.log(3, "rotating 270º" )
				cxt.q()
				.cm(1,0,0,1,xPos,yPos) // this translates state to current position
				.cm(0,1,-1,0,padHeight,0) // this rotates 90º
				.doXObject(page.getResourcesDictionary().addFormXObjectMapping(formIDs[0]));
				cxt.Q();
				xPos = xPos + bigXIncrement;
			}
		}
		// increment y value and reset the x value so next loop is the next row
		yPos = yPos + yIncrement;
		xPos = startXPos
	}
	// Writes the page and ends writing to the file
	pdfWriter.writePage(page).end();
}

// define global variables for both CACC and CACH
var pageWidth = 14.33 * 72,
pageHeight = 22.5 * 72,
padHeight = 2.25 * 72;

// set padWidth based on the shape
var padWidth;
if (shape == "CACH") {
	padWidth = 4 * 72;
}
else {
	padWidth = 3.75 * 72;
}

// run imposition
impose(shape);

// delete printfile once we've imposed
// this is to prevent errors where quark fails to save over existing printfiles
fs.unlinkSync(fileToImpose);
