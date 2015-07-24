var modulesPath = "/Volumes/RESOURCE/node_modules/"

var fs = require('fs');
var hummus = require(modulesPath + 'hummus');
var mv = require(modulesPath + 'mv');
var path = require('path');

//define global variables for both CACP and CHCP

var pageWidth = 14.33 * 72,
pageHeight = 22.5 * 72,
singlePageHeight = 2.25 * 72;

// script expects one parameter to be passed to it: the path to the file to be imposed
var fileToImpose = process.argv[2];

// get the dimensions of the file to be imposed
// need to create a pdf file to check dimensions of a file. we will use the "tempFile" if it's CHCP to create a new clipped pdf or delete it if its CACP
var tempFile = process.env.HOME + "/tempNodeFile.pdf"
var pdfWriter = hummus.createWriter(tempFile);
var dimensions = pdfWriter.getImageDimensions(fileToImpose)
var singlePageWidth = dimensions.width/72

var rotate180 = {transformation:[-1,0,0,-1,4*72,2.25*72]}
var rotate90 = {transformation:[0,1,-1,0,2.25*72,0]}
var rotate270 = {transformation:[0,-1,1,0,0,singlePageWidth*72]}

// if product is 4 inches wide, product is CHCP, so run that imposition
if(singlePageWidth == 4) {
	imposeCHCP();
}
//if product is 3.75 inches wide, product is CACP, so run that imposition
else if(singlePageWidth == 3.75) {
	imposeCACP();
}




/*======= Function Declarations =======*/

//define imposition functions
function imposeCHCP(){
	// builds the finished pdf name based on the path to the file to impose
	var parentFolder = path.dirname(fileToImpose);
	var jobNumber = path.basename(parentFolder)
	var finishedPDFName = parentFolder + "/" + jobNumber + "_CHCP.pdf"
	var page = pdfWriter.createPage(0,0,singlePageWidth*72,singlePageHeight);
	var cxt = pdfWriter.startPageContentContext(page);
	cxt.m(11.484,0).l(276.009,0).l(276.009,81.693).l(288,93.761).l(288,134.525).l(156.137,162).l(133.5,162).l(94.055,154.646).l(90.413,162).l(27.603,162).l(27.609,141.695).l(0,133.147).l(0,96.274).l(11.476,88.978).h().W().n().drawImage(0,0,fileToImpose);
	pdfWriter.writePage(page).end();
	mv(tempFile,fileToImpose,{clobber:true},function(err){
		if(err) callback(err);
		//creates the file for the imposed PDF, a new page, and a content context for writing to.
		var pdfWriter = hummus.createWriter(finishedPDFName);
		var page = pdfWriter.createPage(0,0,pageWidth,pageHeight);
		var cxt = pdfWriter.startPageContentContext(page);
		//these three are the bottom row. they are laid out differently than the rest so I didn't bother making a loop
		cxt.drawImage(79.299,45.012,fileToImpose);
		cxt.drawImage(407.03,9,fileToImpose,rotate180)
		cxt.drawImage(734.758 ,45.012,fileToImpose);

		// this is the loop for the rest of the imposition

		var xPos = 77.75,
		yPos = 207.007,
		startSmallXPos = 77.75,
		startBigXPos = 113.764,
		bigXIncrement = 211.405,
		smallXIncrement = 162.095,
		yIncrement = 274.502,
		columnLimit = 5,
		rowLimit = 5;

		for(i = 0; i < rowLimit; i++) {
			//if its on an even row, do this loop, which starts with the image rotated 90º
			if(i%2 == 0) {
				for(j = 0; j < columnLimit; j++) {
					//draws the image rotated 90 or 270 degrees alternating
					if(j%2 == 0) {
						cxt.drawImage(xPos,yPos,fileToImpose,rotate90);
						// increment x value
						xPos = xPos + bigXIncrement;
					}
					else {
						cxt.drawImage(xPos,yPos,fileToImpose,rotate270);
						// increment x value
						xPos = xPos + smallXIncrement;
					}
				}
			}
			//if it's on an odd row, do this loop, which starts with the image rotated 270º
			else {
				for(j = 0; j < columnLimit; j++) {
					//draws the image rotated 90 or 270 degrees alternating
					if(j%2 == 0) {
						cxt.drawImage(xPos,yPos,fileToImpose,rotate270);
						// increment x value
						xPos = xPos + smallXIncrement;
					}
					else {
						cxt.drawImage(xPos,yPos,fileToImpose,rotate90);
						// increment x value
						xPos = xPos + bigXIncrement;
					}
				}
			}
			//increment y value and reset the x value so next loop is the next row
			yPos = yPos + yIncrement;
			if(i%2 != 0) { // if its on an even row, the first image is placed closer to the x axis
				xPos = startSmallXPos;
			}
			else { // if its on an odd row, the first image is placed further from the x axis
				xPos = startBigXPos;
			}
		}
		// Writes the page and ends writing to the file
		pdfWriter.writePage(page).end();
	});
}

function imposeCACP() {
	//delete the tempfile that we used to determine the dimensions of the pdf to impose
	fs.unlink(tempFile);

	// builds the finished pdf name based on the path to the file to impose
	var parentFolder = path.dirname(fileToImpose);
	var jobNumber = path.basename(parentFolder)
	var finishedPDFName = parentFolder + "/" + jobNumber + "_CACP.pdf"

	//creates the file for the imposed PDF, a new page, and a content context for writing to.
	var pdfWriter = hummus.createWriter(finishedPDFName);
	var page = pdfWriter.createPage(0,0,pageWidth,pageHeight);
	var cxt = pdfWriter.startPageContentContext(page);

	// defining some variables for the imposition loop
	var xPos = 113.759,
	yPos = 8.991,
	startXPos = 113.759,
	bigXIncrement = 211.499,
	smallXIncrement = 162.001,
	yIncrement = 268.483,
	columnLimit = 5,
	rowLimit = 6;

	//the imposition loop. inside loop creates a row, outside loop increments that up and makes all the rows
	for(i = 0; i < rowLimit; i++) {
		for(j = 0; j < columnLimit; j++) {
			//if its on an even column, draw it rotated 90º and increment the small increment
			if(j%2 == 0) {
				cxt.drawImage(xPos,yPos,fileToImpose,rotate270);
				xPos = xPos + smallXIncrement;
			}
			//if its on an odd column, draw it rotated 270º and increment the big increment
			else {
				cxt.drawImage(xPos,yPos,fileToImpose,rotate90);
				xPos = xPos + bigXIncrement;
			}
		}
		//increment y value and reset the x value so next loop is the next row
		yPos = yPos + yIncrement;
		xPos = startXPos
	}
	// Writes the page and ends writing to the file
	pdfWriter.writePage(page).end();
}
