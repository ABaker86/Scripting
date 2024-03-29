function main() {
  createPinterestPins();
}

function getRowToPerformActionOn(sheet){
  var data = sheet.getDataRange().getValues().length;
  //var data = dataRange.getValues();
  for(var row = 2; row < data; row = row + 1){
    // Assuming first row is header
    var shouldPerformAction = sheet.getRange(row, 4).getValue();
    if(shouldPerformAction == 0)
    {
      return row;
    }
  }
  return 0;
}

function createPinterestPins() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();

  var rowIndexToPerformActionOn = getRowToPerformActionOn(sheet);
  var topic = "";
  var pinText = "";

  if(rowIndexToPerformActionOn > 1)
  {
    const COLUMN_SUB_CATEGORY = 2;
    const COLUMN_TITLE = 3; 
    const COLUMN_CREATION_ATTEMPTED = 4;
    const HAS_ATTEMPTED = 1;

    topic = sheet.getRange(rowIndexToPerformActionOn, COLUMN_SUB_CATEGORY).getValue();
    pinText = sheet.getRange(rowIndexToPerformActionOn, COLUMN_TITLE).getValue();
    var imageData = getPexelsImageUrl(topic); // Function to query Pexels API for image URL

    if(imageData.src != null & imageData.src != undefined){
      var imageUrl = imageData.src.large;

      if (imageUrl != null && imageUrl != undefined) {
        var pinImage = createPinPresentation(pinText, imageUrl, imageData.photographer, topic);
        sheet.getRange(rowIndexToPerformActionOn, COLUMN_CREATION_ATTEMPTED).setValue(HAS_ATTEMPTED);
      }
    }
    
  }
}

function getPexelsImageUrl(topic) {
  try{
    console.log('getPexelsImageUrUlr for topic: '+ topic)  ;
    var imageIndex = RAND_NUM(1,15);
    
    var API_KEY = '******************';
    var url = 'https://api.pexels.com/v1/search?query=' + encodeURIComponent(topic) + '&orientation=portrait&per_page='+imageIndex;
    var headers = {
      Authorization: API_KEY
    };
    var response = UrlFetchApp.fetch(url, { headers: headers });
    var data = JSON.parse(response.getContentText());
    if (data.total_results > 0) {
      return data.photos[imageIndex-1];
    } else {
      return null;
    }
  }
  catch(error){
    console.log('getPexelsImageUrUlr: '+ error)  
    return null;
  }
}

function getSlideTemplate() {
  var PIN_TEMPLATE_ID = "******************";

  var source = SlidesApp.openById(PIN_TEMPLATE_ID);

  var numerOfSlides = source.getSlides().length;
  console.log("number of slides in presentation: "+numerOfSlides);

  source.getSlides().forEach(s => s.remove());//remove all slides before we add a new one
  
  source.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  return source.getSlides()[0];
}

function addPinHeadlineText(slide, pinText, topic){
  var fontSize = 60;
  var maxNoCharactersPerLine = 12;
  var textCharacterCount = pinText.length;
  var padding = 100;
  var estimatedShapeHeight = ((textCharacterCount / maxNoCharactersPerLine) * fontSize) + padding;
  var shapeWidth = 500;

  var rand = Math.floor((Math.random()*100)+1);

  // Add a shape for text
  //left, top, width, height
  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 0, 200, shapeWidth, estimatedShapeHeight); // Adjust dimensions as needed
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE)
  var fillColor = shape.getFill().setSolidFill('#000000', 0.75);
  //var fillColor = shape.getFill().setSolidFill('#ffffff', 0.3);
  
  //align left
  //align right
  //center middle
  //center

  if(rand > 50){
    shape.alignOnPage(SlidesApp.AlignmentPosition.HORIZONTAL_CENTER);
  }else if(rand > 50 && rand < 70)
  {
    shape.alignOnPage(SlidesApp.AlignmentPosition.VERTICAL_CENTER);
  }
  else{
    shape.setLeft(750-shapeWidth);
  }
  
  // Get the text range of the shape and set the text
  var textRange = shape.getText();
  textRange.setText(pinText);

  // Set text properties
  textRange.getTextStyle().setFontSize(fontSize); // Adjust font size as needed
  textRange.getTextStyle().setForegroundColor('#ffffff'); // Set font color to white
  textRange.getTextStyle().setBold(true);

  var paragraphs = textRange.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var paragraphStyle = paragraphs[i].getRange().getParagraphStyle();
    paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    //blue: #89CFF0
    //yellow: #FFFF00
    //purple: #8c52ff
    //green: #00bf63
    //black: #000000
    //white: #ffffff
    //should update to make negative words read and positive words green??
    //var lineCount = paragraphs.length-1;=
    const PURPLE = "#8c52ff";
    const WHITE = "#ffffff";
    const GREEN = "#00bf63";

    var rand = RAND_NUM(1,3)
    var color = "#ffffff";
    switch(rand){
      case 1:
        color = PURPLE;
        break;
      case 2: 
        color = GREEN;
      default: 
        color = WHITE;
    }
    paragraphs[i].getRange().getTextStyle().setForegroundColor(color);
   
  } 
}

function createPinPresentation(pinText, imageUrl, name, topic) {
  Logger.log(`Attempting creation for : '${pinText}'`)
  var slide = getSlideTemplate();

  // Add the background image
  var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  //slide.insertImage(imageBlob);
  slide.getBackground().setPictureFill(imageBlob);

  addRandomEffects(slide);
  addAttribution(slide, name);
  //width = 1000 height = 1500 750x1125 points
  addPinHeadlineText(slide, pinText, topic);
  addLogo(slide);
  toPNG(pinText);
  return 0; //should return image or path to image
}

function addRandomEffects(slide){
  //should implement randomness here
  addBackgroundOverlayImage(slide);
  addVignetteLayer(slide);
  addTopBanner(slide);
  addFooterBanner(slide);
  
}

function addTopBanner(slide){
  //left, top, width, height (page dimentions are based on points not pixels)
  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 0, 0, 750, 75);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE)
  var fillColor = shape.getFill().setSolidFill('#000000', 1);
}

function addFooterBanner(slide){
  //left, top, width, height (page dimentions are based on points not pixels) 750x1125
  var footer_height = 75;
  var footer_top = 1125-footer_height;
  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 0, footer_top, 750, footer_height);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE)
  var fillColor = shape.getFill().setSolidFill('#000000', 1);
}

function addVignetteLayer(slide){
  const vignetteImageId = "******************";
  var file = DriveApp.getFileById(vignetteImageId);
  var imageUrl = file.getDownloadUrl();
 
  var x = slide.insertImage(imageUrl);
  x.setWidth(750);
  x.setHeight(1125);
  x.alignOnPage(SlidesApp.AlignmentPosition.CENTER);
}

function RAND_NUM(min,max){
	return Math.floor(Math.random()*(max-min+1))+min;
}

function toPNG(pinText){
  const folderId = "******************";
  const presentationId = ""******************";";
  
  SlidesApp.openById(presentationId).saveAndClose();
  // Get all slides in the presentation
  var slides = Slides.Presentations.get(presentationId).slides;

  // Assuming you want to get the first slide
  var slideId = slides[0].objectId;
  
  // Log the object IDs of all slides
  for (var i = 0; i < slides.length; i++) {
    //Logger.log("Slide " + (i + 1) + " Object ID: " + slides[i].objectId);
    slideId = slides[i].objectId
  }
  var PIN_TEMPLATE_ID = "******************";

  var source = SlidesApp.openById(PIN_TEMPLATE_ID);
  var numerOfSlides = source.getSlides().length;
  console.log("number of slides in presentation: "+numerOfSlides);

  var slideBlob = Slides.Presentations.Pages.getThumbnail(
    presentationId, 
    slideId, 
    {
      "thumbnailProperties.mimeType": "PNG",
      "thumbnailProperties.thumbnailSize": "LARGE",
    }
  );

  var height = 1500;
  var bUrl = slideBlob.contentUrl.replace("s1600", `s${height}`)//reset height to 1500 to fit pin height.
  //aspect ratio should be retained automatically.

  var blob = UrlFetchApp.fetch(bUrl).getBlob();
  var name = pinText
    .replace("'","")
    .replace(",","")
    .replace(".","")
    .replace("!","")
    .replace("?","")
    ;
  
  // Generate a unique file name with timestamp
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  var fileName = name +" "+ timestamp + ".png";

  DriveApp.getFolderById(folderId).createFile(blob).setName(fileName);
    
}

function addBackgroundOverlayImage(slide){
  const imageId = "**************";
  var file = DriveApp.getFileById(imageId);
  var imageUrl = file.getDownloadUrl();
  if(RAND_NUM(1,10)%2 == 0)
  {
    var x = slide.insertImage(imageUrl);
    x.setWidth(750);
    x.setHeight(1125);
    x.alignOnPage(SlidesApp.AlignmentPosition.CENTER);
  }  
}

function addAttribution(slide,name){
  //left, top, width, height (page dimentions are based on points not pixels) 750x1125
  var footer_height = 75;
  var footer_top = 1125-footer_height-50;
  
  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 0, footer_top, 750, footer_height);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE)
  //var fillColor = shape.getFill().setSolidFill('#000000', 1);
  
  var attributionText = `Photo by ${name} on Pexels`;
  // Get the text range of the shape and set the text
  var textRange = shape.getText();
  textRange.setText(attributionText);

  var paragraphs = textRange.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var paragraphStyle = paragraphs[i].getRange().getParagraphStyle();
    paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  }

  // Set text properties
  var fontSize = 14;
  textRange.getTextStyle().setFontSize(fontSize); // Adjust font size as needed
  textRange.getTextStyle().setForegroundColor('#ffffff'); // Set font color to white
  textRange.getTextStyle().setBold(true);
}

function addLogo(slide){
  const imageId = "**********";
  var file = DriveApp.getFileById(imageId);
  var imageUrl = file.getDownloadUrl();

  var dimension = 70;
  
  var x = slide.insertImage(imageUrl);
  x.setWidth(dimension);
  x.setHeight(dimension);
  x.setTop(1125-dimension-5);
  x.alignOnPage(SlidesApp.AlignmentPosition.HORIZONTAL_CENTER);
  
}

function convertBlobToBase64(blob){
	return Utilities.base64Encode(blob.getBytes());
}
