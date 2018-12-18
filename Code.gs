/**
 * In the menu, head to Add-ons > Change document text colour > Choose colour. From the sidebar, choose your colour (the default is red), and when ready click Submit. The sidebar will close when your document has been updated.
 */

function onOpen(e) {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem("Choose colour", "openSidebar")
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("index").setTitle("Choose colour");
  SlidesApp.getUi().showSidebar(html);
}

function changeTextColor(color) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  // loop through slides
  for (var i = 0; i < slides.length; i++) {
    var currentSlide = slides[i];
    var elements = currentSlide.getPageElements();

    // loop through elements on current slide
    for (var j = 0; j < elements.length; j++) {
      var element = elements[j];
      var elementType = element.getPageElementType();
      
      // check if element is a shape type
      if (elementType === element.getPageElementType().SHAPE) {
        var shape = element.asShape();
        
        // check if shape is a textbox shape and contains text
        if (shape.getShapeType() === shape.getShapeType().TEXT_BOX 
          && shape.getText().getLength() > 1) {
            
          shape.getText().getTextStyle().setForegroundColor(color);        
        }
      }
    }
  }
}
