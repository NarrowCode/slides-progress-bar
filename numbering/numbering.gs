/**
 * Copyright 2022 Andreas Steinkellner
 * Copyright Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 * The basic idea for the implementation of this code was obtained
 * from these two tutorial pages:
 * https://developers.google.com/apps-script/add-ons/editors/slides/quickstart/progress-bar
 * and was modified by Andreas Steinkellner to:
 *    o) Add basic configuration
 */

// [START apps_script_slides_progress]

// Internal, do not change!
var TXT_ID = 'RATIO_TEXT_ID';

// Sizes (pixels)
var TXT_WIDTH = 100;
var TXT_HEIGHT = 20;
var FONT_SIZE = 8;
var MARGIN_BOTTOM = 8;
var MARGIN_RIGHT = 8;
var DELIMITER = " / ";    // What to put between the current slide and max

// Color
var FONT_COLOR = '#636363';

// Skip parameters
var SKIP_START = 0;       // Number of slides to skip at the start
var SKIP_END = 0;         // Number of slides to exclude at the end

function onInstall(e) {
  onOpen();
}

function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
    .addItem('Add numbering', 'createNumbering')
    .addItem('Remove numbering', 'deleteNumbering')
    .addToUi();
}

function createNumbering() {
  deleteNumbering();
  var pres = SlidesApp.getActivePresentation();
  var slides = pres.getSlides();

  var count = 1;
  for (var i = SKIP_START; i < slides.length - SKIP_END; ++i) {
    var outStr = count.toString() + DELIMITER + (slides.length - (SKIP_START + SKIP_END)).toString();
    var txtBox = slides[i].insertTextBox(outStr, 0, pres.getPageHeight() - TXT_HEIGHT - MARGIN_BOTTOM, TXT_WIDTH, TXT_HEIGHT);

    // Configure text box
    txtBox.getText().getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    txtBox.getText().getTextStyle().setFontSize(FONT_SIZE).setBold(true).setForegroundColor(FONT_COLOR);
    txtBox.setLeft(pres.getPageWidth() - txtBox.getWidth() - MARGIN_RIGHT);
    txtBox.setLinkUrl(TXT_ID);

    count++;
  }
}

function deleteNumbering() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; ++j) {
      var el = elements[j];
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
          el.asShape().getLink() && el.asShape().getLink().getUrl() === TXT_ID) {
        el.remove();
      }
    }
  }
  return true;
}
// [END apps_script_slides_progress]