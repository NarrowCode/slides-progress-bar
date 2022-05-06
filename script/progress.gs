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
 * https://developers.google.com/apps-script/add-ons/editors/slides/quickstart/translate
 * and was modified by Andreas Steinkellner to:
 *    o) Add additional functionality
 *    o) Simplify configuration
 *    o) Add html config dialog
 */

// [START apps_script_slides_progress]

// Internal, do not change!
var BAR_ID = 'PROGRESS_BAR_ID';
var TXT_ID = 'RATIO_TEXT_ID';

// Sizes (pixels)
var BAR_HEIGHT = 4;
var TXT_WIDTH = 100;
var TXT_HEIGHT = 20;
var FONT_SIZE = 8;
var TEXT_MARGIN = 8;

// Color
var FONT_COLOR = '#636363';

// Booleans
var LEFT_ALIGN = true;
var INCLUDE_PERCENTAGE = true;

// Skip parameters
var SKIP_START = 1; // Number of slides to skip before starting the numbering (0 = start on the first slide)
var SKIP_END = 1;   // Number of slides to exclude at the end (0 = go until the end)


function getDefaultSkipStart() {
  return SKIP_START;
}

function getDefaultSkipEnd() {
  return SKIP_END;
}

function getDefaultBarHeight() {
  return BAR_HEIGHT;
}

function getDefaultFontSize() {
  return FONT_SIZE;
}

function getDefaultFontColor() {
  return FONT_COLOR;
}

function getDefaultLeftAlign() {
  return LEFT_ALIGN;
}

function getDefaultTextMargin() {
  return TEXT_MARGIN;
}

function getDefaultIncludePercentage() {
  return INCLUDE_PERCENTAGE;
}

/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen();
}

/**
 * Trigger for opening a presentation.
 * @param {object} e The onOpen event.
 */
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Create default bar', 'createDefaultBars')
      .addItem('Hide progress bar', 'deleteBars')
      .addItem('Open menu', 'showSidebar')

      .addToUi();
}

function showSidebar() {
  const ui = HtmlService
      .createHtmlOutputFromFile('config')
      .setTitle('Settings');
  SlidesApp.getUi().showSidebar(ui);
}

function createDefaultBars() {
  createBars(SKIP_START, SKIP_END, BAR_HEIGHT, FONT_SIZE, FONT_COLOR, LEFT_ALIGN, TEXT_MARGIN, INCLUDE_PERCENTAGE);
}


/**
 * Creates bars + text with the specified settings
 */
function createBars(skipStart, skipEnd, barHeight, fontSize, fontColor, leftAlign, textMargin, includePercentage) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();

  // Check for nonsensical parameters
  if ((skipStart + skipEnd) > slides.length || barHeight < 0 || fontSize < 4)
    return;
  
  var count = 1;
  for (var i = skipStart; i < slides.length - skipEnd; ++i) {
    var ratioComplete = (count / (slides.length - (skipStart + skipEnd)));
    var x = 0;
    var y = presentation.getPageHeight() - barHeight;
    var barWidth = presentation.getPageWidth() * ratioComplete;
    if (barWidth > 0 && barHeight > 0) {
      var bar = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,
                                      barWidth, barHeight);
      bar.getBorder().setTransparent();
      bar.setLinkUrl(BAR_ID);
    }
    var progTextString = "";
    if (includePercentage)
      progTextString += Math.round(ratioComplete * 100) + "% - ";
    progTextString += count.toString() + " / " + (slides.length - (skipStart + skipEnd)).toString();
    var x = 0;
    if (!leftAlign)
      x = presentation.getPageWidth() - fontSize * progTextString.length;

    var progText = slides[i].insertTextBox(progTextString, x + textMargin, y - fontSize * 2.5, TXT_WIDTH, fontSize * 2.5);
    var insertedText = progText.getText();
    
    if (!leftAlign) {
      insertedText.getParagraphs()[0].getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
      progText.setLeft(presentation.getPageWidth() - progText.getWidth() - textMargin);
    }
    insertedText.getTextStyle()
      .setFontSize(fontSize)
      .setBold(true)
      .setForegroundColor(fontColor)
    progText.setLinkUrl(TXT_ID);
    count++;
  }
}

/**
 * Deletes all progress bar rectangles.
 */
function deleteBars() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; ++j) {
      var el = elements[j];
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
          el.asShape().getLink() &&
          (el.asShape().getLink().getUrl() === BAR_ID || el.asShape().getLink().getUrl() === TXT_ID)) {
        el.remove();
      }
    }
  }
  return true;
}
// [END apps_script_slides_progress]
