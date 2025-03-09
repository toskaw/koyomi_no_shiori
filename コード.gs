/**
 * Callback for rendering the main card.
 * @return {CardService.Card} The card to show the user.
 */
function onHomepage(e) {
  console.log(e.commonEventObject.userLocale);
  var card = null;
  if (e.commonEventObject.userLocale == "ja") {
    card = createHomeCard(e);
  }
  else {
    card = createHomeCard_en(e);
  }
  return card;
}

function onFileScopeGranted(e) {
  var card = null;
  if (e.commonEventObject.userLocale == "ja") {
    card = exec(e);
  }
  else {
    card = exec_en(e);
  }
  return card;
}
function formSubmit(e) {
  var calendar = e.commonEventObject.formInputs.calendar.stringInputs.value[0];
  var date_s = new Date(e.commonEventObject.formInputs.start.dateInput.msSinceEpoch);
  var date_e = new Date(e.commonEventObject.formInputs.end.dateInput.msSinceEpoch);

  var start = Utilities.formatDate(date_s, e.commonEventObject.timeZone, "yyyy-MM-dd");
  var end = Utilities.formatDate(date_e, e.commonEventObject.timeZone, "yyyy-MM-dd");
  var msg = execute(calendar, start, end);
  var locale = e.commonEventObject.userLocale;

  if (locale == "ja") {
    card = createExecCard(e, msg);
  }
  else {
    card = createExecCard_en(e, msg);
  }
  var nav = CardService.newNavigation().updateCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();

}

function execute(id, start, end) {
  console.log(id, start, end);
  var calendar = CalendarApp.getCalendarById(id);

  var slide = SlidesApp.getActivePresentation();

  // テンプレートスライドをコピーする
  //var sorceFile = DriveApp.getFileById(slide.getId());
  //var newFile = sorceFile.makeCopy();
  //newFile.setName(calendar.getName() + start + "-" + end);

  //var newSlide = SlidesApp.openById(newFile.getId());
  var newName = calendar.getName() + start + "-" + end;
  var newSlide = copySlidesFromTemplate(slide, newName);
  //createTitlePage(newSlide, calendar, start, end);
  //createDayPage(newSlide, calendar, start, end);
  //createEventPage(newSlide, calendar, start, end);
  var titlePages = new SinglePage(newSlide);
  var dayPages = new DayPage(newSlide);
  var eventPages = new EventPage(newSlide);

  titlePages.doReplaceText(calendar, start, end);
  dayPages.doReplaceText(calendar, start, end);
  eventPages.doReplaceText(calendar, start, end);
  var url = newSlide.getUrl();

  // いらないページを削除
  //newSlide.getSlides()[2].remove();
  //newSlide.getSlides()[1].remove();

  return '<a href="' + url + '" target="_blank">' + newName + '</a>'
}
/**
 * Callback function for a button action. Instructs Docs to display a
 * permissions dialog to the user, requesting `drive.file` scope for the
 * current file on behalf of this add-on.
 *
 * @param {Object} e The parameters object that contains the document’s ID
 * @return {editorFileScopeActionResponse}
 */
function onRequestFileScopeButtonClicked(e) {
  return CardService.newEditorFileScopeActionResponseBuilder()
    .requestFileScopeForActiveDocument().build();
}

function getAllCalendar() {
  let calendars = CalendarApp.getAllCalendars();
  return calendars;
}

function edit() {
  var card = createEditCard();

  var nav = CardService.newNavigation().pushCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();
}

function edit_en() {
  var card = createEditCard_en();

  var nav = CardService.newNavigation().pushCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();
}

function exec(e) {
  var card = createExecCard(e);
  var nav = CardService.newNavigation().pushCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();
}
function exec_en(e) {
  var card = createExecCard_en(e);
  var nav = CardService.newNavigation().pushCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();

}

function generateCalendarDropdown(fieldName, fieldTitle, calendars) {

  var selectionInput = CardService.newSelectionInput().setTitle(fieldTitle)
    .setFieldName(fieldName)
    .setType(CardService.SelectionInputType.DROPDOWN);

  calendars.forEach((obj) => {
    selectionInput.addItem(obj.getName(), obj.getId(), false);
  })

  return selectionInput;
}


function getClass(classname) {
  return Function('return (' + classname + ')')();
}


function buttonClick(param) {
  var id = param.parameters.id;
  console.log(id);
  var classobj = getClass(id);
  var obj = new classobj();
  var msg = obj.insertTag();
  var card = null;
  var locale = param.commonEventObject.userLocale;
  console.log(locale);
  if (locale == "ja") {
    card = createEditCard(msg);
  }
  else {
    card = createEditCard_en(msg);
  }
  var nav = CardService.newNavigation().updateCard(card);
  return CardService.newActionResponseBuilder()
    .setNavigation(nav)
    .build();

}

function createNewSlide(id, start, end) {
  console.log(id, start, end);
  var calendar = CalendarApp.getCalendarById(id);
  var slide = SlidesApp.getActivePresentation();

  // テンプレートスライドをコピーする
  //var sorceFile = DriveApp.getFileById(slide.getId());
  //var newFile = sorceFile.makeCopy();
  //newFile.setName(calendar.getName() + start + "-" + end);

  //var newSlide = SlidesApp.openById(newFile.getId());
  var newName = calendar.getName() + start + "-" + end;
  var newSlide = copySlidesFromTemplate(slide, newName);
  //createTitlePage(newSlide, calendar, start, end);
  //createDayPage(newSlide, calendar, start, end);
  //createEventPage(newSlide, calendar, start, end);
  var titlePages = new SinglePage(newSlide);
  var dayPages = new DayPage(newSlide);
  var eventPages = new EventPage(newSlide);

  titlePages.doReplaceText(calendar, start, end);
  dayPages.doReplaceText(calendar, start, end);
  eventPages.doReplaceText(calendar, start, end);
  var url = newSlide.getUrl();

  // いらないページを削除
  //newSlide.getSlides()[2].remove();
  //newSlide.getSlides()[1].remove();

  return '<a href="' + url + '" target="_blank">' + newName + '</a>'
}

function createTitlePage(slide, calendar, start, end) {
  var page = slide.getSlides()[0];
  var Title = new CTitle();
  var sdate = new Date(start);
  var events = calendar.getEventsForDay(sdate);
  var event = events[0];
  var date = sdate;

  //文字列置換
  replaceTag(page, calendar, start, end, event, date);
}

function createDayPage(slide, calendar, start, end) {
  var orig = slide.getSlides()[2];

  var sdate = new Date(start);
  var edate = new Date(end);
  var days = Math.ceil((edate - sdate) / (1000 * 60 * 60 * 24));
  var table = new ETable();
  var list = new EList();

  for (var i = 0; i <= days; i++) {
    var date_i = new Date(sdate);
    date_i.setDate(sdate.getDate() + i);
    var events = calendar.getEventsForDay(date_i);
    console.log(date_i);

    var event_page = slide.appendSlide(orig);
    table.replace(event_page, calendar, start, end, events, date_i);
    list.replace(event_page, calendar, start, end, events, date_i);
    replaceTag(event_page, calendar, start, end, events[0], date_i);
  }
}

function createEventPage(slide, calendar, start, end) {
  var orig = slide.getSlides()[2];

  var sdate = new Date(start);
  var edate = new Date(end);
  var days = Math.ceil((edate - sdate) / (1000 * 60 * 60 * 24));

  for (var i = 0; i <= days; i++) {
    var date_i = new Date();
    date_i.setDate(sdate.getDate() + i);
    var events = calendar.getEventsForDay(date_i);
    events.forEach(function (event) {
      var event_page = slide.appendSlide(orig);
      replaceTag(event_page, calendar, start, end, event, date_i);
      var map = new EMap();
      map.replace(event_page, calendar, start, end, event, date_i);
    });
  }
}

function get_youbi(date) {

  // 曜日の日本語化
  // 日本語の曜日格納用
  var todayDayStr = '';

  switch (date.getDay()) {
    case 0: todayDayStr = '日'; break;
    case 1: todayDayStr = '月'; break;
    case 2: todayDayStr = '火'; break;
    case 3: todayDayStr = '水'; break;
    case 4: todayDayStr = '木'; break;
    case 5: todayDayStr = '金'; break;
    case 6: todayDayStr = '土'; break;
  }
  return todayDayStr;
}

function getLocale() {
  console.log(Session.getActiveUserLocale());
  return Session.getActiveUserLocale();
}

function copySlidesFromTemplate(presentation, newName) {
  var src = presentation.getId();
  var copyFile = Drive.Files.copy({ "name": newName }, src);
  var newPresentation = SlidesApp.openById(copyFile.id);
  /*
    // 新しいGoogleスライドを作成
    var newPresentation = SlidesApp.create(newName);
    var deleteSlide = newPresentation.getSlides()[0]; // 1枚目のスライドを取得
    deleteSlide.remove(); // 1枚目のスライドを削除
  
    // テンプレートのスライドを取得し、新しいスライドにコピー
    var slides = presentation.getSlides();
    for (var j = 0; j < slides.length; j++) {
      newPresentation.appendSlide(slides[j]);
    }
  */
  return newPresentation;

}
function replaceTag(page, calendar, start, end, event, date) {
  TagsList.forEach(function (tag) {
    page.replaceAllText(tag.getTag(), tag.getValue(calendar, start, end, event, date));
  });
}


