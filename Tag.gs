class tag {
  constructor(tag) {
    this.name = tag;
  }
  getValue(calendar, start, end, event, date) {
    return this.name;
  }
  getTag() {
    return this.name;
  }
  insertTag() {
    var msg = "";
    var selection = SlidesApp.getActivePresentation().getSelection();
    var stype = selection.getSelectionType();
    if (stype == SlidesApp.SelectionType.TEXT) {
      var range = selection.getTextRange();
      range = range.setText(this.name);
      range.appendText("").select();
    }
    else if (stype == SlidesApp.SelectionType.PAGE_ELEMENT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.SHAPE) {
        var range = item.asShape().getText().setText(this.name);
        range.appendText("").select();
      }
      else {
        msg = "Please select textbox.";
      }
    }
    else if (stype == SlidesApp.SelectionType.TABLE_CELL) {
      var item = selection.getTableCellRange().getTableCells()[0];
      var range = item.getText().setText(this.name);
      range.appendText("").select();
    }
    else {
      msg = "Please select textbox.";
    }
    return msg;
  }
}

const timeZone = Session.getScriptTimeZone();
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

class CTitle extends tag {
  constructor() {
    super("[Ctitle]");
  }
  getValue(calendar, start, end, event, date) {
    return calendar.getName();
  }
}

class CStart extends tag {
  constructor() {
    super("[Cstart]");
  }
  getValue(calendar, start, end, event, date) {
    return start;
  }
}

class CEnd extends tag {
  constructor() {
    super("[Cend]");
  }
  getValue(calendar, start, end, event, date) {
    return end;
  }
}

class CDesc extends tag {
  constructor() {
    super("[Cdesc]");
  }
  getValue(calendar, start, end, event, date) {
    return calendar.getDescription();
  }
}

class CDays extends tag {
  constructor() {
    super("[day]");
  }
  getValue(calendar, start, end, event, date) {
    var s = new Date(start);
    return Math.ceil((date - s) / (1000 * 60 * 60 * 24) + 1);
  }
}

class Cdate extends tag {
  constructor() {
    super("[date]");
  }
  getValue(calendar, start, end, event, date) {
    return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
  }
}

class Cweekday extends tag {
  constructor() {
    super("[weekday]");
  }
  getValue(calendar, start, end, event, date) {
    return Utilities.formatDate(date, timeZone, 'E');
  }
}

class CweekdayJPN extends tag {
  constructor() {
    super("[曜日]");
  }
  getValue(calendar, start, end, event, date) {
    return get_youbi(date);
  }
}

class EStart extends tag {
  constructor() {
    super("[Estart]");
  }
  getValue(calendar, start, end, event, date) {
    var e_start = event.getStartTime();
    return Utilities.formatDate(e_start, timeZone, 'HH:mm');
  }
}

class EEnd extends tag {
  constructor() {
    super("[Eend]");
  }
  getValue(calendar, start, end, event, date) {
    var e_end = event.getEndTime();
    return Utilities.formatDate(e_end, timeZone, 'HH:mm');
  }
}

class EStartDate extends tag {
  constructor() {
    super("[EstartDate]");
  }
  getValue(calendar, start, end, event, date) {
    var e_start = event.getStartTime();
    return Utilities.formatDate(e_start, timeZone, 'yyyy-MM-dd');
  }
}

class EEndDate extends tag {
  constructor() {
    super("[EendDate]");
  }
  getValue(calendar, start, end, event, date) {
    var e_end = event.getEndTime();
    return Utilities.formatDate(e_end, timeZone, 'yyyy-MM-dd');
  }
}
class ETitle extends tag {
  constructor() {
    super("[Etitle]");
  }
  getValue(calendar, start, end, event, date) {
    return event.getTitle();
  }
}

class EDesc extends tag {
  constructor() {
    super("[Edesc]");
  }
  getValue(calendar, start, end, event, date) {
    return event.getDescription();
  }
}

const TagsList = [
  new CTitle(),
  new CStart(),
  new CEnd(),
  new CDesc(),
  new CDays(),
  new Cdate(),
  new Cweekday(),
  new CweekdayJPN(),
  new EStart(),
  new EEnd(),
  new EStartDate(),
  new EEndDate(),
  new ETitle(),
  new EDesc()
];

