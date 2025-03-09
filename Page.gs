class Page {
  constructor(slidefile, tagname) {
    this.tag = tagname;
    var pagelist = [];
    this.slidefile = slidefile;
    var slides = slidefile.getSlides();
    slides.forEach(function (slide) {
      if (slide.getNotesPage().getSpeakerNotesShape().getText().asString().includes(tagname)) {
        pagelist.push(slide);
      }
    });
    this.pageList = pagelist.concat();
  }

  getTag() {
    return this.tag;
  }
  doReplaceText(calendar, start, end) {

  }

  insertTag() {
    // noteにタグを設定する
    var selection = SlidesApp.getActivePresentation().getSelection();
    selection.getCurrentPage().asSlide().getNotesPage().getSpeakerNotesShape().getText().setText(this.tag);
    return "Set SpeakerNote:" + this.tag;
  }
}


class SinglePage extends Page {
  /**
   * コンストラクタ
   * @param {SlidesApp.Presentation} slides 変換するファイル
   */
  constructor(slides = null) {
    if (slides == null) {
      slides = SlidesApp.getActivePresentation();
    }
    super(slides, "[CPage]");
  }

  doReplaceText(calendar, start, end) {
    var sdate = new Date(start);
    var edate = new Date(end);
    // 終了日の翌日0:00
    edate.setDate(edate.getDate() + 1);
    var events = calendar.getEvents(sdate, edate);
    var event = events[0];
    var date = sdate;
    var table = new ETable();
    var list = new EList();


    this.pageList.forEach(function (page) {
      //文字列置換
      table.replace(page, calendar, start, end, events, date);
      list.replace(page, calendar, start, end, events, date);
      replaceTag(page, calendar, start, end, event, date);
    });
  }
}

class DayPage extends Page {
  constructor(slides = null) {
    if (slides == null) {
      slides = SlidesApp.getActivePresentation();
    }
    super(slides, "[DPage]");
  }

  doReplaceText(calendar, start, end) {

    var sdate = new Date(start);
    var edate = new Date(end);
    var days = Math.ceil((edate - sdate) / (1000 * 60 * 60 * 24));
    var table = new ETable();
    var list = new EList();
    var file = this.slidefile;

    this.pageList.forEach(function (orig) {
      for (var i = 0; i <= days; i++) {
        var date_i = new Date(sdate);
        date_i.setDate(sdate.getDate() + i);
        var events = calendar.getEventsForDay(date_i);

        var event_page = file.appendSlide(orig);
        table.replace(event_page, calendar, start, end, events, date_i);
        list.replace(event_page, calendar, start, end, events, date_i);
        replaceTag(event_page, calendar, start, end, events[0], date_i);
      }
      orig.remove();
    });
  }
}

class EventPage extends Page {
  constructor(slides = null) {
    if (slides == null) {
      slides = SlidesApp.getActivePresentation();
    }
    super(slides, "[EPage]");
  }

  doReplaceText(calendar, start, end) {
    var sdate = new Date(start);
    var edate = new Date(end);
    var days = Math.ceil((edate - sdate) / (1000 * 60 * 60 * 24));
    var file = this.slidefile;

    this.pageList.forEach(function (orig) {
      for (var i = 0; i <= days; i++) {
        var date_i = new Date(sdate);
        date_i.setDate(sdate.getDate() + i);
        var events = calendar.getEventsForDay(date_i);
        events.forEach(function (event) {
          var event_page = file.appendSlide(orig);
          replaceTag(event_page, calendar, start, end, event, date_i);
          var map = new EMap();
          map.replace(event_page, calendar, start, end, event, date_i);
        });
      }
      orig.remove();
    });
  }
}
