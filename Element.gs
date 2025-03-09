class Element {
  constructor(name) {
    this.tag = name;
  }
  getTag() {
    return this.tag;
  }
  insertTag() {
    var msg = "set alt text:" + this.tag;
    console.log(this.tag);
    // 代替テキストにタグを設定
    var selection = SlidesApp.getActivePresentation().getSelection();
    var stype = selection.getSelectionType();
    if (stype == SlidesApp.SelectionType.TEXT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      item.setTitle(this.tag);
    }
    else if (stype == SlidesApp.SelectionType.PAGE_ELEMENT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      item.setTitle(this.tag);
    }
    else if (stype == SlidesApp.SelectionType.TABLE_CELL) {
      var item = selection.getTableCellRange().getTableCells()[0];
      item.getParentTable().setTitle(this.tag);
    }
    else {
      msg = "Plese select element."
    }
    return msg;
  }
}

class ETable extends Element {
  constructor() {
    super('[EventTable]');
  }

  replace(slide, calendar, start, end, eventList, date) {
    var tables = slide.getTables();
    var name = this.tag;

    if (tables) {
      tables.forEach(function (table) {
        if (table.getTitle().includes(name)) {
          // 最終行を取得
          var rows = table.getNumRows();
          var temp_row = table.getRow(rows - 1);
          var colums = table.getNumColumns();
          eventList.forEach(function (event) {
            var row = table.appendRow();
            for (var col = 0; col < colums; col++) {
              var text = temp_row.getCell(col).getText().asString();
              row.getCell(col).getText().setText(text);
              replaceTag(row.getCell(col).getText(), calendar, start, end, event, date);
            }
          });
          temp_row.remove();
        }
      });
    }
  }

  insertTag() {
    var msg = "";
    // 選択対象を表に限定
    var selection = SlidesApp.getActivePresentation().getSelection();
    var stype = selection.getSelectionType();
    if (stype == SlidesApp.SelectionType.TEXT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.TABLE) {
        msg = super.insertTag();
      }
      else {
        msg = "Please select table element.";
      }
    }
    else if (stype == SlidesApp.SelectionType.PAGE_ELEMENT) {
      item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.TABLE) {
        msg = super.insertTag();
      }
      else {
        msg = "Please select table element.";
        console.log(item.getPageElementType().toString());
      }
    }
    else if (stype == SlidesApp.SelectionType.TABLE_CELL) {
      msg = super.insertTag();
    }
    else {
      msg = "Please select table element.";
      console.log(stype.toString());
    }
    return msg;
  }
}

class EList extends Element {
  constructor() {
    super('[EventList]');
  }

  /**
   * イベント情報を箇条書きで出力する
   * @param {SlidesApp.Page} slide 出力ページ
   * @param {CalendarApp.Calendar} calendar カレンダー
   * @param {String} start 開始日
   * @param {String} end 終了日
   * @param {CalendarApp.CalendarEvent[]}  eventList 表示するイベントのリスト
   * @param {Date} ページに出力するイベントの日付
   */
  replace(slide, calendar, start, end, eventList, date) {
    var shapes = slide.getShapes();
    var name = this.tag;

    if (shapes) {
      shapes.forEach(function (shape) {
        if (shape.getTitle().includes(name)) {
          // 最終行を取得
          var text = shape.getText();
          if (text.getListStyle().getList()) {
            var textList = text.getListStyle().getList().getListParagraphs();
            var rows = textList.length;
            var temp_row = textList[rows - 1].getRange().asString();
            var newText = text.asString();

            eventList.forEach(function (event) {
              //var row = text.appendRange(temp_row);
              text.setText(newText);
              replaceTag(text, calendar, start, end, event, date);
              newText = text.asString() + temp_row;
            });
          }
          else {
            
          }
        }
      });
    }
  }
  insertTag() {
    var msg = "";
    // 選択対象を図形に限定
    var selection = SlidesApp.getActivePresentation().getSelection();
    var stype = selection.getSelectionType();
    if (stype == SlidesApp.SelectionType.TEXT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.SHAPE && item.asShape().getText().getListStyle().getList()) {
        msg = super.insertTag();
      }
      else {
        msg = "Please select shape element.";
        console.log(item.getPageElementType().toString());
      }
    }
    else if (stype == SlidesApp.SelectionType.PAGE_ELEMENT) {
      item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.SHAPE && item.asShape().getText().getListStyle().getList()) {
        msg = super.insertTag();
      }
      else {
        msg = "Please select shape element.";
        console.log(item.getPageElementType().toString());
      }
    }
    else {
      msg = "Please select shape element.";
      console.log(stype.toString());
    }
    return msg;
  }
}

class EMap extends Element {
  constructor() {
    super('[Emap]');
  }

  replace(slide, calendar, start, end, event, date) {
    var images = slide.getImages();
    var name = this.tag;
    images.forEach(function (image) {
      if (image.getTitle().includes(name)) {
        var height = image.getHeight();
        var width = image.getWidth();

        var address = event.getLocation();
        if (address) {
          var map = Maps.newStaticMap()
            .setLanguage('ja')
            .setSize(width, height)
            .setZoom(15)
            .setCenter(address)
            .addMarker(address);
          var blob = map.getBlob();
          image.replace(blob);
        }
        else {
          image.remove();
        }
      }
    });
  }

  insertTag() {
    var msg = "";
    // 選択対象を画像に限定
    var selection = SlidesApp.getActivePresentation().getSelection();
    var stype = selection.getSelectionType();
    if (stype == SlidesApp.SelectionType.PAGE_ELEMENT) {
      var item = selection.getPageElementRange().getPageElements()[0];
      if (item.getPageElementType() == SlidesApp.PageElementType.IMAGE) {
        msg = super.insertTag();
      }
      else {
        msg = "Please select image.";
        console.log(item.getPageElementType().toString());
      }
    }
    else {
      msg = "Please select image.";
      console.log(stype.toString());
    }
    return msg;
  }
}