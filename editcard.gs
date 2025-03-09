function createEditCard(msg="") {
  var builder = CardService.newCardBuilder();

  // "Edit" tag buttons 
  var editSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setText("先に要素を選択してから取り込む情報をクリックしてください")
      .setWrapText(true)
    );

  var calSection1 = CardService.newCardSection()
    .setHeader("カレンダーの属性");


  calSection1.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('カレンダー名')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CTitle"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('開始日')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CStart"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('終了日')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CEnd"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('説明')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CDesc"}))
      .setDisabled(false))
  )

  var calSection2 = CardService.newCardSection()
    .setHeader("日単位出力ページ用属性");


  calSection2.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('日数')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CDays"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('日付')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "Cdate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('曜日（英語）')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "Cweekday"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('曜日（日本語）')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CweekdayJPN"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('表')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#f4f4f4")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "ETable"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('箇条書き')
      .setBackgroundColor("#f4f4f4")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EList"}))
      .setDisabled(false))
  )

  var calSection3 = CardService.newCardSection()
    .setHeader("イベント属性");


  calSection3.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('開始時刻')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EStart"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('終了時刻')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EEnd"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('開始日')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EStartDate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('終了日')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EEndDate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('イベント名')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "ETitle"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('説明')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EDesc"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('地図')
      .setBackgroundColor("#f4f4f4")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EMap"}))
      .setDisabled(false))
  )

  var calSection4 = CardService.newCardSection()
    .setHeader("ページ属性");

  calSection4.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('タイトル')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "SinglePage"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('日単位')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "DayPage"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('イベント詳細')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EventPage"}))
      .setDisabled(false))
  )

  editSection.setHeader("テンプレート作成");
  createHeader(builder);
  builder.addSection(editSection)
  builder.addSection(calSection1);
  builder.addSection(calSection2);
  builder.addSection(calSection3);
  builder.addSection(calSection4);
  builder.addSection(CardService.newCardSection().addWidget(CardService.newDecoratedText().setText(msg)));

  return builder.build();


}