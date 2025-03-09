function createEditCard_en(msg="") {
  var builder = CardService.newCardBuilder();

  // "Edit" tag buttons 
  var editSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setText("Please select the slide element first and then click on the information you want to import.")
      .setWrapText(true)
    );

  var calSection1 = CardService.newCardSection()
    .setHeader("Calendar attributes");


  calSection1.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Title')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CTitle"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Start date')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CStart"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('End date')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CEnd"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Description')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CDesc"}))
      .setDisabled(false))
  )

  var calSection2 = CardService.newCardSection()
    .setHeader("Attributes for daily output pages");


  calSection2.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Days')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CDays"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Date')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "Cdate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('DOW')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "Cweekday"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('DOW in JPN')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "CweekdayJPN"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Event table')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#f4f4f4")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "ETable"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Event list')
      .setBackgroundColor("#f4f4f4")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EList"}))
      .setDisabled(false))
  )

  var calSection3 = CardService.newCardSection()
    .setHeader("Event attributes");


  calSection3.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Start time')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EStart"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('End time')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EEnd"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Start date')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EStartDate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('End date')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EEndDate"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Title')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "ETitle"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Description')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EDesc"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Map')
      .setBackgroundColor("#f4f4f4")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EMap"}))
      .setDisabled(false))
  )

  var calSection4 = CardService.newCardSection()
    .setHeader("Page attributes");

  calSection4.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Single page')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "SinglePage"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Daily page')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "DayPage"}))
      .setDisabled(false))
    .addButton(CardService.newTextButton()
      .setText('Event page')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#3c9200")
      .setOnClickAction(CardService.newAction().setFunctionName('buttonClick').setParameters({id: "EventPage"}))
      .setDisabled(false))
  )

  editSection.setHeader("Edit template");
  createHeader(builder);
  builder.addSection(editSection)
  builder.addSection(calSection1);
  builder.addSection(calSection2);
  builder.addSection(calSection3);
  builder.addSection(calSection4);
  builder.addSection(CardService.newCardSection().addWidget(CardService.newDecoratedText().setText(msg)));

  return builder.build();


}