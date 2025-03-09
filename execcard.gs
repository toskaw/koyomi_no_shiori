function createExecCard(e, msg = null) {
  var builder = CardService.newCardBuilder();
  var docsEventObject = e['slides'];
  var builder = CardService.newCardBuilder();

  createHeader(builder);

  if (!docsEventObject['addonHasFileScopePermission']) {
    // If the add-on does not have access permission, add a button that
    // allows the user to provide that permission on a per-file basis.
    var cardSection = CardService.newCardSection();
    cardSection.addWidget(
      CardService.newTextParagraph().setText(
        "The add-on needs permission to access this file's quota."));

    var buttonAction = CardService.newAction()
      .setFunctionName("onRequestFileScopeButtonClicked");

    var button = CardService.newTextButton()
      .setText("Request permission")
      .setOnClickAction(buttonAction);

    cardSection.addWidget(button);
    return builder.addSection(cardSection).build();
  }

  var formSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
      .setText("カレンダーと期間を選択してください")
      .setWrapText(true)
    )
    .setHeader("レポート作成");

  var calendars = getAllCalendar();

  formSection.addWidget(generateCalendarDropdown("calendar", "カレンダーを選択してください", calendars));

  var startDate =
    CardService.newDatePicker()
      .setTitle('開始日')
      .setFieldName('start')
      .setValueInMsSinceEpoch(Date.now());
  var endDate =
    CardService.newDatePicker()
      .setTitle('終了日')
      .setFieldName('end')
      .setValueInMsSinceEpoch(Date.now());
  formSection.addWidget(startDate).addWidget(endDate);
  formSection.addWidget(
    CardService.newTextButton().
      setText("Exec").
      setOnClickAction(
        CardService.newAction()
          .setFunctionName("formSubmit")
      )
  );

  builder.addSection(formSection)
  if (msg) {
    builder.addSection(CardService.newCardSection().addWidget(CardService.newDecoratedText().setText(msg)));
  }

  return builder.build();


}