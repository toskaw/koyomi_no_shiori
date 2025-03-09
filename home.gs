
function createHeader(builder) {
  const cardHeader = CardService.newCardHeader()
    .setTitle('暦のしおり')
    .setSubtitle("KOYOMInoSHIORI")
    .setImageUrl('http://koyomi.castanet.tokyo/wp-content/uploads/sites/13/2025/03/miniicon-.png');
  builder.setHeader(cardHeader);
}

function createHomeCard(e) {
  console.log("createHomeCard");
  var hostApp = e['hostApp'];
  var builder = CardService.newCardBuilder();
  createHeader(builder);

  //Buttons section
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('テンプレート作成')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('edit'))
        .setDisabled(false))
      .addButton(CardService.newTextButton()
        .setText('実行')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('exec'))
        .setDisabled(false))));

  return builder.build();

}
