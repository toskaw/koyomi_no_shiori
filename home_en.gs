
function createHeader_en(builder) {
  const cardHeader = CardService.newCardHeader()
    .setTitle('暦のしおり')
    .setSubtitle("KOYOMInoSHIORI")
    .setImageUrl('http://koyomi.castanet.tokyo/wp-content/uploads/sites/13/2025/03/miniicon-.png');
  builder.setHeader(cardHeader);
}

function createHomeCard_en(e) {
  console.log("createHomeCard");
  var hostApp = e['hostApp'];
  var builder = CardService.newCardBuilder();
  createHeader(builder);

  //Buttons section
  builder.addSection(CardService.newCardSection()
    .addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
        .setText('Edit template')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('edit_en'))
        .setDisabled(false))
      .addButton(CardService.newTextButton()
        .setText('Exec')
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(CardService.newAction().setFunctionName('exec_en'))
        .setDisabled(false))));

  return builder.build();

}
