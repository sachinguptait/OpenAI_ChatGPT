function summarizeContent() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = spreadSheet.getSheetByName('Sheet1');
  const numberOfContents = sheetName.getLastRow();
  const completionsLink = "https://api.openai.com/v1/completions";

  let paramsToSummarise = {
    model: "text-davinci-003",
    prompt: "",
    temperature: 0,
    max_tokens: 204,
    top_p: 1.0,
    frequency_penalty: 0.0,
    presence_penalty: 0.0,
  }

  let options = {
    method : 'get',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: {},
    headers: {Authorization: "Bearer <<Token>>"},
  }

  for (let content = 2; content <= numberOfContents ; content++ ) {
    let promptContent = sheetName.getRange(content,1).getValue();
    //console.log("Question- ",promptContent);

    paramsToSummarise.prompt = promptContent + "\n\nBullet points"

   //Convert Java Script value to JSON
    options.payload = JSON.stringify(paramsToSummarise)

    //console.log("Payload",options.payload);

    //Sending ChatGPT Request
    let text = UrlFetchApp.fetch(completionsLink, options).getContentText()
   // console.log("Response-",text);

    let obj = JSON.parse(text)
    sheetName.getRange(content,2).setValue(obj.choices[0].text);
  }

}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ChatGPT Menu')
      .addItem('Execute', 'menuItem1').addToUi();
}

function menuItem1() {
  summarizeContent();
}

