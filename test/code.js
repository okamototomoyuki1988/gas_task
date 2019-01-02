//// 関数を実行するメニューを追加
//function onOpen() {
//  var ui = SpreadsheetApp.getUi();
//  var menu = ui.createMenu('メッセージ表示');
//  menu.addItem('Hello world! 実行', 'myFunction');
//  menu.addToUi();
//}
// 
//function myFunction() {
//  Browser.msgBox('Hello world!');
//}


function toTable() {  
  var from = toN('E');
  var to = toN('DH');
  for (var i = from; i <= to; i++)
  {
  
  var cell = SpreadsheetApp.getActive().getRange("a1:b2"); 
  }
}


function toA(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function toN(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}