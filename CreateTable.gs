function CreateTable() {
  let alertFlag = Browser.msgBox("テーブルシートを作成しますか？", Browser.Buttons.OK_CANCEL);
  if(alertFlag === "ok"){
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("DBInfo");
    
    //テーブル名一覧を取得
    let getTableName = sheet.getRange(3,1,sheet.getLastRow()).getValues();
    
    //空白行の削除
    for(let i = 0 ; i < getTableName.length ; i++){
      if(getTableName[i][0] == "" || getTableName[i][0] == null){
        getTableName.splice(i);
      }
    }
    
    addSheet(getTableName);

    sheet.activate();
  }
}

/**
 * 新規シートの作成
 */
function addSheet(tableName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //現在あるシートを全て取得
  const sheets = spreadsheet.getSheets();
  let sheetsName = [];
  for(let i = 0 ; i < sheets.length ; i ++){
    sheetsName.push(sheets[i].getName());
  }

  let errName = [];
  
  for(let i = 0 ; i < tableName.length ; i++){
    let flag = sheetsName.includes(tableName[i]);
    if(!sheetsName.includes(tableName[i][0])){
      //シートの追加
      const newSheet = spreadsheet.insertSheet();
      newSheet.setName(tableName[i]);
      //テンプレート作成
      let sheet = spreadsheet.getSheetByName(tableName[i]);
      addRules(sheet);
      const header = ["カラム名","データ型","データサイズ","デフォルト値","Extra","NULL","コメント","インデックス","AutoIncrement"];
      sheet.appendRow(header);
      sheetsName.push(tableName[i]);
    }else{
      errName.push(tableName[i]);
    }
  }
  
  if(errName.length != 0){
    let errMsg = ("既に存在するテーブルがあります。\\n[スキップしたテーブル]\\n");
    for(let i = 0 ; i < errName.length ; i++){
      errMsg += `${errName[i]}\\n`;
    }
    
    Browser.msgBox(errMsg);
  }
}

function addRules(sheet) {
  const types = ['int', 'varchar', 'date', 'boolean', 'tinyint', 'smallint', 'mediumint', 'bigint', 'decimal', 'float', 'double', 'real', 'bit', 'serial', 'datetime', 'timestamp', 'time', 'year','char', 'tinytext', 'text', 'mediumtext', 'longtext', 'binary', 'varbinary', 'tinyblob', 'blob', 'mediumblob', 'longblob', 'enum', 'set' , 'json'];
  const defaultValue = ['NULL', 'current_timestamp'];
  const booleanValue = ['TRUE'];
  const extraValue = ['ON UPDATE CURRENT_TIMESTAMP'];
  const indexValue = ['primary', 'unique', 'index', 'fullltext', 'spatial'];

  let rule = SpreadsheetApp.newDataValidation().requireValueInList(types).build();
  let cell = sheet.getRange(2, 2, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);

  rule = SpreadsheetApp.newDataValidation().requireValueInList(defaultValue).build();
  cell = sheet.getRange(2, 4, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);

  rule = SpreadsheetApp.newDataValidation().requireValueInList(extraValue).build();
  cell = sheet.getRange(2, 5, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);

  rule = SpreadsheetApp.newDataValidation().requireValueInList(booleanValue).build();
  cell = sheet.getRange(2, 6, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);
  cell = sheet.getRange(2, 9, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);

  rule = SpreadsheetApp.newDataValidation().requireValueInList(indexValue).build();
  cell = sheet.getRange(2, 8, sheet.getMaxRows(), 1);
  cell.setDataValidation(rule);
}





