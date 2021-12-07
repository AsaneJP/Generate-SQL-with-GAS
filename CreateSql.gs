function CreateSql() {
  let alertFlag = Browser.msgBox("SQLを作成しますか？", Browser.Buttons.OK_CANCEL);
  if(alertFlag === "ok"){
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    //シートが存在しない場合は作成する
    //現在あるシートを全て取得
    const sheets = spreadsheet.getSheets();
    let sheetFlag = true;
    for(let n = 0 ; n < sheets.length ; n ++){
      if(sheets[n].getName() == "exportSQL"){
        sheetFlag = false;
      }
    }

    if(sheetFlag){
      //シートの追加
      const newSheet = spreadsheet.insertSheet();
      newSheet.setName("exportSQL");
    }

    const sheet = spreadsheet.getSheetByName("exportSQL");
    
    let sql = "";
    sql = dbCreate(sql);
    sql = tableCreate(sql);
    sheet.getRange(1, 1).setValue(sql);
    //SQLファイルの移動
    sheet.activate();
    spreadsheet.moveActiveSheet(2);
  }
}

function dbCreate(sql) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("DBInfo");
  //DB名取得
  const dbName = sheet.getRange(1,2).getValue();

  sql += `CREATE DATABASE IF NOT EXISTS \`${dbName}\`;`
  sql += `\nUSE \`${dbName}\`;`;
  return sql;
}

function tableCreate(sql){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("DBInfo");

  //テーブル名一覧を取得
  let getTableName = sheet.getRange(3,1,sheet.getLastRow(),2).getValues();

  //空白行の削除
  for(let n = 0 ; n < getTableName.length ; n++){
    if(getTableName[n][0] == "" || getTableName[n][0] == null){
      getTableName.splice(n);
    }
  }
  console.log(getTableName);

  for(let i = 0 ; i < getTableName.length ; i++){
    let addSql = "";
    
    sql += "\n\n";
    sql += "--\n";
    sql += `-- テーブル：${getTableName[i][0]}\n`;
    sql += "--\n";
    sql += `CREATE TABLE \`${getTableName[i][0]}\` (\n`;
    sheet = spreadsheet.getSheetByName(getTableName[i][0]);

    //作成するTable情報を取得
    let getTableInfo = sheet.getRange(2,1,sheet.getLastRow(),9).getValues();

    //空白行の削除
    for(let n = 0 ; n < getTableInfo.length ; n++){
      if(getTableInfo[n][0] == "" || getTableInfo[n][0] == null || getTableInfo[n][1] == "" || getTableInfo[n][1] == null){
        getTableInfo.splice(n);
      }
    }
    console.log(getTableInfo);
    for(let j = 0 ; j < getTableInfo.length ; j++){
      //カラム名入力
      sql += `  \`${getTableInfo[j][0]}\``;
      
      //データ型入力
      sql += ` ${getTableInfo[j][1]}`;
      
      //データサイズの入力
      if(getTableInfo[j][2] != ""){
        sql += `(${getTableInfo[j][2]})`;
      }

      //デフォルトの入力
      if(getTableInfo[j][3] != ""){
        if(getTableInfo[j][3] == "NULL"){
          sql += " DEFAULT NULL";
        }else if(getTableInfo[j][3] == "current_timestamp"){
          sql += " DEFAULT current_timestamp()";
        }else{
          sql += ` DEFAULT ${getTableInfo[j][3]}`;
        }
      }else if(!getTableInfo[j][5]){
        sql += " NOT NULL";
      }

      if(getTableInfo[j][4] != ""){
        if(getTableInfo[j][4] == "ON UPDATE CURRENT_TIMESTAMP"){
          sql += " ON UPDATE current_timestamp()";
        }
      }

      //コメント入力
      if(getTableInfo[j][6] != ""){
        sql += ` COMMENT '${getTableInfo[j][6]}'`;
      }

      //ALTER TABLE
      if(getTableInfo[j][7] != ""){
        addSql += "\n";
        addSql += "-- インデックスの追加\n";
        if(getTableInfo[j][7] == "primary"){
          addSql += `ALTER TABLE \`${getTableName[i][0]}\`\n`;
          addSql += `  ADD PRIMARY KEY (\`${getTableInfo[j][0]}\`);\n`;
        }
      }
      if(getTableInfo[j][8] != ""){
        addSql += "\n";
        addSql += "-- オートインクリメントの追加\n";
        if(getTableInfo[j][8]){
          addSql += `ALTER TABLE \`${getTableName[i][0]}\`\n`;
          addSql += `  MODIFY \`${getTableInfo[j][0]}\` ${getTableInfo[j][1]}(${getTableInfo[j][2]}) NOT NULL AUTO_INCREMENT;\n`;
        }
      }

      //次のレコードへの追加事項
      if(j < getTableInfo.length - 1){
        sql += ",";
      }
      sql += "\n";
    }
    sql += ")";

    //コメントの追加
    if(getTableName[i][1] != ""){
      sql += ` COMMENT='${getTableName[i][1]}';\n`;
    }else{
      sql += ";\n";
    }

    if(addSql != ""){
      sql += addSql;
    }
  }

  return sql ;
}
