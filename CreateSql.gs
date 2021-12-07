function CreateSql() {
  let alertFlag = Browser.msgBox("SQLを作成しますか？", Browser.Buttons.OK_CANCEL);
  if(alertFlag === "ok"){
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    //シートが存在しない場合は作成する
    //現在あるシートを全て取得
    const sheets = spreadsheet.getSheets();
    let sheetFlag = true;
    for(let i = 0 ; i < sheets.length ; i ++){
      if(sheets[i].getName() == "exportSQL"){
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
    spreadsheet.moveActiveSheet(3);
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
  for(let i = 0 ; i < getTableName.length ; i++){
    if(getTableName[i][0] == "" || getTableName[i][0] == null){
      getTableName.splice(i);
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
    for(let i = 0 ; i < getTableInfo.length ; i++){
      if(getTableInfo[i][0] == "" || getTableInfo[i][0] == null || getTableInfo[i][1] == "" || getTableInfo[i][1] == null){
        getTableInfo.splice(i);
      }
    }
    console.log(getTableInfo);
    for(let i = 0 ; i < getTableInfo.length ; i++){
      //カラム名入力
      sql += `  \`${getTableInfo[i][0]}\``;
      
      //データ型入力
      sql += ` ${getTableInfo[i][1]}`;
      
      //データサイズの入力
      if(getTableInfo[i][2] != ""){
        sql += `(${getTableInfo[i][2]})`;
      }

      //デフォルトの入力
      if(getTableInfo[i][3] != ""){
        if(getTableInfo[i][3] == "NULL"){
          sql += " DEFAULT NULL";
        }else if(getTableInfo[i][3] == "current_timestamp"){
          sql += " DEFAULT current_timestamp()";
        }else{
          sql += ` DEFAULT ${getTableInfo[i][3]}`;
        }
      }else if(!getTableInfo[i][5]){
        sql += " NOT NULL";
      }

      if(getTableInfo[i][4] != ""){
        if(getTableInfo[i][4] == "ON UPDATE CURRENT_TIMESTAMP"){
          sql += " ON UPDATE current_timestamp()";
        }
      }

      //コメント入力
      if(getTableInfo[i][6] != ""){
        sql += ` COMMENT '${getTableInfo[i][6]}'`;
      }

      //ALTER TABLE
      if(getTableInfo[i][7] != ""){
        addSql += "\n";
        addSql += "-- インデックスの追加\n";
        if(getTableInfo[i][7] == "primary"){
          addSql += `ALTER TABLE \`${getTableName[i][0]}\`\n`;
          addSql += `  ADD PRIMARY KEY (\`${getTableInfo[i][0]}\`);\n`;
        }
      }
      if(getTableInfo[i][8] != ""){
        addSql += "\n";
        addSql += "-- オートインクリメントの追加\n";
        if(getTableInfo[i][8]){
          addSql += `ALTER TABLE \`${getTableName[i][0]}\`\n`;
          addSql += `  MODIFY \`${getTableInfo[i][0]}\` ${getTableInfo[i][1]}(${getTableInfo[i][2]}) NOT NULL AUTO_INCREMENT;\n`;
        }
      }

      //次のレコードへの追加事項
      if(i < getTableInfo.length - 1){
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




