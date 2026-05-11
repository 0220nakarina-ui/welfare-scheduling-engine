function generateShift() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

// シート取得(シート使用可能にする)
  const configSheet = ss.getSheetByName("設定");
  const staffSheet = ss.getSheetByName("職員リスト");
  const requestSheet = ss.getSheetByName("希望休");
  const shiftSheet = ss.getSheetByName("シフト");
  const roleSheet = ss.getSheetByName("役割");
  const checkSheet = ss.getSheetByName("チェック");
  const summarySheet = ss.getSheetByName("集計");

// 年月取得
  const year = Number(configSheet.getRange("B1").getValue());
  const month = Number(configSheet.getRange("B2").getValue());

// 閉所日取得
  const closedRaw = String(
    configSheet.getRange("B3").getValue()
  );

  const closedDays = closedRaw
    .split(",")
    .map(x => Number(x.trim()));

  const weekdayNeed =
    Number(configSheet.getRange("B4").getValue());

  const saturdayNeed =
    Number(configSheet.getRange("B5").getValue());

  const sundayNeed =
    Number(configSheet.getRange("B6").getValue());

// 職員データ取得
  const staffData = staffSheet
    .getRange(
      2,
      1,
      staffSheet.getLastRow()-1,
      2
    )
    .getValues();

  const staff =
    staffData.map(r => r[0]);

//　リーダーを1人取得
  const leaders =
    staffData
      .filter(r => r[1] == 1)
      .map(r => r[0]);

// 希望休取得
  const requestData =
    requestSheet
      .getRange(
        2,
        1,
        requestSheet.getLastRow()-1,
        2
      )
      .getValues();

  let requests = {};

  requestData.forEach(r => {

    let name = r[0];

    let days = [];

    if(r[1]){

      days = String(r[1])
        .split(",")
        .map(x => Number(x.trim()));
    }

    requests[name] = days;
  });

// シート初期化(前月データの初期化)
  shiftSheet.clear();
  roleSheet.clear();
  checkSheet.clear();
  summarySheet.clear();

// 月末取得(自動取得、2.4.6.9.11も対象)
  const days =
    new Date(year, month, 0).getDate();


// ヘッダ
  shiftSheet.appendRow([
    "日付",
    "曜日",
    ...staff
  ]);

  roleSheet.appendRow([
    "日付",
    "ファシリ",
    "サポート",
    "リーダー"
  ]);

  checkSheet.appendRow([
    "日付",
    "必要人数",
    "実人数",
    "判定"
  ]);

  summarySheet.appendRow([
    "職員",
    "出勤数"
  ]);

//　勤務回数管理
  let workCount = {};

//　連勤管理(何連勤中か管理)
  let consecutive = {};
  staff.forEach(s => {
    workCount[s] = 0;
    consecutive[s] = 0;
  });

// 第3日曜取得
  let sundays = [];

  for(let d=1; d<=days; d++){

    let dt =
      new Date(year, month-1, d);

    if(dt.getDay() == 0){

      sundays.push(d);
    }
  }

  const thirdSunday =
    sundays.length >= 3
    ? sundays[2]
    : null;

// シフト生成
  for(let day=1; day<=days; day++){
    let row = [];
    let date =
      new Date(year, month-1, day);
    let wd = date.getDay();
    row.push(day);
    row.push(
      ["日","月","火","水","木","金","土"][wd]
    );

// 必要人数決定
    let req = weekdayNeed;

// 閉所
    if(closedDays.includes(day)){
      req = 0;
    }

// 土曜
    else if(wd == 6){
      req = saturdayNeed;
    }

// 日曜
    else if(wd == 0){
      if(day == thirdSunday){
        req = sundayNeed;
      }else{
        req = 0;
      }
    }

// 勤務可能者抽出
    let candidates =
      staff.filter(s => {

        // 希望休
        if(
          requests[s] &&
          requests[s].includes(day)
        ){
          return false;
        }

        // 5連勤
        if(consecutive[s] >= 5){
          return false;
        }
        return true;
      });

    // 少ない人優先
    candidates.sort(
      (a,b)=>
      workCount[a]-workCount[b]
    );

// 出勤者決定
    let workers =
      candidates.slice(0, req);


// リーダー不在時
    if(
      req > 0 &&
      !workers.some(
        w => leaders.includes(w)
      )
    ){

      for(let l of leaders){
        if(
          candidates.includes(l)
        ){
          workers[workers.length-1] = l;
          break;
        }
      }
    }


// シフト入力
    for(let s of staff){
      if(workers.includes(s)){

        // 20%で時差
        let shift =
          Math.random() < 0.2
          ? 1
          : 0;

        row.push(shift);
        workCount[s]++;
        consecutive[s]++;
      }else{
        row.push(2);
        consecutive[s] = 0;
      }
    }

    shiftSheet.appendRow(row);

// 役割決定
    if(req > 0){

      let fac =
        workers[
          Math.floor(
            Math.random()*workers.length
          )
        ];

      let support =
        workers.find(w => w != fac);
      let leader =
        workers.find(
          w => leaders.includes(w)
        );

      roleSheet.appendRow([
        day,
        fac || "",
        support || "",
        leader || ""
      ]);
    }


// チェック
    checkSheet.appendRow([
      day,
      req,
      workers.length,
      req == workers.length
      ? "OK"
      : "NG"
    ]);
  }

// 集計
  for(let s of staff){
    summarySheet.appendRow([
      s,
      workCount[s]
    ]);
  }

 // 色付け
  const range =
    shiftSheet.getDataRange();
  const values =
    range.getValues();
  for(let r=1; r<values.length; r++){
　   for(let c=2; c<values[r].length; c++){
      let cell =
        shiftSheet.getRange(r+1,c+1);
      let v = values[r][c];
      if(v == 0){
　       cell.setBackground("#ffffff");
      }else if(v == 1){
　      cell.setBackground("#fff2cc");
      }else{
        cell.setBackground("#d9d9d9");
      }
    }
  }

  SpreadsheetApp.getUi()
    .alert("シフト生成完了");
}
