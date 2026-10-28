/**
 * 💡 LINEからのイベント（メッセージ送信など）を受け取るWebhookの入口（マルチアカウント完全対応版）
 */
/**
 * 💡 LINEからのイベント（メッセージ送信など）を受け取るWebhookの入口（マルチアカウント完全対応版）
 */
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    var events = contents.events;
    var destination = contents.destination; // 💡 LINEから送られてきた宛先ボットのID (Uから始まる文字列)
    
    // 💡 宛先ID（destination）から、該当するアカウントのアクセストークンを特定する
    var accessToken = getAccessTokenByDestination(destination);
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      
      if (event.type === 'message' && event.message.type === 'text') {
        var userMessage = event.message.text.trim();
        var lineUserId = event.source.userId;
        var replyToken = event.replyToken;
        
        if (userMessage === "登録" || userMessage === "ID確認" || userMessage === "id") {
          saveLineIdLog(lineUserId);
          
          // ▼ アカウント名を特定してURLに付与する ▼
          var accountName = getAccountNameByDestination(destination);
          var webAppUrl = ScriptApp.getService().getUrl();
          var registerUrl = webAppUrl + "?page=register&account=" + encodeURIComponent(accountName);
          
          var messages = [
            { type: "text", text: "お疲れ様です！スタッフ登録を行います。\n\n① まず以下のご自身の【LINE ID】をコピーしてください👇" },
            { type: "text", text: lineUserId },
            { type: "text", text: "② 次に、以下のリンクから登録フォームを開き、コピーしたLINE IDを貼り付けてください。\n\n▼ スタッフ初回登録ページ\n" + registerUrl }
          ];
          
          // 💡 特定したアクセストークンを使って返信する
          sendLineMultipleRepliesWithToken(replyToken, messages, accessToken);
        }
      }
    }
  } catch (error) {
    Logger.log("doPostエラー: " + error.message);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
             .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 💡 Webhookの destination（ボットID）から、アカウントマスタの正しいアクセストークンを引く関数
 * LINEの /v2/bot/info APIを使ってボットIDを自動照合し、シートのE列に自動保存します
 */
function getAccessTokenByDestination(destination) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return "";
    
    var data = sheet.getDataRange().getValues();
    
    // ① まず、すでにスプレッドシートのE列（5列目）に destination（Bot ID）が記録されていれば、それを即座に返す
    if (destination) {
      for (var i = 1; i < data.length; i++) {
        var token = data[i][1] ? String(data[i][1]).trim() : "";
        var storedDest = data[i][4] ? String(data[i][4]).trim() : "";
        
        if (storedDest && storedDest === String(destination).trim()) {
          return token;
        }
      }
      
      // ② E列にまだ記録がない場合、各トークンで LINE API (/v2/bot/info) を叩いて本当のBot IDを調べ、一致するものを探す
      for (var j = 1; j < data.length; j++) {
        var checkToken = data[j][1] ? String(data[j][1]).trim() : "";
        if (!checkToken || checkToken.indexOf("YOUR_") !== -1) continue;
        
        var res = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
          method: "get",
          headers: { "Authorization": "Bearer " + checkToken },
          muteHttpExceptions: true
        });
        
        if (res.getResponseCode() === 200) {
          var info = JSON.parse(res.getContentText());
          var botUserId = info.userId; // このトークンを持つ公式LINEのBot ID
          
          if (botUserId) {
            // 次回から一瞬で判定できるよう、スプレッドシートのE列（5列目）にBot IDを自動記録しておく
            sheet.getRange(j + 1, 5).setValue(botUserId);
            
            if (botUserId === String(destination).trim()) {
              return checkToken; // 宛先と完全一致したトークンを返す！
            }
          }
        }
      }
    }
    
    return ""; // 特定できなかった場合は空文字を返し、送信関数側の順次トライに任せる
  } catch (e) {
    Logger.log("getAccessTokenByDestinationエラー: " + e.message);
    return "";
  }
}

/**
 * 💡 指定したアクセストークンを使用して複数のメッセージを返信する関数
 * （もしトークンが未特定・不一致だった場合は、全アカウントのトークンを順番に試して確実に返信します）
 */
function sendLineMultipleRepliesWithToken(replyToken, messagesArray, accessToken) {
  try {
    var url = "https://api.line.me/v2/bot/message/reply";
    var payload = JSON.stringify({
      replyToken: replyToken,
      messages: messagesArray
    });
    
    // ① 特定されたアクセストークンがある場合は、まずそれで送信を試みる
    if (accessToken && accessToken.indexOf("YOUR_") === -1) {
      var res = UrlFetchApp.fetch(url, {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + accessToken
        },
        payload: payload,
        muteHttpExceptions: true
      });
      
      if (res.getResponseCode() === 200) {
        return; // 送信成功したら終了！
      }
    }
    
    // ② 万が一①で送信できなかった場合（accessTokenが空など）、アカウントマスタの全トークンを順番に試す
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var fallbackToken = data[i][1] ? String(data[i][1]).trim() : "";
      if (!fallbackToken || fallbackToken.indexOf("YOUR_") !== -1 || fallbackToken === accessToken) continue;
      
      var resFallback = UrlFetchApp.fetch(url, {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + fallbackToken
        },
        payload: payload,
        muteHttpExceptions: true
      });
      
      // ステータス200（成功）が返ってきたら、正しく返信できたのでループを抜ける
      if (resFallback.getResponseCode() === 200) {
        break;
      }
    }
  } catch (e) {
    Logger.log("sendLineMultipleRepliesWithTokenエラー: " + e.message);
  }
}

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var page = params.page || "";
  
  // page=staff の場合はスタッフ用ポータルを表示
  if (page === "staff") {
    var staffId = params.staffId || "";
    
    // 💡 【追加】スタッフのステータスが「退職」または無効な場合はアクセスを拒否する
    if (isStaffRetired(staffId)) {
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><meta charset="utf-8"><title>アクセス不可</title></head>' +
        '<body style="font-family: sans-serif; text-align: center; padding-top: 50px; background: #f8fafc; color: #333;">' +
        '<h2 style="color: #e11d48;">🔒 このページにはアクセスできません</h2>' +
        '<p>このスタッフアカウントは現在無効（または退職済み）です。<br>ご不明な点は管理者へお問い合わせください。</p>' +
        '</body></html>'
      ).addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    var template = HtmlService.createTemplateFromFile("staff_view");
    template.staffId = staffId;
    template.staffName = getStaffName(staffId);
    template.mode = params.mode || ""; // 👈 管理者ロック用
    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
    
  // page=register の場合はスタッフ初回登録画面を表示
  if (page === "register") {
    var regTemplate = HtmlService.createTemplateFromFile("staff_register");
    regTemplate.accountName = params.account || "";
    return regTemplate.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 指定がない場合は管理者ダッシュボードを表示
  var adminTemplate = HtmlService.createTemplateFromFile("admin_view");
  return adminTemplate.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * LINE IDをスプレッドシート（LINE連携ログ）に保存する関数
 */
function saveLineIdLog(lineUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LINE連携ログ");
    if (!sheet) {
      sheet = ss.insertSheet("LINE連携ログ");
      sheet.appendRow(["timestamp", "lineUserId", "status"]);
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === lineUserId) {
        return;
      }
    }
    
    sheet.appendRow([new Date(), lineUserId, "未登録"]);
  } catch (e) {
    Logger.log("saveLineIdLogエラー: " + e.message);
  }
}

/**
 * 互換性のための既存関数（他の場所から呼ばれたとき用）
 */
function sendLineReply(replyToken, messageText) {
  sendLineMultipleRepliesWithToken(replyToken, [{ type: "text", text: messageText }], "");
}

function sendLineMultipleReplies(replyToken, messagesArray) {
  sendLineMultipleRepliesWithToken(replyToken, messagesArray, "");
}

/**
 * スタッフIDから名前を引く関数
 */
function getStaffName(staffId) {
  if (!staffId) return "未指定";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("スタッフマスタ");
  if (!sheet) return "スタッフ名不明";
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === staffId) {
      return data[i][1];
    }
  }
  return "登録なしのスタッフ";
}

/**
 * スタッフからの勤怠申請データを保存し、LINEへ通知を送る関数
 */
function saveKintaiData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) {
      throw new Error("「申請データ」シートが見つかりません。");
    }
    
    var requestId = "req_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    
    sheet.appendRow([
      requestId,
      data.staffId,
      data.staffName,
      data.workDate,
      data.startTime,
      data.endTime,
      data.breakTime,
      data.transportCost,
      timestamp
    ]);
    
    sendLineNoticeToStaff(data);
    
    return { success: true, message: "申請が完了し、LINEに控えを通知しました！" };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * スタッフマスタからLINE通知を送る関数
 */
function sendLineNoticeToStaff(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName("スタッフマスタ");
  if (!staffSheet) return;
  
  var staffData = staffSheet.getDataRange().getValues();
  var lineUserId = "";
  var assignedAccount = "";
  
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][0] === data.staffId) {
      lineUserId = staffData[i][2];
      assignedAccount = staffData[i][3];
      break;
    }
  }
  
  if (!lineUserId) return;
  
  var accountSheet = ss.getSheetByName("アカウントマスタ");
  var accessToken = "";
  
  if (accountSheet) {
    var accountData = accountSheet.getDataRange().getValues();
    for (var k = 1; k < accountData.length; k++) {
      if (accountData[k][0] === assignedAccount) {
        accessToken = accountData[k][1];
        break;
      }
    }
  }
  
  if (!accessToken || accessToken.includes("YOUR_")) return;
  
    // 💡 スタッフマスタからそのスタッフ専用のポータルURL（G列 / インデックス6）を取得する
  var staffUrl = "";
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][0] === data.staffId) {
      staffUrl = staffData[i][6] || ""; // G列のスタッフ用URL
      break;
    }
  }

  var messageText = 
    "【勤怠申請の控え】\n" +
    "以下の内容で受け付けました。\n\n" +
    "📅 勤務日: " + data.workDate + "\n" +
    "⏰ 時間: " + data.startTime + " ～ " + data.endTime + "\n" +
    "☕ 休憩: " + data.breakTime + "分\n" +
    "🚃 交通費: " + data.transportCost + "円\n\n" +
    (staffUrl ? "🔗 スタッフ専用ポータルURL:\n" + staffUrl + "\n\n" : "") +
    "お疲れ様でした！";
    
  var url = "https://api.line.me/v2/bot/message/push";
  var payload = {
    to: lineUserId,
    messages: [{ type: "text", text: messageText }]
  };
  
  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + accessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, options);
}

/**
 * 新規スタッフを登録する関数（管理者の手動用）
 */
function registerNewStaff(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      throw new Error("「スタッフマスタ」シートが見つかりません。");
    }
    
    if (!formData.staffName || !formData.displayChar) {
      return { success: false, message: "エラー: スタッフ氏名とカレンダー表示文字は必須です。" };
    }
    
    var existingIds = [];
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) existingIds.push(String(data[i][0]));
    }
    
    var staffId;
    do {
      staffId = "st_" + Utilities.getUuid().substring(0, 8);
    } while (existingIds.indexOf(staffId) !== -1);
    
    var timestamp = new Date();
    
    // 💡 通常のWebアプリURLベースで個別URLを発行する
    var webAppUrl = ScriptApp.getService().getUrl();
    var staffUrl = webAppUrl + "?page=staff&staffId=" + staffId;
    
    sheet.appendRow([
      staffId,
      formData.staffName,
      formData.lineUserId || "", // LINE IDも確実に保存
      formData.assignedAccount,
      "有効",
      timestamp,
      staffUrl,
      formData.displayChar
    ]);
    
    return {
      success: true,
      message: "スタッフ「" + formData.staffName + "」さんを登録しました！",
      staffUrl: staffUrl
    };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * 登録されている全スタッフのリストを取得する関数
 */
function getStaffList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var staffList = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        staffList.push({
          staffId: row[0],
          staffName: row[1],
          lineUserId: row[2],
          assignedAccount: row[3],
          status: row[4],
          staffUrl: row[6] || "",
          displayChar: row[7] || row[1].substring(0, 1)
        });
      }
    }
    return staffList;
  } catch (error) {
    Logger.log("スタッフ一覧取得エラー: " + error.message);
    return [];
  }
}

/**
 * staffIdを指定してスタッフマスタからスタッフを削除する関数
 */
function deleteStaffById(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "スタッフを削除しました" };
      }
    }
    
    return { success: false, message: "削除対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "削除エラー: " + error.message };
  }
}

/**
 * 指定された期間の勤怠データを集計する関数（全体の合計値つき）
 */
function getKintaiSummary(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: false, message: "スタッフマスタが見つかりません" };
    var staffData = staffSheet.getDataRange().getValues();
    
    var staffMap = {};
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      if (sId) {
        staffMap[sId] = {
          staffId: sId,
          staffName: staffData[i][1],
          assignedAccount: staffData[i][3],
          daysCount: 0,
          totalCost: 0,
          totalWorkMinutes: 0,
          totalBreakMinutes: 0,
          totalWorkHoursStr: "0時間0分",
          records: []
        };
      }
    }
    
    var dataSheet = ss.getSheetByName("申請データ");
    if (!dataSheet) return { success: false, message: "申請データシートが見つかりません" };
    var kintaiData = dataSheet.getDataRange().getValues();
    
    var start = startDate ? new Date(startDate) : new Date("2000-01-01");
    var end = endDate ? new Date(endDate) : new Date("2100-12-31");
    end.setHours(23, 59, 59, 999);
    
    var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
    
    // 💡 全体の合計用変数
    var grandTotalWorkMinutes = 0;
    var grandTotalCost = 0;
    
    for (var j = 1; j < kintaiData.length; j++) {
      var row = kintaiData[j];
      var reqId = row[0];
      var sId = row[1];
      var workDateStr = row[3];
      var startTime = row[4];
      var endTime = row[5];
      var breakMins = Number(row[6]) || 0;
      var cost = Number(row[7]) || 0;
      
      if (!reqId || !sId || !workDateStr) continue;
      
      var workDate = new Date(workDateStr);
      if (isNaN(workDate.getTime())) continue;
      
      if (workDate >= start && workDate <= end) {
        if (staffMap[sId]) {
          staffMap[sId].daysCount += 1;
          staffMap[sId].totalCost += cost;
          staffMap[sId].totalBreakMinutes += breakMins;
          
          grandTotalCost += cost; // 全体交通費に加算
          
          var formattedStart = formatTime(startTime);
          var formattedEnd = formatTime(endTime);
          var workMins = 0;
          
          if (formattedStart && formattedEnd) {
            var sParts = formattedStart.split(":");
            var eParts = formattedEnd.split(":");
            var startM = Number(sParts[0]) * 60 + Number(sParts[1]);
            var endM = Number(eParts[0]) * 60 + Number(eParts[1]);
            workMins = (endM - startM) - breakMins;
            if (workMins > 0) {
              staffMap[sId].totalWorkMinutes += workMins;
              grandTotalWorkMinutes += workMins; // 全体労働時間に加算（終了時刻ー開始時刻ー休憩時間）
            }
          }
          
          var m = ('0' + (workDate.getMonth() + 1)).slice(-2);
          var d = ('0' + workDate.getDate()).slice(-2);
          var dayOfWeek = weekDays[workDate.getDay()];
          var formattedDate = m + '/' + d + '(' + dayOfWeek + ')';
          
          staffMap[sId].records.push({
            workDate: formattedDate,
            startTime: formattedStart,
            endTime: formattedEnd,
            breakTime: breakMins,
            transportCost: cost,
            workHoursStr: workMins > 0 ? Math.floor(workMins / 60) + "時間" + (workMins % 60) + "分" : "0時間"
          });
        }
      }
    }
    
    var summaryList = [];
    for (var key in staffMap) {
      var item = staffMap[key];
      var totH = Math.floor(item.totalWorkMinutes / 60);
      var totM = item.totalWorkMinutes % 60;
      item.totalWorkHoursStr = totH + "時間" + totM + "分";
      
      var brH = Math.floor(item.totalBreakMinutes / 60);
      var brM = item.totalBreakMinutes % 60;
      item.totalBreakHoursStr = brH > 0 ? brH + "時間" + brM + "分" : brM + "分";
      
      item.records.sort(function(a, b) {
        return new Date(a.workDate) - new Date(b.workDate);
      });
      summaryList.push(item);
    }
    
    // 全体の総労働時間の文字列表記作成
    var gTotH = Math.floor(grandTotalWorkMinutes / 60);
    var gTotM = grandTotalWorkMinutes % 60;
    var grandTotalWorkHoursStr = gTotH + "時間" + gTotM + "分";
    
    return { 
      success: true, 
      summary: summaryList,
      grandTotal: {
        totalWorkHoursStr: grandTotalWorkHoursStr,
        totalCost: grandTotalCost
      }
    };
  } catch (error) {
    return { success: false, message: "集計エラー: " + error.message };
  }
}

/**
 * 時刻データを "HH:mm" 形式に変換
 */
function formatTime(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var h = ('0' + val.getHours()).slice(-2);
    var m = ('0' + val.getMinutes()).slice(-2);
    return h + ':' + m;
  }
  var str = String(val);
  if (str.indexOf("T") !== -1 || str.indexOf("1899-") !== -1) {
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      var h = ('0' + d.getHours()).slice(-2);
      var m = ('0' + d.getMinutes()).slice(-2);
      return h + ':' + m;
    }
  }
  if (str.length >= 5) {
    return str.substring(0, 5);
  }
  return str;
}

/**
 * LINEアカウント一覧を取得
 */
function getAccountList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return ["アカウントA", "アカウントB"];
    
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) accounts.push(data[i][0]);
    }
    return accounts.length > 0 ? accounts : ["アカウントA", "アカウントB"];
  } catch (error) {
    return ["アカウントA", "アカウントB"];
  }
}

/**
 * シフト予定を取得
 */
function getShiftSchedule(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) return { success: true, shifts: [] };
    
    var data = sheet.getDataRange().getValues();
    var shifts = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var shiftId = row[0];
      var workDate = row[1];
      var staffId = row[2];
      var staffName = row[3];
      var dailyMemo = row[4] || ""; // E列から取得
      
      if (!workDate || !staffId) continue;
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (!workDate || !staffId) continue;
      // ...(中略)...
      if (dateStr.indexOf(yearMonth) === 0) {
        shifts.push({ shiftId: shiftId, workDate: dateStr, staffId: staffId, staffName: staffName, dailyMemo: dailyMemo });
      }
    }

    return { success: true, shifts: shifts };
  } catch (error) {
    return { success: false, message: "シフト取得エラー: " + error.message };
  }
}

/**
 * 日本の祝日マップを取得（計算ロジック内蔵）
 */
function getJapaneseHolidays(yearMonth) {
  try {
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var holidayMap = {};
    
    var fixedHolidays = {
      "1-1": "元日", "2-11": "建国記念の日", "2-23": "天皇誕生日",
      "4-29": "昭和の日", "5-3": "憲法記念日", "5-4": "みどりの日",
      "5-5": "こどもの日", "8-11": "山の日", "11-3": "文化の日", "11-23": "勤労感謝の日"
    };

    var lastDay = new Date(year, month, 0).getDate();
    for (var day = 1; day <= lastDay; day++) {
      var d = new Date(year, month - 1, day);
      var w = d.getDay();
      var m = month;
      
      var key = year + '-' + ('0' + m).slice(-2) + '-' + ('0' + day).slice(-2);
      var hName = "";
      
      var fixKey = m + "-" + day;
      if (fixedHolidays[fixKey]) hName = fixedHolidays[fixKey];
      
      if (m === 1 && w === 1 && day >= 8 && day <= 14) hName = "成人の日";
      if (m === 4 && day === 29) hName = "昭和の日";
      if (m === 7 && w === 1 && day >= 15 && day <= 21) hName = "海の日";
      if (m === 8 && day === 11) hName = "山の日";
      if (m === 9 && w === 1 && day >= 15 && day <= 21) hName = "敬老の日";
      if (m === 10 && w === 1 && day >= 8 && day <= 14) hName = "スポーツの日";
      if (m === 11 && day === 3) hName = "文化の日";
      if (m === 11 && day === 23) hName = "勤労感謝の日";
      
      if (m === 3 && ((year === 2026 && day === 20) || (year !== 2026 && day === 21))) hName = "春分の日";
      if (m === 9 && ((year === 2026 && day === 23) || (year !== 2026 && day === 23))) hName = "秋分の日";

      if (hName) holidayMap[key] = hName;
    }
    return { success: true, holidays: holidayMap };
  } catch (error) {
    return { success: true, holidays: {} };
  }
}

/**
 * 休業日を取得
 */
function getClosedDays(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("休業日");
    if (!sheet) return { success: true, closedDays: [] };
    
    var data = sheet.getDataRange().getValues();
    var closedDays = [];
    
    for (var i = 1; i < data.length; i++) {
      var workDate = data[i][1];
      if (!workDate) continue;
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (dateStr.indexOf(yearMonth) === 0) {
        closedDays.push(dateStr);
      }
    }
    return { success: true, closedDays: closedDays };
  } catch (error) {
    return { success: false, closedDays: [], message: error.message };
  }
}

/**
 * シフト予定と休業日をまとめて一括保存する関数
 */
function saveAllSchedules(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth;
    
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) {
      shiftSheet = ss.insertSheet("シフト予定");
      shiftSheet.appendRow(["shiftId", "workDate", "staffId", "staffName", "dailyMemo"]);
    }
    
    var shiftData = shiftSheet.getDataRange().getValues();
    var newShiftRows = [shiftData[0]]; // ヘッダー行を維持
    
    // 選択された年月以外の既存シフトを保持する
    for (var i = 1; i < shiftData.length; i++) {
      var rDate = new Date(shiftData[i][1]);
      if (isNaN(rDate.getTime())) continue;
      var rYm = rDate.getFullYear() + '-' + ('0' + (rDate.getMonth() + 1)).slice(-2);
      if (rYm !== yearMonth) {
        newShiftRows.push(shiftData[i]);
      }
    }
    
    // 画面側から送られてきた新しいシフトデータを追加（5列：id, workDate, staffId, staffName, dailyMemo）
    (data.shifts || []).forEach(function(sh) {
      var sId = sh.shiftId && sh.shiftId.indexOf("local_") === -1 ? sh.shiftId : "shift_" + Utilities.getUuid().substring(0, 8);
      newShiftRows.push([sId, sh.workDate, sh.staffId, sh.staffName, sh.dailyMemo || ""]);
    });
    
    shiftSheet.clearContents();
    if (newShiftRows.length > 0) {
      shiftSheet.getRange(1, 1, newShiftRows.length, newShiftRows[0].length).setValues(newShiftRows);
    }
    
    var holSheet = ss.getSheetByName("休業日");
    if (!holSheet) {
      holSheet = ss.insertSheet("休業日");
      holSheet.appendRow(["holidayId", "workDate"]);
    }
    
    var holData = holSheet.getDataRange().getValues();
    var newHolRows = [holData[0]];
    
    for (var j = 1; j < holData.length; j++) {
      var hDate = new Date(holData[j][1]);
      if (isNaN(hDate.getTime())) continue;
      var hYm = hDate.getFullYear() + '-' + ('0' + (hDate.getMonth() + 1)).slice(-2);
      if (hYm !== yearMonth) {
        newHolRows.push(holData[j]);
      }
    }
    
    (data.closedDays || []).forEach(function(cDate) {
      var hId = "hol_" + Utilities.getUuid().substring(0, 8);
      newHolRows.push([hId, cDate]);
    });
    
    holSheet.clearContents();
    if (newHolRows.length > 0) {
      holSheet.getRange(1, 1, newHolRows.length, newHolRows[0].length).setValues(newHolRows);
    }
    
    return { success: true, message: "変更内容をすべて保存しました！" };
  } catch (error) {
    return { success: false, message: "保存エラー: " + error.message };
  }
}

/**
 * スタッフの「カレンダー表示文字」を更新する関数
 */
function updateStaffDisplayChar(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.staffId)) {
        sheet.getRange(i + 1, 8).setValue(data.displayChar);
        return { success: true, message: "表示文字を更新しました！" };
      }
    }
    
    return { success: false, message: "対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "更新エラー: " + error.message };
  }
}

/**
 * 指定された年月のフリー予定を取得する関数
 */
function getFreeSchedules(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("フリー予定");
    if (!sheet) {
      return { success: true, freeSchedules: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var freeSchedules = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var scheduleId = row[0];
      var workDate = row[1];
      var title = row[2];
      
      if (!workDate || !title) continue;
      
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (dateStr.indexOf(yearMonth) === 0) {
        freeSchedules.push({
          scheduleId: scheduleId,
          workDate: dateStr,
          title: title
        });
      }
    }
    
    return { success: true, freeSchedules: freeSchedules };
  } catch (error) {
    return { success: false, message: "フリー予定取得エラー: " + error.message };
  }
}

/**
 * シフト予定、休業日、フリー予定をまとめて一括保存する関数
 */
function saveAllSchedulesWithFree(data) {
  try {
    var res = saveAllSchedules(data);
    if (!res.success) return res;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth;
    
    var freeSheet = ss.getSheetByName("フリー予定");
    if (!freeSheet) {
      freeSheet = ss.insertSheet("フリー予定");
      freeSheet.appendRow(["scheduleId", "workDate", "title"]);
    }
    
    var freeData = freeSheet.getDataRange().getValues();
    var newFreeRows = [freeData[0]];
    
    for (var i = 1; i < freeData.length; i++) {
      var fDate = new Date(freeData[i][1]);
      if (isNaN(fDate.getTime())) continue;
      var fYm = fDate.getFullYear() + '-' + ('0' + (fDate.getMonth() + 1)).slice(-2);
      if (fYm !== yearMonth) {
        newFreeRows.push(freeData[i]);
      }
    }
    
    (data.freeSchedules || []).forEach(function(fs) {
      var sId = "free_" + Utilities.getUuid().substring(0, 8);
      newFreeRows.push([sId, fs.workDate, fs.title]);
    });
    
    freeSheet.clearContents();
    if (newFreeRows.length > 0) {
      freeSheet.getRange(1, 1, newFreeRows.length, newFreeRows[0].length).setValues(newFreeRows);
    }
    
    return { success: true, message: "変更内容（予定・フリー項目含む）をすべて保存しました！" };
  } catch (error) {
    return { success: false, message: "一括保存エラー: " + error.message };
  }
}

/**
 * スタッフのステータス（有効/退職）を切り替える関数
 */
function toggleStaffStatus(staffId, newStatus) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        sheet.getRange(i + 1, 5).setValue(newStatus);
        return { success: true, message: "スタッフのステータスを「" + newStatus + "」に変更しました！" };
      }
    }
    
    return { success: false, message: "対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "ステータス変更エラー: " + error.message };
  }
}

/**
 * スタッフからのシフト希望を保存し、管理者向けの新着通知を記録する関数（伝達事項対応版）
 */
function saveShiftPreferences(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      sheet = ss.insertSheet("シフト予定");
      sheet.appendRow(["shiftId", "workDate", "staffId", "staffName"]);
    }
    
    var existingData = sheet.getDataRange().getValues();
    var staffId = data.staffId;
    var staffName = data.staffName;
    var targetRequests = data.dailyRequests || [];
    var memo = data.memo || "";
    
    var newRows = [existingData[0]]; 
    for (var i = 1; i < existingData.length; i++) {
      var row = existingData[i];
      var rDateStr = String(row[1] || "");
      var rStaffId = String(row[2] || "");
      
      var isSameMonth = rDateStr.indexOf(data.yearMonth) === 0;
      if (!(isSameMonth && rStaffId === staffId)) {
        newRows.push(row);
      }
    }
    
    targetRequests.forEach(function(req) {
      var shiftId = "shift_" + Utilities.getUuid().substring(0, 8);
      // E列に req.memo を追加保存
      newRows.push([shiftId, req.date, staffId, staffName, req.memo || ""]);
    });
    
    sheet.clearContents();
    if (newRows.length > 0) {
      sheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    var notifSheet = ss.getSheetByName("希望通知ログ");
    if (!notifSheet) {
      notifSheet = ss.insertSheet("希望通知ログ");
      notifSheet.appendRow(["notifId", "staffId", "staffName", "yearMonth", "timestamp", "status", "memo"]);
    }
    
    var notifId = "notif_" + Utilities.getUuid().substring(0, 8);
    notifSheet.appendRow([
      notifId,
      staffId,
      staffName,
      data.yearMonth,
      new Date(),
      "未読",
      memo
    ]);
    
    return { success: true, message: "シフト希望を提出し、予定に反映しました！" };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * カレンダー上で削除されたシフト希望、または確定シフトをシートから削除する関数
 */
function deleteScheduleOrPreference(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetDate = data.workDate;
    var staffId = data.staffId;
    
    var prefSheet = ss.getSheetByName("希望シフト");
    if (prefSheet) {
      var prefData = prefSheet.getDataRange().getValues();
      for (var i = 0; i < prefData.length; i++) {
        var row = prefData[i];
        if (row[1] === staffId) {
          var datesStr = String(row[4] || "");
          if (datesStr.indexOf(targetDate) !== -1) {
            var datesArr = datesStr.split(",").map(function(d){ return d.trim(); });
            var newDatesArr = datesArr.filter(function(d){ return d !== targetDate; });
            
            if (newDatesArr.length > 0) {
              prefSheet.getRange(i + 1, 5).setValue(newDatesArr.join(","));
            } else {
              prefSheet.deleteRow(i + 1);
            }
          }
        }
      }
    }
    
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (shiftSheet) {
      var shiftData = shiftSheet.getDataRange().getValues();
      for (var j = 1; j < shiftData.length; j++) {
        var sRow = shiftData[j];
        var sDate = new Date(sRow[1]);
        var sDateStr = isNaN(sDate.getTime()) ? String(sRow[1]) : Utilities.formatDate(sDate, "Asia/Tokyo", "yyyy-MM-dd");
        if (sDateStr === targetDate && sRow[2] === staffId) {
          shiftSheet.deleteRow(j + 1);
          break;
        }
      }
    }
    
    return { success: true, message: "データを削除しました" };
  } catch (error) {
    return { success: false, message: "削除エラー: " + error.message };
  }
}

/**
 * 互換性のために空の配列を返すダミー関数
 */
function getShiftPreferences(yearMonth) {
  return { success: true, preferences: [] };
}

/**
 * LINEアカウントのアクセストークンを考慮して一斉または個別にメッセージを送信する関数
 */
function sendLineMessage(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) throw new Error("スタッフマスタが見つかりません。");
    var staffData = staffSheet.getDataRange().getValues();

    var targetStaff = [];
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      var sName = staffData[i][1];
      var lineId = staffData[i][2];
      var account = staffData[i][3];
      var status = staffData[i][4];

      if (sId && lineId && status === "有効") {
        if (data.targetType === "all" || data.targetId === sId) {
          targetStaff.push({ staffId: sId, staffName: sName, lineUserId: lineId, assignedAccount: account });
        }
      }
    }

    if (targetStaff.length === 0) {
      return { success: false, message: "送信対象の有効なスタッフが見つかりませんでした。" };
    }

    var accountSheet = ss.getSheetByName("アカウントマスタ");
    var tokens = {};
    if (accountSheet) {
      var accData = accountSheet.getDataRange().getValues();
      for (var k = 1; k < accData.length; k++) {
        if (accData[k][0]) tokens[accData[k][0]] = accData[k][1];
      }
    }

    var successCount = 0;
    var failCount = 0;

    targetStaff.forEach(function(staff) {
      var token = tokens[staff.assignedAccount];
      if (!token || token.includes("YOUR_")) {
        failCount++;
        return;
      }

      var url = "https://api.line.me/v2/bot/message/push";
      var payload = {
        to: staff.lineUserId,
        messages: [{ type: "text", text: data.messageText }]
      };

      var options = {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        var res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() === 200) {
          successCount++;
        } else {
          failCount++;
        }
      } catch (e) {
        failCount++;
      }
    });

    var histSheet = ss.getSheetByName("連絡履歴");
    if (!histSheet) {
      histSheet = ss.insertSheet("連絡履歴");
      histSheet.appendRow(["historyId", "timestamp", "targetName", "messageText", "resultSummary"]);
    }

    var historyId = "hist_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    var targetNameDisplay = (data.targetType === "all") ? "【一斉送信】全スタッフ" : (targetStaff[0] ? targetStaff[0].staffName : "個別");
    var resultStr = "成功: " + successCount + "件 / 失敗: " + failCount + "件";
    histSheet.appendRow([historyId, timestamp, targetNameDisplay, data.messageText, resultStr]);

    // 💡 ここのダブルクォーテーションを追加
    return { success: true, message: successCount + "名に確認・確定依頼を送信しました！" };
  } catch (error) {
    return { success: false, message: "送信エラー: " + error.message };
  }
}

/**
 * 送信履歴を取得する関数
 */
function getLineHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("連絡履歴");
    if (!sheet) return { success: true, history: [] };

    var data = sheet.getDataRange().getValues();
    var history = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var timeStr = "";
        if (row[1] instanceof Date) {
          timeStr = Utilities.formatDate(row[1], "Asia/Tokyo", "yyyy/MM/dd HH:mm");
        } else {
          timeStr = String(row[1]);
        }

        history.push({
          timestamp: timeStr,
          targetName: row[2],
          messageText: row[3],
          resultSummary: row[4]
        });
      }
    }
    history.reverse();
    return { success: true, history: history };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

/**
 * 過去1か月間の勤怠未提出者をリストアップする関数
 */
function getKintaiPendingStaff() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var today = new Date();
    var startDate = new Date();
    startDate.setDate(today.getDate() - 30);
    
    var todayStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy-MM-dd");
    var startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd");

    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) return { success: true, pendingList: [] };
    var shiftData = shiftSheet.getDataRange().getValues();
    
    var scheduledMap = {}; 
    for (var i = 1; i < shiftData.length; i++) {
      var row = shiftData[i];
      var wDate = row[1];
      var sId = row[2];
      var sName = row[3];
      if (!wDate || !sId) continue;
      
      var d = new Date(wDate);
      if (isNaN(d.getTime())) continue;
      var dStr = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      
      if (dStr >= startStr && dStr <= todayStr) {
        var key = sId + "_" + dStr;
        scheduledMap[key] = {
          staffId: sId,
          staffName: sName,
          targetDate: dStr
        };
      }
    }

    var dataSheet = ss.getSheetByName("申請データ");
    if (dataSheet) {
      var kintaiData = dataSheet.getDataRange().getValues();
      for (var j = 1; j < kintaiData.length; j++) {
        var kRow = kintaiData[j];
        var kStaffId = kRow[1];
        var kDate = kRow[3];
        if (!kStaffId || !kDate) continue;
        
        var kd = new Date(kDate);
        if (isNaN(kd.getTime())) continue;
        var kdStr = Utilities.formatDate(kd, "Asia/Tokyo", "yyyy-MM-dd");
        
        var submittedKey = kStaffId + "_" + kdStr;
        if (scheduledMap[submittedKey]) {
          delete scheduledMap[submittedKey];
        }
      }
    }

    var pendingList = [];
    for (var key in scheduledMap) {
      pendingList.push(scheduledMap[key]);
    }
    
    pendingList.sort(function(a, b) {
      return new Date(a.targetDate) - new Date(b.targetDate);
    });

    return { success: true, pendingList: pendingList };
  } catch (error) {
    return { success: false, message: "未提出者取得エラー: " + error.message, pendingList: [] };
  }
}

/**
 * 未提出者リストをLINEリマインド用に詳細情報つきで取得する関数
 */
function getKintaiPendingForReminder() {
  try {
    var res = getKintaiPendingStaff();
    if (!res.success) return res;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: true, pendingList: [] };
    var staffData = staffSheet.getDataRange().getValues();
    
    var staffLineMap = {};
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      var lineId = staffData[i][2];
      var status = staffData[i][4];
      if (sId && lineId && status === "有効") {
        staffLineMap[sId] = true;
      }
    }

    var validPendingList = [];
    res.pendingList.forEach(function(item) {
      if (staffLineMap[item.staffId]) {
        validPendingList.push(item);
      }
    });

    return { success: true, pendingList: validPendingList };
  } catch (error) {
    return { success: false, message: "リマインド対象取得エラー: " + error.message, pendingList: [] };
  }
}

/**
 * 未提出者へリマインドLINEを送信する関数
 */
function sendKintaiReminderLine(data) {
  try {
    var targets = data.targets || [];
    var messageText = data.messageText;

    if (targets.length === 0) {
      return { success: false, message: "送信対象のスタッフが選択されていません。" };
    }
    if (!messageText) {
      return { success: false, message: "メッセージ本文が入力されていません。" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) throw new Error("スタッフマスタが見つかりません。");
    var staffData = staffSheet.getDataRange().getValues();

    var staffMap = {};
    for (var i = 1; i < staffData.length; i++) {
      if (staffData[i][0]) {
        staffMap[staffData[i][0]] = {
          staffName: staffData[i][1],
          lineUserId: staffData[i][2],
          assignedAccount: staffData[i][3]
        };
      }
    }

    var accountSheet = ss.getSheetByName("アカウントマスタ");
    var tokens = {};
    if (accountSheet) {
      var accData = accountSheet.getDataRange().getValues();
      for (var k = 1; k < accData.length; k++) {
        if (accData[k][0]) tokens[accData[k][0]] = accData[k][1];
      }
    }

    var successCount = 0;
    var failCount = 0;
    var targetNames = [];

    targets.forEach(function(staffId) {
      var staff = staffMap[staffId];
      if (!staff || !staff.lineUserId) {
        failCount++;
        return;
      }

      var token = tokens[staff.assignedAccount];
      if (!token || token.includes("YOUR_")) {
        failCount++;
        return;
      }

      targetNames.push(staff.staffName);

      var url = "https://api.line.me/v2/bot/message/push";
      var payload = {
        to: staff.lineUserId,
        messages: [{ type: "text", text: messageText }]
      };

      var options = {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        var res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() === 200) {
          successCount++;
        } else {
          failCount++;
        }
      } catch (e) {
        failCount++;
      }
    });

    var histSheet = ss.getSheetByName("連絡履歴");
    if (!histSheet) {
      histSheet = ss.insertSheet("連絡履歴");
      histSheet.appendRow(["historyId", "timestamp", "targetName", "messageText", "resultSummary"]);
    }

    var historyId = "hist_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    var targetNameDisplay = "【勤怠リマインド】" + targetNames.join(", ");
    var resultStr = "成功: " + successCount + "件 / 失敗: " + failCount + "件";

    histSheet.appendRow([historyId, timestamp, targetNameDisplay, messageText, resultStr]);

    return { 
      success: true, 
      message: "リマインドを送信しました！（" + resultStr + "）" 
    };

  } catch (error) {
    return { success: false, message: "リマインド送信エラー: " + error.message };
  }
}

/**
 * 管理者向けに未読の「シフト希望申請通知」を取得する関数（新しいもの順・最大20件）
 */
function getUnreadShiftNotifications() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, notifications: [], latestTimestamp: "" };
    
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: true, notifications: [], latestTimestamp: "" };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: true, notifications: [], latestTimestamp: "" };

    var notifications = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var notifId = String(row[0] || "").trim();
      var staffName = String(row[2] || "").trim();
      var yearMonth = String(row[3] || "").trim();
      var timeObj = row[4];
      var status = String(row[5] || "").trim();

      if (notifId && status === "未読") {
        var timeStr = "";
        if (timeObj instanceof Date) {
          timeStr = Utilities.formatDate(timeObj, "Asia/Tokyo", "MM/dd HH:mm");
        } else {
          timeStr = String(timeObj);
        }

        var formattedYm = yearMonth;
        if (yearMonth && yearMonth.indexOf("-") !== -1) {
          var ymParts = yearMonth.split("-");
          formattedYm = ymParts[0] + "年" + Number(ymParts[1]) + "月";
        } else if (yearMonth) {
          var ymDate = new Date(yearMonth);
          if (!isNaN(ymDate.getTime())) {
            formattedYm = Utilities.formatDate(ymDate, "Asia/Tokyo", "yyyy年M月");
          }
        }

        notifications.push({
          notifId: notifId,
          staffName: staffName || "スタッフ",
          yearMonth: formattedYm || "未指定",
          timestamp: timeStr,
          rawTime: timeObj instanceof Date ? timeObj.getTime() : 0
        });
      }
    }

    // 新しい順（降順）にソート
    notifications.sort(function(a, b) {
      return b.rawTime - a.rawTime;
    });

    // 💡 最も新しい通知のタイムスタンプを取得しておく
    var latestTimestamp = notifications.length > 0 ? notifications[0].timestamp : "";

    notifications = notifications.slice(0, 20);

    return { 
      success: true, 
      notifications: notifications, 
      latestTimestamp: latestTimestamp // 👈 追加
    };
  } catch (error) {
    return { success: true, notifications: [], latestTimestamp: "", message: error.message };
  }
}

/**
 * シフト希望通知を「既読（確認済み）」にする関数
 */
function markNotificationAsRead(notifId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: false, message: "通知シートが見つかりません" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(notifId)) {
        sheet.getRange(i + 1, 6).setValue("既読");
        return { success: true, message: "確認済みにしました" };
      }
    }
    return { success: false, message: "該当の通知が見つかりませんでした" };
  } catch (error) {
    return { success: false, message: "エラー: " + error.message };
  }
}

/**
 * 💡 スタッフ自身がフォームから登録を行った際の処理（アカウントごとの上限10名チェック付き）
 */
function registerStaffBySelf(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      throw new Error("「スタッフマスタ」シートが見つかりません。");
    }
    
    if (!data.staffName || !data.displayChar || !data.lineUserId) {
      return { success: false, message: "エラー: 氏名、表示文字、LINE IDは必須です。" };
    }
    
    var assignedAccount = data.assignedAccount || "アカウントA"; 

    // 指定されたアカウントに現在「有効」なスタッフが何名いるかカウントする
    var dataRows = sheet.getDataRange().getValues();
    var activeCount = 0;
    
    for (var i = 1; i < dataRows.length; i++) {
      var rowAccount = dataRows[i][3]; // 割当アカウントの列
      var rowStatus  = dataRows[i][4]; // ステータスの列（有効 / 退職）
      
      if (rowAccount === assignedAccount && rowStatus === "有効") {
        activeCount++;
      }
    }

   var maxLimit = getMaxStaffLimit();
if (activeCount >= maxLimit) {
  return { 
    success: false, 
    message: "⚠️ 登録エラー\n「" + assignedAccount + "」の登録上限（最大" + maxLimit + "名）に達しているため、これ以上登録できません。" 
  };
}
    
    var staffId = "st_" + Utilities.getUuid().substring(0, 6);
    var timestamp = new Date();
    var webAppUrl = ScriptApp.getService().getUrl();
    var staffUrl = webAppUrl + "?page=staff&staffId=" + staffId;
    
    sheet.appendRow([
      staffId,
      data.staffName,
      data.lineUserId,
      assignedAccount,
      "有効",
      timestamp,
      staffUrl,
      data.displayChar
    ]);
    
    sendWelcomeAndTestNotice(data.lineUserId, assignedAccount, data.staffName, staffUrl);
    
    return {
      success: true,
      message: "登録が完了しました！ご自身のLINEに専用ポータルURLを送信しましたのでご確認ください。"
    };
  } catch (error) {
    return { success: false, message: "登録エラー: " + error.message };
  }
}

/**
 * 登録完了時にスタッフへウェルカムメッセージと個別URLをプッシュ通知する関数
 */
function sendWelcomeAndTestNotice(lineUserId, assignedAccount, staffName, staffUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var accountSheet = ss.getSheetByName("アカウントマスタ");
  var accessToken = "";
  
  if (accountSheet) {
    var accountData = accountSheet.getDataRange().getValues();
    for (var k = 1; k < accountData.length; k++) {
      if (accountData[k][0] === assignedAccount) {
        accessToken = accountData[k][1];
        break;
      }
    }
  }
  
  if (!accessToken || accessToken.includes("YOUR_")) return;
  
  var messageText = 
    "🎉 " + staffName + "さん、スタッフ登録が完了しました！\n\n" +
    "これより勤怠管理システムをご利用いただけます。\n" +
    "今後の勤怠申請やシフト申請は、以下の【スタッフ専用ポータルURL】からお願いします。\n\n" +
    "🔗 スタッフ専用ポータルURL:\n" + staffUrl + "\n\n" +
    "お疲れ様でした！";
  
  var url = "https://api.line.me/v2/bot/message/push";
  var payload = {
    to: lineUserId,
    messages: [{ type: "text", text: messageText }]
  };
  
  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + accessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("ウェルカムLINE送信失敗: " + e.message);
  }
}

/**
 * スタッフがLIFF経由で専用URLを開いたときにLINE IDを自動紐付けする関数
 */
function checkAndLinkStaffAccess(staffId, lineUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) return { success: false };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        var currentLineId = data[i][2];
        
        if (!currentLineId && lineUserId) {
          sheet.getRange(i + 1, 3).setValue(lineUserId);
          return { success: true, isNewlyLinked: true };
        }
        
        return { success: true, isNewlyLinked: false };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * LINEアカウントの詳細情報と現在の登録人数を取得する関数
 */
function getAccountDetailsList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    var countsMap = {};
    if (staffSheet) {
      var staffData = staffSheet.getDataRange().getValues();
      for (var s = 1; s < staffData.length; s++) {
        var acc = staffData[s][3];
        var status = staffData[s][4];
        if (acc && status === "有効") {
          countsMap[acc] = (countsMap[acc] || 0) + 1;
        }
      }
    }

    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) {
      return [
        { name: "アカウントA", addUrl: "https://lin.ee/xxxxxA", qrUrl: "", count: countsMap["アカウントA"] || 0 },
        { name: "アカウントB", addUrl: "https://lin.ee/xxxxxB", qrUrl: "", count: countsMap["アカウントB"] || 0 }
      ];
    }
    
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var accName = row[0];
        accounts.push({
          name: accName,
          channelId: row[1] || "",
          addUrl: row[2] || "https://lin.ee/default",
          qrUrl: row[3] || "",
          count: countsMap[accName] || 0
        });
      }
    }
    return accounts.length > 0 ? accounts : [{ name: "アカウントA", addUrl: "", qrUrl: "", count: 0 }];
  } catch (error) {
    return [];
  }
}

/**
 * 💡 スタッフ自身が自分の直近の勤怠申請履歴を確認するための関数（最大20件）
 */
function getStaffKintaiHistory(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return { success: true, history: [] };
    
    var data = sheet.getDataRange().getValues();
    var history = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var requestId = row[0];
      var sId = row[1];
      
      if (String(sId) === String(staffId)) {
        var workDate = formatTimeOrDate(row[3]);
        var startTime = formatTime(row[4]);
        var endTime = formatTime(row[5]);
        var breakTime = row[6];
        var cost = Number(row[7]) || 0;
        
        history.push({
          requestId: requestId,
          workDate: workDate,
          startTime: startTime,
          endTime: endTime,
          breakTime: breakTime,
          transportCost: cost
        });
        
        if (history.length >= 20) break;
      }
    }
    
    return { success: true, history: history };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

/**
 * 💡 スタッフが自分の申請履歴を修正してスプレッドシートを更新する関数
 */
function updateKintaiData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) {
      throw new Error("「申請データ」シートが見つかりません。");
    }
    
    var sheetData = sheet.getDataRange().getValues();
    var targetRow = -1;
    
    for (var i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][0]) === String(data.requestId)) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, message: "対象の申請データが見つかりませんでした。" };
    }
    
    sheet.getRange(targetRow, 4).setValue(data.workDate);
    sheet.getRange(targetRow, 5).setValue(data.startTime);
    sheet.getRange(targetRow, 6).setValue(data.endTime);
    sheet.getRange(targetRow, 7).setValue(data.breakTime);
    sheet.getRange(targetRow, 8).setValue(data.transportCost);
    
    return { success: true, message: "申請内容を修正しました！" };
  } catch (error) {
    return { success: false, message: "修正エラー: " + error.message };
  }
}

function formatTimeOrDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy-MM-dd");
  }
  var str = String(val);
  if (str.indexOf("T") !== -1) {
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
    }
  }
  return str.substring(0, 10);
}

function formatTimestamp(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  return String(val);
}

/**
 * 💡 スタッフが自分のシフト希望提出履歴（日付ごと＆伝達事項）を確認するための関数
 */
function getStaffShiftHistory(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) return { success: true, history: [] };
    
    var shiftData = shiftSheet.getDataRange().getValues();
    var historyList = [];
    
    // シフト予定シートから該当スタッフのデータ（日付とE列のメモ）をすべて取得する
    for (var i = 1; i < shiftData.length; i++) {
      var row = shiftData[i];
      var workDateStr = String(row[1] || "");
      var rowStaffId = String(row[2] || "");
      var memo = String(row[4] || ""); // E列のメモ
      
      if (rowStaffId === String(staffId) && workDateStr) {
        var d = new Date(workDateStr);
        var formattedDate = workDateStr;
        if (!isNaN(d.getTime())) {
          var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
          var m = ('0' + (d.getMonth() + 1)).slice(-2);
          var day = ('0' + d.getDate()).slice(-2);
          var dayOfWeek = weekDays[d.getDay()];
          formattedDate = m + '/' + day + '(' + dayOfWeek + ')';
        }
        
        historyList.push({
          workDate: formattedDate,
          rawDate: workDateStr,
          memo: memo
        });
      }
    }
    
  // 💡 日付の新しい順（降順）に並び替え
    historyList.sort(function(a, b) {
      return new Date(b.rawDate) - new Date(a.rawDate);
    });
    
    return { success: true, history: historyList };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

/**
 * 💡 Webhookの destination からアカウント名を取得する関数
 */
function getAccountNameByDestination(destination) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return "アカウントA";
    
    var data = sheet.getDataRange().getValues();
    if (destination) {
      for (var i = 1; i < data.length; i++) {
        var storedDest = data[i][4] ? String(data[i][4]).trim() : "";
        if (storedDest === String(destination).trim()) {
          return data[i][0]; // A列のアカウント名を返す
        }
      }
    }
    return data.length > 1 ? data[1][0] : "アカウントA";
  } catch (e) {
    return "アカウントA";
  }
}

/**
 * 💡 指定したスタッフの過去1か月間における「シフト予定があるのに未提出の日」のリストを取得する関数（日付順にソート）
 */
function getStaffPendingDates(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var today = new Date();
    var startDate = new Date();
    startDate.setDate(today.getDate() - 30);
    
    var todayStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy-MM-dd");
    var startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd");

    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) return { success: true, pendingDates: [] };
    var shiftData = shiftSheet.getDataRange().getValues();
    
    var scheduledDates = [];
    for (var i = 1; i < shiftData.length; i++) {
      var row = shiftData[i];
      var wDate = row[1];
      var sId = row[2];
      if (!wDate || String(sId) !== String(staffId)) continue;
      
      var d = new Date(wDate);
      if (isNaN(d.getTime())) continue;
      var dStr = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      
      if (dStr >= startStr && dStr <= todayStr) {
        scheduledDates.push(dStr);
      }
    }

    var dataSheet = ss.getSheetByName("申請データ");
    if (dataSheet) {
      var kintaiData = dataSheet.getDataRange().getValues();
      for (var j = 1; j < kintaiData.length; j++) {
        var kRow = kintaiData[j];
        var kStaffId = kRow[1];
        var kDate = kRow[3];
        if (String(kStaffId) !== String(staffId) || !kDate) continue;
        
        var kd = new Date(kDate);
        if (isNaN(kd.getTime())) continue;
        var kdStr = Utilities.formatDate(kd, "Asia/Tokyo", "yyyy-MM-dd");
        
        var index = scheduledDates.indexOf(kdStr);
        if (index !== -1) {
          scheduledDates.splice(index, 1); // 申請済みなら除外
        }
      }
    }

    // 💡 日付を古い順（昇順）に並び替える
    scheduledDates.sort(function(a, b) {
      return new Date(a) - new Date(b);
    });

    return { success: true, pendingDates: scheduledDates };
  } catch (error) {
    return { success: false, pendingDates: [], message: error.message };
  }
}

/**
 * 💡 締め状況の取得 ＆ 詳細コンソールログ出力版
 */
function getClosingStatus(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = String(data.yearMonth || "").trim(); // 例: "2026-10"
    var half = String(data.half || "").trim(); // "first" または "second"
    
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var startDateStr = "";
    var endDateStr = "";
    
    if (half === "first") {
      startDateStr = yearMonth + "-01";
      endDateStr = yearMonth + "-15";
    } else {
      startDateStr = yearMonth + "-16";
      var lastDayObj = new Date(year, month, 0);
      endDateStr = yearMonth + "-" + String(lastDayObj.getDate()).padStart(2, '0');
    }

    console.log("==================================================");
    console.log("🔍 [デバッグ開始] 対象期間: " + startDateStr + " 〜 " + endDateStr);
    console.log("==================================================");

    // 日付を安全に "YYYY-MM-DD" 文字列に変換するヘルパー
    function parseDateToString(raw) {
      if (!raw) return "";
      if (raw instanceof Date) {
        return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      }
      var str = String(raw).trim();
      if (str.length >= 10 && str.indexOf("-") !== -1) {
        return str.substring(0, 10);
      }
      var d = new Date(raw);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      }
      return str;
    }

    // 1. スタッフマスタの検証
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) {
      console.log("❌ エラー: 「スタッフマスタ」シートが見つかりません");
      return { success: false, message: "スタッフマスタが見つかりません" };
    }
    var staffData = staffSheet.getDataRange().getValues();
    
    var staffList = [];
    for (var i = 1; i < staffData.length; i++) {
      var sId = String(staffData[i][0] || "").trim().toLowerCase();
      var sName = String(staffData[i][1] || "").trim();
      var sStatus = String(staffData[i][4] || "").trim();
      console.log("👤 スタッフ行[" + i + "] ID(小文字): [" + sId + "] 名前: [" + sName + "] ステータス: [" + sStatus + "]");
      
      if (sId && sStatus === "有効") {
        staffList.push({
          staffId: sId,
          originalId: String(staffData[i][0] || "").trim(),
          staffName: sName
        });
      }
    }
    console.log("✅ 有効なスタッフ数: " + staffList.length + "名");

    // 2. 「シフト予定」シートの検証
    var shiftSheet = ss.getSheetByName("シフト予定");
    var shiftMap = {}; 
    if (shiftSheet) {
      var shiftData = shiftSheet.getDataRange().getValues();
      console.log("📅 シフト予定シートの総行数: " + shiftData.length);
      for (var s = 1; s < shiftData.length; s++) {
        var rawDate = shiftData[s][1]; // B列：勤務日
        var sId = String(shiftData[s][2] || "").trim().toLowerCase(); // C列：スタッフID
        if (!rawDate || !sId) continue;
        
        var sDateStr = parseDateToString(rawDate);
        console.log("  -> シフト行[" + s + "] 日付変換: [" + sDateStr + "] スタッフID: [" + sId + "]");

        if (sDateStr >= startDateStr && sDateStr <= endDateStr) {
          if (!shiftMap[sId]) shiftMap[sId] = [];
          if (shiftMap[sId].indexOf(sDateStr) === -1) {
            shiftMap[sId].push(sDateStr);
            console.log("    🎯 【予定マッチ】スタッフ [" + sId + "] の日付 [" + sDateStr + "] を予定日数に加算します");
          }
        }
      }
    } else {
      console.log("⚠️ 警告: 「シフト予定」シートが見つかりません");
    }

    // 3. 「申請データ」シートの検証
    var dataSheet = ss.getSheetByName("申請データ");
    var kintaiMap = {};
    if (dataSheet) {
      var kData = dataSheet.getDataRange().getValues();
      console.log("📋 申請データシートの総行数: " + kData.length);
      for (var k = 1; k < kData.length; k++) {
        var rawKDate = kData[k][3]; // D列：勤務日
        var kId = String(kData[k][1] || "").trim().toLowerCase(); // B列：スタッフID
        if (!rawKDate || !kId) continue;
        
        var kDateStr = parseDateToString(rawKDate);
        console.log("  -> 申請行[" + k + "] 日付変換: [" + kDateStr + "] スタッフID: [" + kId + "]");

        if (kDateStr >= startDateStr && kDateStr <= endDateStr) {
          if (!kintaiMap[kId]) kintaiMap[kId] = [];
          if (kintaiMap[kId].indexOf(kDateStr) === -1) {
            kintaiMap[kId].push(kDateStr);
            console.log("    🎯 【申請マッチ】スタッフ [" + kId + "] の日付 [" + kDateStr + "] を申請日数に加算します");
          }
        }
      }
    } else {
      console.log("⚠️ 警告: 「申請データ」シートが見つかりません");
    }

    // 4. 「締め・確認管理」シートの検証
    var closingSheet = ss.getSheetByName("締め・確認管理");
    if (!closingSheet) {
      closingSheet = ss.insertSheet("締め・確認管理");
      closingSheet.appendRow(["staffId", "periodKey", "status", "updatedAt"]);
    }
    
    var statusMap = {}; 
    var cData = closingSheet.getDataRange().getValues();
    var periodKey = yearMonth + "_" + half;
    for (var c = 1; c < cData.length; c++) {
      var cStaffId = String(cData[c][0] || "").trim().toLowerCase();
      var cPeriod = String(cData[c][1] || "").trim();
      if (cPeriod === periodKey) {
        statusMap[cStaffId] = String(cData[c][2] || "").trim();
      }
    }
    
    var resultList = [];
    var allConfirmed = true;
    
    staffList.forEach(function(staff) {
      var scheduledDays = shiftMap[staff.staffId] ? shiftMap[staff.staffId].length : 0;
      var kintaiDays = kintaiMap[staff.staffId] ? kintaiMap[staff.staffId].length : 0;
      var status = statusMap[staff.staffId] || "確認待ち";
      
      console.log("📊 集計結果 -> スタッフ: " + staff.staffName + " (ID: " + staff.staffId + ") | 予定日数: " + scheduledDays + " | 申請日数: " + kintaiDays);

      if (status !== "確定済み") {
        allConfirmed = false;
      }
      
      resultList.push({
        staffId: staff.originalId,
        staffName: staff.staffName,
        scheduledDays: scheduledDays,
        kintaiDays: kintaiDays,
        status: status
      });
    });
    
    console.log("==================================================");
    return {
      success: true,
      startDate: startDateStr,
      endDate: endDateStr,
      resultList: resultList,
      allConfirmed: allConfirmed
    };
  } catch (error) {
    console.log("❌ 致命的エラー: " + error.message);
    return { success: false, message: "締め状況取得エラー: " + error.message };
  }
}

/**
 * 💡 締め確認・確定依頼のLINEメッセージを送信する関数
 */
function sendClosingReminderLine(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth, half = data.half, targetStaffIds = data.staffIds || [], targetMode = data.targetMode || "";
    var periodLabel = yearMonth.replace("-", "年") + "月度 " + (half === "first" ? "前半(1〜15日)" : "後半(16〜末日)");
    
    var statusRes = getClosingStatus({ yearMonth: yearMonth, half: half });
    if (!statusRes.success) return { success: false, message: statusRes.message };
    
    var sendTargets = [];
    statusRes.resultList.forEach(function(item) {
      if (item.status !== "確定済み" && (targetMode === "unconfirmed" || targetStaffIds.indexOf(item.staffId) !== -1)) {
        sendTargets.push(item.staffId);
      }
    });
    
    if (sendTargets.length === 0) return { success: false, message: "送信対象のスタッフがいません。" };
    
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    var staffData = staffSheet.getDataRange().getValues();
    var staffMap = {};
    var targetNames = []; // 💡 履歴用のスタッフ名収集配列
    for (var i = 1; i < staffData.length; i++) {
      if (staffData[i][0]) staffMap[staffData[i][0]] = { staffName: staffData[i][1], lineUserId: staffData[i][2], assignedAccount: staffData[i][3] };
    }
    
    var accountSheet = ss.getSheetByName("アカウントマスタ");
    var tokens = {};
    if (accountSheet) {
      var accData = accountSheet.getDataRange().getValues();
      for (var k = 1; k < accData.length; k++) {
        if (accData[k][0]) tokens[accData[k][0]] = accData[k][1];
      }
    }
    
    var messageText = "【勤怠の確認・確定のお願い】\n\nお疲れ様です！\n「" + periodLabel + "」の勤務実績について確認期間となりました。\n\nスタッフ専用ポータルを開き、申請内容に誤りがないかご確認の上、【内容を確定する】ボタンを押してご報告をお願いいたします！";
    var successCount = 0;
    var failCount = 0;
    
    sendTargets.forEach(function(sId) {
      var st = staffMap[sId];
      if (!st || !st.lineUserId) { failCount++; return; }
      var token = tokens[st.assignedAccount];
      if (!token || token.includes("YOUR_")) { failCount++; return; }
      
      targetNames.push(st.staffName);

      var res = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
        method: "post",
        headers: { "Content-Type": "application/json; charset=UTF-8", "Authorization": "Bearer " + token },
        payload: JSON.stringify({ to: st.lineUserId, messages: [{ type: "text", text: messageText }] }),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) successCount++; else failCount++;
    });
    
    // 💡 【追加】送信履歴（「連絡履歴」シート）へ自動保存する処理
    var histSheet = ss.getSheetByName("連絡履歴");
    if (!histSheet) {
      histSheet = ss.insertSheet("連絡履歴");
      histSheet.appendRow(["historyId", "timestamp", "targetName", "messageText", "resultSummary"]);
    }
    var historyId = "hist_" + Utilities.getUuid().substring(0, 8);
    var targetNameDisplay = "【締め確認依頼】" + (targetMode === "unconfirmed" ? "未確定者一斉" : targetNames.join(", "));
    var resultStr = "成功: " + successCount + "件 / 失敗: " + failCount + "件";
    histSheet.appendRow([historyId, new Date(), targetNameDisplay, messageText, resultStr]);

    return { success: true, message: successCount + "名に確認・確定依頼を送信し、履歴に記録しました！" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * 💡 スタッフ側：指定期間の締めステータスを取得する関数
 */
function getStaffClosingStatus(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffId = String(data.staffId || "").trim();
    var yearMonth = String(data.yearMonth || "").trim(); // "2026-10"
    var half = String(data.half || "").trim(); // "first" または "second"
    
    var periodKey = yearMonth + "_" + half;
    
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var startDateStr = "";
    var endDateStr = "";
    
    if (half === "first") {
      startDateStr = yearMonth + "-01";
      endDateStr = yearMonth + "-15";
    } else {
      startDateStr = yearMonth + "-16";
      var lastDayObj = new Date(year, month, 0);
      endDateStr = yearMonth + "-" + String(lastDayObj.getDate()).padStart(2, '0');
    }
    
    var dataSheet = ss.getSheetByName("申請データ");
    var kintaiDays = 0;
    var totalMinutes = 0;
    var totalCost = 0;
    var records = []; // 💡 明細を格納する配列
    
    function parseDate(raw) {
      if (!raw) return "";
      if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      var d = new Date(raw);
      return !isNaN(d.getTime()) ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd") : String(raw).substring(0, 10);
    }
    
    if (dataSheet) {
      var kData = dataSheet.getDataRange().getValues();
      for (var k = 1; k < kData.length; k++) {
        var kId = String(kData[k][1] || "").trim();
        var kDateStr = parseDate(kData[k][3]);
        if (kId === staffId && kDateStr >= startDateStr && kDateStr <= endDateStr) {
          kintaiDays++;
          var cost = Number(kData[k][7]) || 0;
          totalCost += cost;
          
          var startTime = formatTime(kData[k][4]);
          var endTime = formatTime(kData[k][5]);
          var breakMins = Number(kData[k][6]) || 0;
          var workMins = 0;
          if (startTime && endTime) {
            var sParts = startTime.split(":");
            var eParts = endTime.split(":");
            var startM = Number(sParts[0]) * 60 + Number(sParts[1]);
            var endM = Number(eParts[0]) * 60 + Number(eParts[1]);
            workMins = (endM - startM) - breakMins;
            if (workMins > 0) totalMinutes += workMins;
          }
          
          records.push({
            workDate: kDateStr,
            startTime: startTime,
            endTime: endTime,
            breakTime: breakMins,
            transportCost: cost
          });
        }
      }
    }
    
    records.sort(function(a, b) { return a.workDate > b.workDate ? 1 : -1; });
    
    var closingSheet = ss.getSheetByName("締め・確認管理");
    if (!closingSheet) {
      closingSheet = ss.insertSheet("締め・確認管理");
      closingSheet.appendRow(["staffId", "periodKey", "status", "updatedAt"]);
    }
    
    var status = "確認待ち";
    var cData = closingSheet.getDataRange().getValues();
    for (var c = 1; c < cData.length; c++) {
      if (String(cData[c][0]).trim() === staffId && String(cData[c][1]).trim() === periodKey) {
        status = String(cData[c][2]).trim();
        break;
      }
    }
    
    var totH = Math.floor(totalMinutes / 60);
    var totM = totalMinutes % 60;
    
    // 💡 スタッフごとの適用時給と給与総額を計算
    var hourlyWage = getEffectiveHourlyWage(staffId, yearMonth, half);
    var totalHours = totalMinutes / 60;
    var estimatedSalary = Math.round(totalHours * hourlyWage);

    return {
      success: true,
      startDate: startDateStr,
      endDate: endDateStr,
      kintaiDays: kintaiDays,
      totalWorkHoursStr: totH + "時間" + totM + "分",
      totalCost: totalCost,
      hourlyWage: hourlyWage,           // 👈 ここを追加
      estimatedSalary: estimatedSalary, // 👈 ここを追加
      status: status,
      records: records
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * 💡 スタッフ側：指定期間の内容を確定する関数
 */
function confirmStaffClosingStatus(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffId = String(data.staffId || "").trim();
    var yearMonth = String(data.yearMonth || "").trim();
    var half = String(data.half || "").trim();
    var periodKey = yearMonth + "_" + half;
    
    var closingSheet = ss.getSheetByName("締め・確認管理");
    if (!closingSheet) {
      closingSheet = ss.insertSheet("締め・確認管理");
      closingSheet.appendRow(["staffId", "periodKey", "status", "updatedAt"]);
    }
    
    var cData = closingSheet.getDataRange().getValues();
    var targetRow = -1;
    
    for (var c = 1; c < cData.length; c++) {
      if (String(cData[c][0]).trim() === staffId && String(cData[c][1]).trim() === periodKey) {
        targetRow = c + 1;
        break;
      }
    }
    
    var timestamp = new Date();
    if (targetRow !== -1) {
      closingSheet.getRange(targetRow, 3).setValue("確定済み");
      closingSheet.getRange(targetRow, 4).setValue(timestamp);
    } else {
      closingSheet.appendRow([staffId, periodKey, "確定済み", timestamp]);
    }
    
    return { success: true, message: "お疲れ様でした！勤務内容を確定し、管理者に報告しました。" };
  } catch (error) {
    return { success: false, message: "確定エラー: " + error.message };
  }
}

/**
 * 💡 指定したスタッフの対象月における既存のシフト希望（申請済み日）のリストを取得する関数
 */
function getStaffSubmittedShiftDates(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) return { success: true, submittedDates: [] };
    
    var shiftData = sheet.getDataRange().getValues();
    var staffId = String(data.staffId || "").trim();
    var targetYearMonth = String(data.yearMonth || "").trim(); // 例: "2026-10"
    var submittedDates = [];
    
    for (var i = 1; i < shiftData.length; i++) {
      var rawDate = shiftData[i][1];
      var rowStaffId = String(shiftData[i][2] || "").trim();
      
      if (!rawDate || rowStaffId !== staffId) continue;
      
      var dateStr = "";
      if (rawDate instanceof Date) {
        dateStr = Utilities.formatDate(rawDate, "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        var d = new Date(rawDate);
        if (!isNaN(d.getTime())) {
          dateStr = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
        } else {
          dateStr = String(rawDate).substring(0, 10);
        }
      }
      
      if (dateStr.indexOf(targetYearMonth) === 0) {
        if (submittedDates.indexOf(dateStr) === -1) {
          submittedDates.push(dateStr);
        }
      }
    }
    
    return { success: true, submittedDates: submittedDates };
  } catch (error) {
    return { success: false, submittedDates: [], message: error.message };
  }
}

/**
 * 💡 指定された年月のシフト情報をSlack共有用テキストに変換して返す関数（スタッフのシフトのみ）
 */
function getSlackShiftText(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!yearMonth) return { success: false, message: "対象月が指定されていません" };
    
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    
    // タイトル行（例: "10月シフト"）
    var slackText = month + "月シフト\n";
    
    var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
    var lastDay = new Date(year, month, 0).getDate();
    
    // 1. シフト予定データを取得
    var shiftSheet = ss.getSheetByName("シフト予定");
    var shiftMap = {};
    if (shiftSheet) {
      var shiftData = shiftSheet.getDataRange().getValues();
      for (var i = 1; i < shiftData.length; i++) {
        var workDate = shiftData[i][1];
        var staffName = String(shiftData[i][3] || "").trim();
        var dailyMemo = String(shiftData[i][4] || "").trim(); // 例: "pm" など
        
        if (!workDate || !staffName) continue;
        var d = new Date(workDate);
        if (isNaN(d.getTime())) continue;
        
        var y = d.getFullYear();
        var m = ('0' + (d.getMonth() + 1)).slice(-2);
        var day = ('0' + d.getDate()).slice(-2);
        var dateStr = y + '-' + m + '-' + day;
        
        if (dateStr.indexOf(yearMonth) === 0) {
          if (!shiftMap[dateStr]) shiftMap[dateStr] = [];
          var displayName = dailyMemo ? (staffName + "（" + dailyMemo + "）") : staffName;
          shiftMap[dateStr].push(displayName);
        }
      }
    }
    
    // 2. 1日から末日までループしてスタッフのシフトのみテキストを構築
    var hasContent = false;
    for (var day = 1; day <= lastDay; day++) {
      var mStr = String(month).padStart(2, '0');
      var dStr = String(day).padStart(2, '0');
      var dateStr = year + '-' + mStr + '-' + dStr;
      
      var dObj = new Date(year, month - 1, day);
      var wStr = weekDays[dObj.getDay()];
      
      // スタッフのシフトが存在する場合のみ行を出力
      if (shiftMap[dateStr] && shiftMap[dateStr].length > 0) {
        slackText += day + "(" + wStr + ")" + shiftMap[dateStr].join("、") + "\n";
        hasContent = true;
      }
    }
    
    return { success: true, text: slackText.trim() };
  } catch (error) {
    return { success: false, message: "Slackテキスト生成エラー: " + error.message };
  }
}

/**
 * 💡 指定したスタッフの指定期（例: 2026-10_first）における有効な時給を引く関数
 * （設定がない場合は、直近の過去に遡って設定された時給を自動引き継ぎする）
 */
function getEffectiveHourlyWage(staffId, yearMonth, half) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("時給管理");
    if (!sheet) return 1000; // シートがない場合のデフォルト時給（例: 1000円）
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 1000;
    
    // 対象の期を数値化して比較しやすくする（例: "2026-10_first" -> 年月と前半後半）
    var targetYm = String(yearMonth).trim();
    var targetHalf = String(half).trim();
    var targetValue = Number(targetYm.replace("-", "")) * 2 + (targetHalf === "second" ? 1 : 0);
    
    var staffWageMap = {}; // this staff's wage history
    var staffDefaultWage = 1000;
    
    for (var i = 1; i < data.length; i++) {
      var sId = String(data[i][0] || "").trim();
      var pKey = String(data[i][1] || "").trim(); // "2026-10_first"
      var wage = Number(data[i][2]) || 0;
      
      if (sId === String(staffId).trim()) {
        if (pKey === "default") {
          staffDefaultWage = wage;
        } else if (pKey.indexOf("_") !== -1) {
          var parts = pKey.split("_");
          var pYm = parts[0];
          var pHalf = parts[1];
          var pVal = Number(pYm.replace("-", "")) * 2 + (pHalf === "second" ? 1 : 0);
          staffWageMap[pVal] = wage;
        }
      }
    }
    
    // ① 完全一致する期があればそれを返す
    var exactVal = targetValue;
    if (staffWageMap[exactVal] !== undefined) {
      return staffWageMap[exactVal];
    }
    
    // ② なければ、過去に向かって最も近い設定済み時給を探して引き継ぐ
    var bestVal = -1;
    var bestWage = -1;
    for (var key in staffWageMap) {
      var kNum = Number(key);
      if (kNum <= targetValue && kNum > bestVal) {
        bestVal = kNum;
        bestWage = staffWageMap[kNum];
      }
    }
    
    if (bestWage !== -1) {
      return bestWage;
    }
    
    return staffDefaultWage;
  } catch (e) {
    return 1000;
  }
}

/**
 * 💡 管理者側：スタッフの時給を保存・更新する関数
 */
function saveStaffHourlyWage(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("時給管理");
    if (!sheet) {
      sheet = ss.insertSheet("時給管理");
      sheet.appendRow(["staffId", "periodKey", "hourlyWage"]);
    }
    
    var staffId = String(data.staffId).trim();
    var periodKey = String(data.periodKey).trim(); // 例: "2026-10_first" または "default"
    var hourlyWage = Number(data.hourlyWage) || 0;
    
    var values = sheet.getDataRange().getValues();
    var targetRow = -1;
    
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === staffId && String(values[i][1]).trim() === periodKey) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow !== -1) {
      sheet.getRange(targetRow, 3).setValue(hourlyWage);
    } else {
      sheet.appendRow([staffId, periodKey, hourlyWage]);
    }
    
    return { success: true, message: "時給を保存しました！" };
  } catch (error) {
    return { success: false, message: "保存エラー: " + error.message };
  }
}

/**
 * 💡 指定した期の全スタッフの時給リストを取得する関数
 */
function getAllStaffWagesForPeriod(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: false, message: "スタッフマスタが見つかりません" };
    
    var staffData = staffSheet.getDataRange().getValues();
    var activeStaff = [];
    
    for (var i = 1; i < staffData.length; i++) {
      var sId = String(staffData[i][0] || "").trim();
      var sName = String(staffData[i][1] || "").trim();
      var status = String(staffData[i][4] || "").trim();
      
      if (sId && status === "有効") {
        var wage = getEffectiveHourlyWage(sId, data.yearMonth, data.half);
        activeStaff.push({
          staffId: sId,
          staffName: sName,
          wage: wage
        });
      }
    }
    
    return { success: true, staffList: activeStaff };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * 💡 スタッフのステータス（有効/退職）を切り替える関数
 */
function changeStatus(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) return { success: false, message: "スタッフマスタが見つかりません" };
    
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(data.staffId).trim()) {
        sheet.getRange(i + 1, 5).setValue(data.newStatus); // E列(5列目)がステータス
        return { success: true, message: "スタッフのステータスを「" + data.newStatus + "」に変更しました！" };
      }
    }
    return { success: false, message: "対象スタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "エラー: " + error.message };
  }
}

/**
 * 💡 スタッフが退職（無効）状態かどうかを判定する関数
 */
function isStaffRetired(staffId) {
  if (!staffId) return true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("スタッフマスタ");
  if (!sheet) return false;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(staffId).trim()) {
      var status = String(data[i][4] || "").trim(); // E列(5列目)がステータス
      return (status === "退職");
    }
  }
  return false; // スタッフが見つからない場合は一応開けるようにするか、お好みで
}

/**
 * 💡 スプレッドシートの「設定」シートから登録上限人数を取得する関数
 */
function getMaxStaffLimit() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("設定");
    if (!sheet) {
      // 「設定」シートがなければ自動作成してデフォルト9をセット
      sheet = ss.insertSheet("設定");
      sheet.appendRow(["設定キー", "設定値"]);
      sheet.appendRow(["MAX_STAFF_LIMIT", 9]);
      return 9;
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "MAX_STAFF_LIMIT") {
        var val = Number(data[i][1]);
        return !isNaN(val) && val > 0 ? val : 9;
      }
    }
    
    // 見つからない場合は行を追加して9を返す
    sheet.appendRow(["MAX_STAFF_LIMIT", 9]);
    return 9;
  } catch (e) {
    return 9;
  }
}

/**
 * 💡 ダッシュボードからの変更を「設定」シートに保存する関数
 */
function setMaxStaffLimit(newLimit) {
  try {
    var limitNum = Number(newLimit);
    if (isNaN(limitNum) || limitNum < 1) {
      return { success: false, message: "有効な数値を入力してください。" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("設定");
    if (!sheet) {
      sheet = ss.insertSheet("設定");
      sheet.appendRow(["設定キー", "設定値"]);
    }
    
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "MAX_STAFF_LIMIT") {
        sheet.getRange(i + 1, 2).setValue(limitNum);
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow(["MAX_STAFF_LIMIT", limitNum]);
    }
    
    return { success: { limit: limitNum }, message: "登録上限を「" + limitNum + "人」に更新しました！" };
  } catch (e) {
    return { success: false, message: "保存エラー: " + e.message };
  }
}

/**
 * 💡 指定したスタッフ・日付の既存申請件数を数える関数
 */
function checkKintaiDuplication(staffId, workDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    var count = 0;
    
    // 日付を安全に比較するためのフォーマット関数
    function parseDateStr(raw) {
      if (!raw) return "";
      if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      var str = String(raw).trim();
      if (str.length >= 10) return str.substring(0, 10);
      var d = new Date(raw);
      return !isNaN(d.getTime()) ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd") : str;
    }
    
    var targetDateStr = parseDateStr(workDate);
    
    for (var i = 1; i < data.length; i++) {
      var rowStaffId = String(data[i][1] || "").trim();
      var rowDateStr = parseDateStr(data[i][3]);
      
      if (rowStaffId === String(staffId).trim() && rowDateStr === targetDateStr) {
        count++;
      }
    }
    return count;
  } catch (e) {
    return 0;
  }
}

/**
 * 💡 管理者ダッシュボード用：直近で同じ日に複数回（2回以上）申請している「重複勤務」を検出する関数
 */
function getDuplicateKintaiAlerts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return { success: true, duplicates: [] };
    
    var data = sheet.getDataRange().getValues();
    var map = {}; // "staffId_date" -> { staffName, workDate, count }
    
    function parseDateStr(raw) {
      if (!raw) return "";
      if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      var str = String(raw).trim();
      if (str.length >= 10) return str.substring(0, 10);
      var d = new Date(raw);
      return !isNaN(d.getTime()) ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd") : str;
    }
    
    for (var i = 1; i < data.length; i++) {
      var sId = String(data[i][1] || "").trim();
      var sName = String(data[i][2] || "").trim();
      var wDate = parseDateStr(data[i][3]);
      
      if (!sId || !wDate) continue;
      
      var key = sId + "_" + wDate;
      if (!map[key]) {
        map[key] = { staffId: sId, staffName: sName, workDate: wDate, count: 0 };
      }
      map[key].count++;
    }
    
    var duplicateList = [];
    for (var key in map) {
      if (map[key].count > 1) {
        duplicateList.push(map[key]);
      }
    }
    
    // 日付の新しい順にソート
    duplicateList.sort(function(a, b) {
      return new Date(b.workDate) - new Date(a.workDate);
    });
    
    return { success: true, duplicates: duplicateList.slice(0, 15) }; // 最大15件
  } catch (e) {
    return { success: false, duplicates: [], message: e.message };
  }
}

/**
 * 💡 指定したスタッフ・日付の既存申請件数を数える関数
 */
function checkKintaiDuplication(staffId, workDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    var count = 0;
    
    function parseDateStr(raw) {
      if (!raw) return "";
      if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      var str = String(raw).trim();
      if (str.length >= 10) return str.substring(0, 10);
      var d = new Date(raw);
      return !isNaN(d.getTime()) ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd") : str;
    }
    
    var targetDateStr = parseDateStr(workDate);
    
    for (var i = 1; i < data.length; i++) {
      var rowStaffId = String(data[i][1] || "").trim();
      var rowDateStr = parseDateStr(data[i][3]);
      
      if (rowStaffId === String(staffId).trim() && rowDateStr === targetDateStr) {
        count++;
      }
    }
    return count;
  } catch (e) {
    return 0;
  }
}

/**
 * 💡 管理者ダッシュボード用：直近で同じ日に複数回（2回以上）申請している「重複勤務」を検出する関数
 */
function getDuplicateKintaiAlerts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return { success: true, duplicates: [] };
    
    var data = sheet.getDataRange().getValues();
    var map = {}; 
    
    function parseDateStr(raw) {
      if (!raw) return "";
      if (raw instanceof Date) return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
      var str = String(raw).trim();
      if (str.length >= 10) return str.substring(0, 10);
      var d = new Date(raw);
      return !isNaN(d.getTime()) ? Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd") : str;
    }
    
    for (var i = 1; i < data.length; i++) {
      var sId = String(data[i][1] || "").trim();
      var sName = String(data[i][2] || "").trim();
      var wDate = parseDateStr(data[i][3]);
      
      if (!sId || !wDate) continue;
      
      var key = sId + "_" + wDate;
      if (!map[key]) {
        map[key] = { staffId: sId, staffName: sName, workDate: wDate, count: 0 };
      }
      map[key].count++;
    }
    
    var duplicateList = [];
    for (var key in map) {
      if (map[key].count > 1) {
        duplicateList.push(map[key]);
      }
    }
    
    duplicateList.sort(function(a, b) {
      return new Date(b.workDate) - new Date(a.workDate);
    });
    
    return { success: true, duplicates: duplicateList.slice(0, 15) };
  } catch (e) {
    return { success: false, duplicates: [], message: e.message };
  }
}

/**
 * 💡 管理画面の手動登録タブから呼び出される登録関数
 */
function registerStaffByAdmin(data) {
  return registerNewStaff(data);
}