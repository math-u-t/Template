function send_Scheduled_Messages() {
  const startTime = new Date().getTime();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const currentTime = new Date();
  const rowsToDelete = [];
  const messagesToSend = [];

  for (let i = 1; i < data.length; i++) {
    const elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > 5) {
      throw new Error('Execution time exceeded 5 seconds, stopping script.');
    }

    const row = data[i];
    const email = row[1];
    const message = row[3];
    const sendTime = row[5] ? new Date(row[5]) : null;
    const webhookUrl = row[6];

    if (!sendTime || (currentTime - sendTime) / (1000 * 60) >= 2) {
      rowsToDelete.push(i + 1);
      continue;
    }

    if (!message || !webhookUrl) {
      continue;
    }

    if (sendTime <= currentTime) {
      let finalMessage = message;

      if (finalMessage.includes('<hide>')) {
        finalMessage = finalMessage.replace(/<hide>$/, '').trim();

      } else {
        if (email) {
          finalMessage += `\n《${email}》\n`;
        }
      }

      messagesToSend.push({ url: webhookUrl, message: finalMessage, rowIndex: i + 1 });
    }
  }

  messagesToSend.forEach(({ url, message, rowIndex }) => {
    const payload = JSON.stringify({ text: message });
    const options = {
      method: "post",
      contentType: "application/json",
      payload: payload,
    };

    try {
      UrlFetchApp.fetch(url, options);
      rowsToDelete.push(rowIndex);
    } catch (error) {
    }
  });

  if (rowsToDelete.length > 0) {
    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  }
}

// エラーの詳細
function send_Custom_Email() {
  const startTime = new Date().getTime();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(lastRow, 1, 1, 7).getValues()[0];
  if (data[2] !== '送信予約') return;
  var scheduledTime = new Date(data[5]);
  var now = new Date();
  var diffMinutes = (scheduledTime - now) / 60000;
  var elapsedTime = (new Date().getTime() - startTime) / 1000;
  if (elapsedTime > 5) {
    throw new Error('Execution time exceeded 5 seconds, stopping script.');
  }

  if (data[3].length > 4000) {
    send_Error_Email(data[1], "送信内容が4000文字を超えています。\n送信内容を短くしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  elapsedTime = (new Date().getTime() - startTime) / 1000;
  if (elapsedTime > 5) {
    throw new Error('Execution time exceeded 5 seconds, stopping script.');
  }

  if (isNaN(scheduledTime.getTime()) || diffMinutes < -2) {
    send_Error_Email(data[1], "送信予定時間が過去です。\n送信予定時間は未来になるようにしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  if (diffMinutes > 60 * 24 * 30) {
    send_Error_Email(data[1], "送信予定時間が1か月以上先です。\n送信予定時間は1か月よりも最近にしてください。");
    sheet.deleteRow(lastRow);
    return;
  }

  var userCode = generate_Random_Code(30);
  sheet.getRange(lastRow, 5).setValue(userCode);
  var emailContent = create_Email_Content(data, userCode, scheduledTime);

  elapsedTime = (new Date().getTime() - startTime) / 1000;
  if (elapsedTime > 5) {
    throw new Error('Execution time exceeded 5 seconds, stopping script.');
  }

  try {
    GmailApp.sendEmail(data[1], "自動送信予約の詳細", emailContent, { name: "予約の確認（自動）" });
    sheet.getRange(lastRow, 3).setValue('送信済み');
  } catch (e) {
  }
}

// エラーメッセージ送信
function send_Error_Email(recipient, errorMessage) {
  var subject = "予約できませんでした";
  var body = "詳細：" + errorMessage +
    "\nもう一度予約を行ってください。" +
    "\n\nこのメッセージは自動送信されています。" +
    "\n返信しないでください。";
  generalization_send_Email(recipient, subject, body);
}

// 成功メッセージ送信
function create_Email_Content(data, userCode, scheduledTime) {
  return [
    "予約者 : " + data[1],
    "送信内容 : " + data[3],
    "送信予定時間 : " + format_Date(scheduledTime),
    "webhookURL : " + data[6],
    "予約コード : " + userCode,
    "予約が完了しました。",
    "\nご利用ありがとうございます。",
    "このメッセージは自動で送信されています。",
    "返信しないでください。"
  ].join("\n");
}

// 日付形式変換
function format_Date(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');
  const seconds = date.getSeconds().toString().padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// 予約コード生成
function generate_Random_Code(length) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  const charsLength = chars.length;
  let code = new Array(length);
  
  for (let i = 0; i < length; i++) {
    code[i] = chars[Math.floor(Math.random() * charsLength)];
  }

  return code.join('');
}

// 一般化送信
function generalization_send_Email(email, subject, body) {
  try {
    GmailApp.sendEmail(email, subject, body, { name: "予約の確認《自動送信》" });
  } catch (e) {
  }
}
