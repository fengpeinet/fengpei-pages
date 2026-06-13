/**
 * FENGPEINET 說明會推廣報名系統 - 圖文配置 + 完整日誌版
 */

// ==========================================
//   一、 營運設定區 (連結與圖片)
// ==========================================

const ZOOM_LINK_6_16 = "https://us06web.zoom.us/j/3909390259?omn=84699032198";
const ZOOM_LINK_6_18 = "https://us06web.zoom.us/j/3909390259?omn=86322017358";

const HEAD_IMAGE_URL = "https://lh3.googleusercontent.com/d/1XG6Co4hBQ0r0Yk5c5Zc-I3zhHftUYt3B";
const FOOTER_IMAGE_URL = "https://lh3.googleusercontent.com/d/1MAo2woNkcN7-LYjjoL6ZMdlFIToLyz-s";
const POSTER_IMAGE_URL = "https://lh3.googleusercontent.com/d/1-unnfEgNAHtFiD9cAhKPLy_cpGVSVJJ0";
const CONTENT_IMAGE_URL = "https://lh3.googleusercontent.com/d/1njxgyVqOTcnSp4KdC7jk4uWeouwhMZDw";
const SHEET_NAME = "工作表1";


// ==========================================
//   二、 排程發信功能 (含詳細日誌)
// ==========================================

function sendScheduledEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  console.log("--- 執行排程掃描開始 ---");

  if (!sheet) {
    console.error("錯誤：找不到工作表 [" + SHEET_NAME + "]");
    return;
  }

  const data = sheet.getDataRange().getValues();
  console.log("偵測到總列數 (含標題): " + data.length);

  if (data.length <= 1) {
    console.log("目前沒有報名資料。");
    return;
  }

  const today = new Date();
  const timeZone = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(today, timeZone, "yyyy-M-d");

  const tomorrow = new Date();
  tomorrow.setDate(today.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, timeZone, "yyyy-M-d");

  console.log("今日基準日期: " + todayStr);
  console.log("明日基準日期: " + tomorrowStr);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const userName = row[1];
    const userEmail = row[3];
    const rawDate = row[6];
    const isDayBeforeReminded = row[7];
    const isDayOfReminded = row[8];

    if (!userName || !userEmail || !rawDate) {
      console.log("第 " + (i + 1) + " 列資料不完整，跳過。");
      continue;
    }

    // 日期標準化（用於比對）
    let targetDateStr = "";
    if (rawDate instanceof Date) {
      targetDateStr = Utilities.formatDate(rawDate, timeZone, "yyyy-M-d");
    } else {
      targetDateStr = rawDate.toString().replace(/-/g, "/").split(" ")[0];
      let tempDate = new Date(targetDateStr);
      if (!isNaN(tempDate.getTime())) {
        targetDateStr = Utilities.formatDate(tempDate, timeZone, "yyyy-M-d");
      }
    }

    // 顯示用日期（保留時間，去除"台北"）
    let displayDate = "";
    if (rawDate instanceof Date) {
      displayDate = Utilities.formatDate(rawDate, timeZone, "yyyy/MM/dd HH:mm");
    } else {
      displayDate = rawDate.toString().replace(/台北/g, "").trim();
    }

    console.log("正在檢查: " + userName + " | 場次: " + targetDateStr);

    let zoomLink = targetDateStr.includes("-6-18") ? ZOOM_LINK_6_18 : ZOOM_LINK_6_16;

    // --- 今日提醒 ---
    if (targetDateStr === todayStr) {
      if (!isDayOfReminded || isDayOfReminded === "FALSE" || isDayOfReminded === false) {
        sendReminderEmail(userEmail, userName, displayDate, zoomLink, "今天");
        sheet.getRange(i + 1, 9).setValue(true);
        console.log("   >>> [成功] 已寄出今日提醒信");
      } else {
        console.log("   (今日提醒已寄過，跳過)");
      }
    }

    // --- 明日提醒 ---
    if (targetDateStr === tomorrowStr) {
      if (!isDayBeforeReminded || isDayBeforeReminded === "FALSE" || isDayBeforeReminded === false) {
        sendReminderEmail(userEmail, userName, displayDate, zoomLink, "明天");
        sheet.getRange(i + 1, 8).setValue(true);
        console.log("   >>> [成功] 已寄出明日提醒信");
      } else {
        console.log("   (明日提醒已寄過，跳過)");
      }
    }
  }
  SpreadsheetApp.flush();
  console.log("--- 掃描結束 ---");
}


// ==========================================
//   三、 精美信件 HTML 排版
// ==========================================

function sendReminderEmail(to, name, displayDate, link, dayWord) {
  const subject = `📣 ${dayWord}見！多品牌分潤創業系統分享會 活動提醒`;
  const footerUrl = "https://lh3.googleusercontent.com/d/1MAo2woNkcN7-LYjjoL6ZMdlFIToLyz-s";
  const headUrl = "https://lh3.googleusercontent.com/d/1XG6Co4hBQ0r0Yk5c5Zc-I3zhHftUYt3B";
  const brandLogoUrl = "https://lh3.googleusercontent.com/d/12c_FvDGcLxbMlDR4caLLjeElunPRiOrB";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">

      <img src="${headUrl}" alt="多品牌分潤創業系統分享會" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 24px;">

      <p style="margin: 0 0 8px 0;"><strong>${name}</strong> 您好：</p>
      <p style="margin: 0 0 20px 0;">多品牌分潤創業系統分享會<strong>${dayWord}</strong>就要開始囉！</p>

      <!-- 活動資訊 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 活動資訊</p>
      <p style="margin: 4px 0 4px 16px;">活動名稱：<strong>多品牌分潤創業系統分享會</strong></p>
      <p style="margin: 4px 0 4px 16px;">活動時間：<strong>${displayDate}</strong></p>
      <p style="margin: 4px 0 20px 16px;">活動形式：線上 Zoom</p>

      <!-- 你將會了解 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 這場分享會，你將會了解</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <li style="margin-bottom: 6px;">多品牌分潤的運作邏輯，如何透過多元品牌建立<strong>穩定收入來源</strong></li>
        <li style="margin-bottom: 6px;">從零開始搭建屬於自己的分潤系統，<strong>不需要囤貨、不需要技術背景</strong></li>
        <li style="margin-bottom: 6px;">一般人也能透過這套系統創造<strong>副業收入</strong>，甚至建立被動收益</li>
        <li style="margin-bottom: 6px;">現場開放提問，協助你釐清<strong>最適合自己的起步方式</strong></li>
      </ul>

      <!-- 橘色框：三步驟 + Zoom 按鈕 -->
      <div style="border: 2px solid #ff6600; border-radius: 8px; padding: 16px 20px; margin: 0 0 20px 0; background-color: #fff8f0;">
        <p style="margin: 0 0 10px 0; color: #ff6600; font-weight: bold;">請依照下面簡單三步驟就可參加：</p>
        <ol style="margin: 0 0 20px 0; padding-left: 20px;">
          <li style="margin-bottom: 8px;"><strong>下載會議軟體：</strong>我們使用「Zoom」作為線上遠端教學軟體。</li>
          <li style="margin-bottom: 8px;"><strong>點擊下方按鈕：</strong>直接進入線上研討會教室。</li>
          <li style="margin-bottom: 8px;"><strong>準時出席：</strong>活動當天請提前 10 分鐘進入測試設備。</li>
        </ol>
        <p style="text-align: center; margin: 0;">
          <a href="${link}" style="background-color: #ff6600; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 16px;">立即進入 Zoom 研討會</a>
        </p>
      </div>

      <!-- 參加須知 -->
      <p style="font-weight: bold; margin: 0 0 8px 0;">✅ 參加須知</p>
      <ul style="margin: 0 0 20px 0; padding-left: 20px;">
        <li style="margin-bottom: 6px;">請於活動開始前 <strong>10 分鐘</strong>進入會議室，以免錯過開場內容</li>
        <li style="margin-bottom: 6px;">線上分享，請提前確認網路與設備，<strong>建議使用電腦觀看效果較佳</strong></li>
        <li style="margin-bottom: 6px;">建議準備筆記工具，活動中會分享重點內容與資源連結</li>
      </ul>

      <p style="margin: 0 0 20px 0;">期待這場分享會能帶給您新的啟發與方向，我們活動當天見！</p>

      <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

      <table style="width: 100%; max-width: 600px; margin: 0 0 24px 0; border-collapse: collapse;">
        <tr>
          <td style="width: 180px; vertical-align: middle; padding-right: 16px;">
            <img src="${brandLogoUrl}" alt="風霈國際傳媒" style="display: block; width: 160px; height: auto;">
          </td>
          <td style="vertical-align: middle;">
            <p style="margin: 0 0 6px 0; font-size: 14px;">如有問題歡迎詢問官方 LINE：<a href="https://lin.ee/RHr2pNi">https://lin.ee/RHr2pNi</a></p>
            <p style="margin: 0; color: #999; font-size: 13px;">※ 此信件為系統自動發送，請勿直接回覆此信件。</p>
          </td>
        </tr>
      </table>

      <img src="${footerUrl}" alt="" style="width: 100%; max-width: 600px; display: block;">

    </div>
  `;
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody });
}

function sendRegistrationSuccessEmail(to, name, date) {
  const subject = "✅ 您已成功報名「多品牌分潤創業系統分享會」";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', Arial, sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">

      <img src="${POSTER_IMAGE_URL}" alt="多品牌分潤創業系統分享會" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 24px;">

      <p style="margin: 0 0 8px 0;"><strong>${name}</strong> 您好：</p>
      <p style="margin: 0 0 20px 0;">感謝您報名 <strong>多品牌分潤創業系統分享會</strong>，您的報名已成功完成！</p>

      <div style="background-color: #f9f9f9; border-left: 4px solid #c8a84b; padding: 16px 20px; border-radius: 4px; margin: 0 0 20px 0;">
        <p style="margin: 0 0 8px 0; color: #a07830; font-weight: bold;">📋 您的報名資訊</p>
        <p style="margin: 4px 0;">報名姓名：<strong>${name}</strong></p>
        <p style="margin: 4px 0;">活動名稱：<strong>多品牌分潤創業系統分享會</strong></p>
        <p style="margin: 4px 0;">場次時間：<strong>${date}</strong></p>
      </div>

      <p style="margin: 0 0 16px 0;">我們將在活動<strong>前一天與當天</strong>，另行寄送 Zoom 會議連結與活動提醒，請留意信箱！</p>

      <p style="margin: 0 0 20px 0;">如有任何問題，歡迎透過官方 LINE 聯繫我們：<a href="https://lin.ee/RHr2pNi" style="color: #06C755;">https://lin.ee/RHr2pNi</a></p>

      <img src="${CONTENT_IMAGE_URL}" alt="活動內容" style="width: 100%; max-width: 600px; display: block; border-radius: 8px; margin-bottom: 20px;">

      <p style="margin: 0 0 4px 0;">期待在研討會上與您相見！</p>
      <p style="margin: 0;"><strong>風霈國際傳媒 敬上</strong></p>

    </div>
  `;
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody });
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const data = JSON.parse(e.postData.contents);
    sheet.appendRow([new Date(), data.name, data.phone, data.email, data.lineId || '', data.referrer || '', data.sessionDate, false, false]);
    sendRegistrationSuccessEmail(data.email, data.name, data.sessionDate);
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  } catch (f) { return ContentService.createTextOutput(f.toString()); }
}
