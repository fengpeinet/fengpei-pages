/**
 * FENGPEINET 說明會推廣報名系統 - 圖文配置 + 完整日誌版
 */

// ==========================================
//   一、 營運設定區 (連結與圖片)
// ==========================================

// 【修改點 1】 設定不同日期的 Zoom 連結
const ZOOM_LINK_6_16 = "https://zoom.us/j/YOUR_ZOOM_ID_1";
const ZOOM_LINK_6_18 = "https://zoom.us/j/YOUR_ZOOM_ID_2";

/**
 * 【修改點 2】 頂部海報圖片
 */
const POSTER_IMAGE_URL = "https://lh3.googleusercontent.com/d/1YTTlQtkoCd3KL00IVZda3j_S61NkM0wu";

/**
 * 【修改點 3】 內容區圖片 (文字下方)
 */
const CONTENT_IMAGE_URL = "https://lh3.googleusercontent.com/d/1njxgyVqOTcnSp4KdC7jk4uWeouwhMZDw";

// 【修改點 4】 指定工作表名稱
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
    const userName = row[1];    // B 欄
    const userEmail = row[3];   // D 欄
    const rawDate = row[6];     // G 欄
    const isDayBeforeReminded = row[7]; // H 欄
    const isDayOfReminded = row[8];     // I 欄

    if (!userName || !userEmail || !rawDate) {
      console.log("第 " + (i + 1) + " 列資料不完整，跳過。");
      continue;
    }

    // 日期標準化
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

    console.log("正在檢查: " + userName + " | 場次: " + targetDateStr);

    let zoomLink = targetDateStr.includes("-6-18") ? ZOOM_LINK_6_18 : ZOOM_LINK_6_16;

    // --- 檢查「今日提醒」 ---
    if (targetDateStr === todayStr) {
      if (!isDayOfReminded || isDayOfReminded === "FALSE" || isDayOfReminded === false) {
        sendReminderEmail(userEmail, userName, targetDateStr, zoomLink, "【今日提醒】");
        sheet.getRange(i + 1, 9).setValue(true);
        console.log("   >>> [成功] 已寄出今日提醒信");
      } else {
        console.log("   (今日提醒已寄過，跳過)");
      }
    }

    // --- 檢查「明日提醒」 ---
    if (targetDateStr === tomorrowStr) {
      if (!isDayBeforeReminded || isDayBeforeReminded === "FALSE" || isDayBeforeReminded === false) {
        sendReminderEmail(userEmail, userName, targetDateStr, zoomLink, "【明日提醒】");
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

function sendReminderEmail(to, name, date, link, prefix) {
  const subject = prefix + " 風霈學院 × 蝦皮 AI 攻略班說明會連結";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">
      <img src="${POSTER_IMAGE_URL}" alt="海報" style="width: 100%; max-width: 600px; border-radius: 8px; margin-bottom: 20px;">
      <p>親愛的 <strong>${name}</strong> 同學好：</p>
      <p>本週的<strong>「風霈學院 × 蝦皮 AI 攻略班」</strong>線上研討會即將在 ${prefix.includes("今日") ? "今天" : "明天"} 開始了！</p>
      <p>面對技術的不斷快速更迭，你是不是也常常因為要追新工具而感到焦慮呢？<br>其實，在變動中我們更應該找到那些「不變」的核心！</p>
      <p>這場分享將帶你理解大腦處理資訊的天性，拆解這些困境：</p>
      <ul style="color: #555;">
        <li>為什麼明明學了方法、拿了證照，遇到真實問題卻解決不了？</li>
        <li>為什麼越優化反而越忙？</li>
        <li>如何找回在雜事中被磨損的專注力？</li>
      </ul>
      <p>當你掌握了人類大腦不變的運作規律，才能在科技浪潮中精準運用 AI 輔助，建立一套穩健的模式。</p>
      <img src="${CONTENT_IMAGE_URL}" alt="內容" style="width: 100%; max-width: 600px; border-radius: 8px; margin: 20px 0;">
      <div style="background-color: #fff2f0; padding: 20px; border-radius: 10px; border: 1px solid #ffccc7;">
        <p style="margin-top: 0; color: #ff4d4f;"><strong>請依照下面簡單三步驟就可參加：</strong></p>
        <ol>
          <li><strong>下載會議軟體：</strong>我們使用「Zoom」作為線上遠端教學軟體。</li>
          <li><strong>點擊下方按鈕：</strong>直接進入線上研討會教室。</li>
          <li><strong>準時出席：</strong>活動當天請提前 10 分鐘進入測試設備。</li>
        </ol>
        <p style="text-align: center; margin-bottom: 0; margin-top: 20px;">
          <a href="${link}" style="background-color: #ff4d4f; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: bold; font-size: 18px;">立即進入 Zoom 研討會</a>
        </p>
      </div>
      <p style="margin-top: 30px;">怕沒有跟上的活動嗎？快快加入我們的「風霈行事曆」～</p>
      <p><strong>風霈學院 FENGPEINET 敬上</strong></p>
    </div>
  `;
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody });
}

function sendRegistrationSuccessEmail(to, name, date) {
  const subject = "【風霈學院】恭喜您報名成功！";
  const htmlBody = `
    <div style="font-family: 'Microsoft JhengHei', sans-serif; color: #333; line-height: 1.8; max-width: 600px; margin: 0 auto;">
      <img src="${POSTER_IMAGE_URL}" alt="海報" style="width: 100%; max-width: 600px; border-radius: 8px; margin-bottom: 20px;">
      <h2 style="color: #ff4d4f;">報名成功通知</h2>
      <p>親愛的 <strong>${name}</strong> 同學您好：</p>
      <p>我們已收到您的報名資訊，確認報名場次為：<strong>${date}</strong></p>
      <p>我們將會在說明會<strong>前一天與當天</strong>寄送 Zoom 會議連結給您。</p>
      <img src="${CONTENT_IMAGE_URL}" alt="內容" style="width: 100%; max-width: 600px; border-radius: 8px; margin-top: 20px;">
      <p style="margin-top: 20px;">祝您有美好的一天！<br><strong>風霈學院 FENGPEINET 團隊 敬上</strong></p>
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