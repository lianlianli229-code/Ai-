// ============================================================================
// Google Apps Script - 绑定至表格：AI生图Skill词库表 (v3 多用户 Google 账号身份版)
// 新增功能：
// 1. 支持以 Google 邮箱账号作为个人唯一身份，不同账号的「我的技能炉」完全独立隔离
// 2. 新增【12、用户技能炉档案】子表，按 Google 邮箱独立存储每位用户的技能炉收藏与昵称
// 3. 谁上传的技能只自动收入上传者自己的技能炉，不会进入其他人的技能炉
// ============================================================================

const CATEGORY_MAP = {
  "口播类": "1、口播类",
  "特效类": "2、特效类",
  "圣经故事类": "3、圣经故事类",
  "实拍类": "4、实拍类",
  "卡片类": "5、卡片类",
  "灾难借势类": "6、灾难借势类",
  "剧情类": "7、剧情类",
  "封面图、风景图": "8、封面图、风景图",
  "人设头像": "9、人设头像"
};

const WEB_SKILL_SHEET_NAME = "10、网站技能投稿区";
const HELP_SHEET_NAME = "11、求助互助区";
const USER_FURNACE_SHEET_NAME = "12、用户技能炉档案";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔥 AI技能社区后台")
    .addItem("🚀 一键同步表格数据到前端网站", "syncSheetToFrontend")
    .addItem("🗑️ 前端更新已删除内容", "syncDeletedToFrontend")
    .addSeparator()
    .addItem("📂 初始化/检查【网站投稿、求助与用户档案表】", "initReceivingSheets")
    .addToUi();
}

function normalizeDriveUrl(rawUrl) {
  const url = String(rawUrl || "").trim();
  if (!url) return "";
  const matchD = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (matchD && matchD[1]) return "https://lh3.googleusercontent.com/d/" + matchD[1] + "=w1000";
  const matchId = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (matchId && matchId[1]) return "https://lh3.googleusercontent.com/d/" + matchId[1] + "=w1000";
  return url;
}

function getOrCreateWebSkillSheet(ss) {
  let sheet = ss.getSheetByName(WEB_SKILL_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(WEB_SKILL_SHEET_NAME);
    sheet.appendRow(["状态", "所属分类", "Skill 标题", "Google Drive 图片链接", "完整提示词 (Prompt)", "效果简述", "使用方法/工作流", "适用模型", "上传者 Google 账号", "提交时间"]);
    sheet.getRange(1, 1, 1, 10).setBackground("#292821").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateHelpSheet(ss) {
  let sheet = ss.getSheetByName(HELP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HELP_SHEET_NAME);
    sheet.appendRow(["状态", "求助标题", "所属分类", "求助人 (Google账号)", "参考图 Google Drive 链接", "痛点详细描述", "社区助力回复 (JSON)", "发布时间"]);
    sheet.getRange(1, 1, 1, 8).setBackground("#ea580c").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateUserSheet(ss) {
  let sheet = ss.getSheetByName(USER_FURNACE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USER_FURNACE_SHEET_NAME);
    sheet.appendRow(["Google 邮箱账号", "用户昵称", "个人技能炉收藏列表 (JSON)", "最后同步时间"]);
    sheet.getRange(1, 1, 1, 4).setBackground("#15803d").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 读取所有用户的技能炉映射表：{ "user@gmail.com": { name: "...", savedTitles: [...] } }
function getAllUserFurnaces(ss) {
  const sheet = getOrCreateUserSheet(ss);
  const rows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const email = String(rows[i][0] || "").trim().toLowerCase();
    if (!email) continue;
    let savedTitles = [];
    try {
      savedTitles = JSON.parse(rows[i][2] || "[]");
    } catch (e) {
      savedTitles = [];
    }
    map[email] = {
      name: String(rows[i][1] || email.split("@")[0]),
      savedTitles: Array.isArray(savedTitles) ? savedTitles : []
    };
  }
  return map;
}

// 保存指定 Google 账号的技能炉列表
function saveUserFurnace(ss, email, name, savedTitles) {
  const cleanEmail = String(email || "").trim().toLowerCase();
  if (!cleanEmail) return;
  const sheet = getOrCreateUserSheet(ss);
  const rows = sheet.getDataRange().getValues();
  const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  const jsonStr = JSON.stringify(Array.isArray(savedTitles) ? savedTitles : []);

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim().toLowerCase() === cleanEmail) {
      if (name) sheet.getRange(i + 1, 2).setValue(name);
      sheet.getRange(i + 1, 3).setValue(jsonStr);
      sheet.getRange(i + 1, 4).setValue(nowStr);
      return;
    }
  }
  sheet.appendRow([cleanEmail, name || cleanEmail.split("@")[0], jsonStr, nowStr]);
}

function initReceivingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateWebSkillSheet(ss);
  getOrCreateHelpSheet(ss);
  getOrCreateUserSheet(ss);
  SpreadsheetApp.getUi().alert("✅ 初始化完成！\n\n已创建/检查以下专属工作表：\n1. 【10、网站技能投稿区】\n2. 【11、求助互助区】\n3. 【12、用户技能炉档案】（按 Google 账号独立存储个人技能炉）");
}

function syncSheetToFrontend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let movedCount = 0;

  Object.keys(CATEGORY_MAP).forEach(cat => {
    const sheet = ss.getSheetByName(CATEGORY_MAP[cat]);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const maxScan = Math.min(lastRow, 8);

    for (let r = 3; r <= maxScan; r++) {
      const rowVals = sheet.getRange(r, 1, 1, 8).getValues()[0];
      const title = String(rowVals[1] || "").trim();
      const prompt = String(rowVals[3] || "").trim();
      if (!title && !prompt) continue;

      const cleanImg = normalizeDriveUrl(rowVals[2]);
      const targetRow = Math.max(sheet.getLastRow() + 1, 11);
      sheet.getRange(targetRow, 1, 1, 8).setValues([[
        "已更新",
        title,
        cleanImg || rowVals[2],
        prompt,
        rowVals[4] || "",
        rowVals[5] || cat,
        rowVals[6] || "Midjourney / Flux",
        rowVals[7] || "入门"
      ]]);

      sheet.getRange(r, 1, 1, 8).setValues([["未更新", "", "", "", "", "", "", ""]]);
      movedCount++;
    }

    const newLastRow = sheet.getLastRow();
    if (newLastRow >= 10) {
      for (let r = 10; r <= newLastRow; r++) {
        const t = String(sheet.getRange(r, 2).getValue() || "").trim();
        if (t) sheet.getRange(r, 1).setValue("已更新");
      }
    }
  });

  getOrCreateWebSkillSheet(ss);
  getOrCreateHelpSheet(ss);
  getOrCreateUserSheet(ss);
  SpreadsheetApp.getUi().alert("🚀 同步成功！\n\n本次共将 " + movedCount + " 条【未更新栏】技能正式归档至【已更新栏（线上资源库）】。\n现在刷新前端网页即可立即看到最新数据！");
}

// 按钮2：前端更新已删除内容（清理在分类表或投稿区被删除的技能，并同步清理用户技能炉中的失效收藏）
function syncDeletedToFrontend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const webSheet = getOrCreateWebSkillSheet(ss);
  const userSheet = getOrCreateUserSheet(ss);
  const props = PropertiesService.getScriptProperties();

  let knownWebTitles = [];
  try {
    knownWebTitles = JSON.parse(props.getProperty("KNOWN_WEB_TITLES") || "[]");
  } catch (e) {
    knownWebTitles = [];
  }
  const knownWebSet = {};
  knownWebTitles.forEach(t => { if (t) knownWebSet[String(t).trim()] = true; });

  // 1. 先扫描 10、网站技能投稿区 当前存在的标题（同时清理空白行）
  const currentWebTitles = {};
  const wLast = webSheet.getLastRow();
  let removedFromWeb = 0;
  for (let r = wLast; r >= 2; r--) {
    const t = String(webSheet.getRange(r, 3).getValue() || "").trim();
    const p = String(webSheet.getRange(r, 5).getValue() || "").trim();
    if (!t && !p) {
      webSheet.deleteRow(r);
      removedFromWeb++;
    } else if (t) {
      currentWebTitles[t] = true;
    }
  }

  // 2. 扫描 1~9 分类表：如果某条网站投稿在【10、网站技能投稿区】已被删除，则同步从 1~9 分类表删除
  const currentCatTitles = {};
  let removedFromCat = 0;
  Object.keys(CATEGORY_MAP).forEach(cat => {
    const sheet = ss.getSheetByName(CATEGORY_MAP[cat]);
    if (!sheet) return;
    const cLast = sheet.getLastRow();
    for (let r = cLast; r >= 10; r--) {
      const t = String(sheet.getRange(r, 2).getValue() || "").trim();
      const p = String(sheet.getRange(r, 4).getValue() || "").trim();
      if (r >= 11 && !t && !p) {
        sheet.deleteRow(r);
        removedFromCat++;
        continue;
      }
      if (!t) continue;
      if (cat !== "口播类" && t === "真人口播演播厅微表情" && r === 10) continue;

      // 如果是第11行及以后的新增行，且在【10、网站技能投稿区】中已不存在（被后台删除），则从分类表中同步删除
      if (r >= 11 && (knownWebSet[t] || !currentWebTitles[t])) {
        if (!currentWebTitles[t]) {
          sheet.deleteRow(r);
          removedFromCat++;
          continue;
        }
      }
      currentCatTitles[t] = true;
    }
  });

  // 3. 反过来检查【10、网站技能投稿区】：如果某条技能在 1~9 分类表中已被删除，则同步从【10、网站技能投稿区】删除
  const wLastAfter = webSheet.getLastRow();
  for (let r = wLastAfter; r >= 2; r--) {
    const t = String(webSheet.getRange(r, 3).getValue() || "").trim();
    if (t && !currentCatTitles[t]) {
      webSheet.deleteRow(r);
      delete currentWebTitles[t];
      removedFromWeb++;
    }
  }

  // 更新已记录的网站投稿标题集合
  props.setProperty("KNOWN_WEB_TITLES", JSON.stringify(Object.keys(currentWebTitles)));

  // 4. 同步清理【12、用户技能炉档案】中已被后台删除的技能
  const uLast = userSheet.getLastRow();
  let cleanedUsers = 0;
  for (let r = 2; r <= uLast; r++) {
    let savedArr = [];
    try {
      savedArr = JSON.parse(userSheet.getRange(r, 3).getValue() || "[]");
    } catch (e) {
      savedArr = [];
    }
    if (Array.isArray(savedArr) && savedArr.length > 0) {
      const filtered = savedArr.filter(title => currentCatTitles[String(title).trim()] || currentWebTitles[String(title).trim()]);
      if (filtered.length !== savedArr.length) {
        userSheet.getRange(r, 3).setValue(JSON.stringify(filtered));
        cleanedUsers++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(
    "🗑️ 前端已删除内容更新完成！\n\n" +
    "• 同步清理分类表残留行：" + removedFromCat + " 条\n" +
    "• 同步清理投稿区残留行：" + removedFromWeb + " 条\n" +
    "• 同步移除用户技能炉失效收藏：" + cleanedUsers + " 位用户\n\n" +
    "现在去前端网页刷新，已删除的技能会立即从页面上消失！"
  );
}

function handleAction(ss, body) {
  const action = body.action;
  const userEmail = String(body.userEmail || "").trim().toLowerCase();
  const userName = String(body.userName || "").trim();

  if (action === "addSkill") {
    const cat = body.category || "口播类";
    const imgUrl = normalizeDriveUrl(body.img);
    const tagsStr = Array.isArray(body.tags) ? body.tags.join(", ") : String(body.tags || cat);
    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    const ownerLabel = userEmail ? (userName ? userName + " (" + userEmail + ")" : userEmail) : (body.author || "社区共创");

    const webSheet = getOrCreateWebSkillSheet(ss);
    webSheet.appendRow([
      "已更新",
      cat,
      body.title || "",
      imgUrl,
      body.prompt || "",
      body.desc || "",
      body.workflow || "",
      body.model || "Midjourney / Flux",
      ownerLabel,
      nowStr
    ]);

    const catSheetName = CATEGORY_MAP[cat] || "1、口播类";
    const catSheet = ss.getSheetByName(catSheetName);
    if (catSheet) {
      const targetRow = Math.max(catSheet.getLastRow() + 1, 11);
      catSheet.getRange(targetRow, 1, 1, 8).setValues([[
        "已更新",
        body.title || "",
        imgUrl,
        body.prompt || "",
        body.desc || "",
        tagsStr,
        body.model || "Midjourney / Flux",
        body.difficulty || "入门"
      ]]);
    }

    // 记录网站投稿标题，方便后续双向删除比对
    try {
      const props = PropertiesService.getScriptProperties();
      const kw = JSON.parse(props.getProperty("KNOWN_WEB_TITLES") || "[]");
      const cleanT = String(body.title || "").trim();
      if (cleanT && kw.indexOf(cleanT) === -1) {
        kw.push(cleanT);
        props.setProperty("KNOWN_WEB_TITLES", JSON.stringify(kw));
      }
    } catch (e) {}

    // 如果传了当前用户的 savedTitles，仅更新该上传者自己的技能炉
    if (userEmail && Array.isArray(body.savedTitles)) {
      saveUserFurnace(ss, userEmail, userName, body.savedTitles);
    }
    return { ok: true, img: imgUrl };
  }

  if (action === "addHelp") {
    const helpSheet = getOrCreateHelpSheet(ss);
    const refImg = normalizeDriveUrl(body.refImg);
    const nowStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    helpSheet.appendRow([
      "待助力",
      body.title || "",
      body.category || "口播类",
      body.author || userName || userEmail || "社区伙伴",
      refImg,
      body.desc || "",
      JSON.stringify(body.replies || []),
      nowStr
    ]);
    return { ok: true, rowNumber: helpSheet.getLastRow(), refImg: refImg };
  }

  if (action === "replyHelp") {
    const helpSheet = getOrCreateHelpSheet(ss);
    const hRows = helpSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < hRows.length; i++) {
      if (String(hRows[i][1]).trim() === String(body.title || "").trim()) {
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow < 2 && body.rowNumber >= 2 && body.rowNumber <= hRows.length) {
      targetRow = Number(body.rowNumber);
    }
    if (targetRow >= 2) {
      let replies = [];
      try {
        replies = JSON.parse(helpSheet.getRange(targetRow, 7).getValue() || "[]");
      } catch (err) {
        replies = [];
      }
      if (body.reply) replies.push(body.reply);
      helpSheet.getRange(targetRow, 1).setValue("已助力");
      helpSheet.getRange(targetRow, 7).setValue(JSON.stringify(replies));
    }
    return { ok: true, rowNumber: targetRow };
  }

  // 按 Google 邮箱账号保存该用户个人的技能炉列表
  if (action === "syncSaved") {
    if (userEmail && Array.isArray(body.savedTitles)) {
      saveUserFurnace(ss, userEmail, userName, body.savedTitles);
    }
    return { ok: true };
  }

  return { ok: false, error: "Unknown action: " + action };
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 一键读取当前浏览器 Google 账号（支持 JSON 接口与弹窗自动回传）
  if (e && e.parameter && e.parameter.action === "whoami") {
    const activeEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, email: activeEmail }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (e && e.parameter && e.parameter.action === "authPopup") {
    const activeEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Google 账号身份识别</title>
      <style>body{font-family:sans-serif;background:#f7f5ef;display:flex;align-items:center;justify-content:center;height:90vh;margin:0;text-align:center}
      .box{background:#fff;padding:24px;border-radius:14px;box-shadow:0 8px 24px rgba(0,0,0,.1);max-width:360px;width:90%}
      button{background:#ea580c;color:#fff;border:0;border-radius:30px;padding:10px 20px;font-weight:700;cursor:pointer;margin-top:12px}</style></head>
      <body><div class="box">
      <h3>🔑 Google 账号识别</h3>
      <p id="msg">${activeEmail ? "已识别到当前账号：<br><b>" + activeEmail + "</b>" : "由于浏览器隐私隔离，请在网页登录框直接输入您的 Gmail 邮箱完成绑定。"}</p>
      <button onclick="sendAndClose()">确认并返回网站</button>
      </div><script>
      function sendAndClose(){
        if(window.opener){window.opener.postMessage({type:"GOOGLE_AUTH_EMAIL",email:"${activeEmail}"},"*");}
        window.close();
      }
      if("${activeEmail}"){setTimeout(sendAndClose,600);}
      <\/script></body></html>`;
    return HtmlService.createHtmlOutput(html);
  }

  if (e && e.parameter && e.parameter.payload) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);
    try {
      const body = JSON.parse(e.parameter.payload);
      const res = handleAction(ss, body);
      return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
    } finally {
      lock.releaseLock();
    }
  }

  const data = [];
  const seenTitles = {};
  let idCounter = 100;

  let knownWebSet = {};
  try {
    const kw = JSON.parse(PropertiesService.getScriptProperties().getProperty("KNOWN_WEB_TITLES") || "[]");
    kw.forEach(t => { if (t) knownWebSet[String(t).trim()] = true; });
  } catch (e) {}

  // 先扫描 1~9 分类表当前实际存在的标题集合
  const catTitlesExist = {};
  Object.keys(CATEGORY_MAP).forEach(cat => {
    const sheet = ss.getSheetByName(CATEGORY_MAP[cat]);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 2; i < rows.length; i++) {
      if (i === 7 || i === 8) continue;
      const r = rows[i];
      const title = String(r[1] || "").trim();
      const prompt = String(r[3] || "").trim();
      if (!title && !prompt) continue;
      if (title.indexOf("【") === 0 || title === "Skill 标题") continue;
      if (cat !== "口播类" && title === "真人口播演播厅微表情" && i === 9) continue;
      if (title) catTitlesExist[title] = true;
    }
  });

  // 1. 读取「10、网站技能投稿区」（仅当在分类表中未被删除时展示）
  const webSheet = getOrCreateWebSkillSheet(ss);
  const wRows = webSheet.getDataRange().getValues();
  const webTitlesExist = {};
  for (let i = wRows.length - 1; i >= 1; i--) {
    const r = wRows[i];
    const title = String(r[2] || "").trim();
    const prompt = String(r[4] || "").trim();
    if (!title) continue;
    webTitlesExist[title] = true;
    if (!catTitlesExist[title] || seenTitles[title]) continue;
    seenTitles[title] = true;
    const cat = String(r[1] || "口播类").trim();
    const ownerRaw = String(r[8] || "社区共创").trim();
    const authorDisplay = (ownerRaw === "已收入技能炉" || !ownerRaw) ? "社区共创" : ownerRaw;

    data.push({
      id: idCounter++,
      source: "website",
      status: String(r[0] || "已更新"),
      category: cat,
      title: title,
      img: normalizeDriveUrl(r[3]),
      prompt: prompt,
      desc: String(r[5] || "社区共创精选生图技能包"),
      workflow: String(r[6] || "直接复制 Prompt 并按需替换主体词与比例。"),
      model: String(r[7] || "Midjourney / Flux"),
      author: authorDisplay,
      ownerEmail: ownerRaw.indexOf("@") !== -1 ? ownerRaw : "",
      tags: [cat, "网站共创"]
    });
  }

  // 2. 遍历 1~9 分类表
  Object.keys(CATEGORY_MAP).forEach(cat => {
    const sheetName = CATEGORY_MAP[cat];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    for (let i = 2; i < rows.length; i++) {
      if (i === 7 || i === 8) continue;
      const r = rows[i];
      const title = String(r[1] || "").trim();
      const prompt = String(r[3] || "").trim();
      if (!title && !prompt) continue;
      if (title.indexOf("【") === 0 || title === "Skill 标题") continue;
      if (cat !== "口播类" && title === "真人口播演播厅微表情" && i === 9) continue;
      if (knownWebSet[title] && !webTitlesExist[title]) continue;
      if (seenTitles[title]) continue;
      seenTitles[title] = true;

      const tagsStr = String(r[5] || "").trim();
      const tags = tagsStr ? tagsStr.split(/[,，、\s]+/).filter(Boolean) : [cat];

      data.push({
        id: idCounter++,
        sheet: sheetName,
        rowNumber: i + 1,
        status: String(r[0] || "已更新").trim(),
        title: title,
        category: cat,
        img: normalizeDriveUrl(r[2]),
        prompt: prompt,
        desc: String(r[4] || "").trim(),
        tags: tags,
        model: String(r[6] || "Midjourney / Flux").trim(),
        difficulty: String(r[7] || "入门").trim()
      });
    }
  });

  // 3. 读取「11、求助互助区」
  const helps = [];
  const helpSheet = getOrCreateHelpSheet(ss);
  const hRows = helpSheet.getDataRange().getValues();
  for (let i = hRows.length - 1; i >= 1; i--) {
    const r = hRows[i];
    const title = String(r[1] || "").trim();
    if (!title) continue;
    let replies = [];
    try {
      replies = JSON.parse(r[6] || "[]");
    } catch (err) {
      replies = [];
    }
    helps.push({
      id: 2000 + i,
      rowNumber: i + 1,
      status: String(r[0] || "待助力"),
      title: title,
      category: String(r[2] || "口播类"),
      author: String(r[3] || "社区伙伴"),
      refImg: normalizeDriveUrl(r[4]),
      desc: String(r[5] || ""),
      replies: replies
    });
  }

  // 4. 读取按 Google 邮箱隔离的「12、用户技能炉档案」
  const userFurnaces = getAllUserFurnaces(ss);

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      version: "v3",
      data: data,
      helps: helps,
      userFurnaces: userFurnaces
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = JSON.parse(e.postData.contents || "{}");
    const res = handleAction(ss, body);
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
