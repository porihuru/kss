(function () {
  "use strict";

  var BUILD = "Schema2 / 20260904-5";
  var PREFIX = "STI-AES-CBC-SHA256-v1:";
  var IV_LENGTH = 16;
  var state = { rows: [], logs: [], itemCount: 0 };

  function byId(id) { return document.getElementById(id); }
  function has(obj, key) { return Object.prototype.hasOwnProperty.call(obj, key); }
  function trimText(value) { return String(value === undefined || value === null ? "" : value).replace(/^\s+|\s+$/g, ""); }
  function trimSlash(value) { return String(value || "").replace(/\/+$/, ""); }
  function currentOrigin() { return window.location.protocol + "//" + window.location.host; }
  function itemId(item) { return item.Id !== undefined ? item.Id : item.id; }
  function itemSchema(item) { return item.SchemaVersion !== undefined ? item.SchemaVersion : item.schemaVersion; }
  function itemPayload(item) { return item.EncryptedPayload !== undefined ? item.EncryptedPayload : item.encryptedPayload; }

  function setStatus(message, kind) {
    var node = byId("status");
    node.className = "status" + (kind ? " " + kind : "");
    node.textContent = message;
  }

  function setBusy(busy) {
    byId("runButton").disabled = busy;
    byId("clearButton").disabled = busy;
  }

  function resolveNextUrl(next, site) {
    if (next.indexOf("http://") === 0 || next.indexOf("https://") === 0) { return next; }
    if (next.indexOf("/") === 0) { return currentOrigin() + next; }
    return site + "/" + next;
  }

  function requestJson(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) { return; }
      if (xhr.status >= 200 && xhr.status < 300) {
        try { callback(null, JSON.parse(xhr.responseText)); }
        catch (e) { callback(new Error("SharePointの応答をJSONとして解釈できません。")); }
      } else {
        callback(new Error("SharePoint取得に失敗しました (HTTP " + xhr.status + ")"));
      }
    };
    xhr.onerror = function () { callback(new Error("SharePointへの接続に失敗しました。")); };
    xhr.send(null);
  }

  function requestAllPages(firstUrl, result, site, callback) {
    requestJson(firstUrl, function (error, data) {
      var body, page, next, i;
      if (error) { callback(error); return; }
      body = data.d || data;
      page = body.results || body.value || [];
      next = body.__next || body["@odata.nextLink"] || null;
      for (i = 0; i < page.length; i += 1) { result.push(page[i]); }
      if (next) { requestAllPages(resolveNextUrl(next, site), result, site, callback); }
      else { callback(null); }
    });
  }

  function getSubtle() {
    var provider = window.crypto && window.crypto.subtle ? window.crypto : window.msCrypto;
    return provider && provider.subtle ? provider.subtle : null;
  }

  function finishCrypto(operation, success, failure) {
    if (!operation) { failure(new Error("Web Crypto APIを利用できません。")); return; }
    if (typeof operation.then === "function") {
      operation.then(success, failure);
    } else {
      operation.oncomplete = function (event) {
        success(event.target && event.target.result !== undefined ? event.target.result : event.result);
      };
      operation.onerror = function () { failure(new Error("暗号処理に失敗しました。")); };
      operation.onabort = function () { failure(new Error("暗号処理が中断されました。")); };
    }
  }

  function utf8ToArrayBuffer(value) {
    var encoded = encodeURIComponent(String(value));
    var bytes = [];
    var i, code;
    for (i = 0; i < encoded.length; i += 1) {
      if (encoded.charAt(i) === "%") {
        code = parseInt(encoded.substr(i + 1, 2), 16);
        bytes.push(code);
        i += 2;
      } else {
        bytes.push(encoded.charCodeAt(i));
      }
    }
    return new Uint8Array(bytes).buffer;
  }

  function decodeUtf8(buffer) {
    var bytes = new Uint8Array(buffer);
    var encoded = "";
    var i, hex;
    for (i = 0; i < bytes.length; i += 1) {
      hex = bytes[i].toString(16);
      encoded += "%" + (hex.length === 1 ? "0" : "") + hex;
    }
    return decodeURIComponent(encoded);
  }

  function base64ToArrayBuffer(value) {
    var binary = window.atob(String(value || "").replace(/\s/g, ""));
    var bytes = new Uint8Array(binary.length);
    var i;
    for (i = 0; i < binary.length; i += 1) { bytes[i] = binary.charCodeAt(i); }
    return bytes.buffer;
  }

  function copyBytes(bytes, start, end) {
    var result = new Uint8Array(end - start);
    var i;
    for (i = start; i < end; i += 1) { result[i - start] = bytes[i]; }
    return result;
  }

  function htmlToText(value) {
    var raw = String(value === undefined || value === null ? "" : value);
    var holder;
    raw = raw.replace(/^\uFEFF/, "");
    raw = raw.replace(/[\u200B\u200C\u200D\u2060]/g, "");
    raw = raw.replace(/[\u2010\u2011\u2012\u2013\u2014\u2212\uFF0D]/g, "-");
    raw = raw.replace(/\uFF1A/g, ":");
    try {
      holder = document.createElement("div");
      holder.innerHTML = raw;
      raw = holder.textContent !== undefined ? holder.textContent : holder.innerText;
    } catch (e) {
      raw = raw.replace(/<[^>]*>/g, " ");
      raw = raw.replace(/&nbsp;/gi, " ");
      raw = raw.replace(/&amp;/gi, "&");
    }
    return String(raw || "");
  }

  function payloadPreview(value) {
    var text = htmlToText(value).replace(/[\r\n\t]/g, " ").replace(/\s+/g, " ");
    if (text.length > 80) { text = text.substr(0, 80) + "..."; }
    return text || "(空)";
  }

  function extractEncodedPayload(payload) {
    var text = htmlToText(payload);
    var compact = text.replace(/[\s\u00A0\u3000]+/g, "");
    var index = compact.indexOf(PREFIX);
    var tail, match, encoded;
    if (index < 0) {
      throw new Error("AES-v1プレフィックスを検出できません。受信値先頭: " + payloadPreview(payload));
    }
    tail = compact.substr(index + PREFIX.length);
    match = /^([A-Za-z0-9+\/=]+)/.exec(tail);
    if (!match || !match[1]) {
      throw new Error("EncryptedPayloadのBase64部分を検出できません。受信値先頭: " + payloadPreview(payload));
    }
    encoded = match[1];
    if (encoded.length % 4 !== 0) {
      throw new Error("EncryptedPayloadのBase64長が不正です。長さ=" + encoded.length);
    }
    return encoded;
  }

  function deriveDecryptionKey(callback) {
    var subtle = getSubtle();
    var config = window.STI_DECRYPTION_CONFIG;
    var digestOperation, importOperation;
    if (!subtle) { callback(new Error("Web Crypto APIを利用できません。")); return; }
    if (!config || config.secretKey === undefined) {
      callback(new Error("STI_DECRYPTION_CONFIG.secretKeyを設定してください。"));
      return;
    }
    try {
      digestOperation = subtle.digest({ name: "SHA-256" }, utf8ToArrayBuffer(String(config.secretKey)));
      finishCrypto(digestOperation, function (digest) {
        try {
          importOperation = subtle.importKey("raw", digest, { name: "AES-CBC" }, false, ["decrypt"]);
          finishCrypto(importOperation, function (key) { callback(null, key); }, function () {
            callback(new Error("AES鍵の読み込みに失敗しました。"));
          });
        } catch (e) { callback(new Error("AES鍵の読み込みに失敗しました。")); }
      }, function () { callback(new Error("共有秘密のSHA-256ハッシュ生成に失敗しました。")); });
    } catch (e2) { callback(new Error("AES鍵の生成に失敗しました。")); }
  }

  function decryptPayload(key, payload, callback) {
    var encoded, bytes, iv, ciphertext, operation;
    try {
      encoded = extractEncodedPayload(payload);
      bytes = new Uint8Array(base64ToArrayBuffer(encoded));
      if (bytes.length <= IV_LENGTH || (bytes.length - IV_LENGTH) % 16 !== 0) {
        callback(new Error("EncryptedPayloadのデータ長が不正です。"));
        return;
      }
      iv = copyBytes(bytes, 0, IV_LENGTH);
      ciphertext = copyBytes(bytes, IV_LENGTH, bytes.length);
      operation = getSubtle().decrypt({ name: "AES-CBC", iv: iv }, key, ciphertext.buffer);
      finishCrypto(operation, function (plainBuffer) {
        try { callback(null, JSON.parse(decodeUtf8(plainBuffer))); }
        catch (e) { callback(new Error("復号後JSONの解析に失敗しました。")); }
      }, function () { callback(new Error("AES-CBC復号に失敗しました。")); });
    } catch (e2) {
      callback(e2 && e2.message ? e2 : new Error("EncryptedPayloadを復号できません。"));
    }
  }

  function formatJst(isoText) {
    var date = new Date(isoText);
    var jst, y, m, d, hh, mm, ss;
    function pad(value) { return value < 10 ? "0" + value : String(value); }
    if (isNaN(date.getTime())) { return ""; }
    jst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
    y = jst.getUTCFullYear();
    m = pad(jst.getUTCMonth() + 1);
    d = pad(jst.getUTCDate());
    hh = pad(jst.getUTCHours());
    mm = pad(jst.getUTCMinutes());
    ss = pad(jst.getUTCSeconds());
    return y + "/" + m + "/" + d + " " + hh + ":" + mm + ":" + ss;
  }

  function addLog(logs, level, id, message) {
    logs.push({ level: level, itemId: id, message: message });
  }

  function analyzeItems(items, key, callback) {
    var logs = [], seenItems = {}, userRows = {};
    var index = 0, skipped = 0, legacyCount = 0, validCount = 0;

    function next() {
      var item, id, schema, payload;
      if (index >= items.length) { finish(); return; }
      item = items[index];
      index += 1;
      id = itemId(item);

      if (id === undefined || id === null || id === "" || has(seenItems, String(id))) {
        skipped += 1;
        addLog(logs, "skip", id, "SharePoint項目IDが空、または重複しています。");
        next();
        return;
      }
      seenItems[String(id)] = true;

      schema = String(itemSchema(item));
      if (schema !== "2") {
        legacyCount += 1;
        next();
        return;
      }

      payload = itemPayload(item);
      if (payload === undefined || payload === null || trimText(payload) === "") {
        skipped += 1;
        addLog(logs, "skip", id, "EncryptedPayloadが空です。");
        next();
        return;
      }

      decryptPayload(key, payload, function (error, parsed) {
        var uid, count, date, normalized, row, existing;
        if (error) {
          skipped += 1;
          addLog(logs, "skip", id, error.message);
          next();
          return;
        }
        if (!parsed || parsed.eventType !== "question_start_summary") {
          skipped += 1;
          addLog(logs, "skip", id, "eventTypeがquestion_start_summaryではありません。");
          next();
          return;
        }

        uid = trimText(parsed.userId);
        count = parsed.questionStartCount;
        date = new Date(parsed.lastQuestionStartUtc);
        if (!uid) {
          skipped += 1;
          addLog(logs, "skip", id, "userIdがありません。");
          next();
          return;
        }
        if (typeof count !== "number" || !isFinite(count) || count < 0 || Math.floor(count) !== count) {
          skipped += 1;
          addLog(logs, "skip", id, "questionStartCountが0以上の整数ではありません。");
          next();
          return;
        }
        if (isNaN(date.getTime())) {
          skipped += 1;
          addLog(logs, "skip", id, "lastQuestionStartUtcが不正です。");
          next();
          return;
        }

        normalized = date.toISOString();
        row = {
          itemId: id,
          userId: uid,
          userName: "未照合",
          questionStartCount: count,
          lastQuestionStartUtc: normalized,
          lastQuestionStartJst: formatJst(normalized),
          lastTime: date.getTime()
        };

        if (has(userRows, uid)) {
          existing = userRows[uid];
          skipped += 1;
          addLog(logs, "warn", id, "同一userIdのSchemaVersion=2項目が複数あります。最終開始日時が新しい項目を採用します。");
          if (row.lastTime > existing.lastTime) { userRows[uid] = row; }
        } else {
          userRows[uid] = row;
          validCount += 1;
        }
        next();
      });
    }

    function finish() {
      var rows = [], total = 0, uid, row;
      for (uid in userRows) {
        if (has(userRows, uid)) {
          row = userRows[uid];
          rows.push(row);
          total += row.questionStartCount;
        }
      }
      rows.sort(function (left, right) {
        var ln = parseInt(left.userId, 10), rn = parseInt(right.userId, 10);
        if (!isNaN(ln) && !isNaN(rn)) { return ln - rn; }
        return String(left.userId).localeCompare(String(right.userId));
      });
      callback({
        results: rows,
        logs: logs,
        stats: {
          totalItems: items.length,
          userCount: rows.length,
          totalQuestionStarts: total,
          skippedCount: skipped,
          legacyCount: legacyCount,
          validCount: validCount
        }
      });
    }

    next();
  }

  function resolveUserNames(site, response, callback) {
    var rows = response.results || [];
    var index = 0;

    function next() {
      var row, uid, url;
      if (index >= rows.length) { callback(response); return; }
      row = rows[index];
      index += 1;
      uid = trimText(row.userId);
      if (!/^\d+$/.test(uid)) {
        addLog(response.logs, "warn", row.itemId, "userIdが数値ではないため氏名照合をスキップしました: " + uid);
        next();
        return;
      }

      url = site + "/_api/web/siteusers/getbyid(" + parseInt(uid, 10) + ")?$select=Id,Title,LoginName,Email";
      requestJson(url, function (error, data) {
        var user, name;
        if (error) {
          addLog(response.logs, "warn", row.itemId, "SharePointユーザーID " + uid + " の氏名取得に失敗しました。");
          next();
          return;
        }
        user = data && data.d ? data.d : data;
        name = user ? (user.Title || user.LoginName || user.Email || "") : "";
        if (name) { row.userName = String(name); }
        else { addLog(response.logs, "warn", row.itemId, "SharePointユーザーID " + uid + " の氏名が空です。"); }
        next();
      });
    }

    next();
  }

  function formatLogs(logs) {
    var lines = ["Analyzer " + BUILD], i, log;
    for (i = 0; i < logs.length; i += 1) {
      log = logs[i];
      lines.push("[" + (log.level || "info") + "] " +
        (log.itemId !== undefined && log.itemId !== null ? "Item " + log.itemId + ": " : "") + log.message);
    }
    if (logs.length === 0) { lines.push("エラー・スキップはありません。"); }
    return lines.join("\n");
  }

  function clearResults() {
    state = { rows: [], logs: [], itemCount: 0 };
    byId("resultBody").innerHTML = '<tr class="empty-row"><td colspan="4">まだ集計結果がありません。</td></tr>';
    byId("logOutput").textContent = "Analyzer " + BUILD + "\nログはここに表示されます。";
    byId("logSummary").textContent = "0件";
    byId("itemCount").textContent = "—";
    byId("userCount").textContent = "—";
    byId("startCount").textContent = "—";
    byId("errorCount").textContent = "—";
    byId("csvButton").disabled = true;
    byId("jsonButton").disabled = true;
  }

  function renderResults(response) {
    var rows = response.results || [];
    var logs = response.logs || [];
    var stats = response.stats || {};
    var body = byId("resultBody"), i, tr, td;

    state.rows = rows;
    state.logs = logs;
    state.itemCount = stats.totalItems || 0;
    body.innerHTML = "";

    if (!rows.length) {
      body.innerHTML = '<tr class="empty-row"><td colspan="4">SchemaVersion=2 の有効な利用記録がありません。</td></tr>';
    }

    for (i = 0; i < rows.length; i += 1) {
      tr = document.createElement("tr");
      td = document.createElement("td"); td.textContent = rows[i].userId; tr.appendChild(td);
      td = document.createElement("td"); td.textContent = rows[i].userName || "未照合"; tr.appendChild(td);
      td = document.createElement("td"); td.textContent = rows[i].questionStartCount; tr.appendChild(td);
      td = document.createElement("td"); td.textContent = rows[i].lastQuestionStartJst; tr.appendChild(td);
      body.appendChild(tr);
    }

    byId("itemCount").textContent = String(stats.totalItems || 0);
    byId("userCount").textContent = String(rows.length);
    byId("startCount").textContent = String(stats.totalQuestionStarts || 0);
    byId("errorCount").textContent = String((stats.skippedCount || 0) + (stats.legacyCount || 0));
    byId("logSummary").textContent = String(logs.length) + "件";
    byId("logOutput").textContent = formatLogs(logs);
    byId("csvButton").disabled = !rows.length;
    byId("jsonButton").disabled = !rows.length;
  }

  function run() {
    var site = trimSlash(byId("siteUrl").value.replace(/\s+/g, ""));
    var rawItems = [];
    var itemsUrl;
    if (!site) { setStatus("SharePointサイトURLを入力してください。", "error"); return; }

    clearResults();
    setBusy(true);
    setStatus("stiuseraccess を取得中…", "busy");
    itemsUrl = site + "/_api/web/lists/getbytitle('stiuseraccess')/items?$select=Id,Title,SchemaVersion,EncryptedPayload&$top=5000";

    requestAllPages(itemsUrl, rawItems, site, function (itemError) {
      if (itemError) { setBusy(false); setStatus(itemError.message, "error"); return; }
      setStatus("SchemaVersion=2 を復号中…", "busy");
      deriveDecryptionKey(function (keyError, key) {
        if (keyError) { setBusy(false); setStatus(keyError.message, "error"); return; }
        analyzeItems(rawItems, key, function (response) {
          setStatus("SharePointユーザーIDから氏名を照合中…", "busy");
          resolveUserNames(site, response, function (namedResponse) {
            var stats = namedResponse.stats || {};
            setBusy(false);
            renderResults(namedResponse);
            setStatus(
              "集計完了: 利用者 " + (stats.userCount || 0) + "人 / 問題開始 " +
              (stats.totalQuestionStarts || 0) + "回 / 旧形式除外 " + (stats.legacyCount || 0) + "件",
              "success"
            );
          });
        });
      });
    });
  }

  function csvEscape(value) {
    return '"' + String(value === null || value === undefined ? "" : value).replace(/"/g, '""') + '"';
  }

  function download(name, content, type) {
    var blob = new Blob([content], { type: type });
    var link;
    if (window.navigator && window.navigator.msSaveBlob) {
      window.navigator.msSaveBlob(blob, name);
      return;
    }
    link = document.createElement("a");
    link.href = window.URL.createObjectURL(blob);
    link.download = name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(function () { window.URL.revokeObjectURL(link.href); }, 1000);
  }

  function exportCsv() {
    var text = "\ufeffuserId,userName,questionStartCount,lastQuestionStartUtc,lastQuestionStartJst\r\n";
    var i, row;
    for (i = 0; i < state.rows.length; i += 1) {
      row = state.rows[i];
      text += [
        csvEscape(row.userId),
        csvEscape(row.userName),
        csvEscape(row.questionStartCount),
        csvEscape(row.lastQuestionStartUtc),
        csvEscape(row.lastQuestionStartJst)
      ].join(",") + "\r\n";
    }
    download("stiuseraccess-summary.csv", text, "text/csv;charset=utf-8");
  }

  function exportJson() {
    download("stiuseraccess-summary.json", JSON.stringify(state.rows, null, 2), "application/json;charset=utf-8");
  }

  byId("runButton").onclick = run;
  byId("clearButton").onclick = function () { clearResults(); setStatus("待機中"); };
  byId("csvButton").onclick = exportCsv;
  byId("jsonButton").onclick = exportJson;
  clearResults();
}());
