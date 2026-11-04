/* global Office */
/*
 * 送信時チェック v6 (OnMessageSend / Smart Alerts)
 * - 何があっても 2.5 秒以内に必ず event.completed を返す（固まり防止）
 * - 確認状態は sessionData（メイン）→ customProperties（予備）の順に読む
 */
var CONFIRMATION_KEY = "send_confirmation_verified";
var CONFIRMATION_TIME_KEY = "send_confirmation_timestamp";
var VALID_MS = 5 * 60 * 1000;     // 確認の有効期限 5分
var WATCHDOG_MS = 2500;           // 全体の待ち上限
var VER = "v6";

function isFresh(ts) {
  var t = Number(ts);
  return !!ts && !isNaN(t) && (Date.now() - t) < VALID_MS;
}

function onMessageSendHandler(event) {
  var finished = false;
  function finish(allow, msg) {
    if (finished) return;
    finished = true;
    try {
      if (allow) {
        event.completed({ allowEvent: true });
      } else {
        event.completed({ allowEvent: false, errorMessage: msg, cancelLabel: "メール作成に戻る" });
      }
    } catch (e) {
      try { event.completed({ allowEvent: allow }); } catch (e2) {}
    }
  }

  var MSG_NOT_CONFIRMED =
    "【誤送信防止】送信前確認が完了していません。\n" +
    "リボンの「送信前確認画面」で5項目にチェックし、「確認完了 (送信許可)」を押してから再度送信してください。（有効期限5分）";

  // 見張り役：時間内に判定できなければ止めて理由を出す
  try {
    setTimeout(function () {
      finish(false, "【誤送信防止】確認状態を読み取れませんでした（" + VER + " timeout）。\n" +
        "送信前確認画面で「確認完了」をもう一度押してから送信してください。");
    }, WATCHDOG_MS);
  } catch (e) {}

  try {
    var item = Office.context.mailbox.item;
    if (!item) { finish(true); return; }

    // 1) sessionData（タスクペインと送信時チェックで共有される）
    var trySessionData = function (next) {
      try {
        if (!item.sessionData) { next(); return; }
        item.sessionData.getAsync(CONFIRMATION_TIME_KEY, function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded && isFresh(r.value)) {
            finish(true);
          } else {
            next();
          }
        });
      } catch (e) { next(); }
    };

    // 2) customProperties（予備）
    var tryCustomProps = function () {
      try {
        item.loadCustomPropertiesAsync(function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            var v = r.value.get(CONFIRMATION_KEY);
            var ts = r.value.get(CONFIRMATION_TIME_KEY);
            if ((v === true || v === "true") && isFresh(ts)) { finish(true); return; }
          }
          finish(false, MSG_NOT_CONFIRMED);
        });
      } catch (e) {
        finish(false, MSG_NOT_CONFIRMED);
      }
    };

    trySessionData(tryCustomProps);
  } catch (e) {
    finish(false, "送信前確認の処理中にエラーが発生しました（" + VER + "）。");
  }
}

// 登録（トップレベルで実行。Office.onReady でも念のため再登録）
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
if (Office.onReady) {
  Office.onReady(function () {
    try { Office.actions.associate("onMessageSendHandler", onMessageSendHandler); } catch (e) {}
  });
}
