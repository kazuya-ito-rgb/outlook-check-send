/* global Office */
var CONFIRMATION_KEY = "send_confirmation_verified";
var CONFIRMATION_TIME_KEY = "send_confirmation_timestamp";
var VALID_MS = 5 * 60 * 1000;   // 有効期限 5分
var TIMEOUT_MS = 3000;          // 読み込み待ちの上限 3秒

// タイムアウト付きで非同期処理を実行（固まり防止）
function withTimeout(fn) {
  return new Promise(function (resolve) {
    var done = false;
    function finish(v) { if (!done) { done = true; resolve(v); } }
    if (typeof setTimeout === "function") {
      setTimeout(function () { finish(false); }, TIMEOUT_MS);
    }
    try { fn(finish); } catch (e) { finish(false); }
  });
}

function isFresh(ts) {
  var t = Number(ts);
  return !!ts && !isNaN(t) && (Date.now() - t) < VALID_MS;
}

// sessionData から確認時刻を読む
function readSessionData(item) {
  return withTimeout(function (finish) {
    if (!item.sessionData) { finish(false); return; }
    item.sessionData.getAsync(CONFIRMATION_TIME_KEY, function (r) {
      finish(r.status === Office.AsyncResultStatus.Succeeded && isFresh(r.value));
    });
  });
}

// customProperties から確認状態を読む（現行タスクペイン互換）
function readCustomProps(item) {
  return withTimeout(function (finish) {
    item.loadCustomPropertiesAsync(function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { finish(false); return; }
      var verified = r.value.get(CONFIRMATION_KEY);
      var ts = r.value.get(CONFIRMATION_TIME_KEY);
      finish((verified === true || verified === "true") && isFresh(ts));
    });
  });
}

function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox.item;
    if (!item) { event.completed({ allowEvent: true }); return; }

    Promise.all([readSessionData(item), readCustomProps(item)]).then(function (res) {
      if (res[0] || res[1]) {
        event.completed({ allowEvent: true });
      } else {
        event.completed({
          allowEvent: false,
          errorMessage: "【誤送信防止：送信前確認が完了していません】\n" +
            "リボンの「送信前確認画面」を開き、5項目を確認してチェックを入れ、" +
            "「確認完了 (送信許可)」を押してから再度送信してください。（有効期限5分）",
          cancelLabel: "メール作成に戻る"
        });
      }
    }, function () {
      event.completed({ allowEvent: false, errorMessage: "送信確認の検証中にエラーが発生しました。" });
    });
  } catch (e) {
    event.completed({ allowEvent: false, errorMessage: "送信確認処理中に例外が発生しました。" });
  }
}

// if で包まない（登録に失敗したらエラーとして見えるようにする）
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
