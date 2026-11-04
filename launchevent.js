/* global Office */

var CODE_VERSION = "v3-diag";
var CONFIRMATION_KEY = "send_confirmation_timestamp";
var VALID_MS = 5 * 60 * 1000;      // 確認の有効期限: 5分
var CHECK_TIMEOUT_MS = 3000;       // 応答がなくても、この時間で必ず結果を返す

var BLOCK_MESSAGE =
    "【誤送信防止：送信前確認が完了していません】\n" +
    "リボンの「送信前確認画面」で5項目にチェックし、「確認完了 (送信許可)」を押してから、再度「送信」してください。";

function onMessageSendHandler(event) {
    var finished = false;
    var timer = null;
    var diag = { item: "?", sd: "?", read: "?" };

    function tag() {
        return "\n[" + CODE_VERSION + "] item=" + diag.item + " sessionData=" + diag.sd + " read=" + diag.read;
    }
    function finish(allow, msg) {
        if (finished) { return; }
        finished = true;
        if (timer) { clearTimeout(timer); }
        if (allow) {
            event.completed({ allowEvent: true });
        } else {
            event.completed({ allowEvent: false, errorMessage: msg + tag(), cancelLabel: "メール作成に戻る" });
        }
    }

    timer = setTimeout(function () {
        if (diag.read === "?") { diag.read = "timeout"; }
        finish(false, "送信前確認の状態を取得できませんでした。もう一度「送信」を押してください。");
    }, CHECK_TIMEOUT_MS);

    try {
        var item = Office.context.mailbox.item;
        diag.item = item ? "yes" : "no";
        if (!item) { finish(false, BLOCK_MESSAGE); return; }

        diag.sd = (item.sessionData && item.sessionData.getAsync) ? "yes" : "no";
        if (diag.sd !== "yes") { diag.read = "n/a"; finish(false, BLOCK_MESSAGE); return; }

        item.sessionData.getAsync(CONFIRMATION_KEY, function (r) {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
                diag.read = "error:" + ((r.error && r.error.message) || "unknown");
                finish(false, BLOCK_MESSAGE);
                return;
            }
            var ts = Number(r.value);
            if (!r.value) { diag.read = "empty"; finish(false, BLOCK_MESSAGE); return; }
            var age = Date.now() - ts;
            diag.read = "ts_age_sec=" + Math.round(age / 1000);
            if (age >= 0 && age < VALID_MS) { finish(true); }
            else { finish(false, BLOCK_MESSAGE); }
        });
    } catch (err) {
        diag.read = "exception:" + (err && err.message);
        finish(false, "送信確認処理中にエラーが発生しました。");
    }
}

// Office.actions.associate は Office.onReady の外（トップレベル）で呼び出す
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
if (typeof window !== "undefined") {
    window.onMessageSendHandler = onMessageSendHandler;
}
