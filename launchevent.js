/* global Office */

// 確認済みフラグは item.sessionData（要件セット1.14）に「確認時刻(ms)」として保存する
var CONFIRMATION_KEY = "send_confirmation_timestamp";
var VALID_MS = 5 * 60 * 1000;      // 確認の有効期限: 5分
var CHECK_TIMEOUT_MS = 4000;       // 確認処理の打ち切り時間（応答がない場合に必ず結果を返す）

var BLOCK_MESSAGE =
    "【誤送信防止：送信前確認が完了していません】\n\n" +
    "リボンの「送信前確認画面」を開き、5項目にチェックして「確認完了 (送信許可)」を押してから、再度「送信」してください。";

/**
 * 確認時刻を読み出す。取得できなければ null を返す。
 */
function readConfirmTime(item, cb) {
    try {
        if (item && item.sessionData && item.sessionData.getAsync) {
            item.sessionData.getAsync(CONFIRMATION_KEY, function (r) {
                if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
                    cb(Number(r.value));
                } else {
                    cb(null);
                }
            });
        } else {
            cb(null);
        }
    } catch (e) {
        cb(null);
    }
}

/**
 * OnMessageSend イベントハンドラー（Smart Alerts）
 * どの経路でも event.completed を一度だけ必ず呼ぶ。
 */
function onMessageSendHandler(event) {
    var finished = false;
    var timer = null;

    function finish(options) {
        if (finished) { return; }
        finished = true;
        if (timer) { clearTimeout(timer); }
        event.completed(options);
    }

    // 応答が返らない場合でも「処理中」のまま止まらないようにする
    timer = setTimeout(function () {
        finish({
            allowEvent: false,
            errorMessage: "送信前確認の状態を取得できませんでした。もう一度「送信」を押してください。繰り返す場合は「送信前確認画面」で再度「確認完了」を押してください。",
            cancelLabel: "メール作成に戻る"
        });
    }, CHECK_TIMEOUT_MS);

    try {
        var item = Office.context.mailbox.item;
        if (!item) {
            finish({ allowEvent: false, errorMessage: BLOCK_MESSAGE, cancelLabel: "メール作成に戻る" });
            return;
        }
        readConfirmTime(item, function (ts) {
            if (ts && (Date.now() - ts) >= 0 && (Date.now() - ts) < VALID_MS) {
                finish({ allowEvent: true });
            } else {
                finish({ allowEvent: false, errorMessage: BLOCK_MESSAGE, cancelLabel: "メール作成に戻る" });
            }
        });
    } catch (err) {
        finish({ allowEvent: false, errorMessage: "送信確認処理中にエラーが発生しました。もう一度お試しください。" });
    }
}

// 【最重要】Office.actions.associate は Office.onReady の外（トップレベル）で呼び出す
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
if (typeof window !== "undefined") {
    window.onMessageSendHandler = onMessageSendHandler;
}
