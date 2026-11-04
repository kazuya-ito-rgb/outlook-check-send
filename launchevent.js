/* global Office */

var CONFIRMATION_KEY = "send_confirmation_verified";
var CONFIRMATION_TIME_KEY = "send_confirmation_timestamp";

/**
 * 送信確認が完了しているか判定（有効期限: 5分以内）
 */
function isSendConfirmed(item) {
    return new Promise(function(resolve) {
        try {
            item.loadCustomPropertiesAsync(function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    var customProps = result.value;
                    var verified = customProps.get(CONFIRMATION_KEY);
                    var timestamp = customProps.get(CONFIRMATION_TIME_KEY);
                    if (verified === true && timestamp) {
                        var elapsed = Date.now() - Number(timestamp);
                        if (elapsed < 5 * 60 * 1000) {
                            resolve(true);
                            return;
                        }
                    }
                    resolve(false);
                } else {
                    resolve(false);
                }
            });
        } catch (e) {
            resolve(false);
        }
    });
}

/**
 * OnMessageSend イベントハンドラー（Smart Alerts）
 */
function onMessageSendHandler(event) {
    try {
        var item = Office.context.mailbox.item;
        if (!item) {
            event.completed({ allowEvent: true });
            return;
        }

        isSendConfirmed(item).then(function(confirmed) {
            if (confirmed) {
                // 確認完了済み -> 送信許可
                event.completed({ allowEvent: true });
            } else {
                // 未確認 -> 送信をブロックし、確認を要求
                var msg = "【誤送信防止：送信前確認が完了していません】\n\n" +
                          "メール作成画面の上部リボンにある「送信前確認画面」を開き、宛先・社外アドレス・添付ファイル・件名・本文の5項目を確認してチェックを入れてください。\n\n" +
                          "「確認完了 (送信許可)」ボタンを押した後、再度「送信」を行ってください。";

                event.completed({
                    allowEvent: false,
                    errorMessage: msg,
                    cancelLabel: "メール作成に戻る"
                });
            }
        }).catch(function(err) {
            event.completed({
                allowEvent: false,
                errorMessage: "送信確認の検証中にエラーが発生しました。"
            });
        });
    } catch (err) {
        event.completed({
            allowEvent: false,
            errorMessage: "送信確認処理中に例外が発生しました。"
        });
    }
}

// 【最重要】Office.actions.associate は Office.onReady の外（トップレベル）で呼び出す
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}

// グローバルスコープにも定義
if (typeof window !== "undefined") {
    window.onMessageSendHandler = onMessageSendHandler;
}
