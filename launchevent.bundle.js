// Smart Alerts: 送信前確認ダイアログ
// Office.MailboxEnums は使わず、シンプルに即座に完了させる

function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "本当にこのメールを送信しますか？\n宛先・件名・添付ファイルをご確認ください。\n問題なければ [そのまま送信する] を押してください。",
    cancelLabel: "送信しない"
  });
}

// Officeランタイムへの関数登録
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function() {
    if (Office.actions) {
      Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    }
    if (typeof window !== "undefined") {
      window["onMessageSendHandler"] = onMessageSendHandler;
    }
  });
} else {
  if (typeof window !== "undefined") {
    window["onMessageSendHandler"] = onMessageSendHandler;
  }
}
