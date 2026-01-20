/* global document, Office, Word */

let pollTimer = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    const btn = document.getElementById("btnFormatSelection");
    const status = document.getElementById("format-status");

    // クリック/Enter/Space で実行（ボタンっぽく）
    btn.addEventListener("click", () => onPressFormat());
    btn.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        onPressFormat();
      }
    });

    // 選択がある時だけ有効化するため、定期的に選択を監視
    startSelectionPolling();

    // 初期表示
    setEnabled(false, "本文で文字を選択すると、ボタンが有効になります。");

    function setEnabled(enabled, message) {
      btn.setAttribute("aria-disabled", enabled ? "false" : "true");
      btn.style.pointerEvents = enabled ? "auto" : "none";
      if (status) status.textContent = message || "";
    }

    async function startSelectionPolling() {
      if (pollTimer) clearInterval(pollTimer);

      pollTimer = setInterval(async () => {
        try {
          const hasSelection = await selectionHasText();
          if (hasSelection) {
            setEnabled(true, "準備OK：青いボタンを押すと整形します。");
          } else {
            setEnabled(false, "本文で整形したい文字を選択してください。");
          }
        } catch {
          // たまに選択取得に失敗しても無視（UIが壊れない方が大事）
        }
      }, 600);
    }

    async function selectionHasText() {
      return Word.run(async (context) => {
        const r = context.document.getSelection();
        r.load("text");
        await context.sync();
        return (r.text || "").trim().length > 0;
      });
    }

    async function onPressFormat() {
      // ボタンが無効状態なら何もしない
      if (btn.getAttribute("aria-disabled") === "true") return;

      setEnabled(true, "整形中…");

      try {
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          range.load("text");
          await context.sync();

          const original = range.text || "";
          if (!original.trim()) {
            setEnabled(false, "本文で整形したい文字を選択してください。");
            return;
          }

          const formatted = formatJapaneseText(original);
          range.insertText(formatted, Word.InsertLocation.replace);

          await context.sync();
        });

        setEnabled(true, "整形が完了しました。");
      } catch (e) {
        console.error(e);
        setEnabled(true, "エラーが発生しました。もう一度お試しください。");
        alert("整形に失敗しました: " + (e?.message ?? e));
      }
    }
  }
});

function formatJapaneseText(input) {
  let s = toHalfwidthAsciiAndSpace(input);
  s = s.replace(/\uFF70/g, "ー"); // 半角長音 → 全角
  s = halfwidthKatakanaToFullwidth(s);
  return s;
}

function toHalfwidthAsciiAndSpace(str) {
  let out = "";
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);

    if (code === 0x3000) {
      out += " ";
      continue;
    }

    if (code >= 0xff01 && code <= 0xff5e) {
      out += String.fromCharCode(code - 0xfee0);
      continue;
    }

    out += str[i];
  }
  return out;
}

function halfwidthKatakanaToFullwidth(str) {
  const baseMap = {
    "ｱ":"ア","ｲ":"イ","ｳ":"ウ","ｴ":"エ","ｵ":"オ",
    "ｶ":"カ","ｷ":"キ","ｸ":"ク","ｹ":"ケ","ｺ":"コ",
    "ｻ":"サ","ｼ":"シ","ｽ":"ス","ｾ":"セ","ｿ":"ソ",
    "ﾀ":"タ","ﾁ":"チ","ﾂ":"ツ","ﾃ":"テ","ﾄ":"ト",
    "ﾅ":"ナ","ﾆ":"ニ","ﾇ":"ヌ","ﾈ":"ネ","ﾉ":"ノ",
    "ﾊ":"ハ","ﾋ":"ヒ","ﾌ":"フ","ﾍ":"ヘ","ﾎ":"ホ",
    "ﾏ":"マ","ﾐ":"ミ","ﾑ":"ム","ﾒ":"メ","ﾓ":"モ",
    "ﾔ":"ヤ","ﾕ":"ユ","ﾖ":"ヨ",
    "ﾗ":"ラ","ﾘ":"リ","ﾙ":"ル","ﾚ":"レ","ﾛ":"ロ",
    "ﾜ":"ワ","ｦ":"ヲ","ﾝ":"ン",
    "ｧ":"ァ","ｨ":"ィ","ｩ":"ゥ","ｪ":"ェ","ｫ":"ォ",
    "ｯ":"ッ","ｬ":"ャ","ｭ":"ュ","ｮ":"ョ",
    "｡":"。","､":"、","ｰ":"ー","｢":"「","｣":"」","･":"・"
  };

  const dakutenMap = {
    "カ":"ガ","キ":"ギ","ク":"グ","ケ":"ゲ","コ":"ゴ",
    "サ":"ザ","シ":"ジ","ス":"ズ","セ":"ゼ","ソ":"ゾ",
    "タ":"ダ","チ":"ヂ","ツ":"ヅ","テ":"デ","ト":"ド",
    "ハ":"バ","ヒ":"ビ","フ":"ブ","ヘ":"ベ","ホ":"ボ",
    "ウ":"ヴ"
  };

  const handakutenMap = {
    "ハ":"パ","ヒ":"ピ","フ":"プ","ヘ":"ペ","ホ":"ポ"
  };

  let out = "";
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    const next = str[i + 1];

    if (ch === "ﾞ" || ch === "ﾟ") continue;

    if (baseMap[ch]) {
      let full = baseMap[ch];

      if (next === "ﾞ" && dakutenMap[full]) {
        full = dakutenMap[full];
        i++;
      } else if (next === "ﾟ" && handakutenMap[full]) {
        full = handakutenMap[full];
        i++;
      }

      out += full;
    } else {
      out += ch;
    }
  }
  return out;
}
