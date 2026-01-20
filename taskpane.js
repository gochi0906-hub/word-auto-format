/* global Office, Word */

Office.onReady(() => {
  const btn = document.getElementById("runBtn");
  const status = document.getElementById("status");

  const setStatus = (msg) => { status.textContent = msg; };

  btn.addEventListener("click", async () => {
    try {
      setStatus("整形中…");

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("text");
        await context.sync();

        const original = range.text || "";
        if (!original.trim()) {
          setStatus("本文で文字を選択してください。");
          return;
        }

        const formatted = formatText(original);

        if (formatted === original) {
          setStatus("変更はありませんでした。");
          return;
        }

        range.insertText(formatted, Word.InsertLocation.replace);
        await context.sync();

        setStatus("整形が完了しました。");
      });

    } catch (e) {
      console.error(e);
      setStatus("エラーが発生しました。選択範囲があるか確認してください。");
    }
  });
});

/**
 * 体裁整形ルール
 * - 全角英数字 → 半角
 * - 半角カナ → 全角カナ（濁点/半濁点も対応）
 * - ｰ 等 → ー
 */
function formatText(input) {
  let s = input;

  // 1) 全角英数字 -> 半角
  s = s.replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  s = s.replace(/　/g, " "); // 全角スペース -> 半角スペース（必要なら）

  // 2) 半角カナ -> 全角カナ（濁点/半濁点込み）
  s = hankakuKanaToZenkaku(s);

  // 3) 長音/ハイフンっぽいものを統一
  // ｰ(FF70)､ －(FF0D)､ ﹣(FE63)､ ｰ etc -> ー
  s = s.replace(/[\uFF70\uFF0D\uFE63\u2010\u2011\u2012\u2013\u2212]/g, "ー");

  return s;
}

/**
 * 半角カナ -> 全角カナ
 * ざっくり実務向け：主要範囲と濁点/半濁点を処理
 */
function hankakuKanaToZenkaku(str) {
  // 半角カナ -> 全角カナ 基本マップ
  const map = {
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
    "ｬ":"ャ","ｭ":"ュ","ｮ":"ョ","ｯ":"ッ",
    "｡":"。","､":"、","･":"・","｢":"「","｣":"」",
    "ｰ":"ー"
  };

  // 濁点/半濁点
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

  // まず基本置換
  let out = str.replace(/[｡-ﾟ]/g, (ch) => map[ch] || ch);

  // ﾞ(FF9E) と ﾟ(FF9F) を前文字に合成
  // ※基本置換後にも残る可能性があるので、合成を後処理
  out = out
    .replace(/([カ-トハヒフヘホウ])ﾞ/g, (m, p1) => dakutenMap[p1] || (p1 + "゛"))
    .replace(/([ハヒフヘホ])ﾟ/g, (m, p1) => handakutenMap[p1] || (p1 + "゜"));

  // 置換できずに残った半角濁点/半濁点は削除（必要なら）
  out = out.replace(/[ﾞﾟ]/g, "");

  return out;
}
