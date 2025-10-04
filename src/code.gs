/** =========================
 * Before/After Sheet Generator
 * フォーム送信 → スライド生成 → PDF保存 → シートにURL書き戻し
 * ========================= */

const CONFIG = {
  TEMPLATE_ID: 'INPUT YOUR TENMPLATE ID',
  OUTPUT_FOLDER_ID: 'INPUT YOUR FOLDER ID',
  SHEET_NAME: 'フォームの回答 1', // フォーム連携シート名（変えていなければデフォルトこれ）

  // フォームの質問タイトル（フォーム側と“完全一致”にする）
  Q: {
    BEFORE: 'ビフォー写真',
    AFTER: 'アフター写真',
    NAME: '実施者',
    PLACE: '改善場所',
    DATE: '改善日時',
    COMMENT: 'コメント'
  },

  // スライドの差し込みマーカー
  MARKER: {
    DATE: '{{DATE}}',
    NAME: '{{NAME}}',
    PLACE: '{{PLACE}}',
    COMMENT: '{{COMMENT}}',
    BEFORE: '[[BEFORE_IMG]]',
    AFTER: '[[AFTER_IMG]]'
  },

  // シートに書き戻す列名（無ければ自動で追加）
  WRITEBACK_COLS: {
    SLIDES_URL: 'スライドURL',
    PDF_URL: 'PDF URL'
  }
};

/**
 * スプレッドシートの「フォーム送信時」トリガーで動く想定
 * （Apps Script 画面の“トリガー”で onFormSubmit を「スプレッドシート/フォーム送信時」に設定）
 */
function onFormSubmit(e) {
  try {
    const nv = e.namedValues || {};
    const row = e.range.getRow();

    const name = pickFirst(nv[CONFIG.Q.NAME]);
    const place = pickFirst(nv[CONFIG.Q.PLACE]);
    const comment = pickFirst(nv[CONFIG.Q.COMMENT]);

    // 日時：優先順位 ①改善日時（入力あり）→ ②タイムスタンプ → ③現在時刻
    const dateStrInput = pickFirst(nv[CONFIG.Q.DATE]);
    const timestamp = pickFirst(nv['タイムスタンプ']); // シート1列目（日本語環境）
    const dateObj = parseDateFlexible_(dateStrInput) || parseDateFlexible_(timestamp) || new Date();
    const dateStrJP = formatDateJP_(dateObj); // 例: 2025/9/24

    // 画像ファイルID（フォームのファイルアップロード回答のURLから抽出）
    const beforeUrl = pickFirst(nv[CONFIG.Q.BEFORE]);
    const afterUrl  = pickFirst(nv[CONFIG.Q.AFTER]);
    const beforeId  = extractDriveFileId_(beforeUrl);
    const afterId   = extractDriveFileId_(afterUrl);

    if (!beforeId || !afterId) {
      throw new Error('画像のファイルIDが取得できませんでした。フォームのファイルアップロード設定と質問タイトルを確認してください。');
    }

    const beforeBlob = DriveApp.getFileById(beforeId).getBlob();
    const afterBlob  = DriveApp.getFileById(afterId).getBlob();

    // テンプレ複製 → 差し込み
    const outFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    const outName = `BeforeAfter_${formatDateForFile_(dateObj)}_${name || 'noname'}_${place || 'noplace'}`;

    const newFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID).makeCopy(outName, outFolder);
    const presId = newFile.getId();
    const pres = SlidesApp.openById(presId);
    const slide = pres.getSlides()[0];

    // テキスト差し込み
    pres.replaceAllText(CONFIG.MARKER.DATE, dateStrJP);
    pres.replaceAllText(CONFIG.MARKER.NAME, name || '');
    pres.replaceAllText(CONFIG.MARKER.PLACE, place || '');
    pres.replaceAllText(CONFIG.MARKER.COMMENT, comment || '');

    // 画像差し込み（枠=マーカー文字の入った図形）
    insertImageAtMarkerBox_(slide, CONFIG.MARKER.BEFORE, beforeBlob);
    insertImageAtMarkerBox_(slide, CONFIG.MARKER.AFTER,  afterBlob);

    pres.saveAndClose();

    // PDF化
    const pdfBlob = newFile.getAs(MimeType.PDF).setName(`${outName}.pdf`);
    const pdfFile = outFolder.createFile(pdfBlob);

    // URL書き戻し
    const sheet = e.source.getSheetByName(CONFIG.SHEET_NAME) || e.source.getActiveSheet();
    const colSlides = findOrCreateHeaderColumn_(sheet, CONFIG.WRITEBACK_COLS.SLIDES_URL);
    const colPdf    = findOrCreateHeaderColumn_(sheet, CONFIG.WRITEBACK_COLS.PDF_URL);

    sheet.getRange(row, colSlides).setValue(newFile.getUrl());
    sheet.getRange(row, colPdf).setValue(pdfFile.getUrl());
  } catch (err) {
    // エラー内容をログ＆シート最終列に出す（任意）
    console.error(err);
    const sheet = e?.source?.getSheetByName(CONFIG.SHEET_NAME) || e?.source?.getActiveSheet();
    if (sheet) {
      const row = e.range.getRow();
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue('エラー');
      sheet.getRange(row, col).setValue(String(err));
    }
    throw err; // 実行ログにも残す
  }
}

/* ---------- 以下、ユーティリティ群（初心者は触らなくてOK） ---------- */

// 配列や文字列から先頭要素を取る（フォームの namedValues は配列）
function pickFirst(v) {
  if (Array.isArray(v)) return v[0] || '';
  return v || '';
}

// 様々なGoogle DriveリンクからファイルIDを抽出
function extractDriveFileId_(url) {
  if (!url) return '';
  // 代表的なパターンに全部マッチするよう、25文字以上の英数/ハイフン/アンダースコアを拾う
  const m = String(url).match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

// 「YYYY/M/D」
function formatDateJP_(d) {
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

// ファイル名向け「YYYYMMDD_HHmm」
function formatDateForFile_(d) {
  const z2 = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${z2(d.getMonth()+1)}${z2(d.getDate())}_${z2(d.getHours())}${z2(d.getMinutes())}`;
}

// シートの見出し行に指定ヘッダが無ければ追加して、その列番号を返す（1始まり）
function findOrCreateHeaderColumn_(sheet, headerName) {
  const lastCol = sheet.getLastColumn() || 1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(headerName);
  if (idx >= 0) return idx + 1;
  // 右端に追加
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

// 文字入り図形（マーカー）を探して、同じ位置サイズに画像を“等倍比でフィット＆中央寄せ”で配置
function insertImageAtMarkerBox_(slide, markerText, blob) {
  const el = findShapeByText_(slide, markerText);
  if (!el) throw new Error(`スライド上にマーカー「${markerText}」の図形が見つかりません。`);

  const box = {
    left: el.getLeft(),
    top: el.getTop(),
    width: el.getWidth(),
    height: el.getHeight()
  };

  // マーカー図形は削除してから画像を置く
  el.remove();

  // ✅ 位置・サイズ指定なしでいったん挿入（これで元画像の実寸を取得できる）
  const img = slide.insertImage(blob);

  // 元画像サイズ（スライド上の初期表示サイズ）を取得
  const iw = img.getWidth();
  const ih = img.getHeight();

  // “箱に収まる（contain）”比率で縮尺
  const scale = Math.min(box.width / iw, box.height / ih);
  img.setWidth(iw * scale);
  img.setHeight(ih * scale);

  // 箱の中央に配置
  const newLeft = box.left + (box.width - img.getWidth()) / 2;
  const newTop  = box.top  + (box.height - img.getHeight()) / 2;
  img.setLeft(newLeft);
  img.setTop(newTop);
}


// スライド内の図形テキストに markerText を含む最初の要素を返す
function findShapeByText_(slide, markerText) {
  const els = slide.getPageElements();
  for (const el of els) {
    if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const txt = el.asShape().getText().asString();
      if (txt && txt.indexOf(markerText) !== -1) return el;
    }
  }
  return null;
}

// ゆるく日時文字列をDateに（空や不正は null）
function parseDateFlexible_(v) {
  if (!v) return null;
  try {
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  } catch (_) { return null; }
}
