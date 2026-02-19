/**
 * 【設定】
 * 1. 読み込み元フォルダID
 * 2. 出力先フォルダID
 * 3. 無視するタブ名の正規表現リスト
 * 4. タイムアウト設定
 */
const CONFIG = {
  INPUT_FOLDER_ID: '1pfoAoJ_tBA0YphjMhz-vDM1muPVKBmZp',
  OUTPUT_FOLDER_ID: '1KzbdiJu1IMOMYWDxjU5VWbqXOxt0b9-p',

  // 出力ファイル名の接頭辞
  OUTPUT_FILE_PREFIX: '【選別済】',

  // タイムアウト設定 (ミリ秒)
  // 20分 = 20 * 60 * 1000
  TIME_LIMIT: 20 * 60 * 1000,

  // 無視する（削除する）タブ名の正規表現リスト
  IGNORE_TAB_PATTERNS: [
    /^はじめに$/,
    /^学び関連$/,
    /^労務関連（勤怠・稼働のつけ方）$/,
    /^週次会議【↓】$/,
    /^インフラ構築スケジュール$/,
    /.*参考.*/
  ]
};

/**
 * プロパティキーの定義
 */
const PROP_KEYS = {
  CURRENT_INDEX: 'PROCESS_INDEX',     // 現在の処理インデックス
  SOURCE_LIST: 'PROCESS_SOURCE_LIST', // 元ファイルのIDリスト
  IS_RUNNING: 'PROCESS_IS_RUNNING'    // 実行中フラグ
};

/**
 * メイン関数: ファイルを1つずつ処理して不要なタブを除去する
 */
function Optimize_Weekly_MTG() {
  const props = PropertiesService.getScriptProperties();
  const isRunning = props.getProperty(PROP_KEYS.IS_RUNNING);

  // 初回実行時（またはリセット後）
  if (!isRunning) {
    console.log("--- 処理を開始します（不要タブ除去モード） ---");

    // ファイルリストの取得
    const inputFolder = DriveApp.getFolderById(CONFIG.INPUT_FOLDER_ID);
    const files = inputFolder.getFilesByType(MimeType.GOOGLE_DOCS);
    const sourceList = [];
    while (files.hasNext()) {
      sourceList.push(files.next().getId());
    }

    if (sourceList.length === 0) {
      console.warn("対象フォルダにGoogleドキュメントが見つかりませんでした。");
      return;
    }

    // 初期化
    props.setProperty(PROP_KEYS.SOURCE_LIST, JSON.stringify(sourceList));
    props.setProperty(PROP_KEYS.CURRENT_INDEX, '0');
    props.setProperty(PROP_KEYS.IS_RUNNING, 'true');

    console.log(`対象ファイル数: ${sourceList.length} 件`);
  }

  runProcessingPhase();
}

/**
 * 処理フェーズ: リスト順にファイルを処理し、出力フォルダへ保存する
 */
function runProcessingPhase() {
  const props = PropertiesService.getScriptProperties();
  const sourceList = JSON.parse(props.getProperty(PROP_KEYS.SOURCE_LIST) || '[]');
  let currentIndex = parseInt(props.getProperty(PROP_KEYS.CURRENT_INDEX) || '0');

  const outputFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const total = sourceList.length;

  const startTimePhase = Date.now();

  console.log(`=== 処理開始: ${currentIndex + 1} / ${total} 件目から ===`);

  while (currentIndex < total) {
    // 時間制限チェック (CONFIGの定数を使用)
    if (Date.now() - startTimePhase > CONFIG.TIME_LIMIT) {
      console.log(`--- 時間制限のため中断します（次回は ${currentIndex + 1} 件目から再開） ---`);
      return;
    }

    const sourceId = sourceList[currentIndex];

    // 万が一ファイルが削除されている場合などを考慮
    try {
      const sourceFile = DriveApp.getFileById(sourceId);
      const sourceName = sourceFile.getName();
      const outputName = `${CONFIG.OUTPUT_FILE_PREFIX}${sourceName}`;

      console.log(`処理中 (${currentIndex + 1}/${total}): ${sourceName}`);
      const startTime = Date.now();

      // 新しいドキュメントを作成 (作成直後はルートフォルダに配置される)
      const newDoc = DocumentApp.create(outputName);
      const newId = newDoc.getId();
      const newFile = DriveApp.getFileById(newId);

      // 指定フォルダへ移動 (これでOUTPUT_FOLDER_ID配下に配置)
      newFile.moveTo(outputFolder);

      const targetBody = newDoc.getBody();
      const sourceDoc = DocumentApp.openById(sourceId);

      // ★ここでタブの選別処理を実行
      processTabsToTarget(sourceDoc, targetBody);

      newDoc.saveAndClose();

      const duration = ((Date.now() - startTime) / 1000).toFixed(2);
      console.log(`    -> 完了 (時間: ${duration}秒)`);

    } catch (e) {
      console.error(`    !! エラー発生 (Index: ${currentIndex}): ${e.message}`);
      // エラーでも止まらず次へ進む
    }

    // インデックスを進めて保存
    currentIndex++;
    props.setProperty(PROP_KEYS.CURRENT_INDEX, currentIndex.toString());
  }

  // 全件完了
  console.log("=== 全ファイルの処理が完了しました ===");
  finalizeProcess();
}

/**
 * 完了後の処理
 */
function finalizeProcess() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log("進捗データをリセットしました。");
}

/**
 * 進捗リセット（手動実行用）
 */
function resetProgress() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  console.log("実行進捗をリセットしました。最初からやり直せます。");
}

// ==========================================
// ヘルパー関数（ロジック部分は変更なし）
// ==========================================

function processTabsToTarget(sourceDoc, targetBody) {
  const tabs = sourceDoc.getTabs();
  if (tabs && tabs.length > 0) {
    processTabsRecursively(tabs, targetBody, 0);
  } else {
    // タブがない古い形式のドキュメントの場合
    appendBodyToTarget(sourceDoc.getBody(), targetBody);
  }
}

function processTabsRecursively(tabs, targetBody, level) {
  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    const title = tab.getTitle();

    // 無視リストに含まれるかチェック
    const shouldSkip = CONFIG.IGNORE_TAB_PATTERNS.some(pattern => pattern.test(title));
    if (shouldSkip) {
      console.log(`    [SKIP] タブ除外: ${title}`);
      continue;
    }

    // タブ名をドキュメント内に見出しとして残す（不要ならこのブロックを削除）
    const headingLevel = Math.min(level + 1, 6); // Heading 1 から開始
    const tabTitle = targetBody.appendParagraph(title);
    tabTitle.setHeading(getHeadingTypeByLevel(headingLevel));

    if (tab.getType() === DocumentApp.TabType.DOCUMENT_TAB) {
      const docTab = tab.asDocumentTab();
      const body = docTab.getBody();
      if (body) appendBodyToTarget(body, targetBody);
    }

    const childTabs = tab.getChildTabs();
    if (childTabs && childTabs.length > 0) {
      processTabsRecursively(childTabs, targetBody, level + 1);
    }
  }
}

function appendBodyToTarget(sourceBody, targetBody) {
  const numChildren = sourceBody.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = sourceBody.getChild(i).copy();
    const type = element.getType();

    try {
      switch (type) {
        case DocumentApp.ElementType.PARAGRAPH:
          targetBody.appendParagraph(element.asParagraph());
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          targetBody.appendListItem(element.asListItem());
          break;
        case DocumentApp.ElementType.TABLE:
          targetBody.appendTable(element.asTable());
          break;
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          targetBody.appendHorizontalRule();
          break;
        case DocumentApp.ElementType.PAGE_BREAK:
          targetBody.appendPageBreak();
          break;
        case DocumentApp.ElementType.TABLE_OF_CONTENTS:
          targetBody.appendTableOfContents(element.asTableOfContents());
          break;
        case DocumentApp.ElementType.INLINE_IMAGE:
          targetBody.appendImage(element.asInlineImage());
          break;
        default:
          break;
      }
    } catch (e) {
      console.warn(`要素コピー失敗 (Type: ${type}): ${e.message}`);
    }
  }
}

function getHeadingTypeByLevel(level) {
  const types = [
    DocumentApp.ParagraphHeading.HEADING1,
    DocumentApp.ParagraphHeading.HEADING2,
    DocumentApp.ParagraphHeading.HEADING3,
    DocumentApp.ParagraphHeading.HEADING4,
    DocumentApp.ParagraphHeading.HEADING5,
    DocumentApp.ParagraphHeading.HEADING6
  ];
  return types[level - 1] || DocumentApp.ParagraphHeading.NORMAL;
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('選別ツール')
    .addItem('選別処理を実行/再開', 'mergeGoogleDocsConsole')
    .addSeparator()
    .addItem('進捗をリセット', 'resetProgress')
    .addToUi();
}