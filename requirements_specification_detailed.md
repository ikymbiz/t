# 統合コックピット HTA 詳細要件定義書（完全版）

---

## Module-01: アプリケーション基盤

### M01-01: HTA宣言・ウィンドウ制御

**要件:**
- `<HTA:APPLICATION>` タグによるウィンドウ属性を定義する
- 属性一覧:
  - `SINGLEINSTANCE="yes"` — 多重起動禁止（2回目の起動は既存ウィンドウをアクティブに）
  - `WINDOWSTATE="maximize"` — 起動時に最大化
  - `BORDER="thick"` — リサイズ可能な太枠
  - `SCROLL="no"` — body自体のスクロールバー非表示（各ペイン内でスクロール）
  - `CAPTION="yes"` — タイトルバー表示
  - `SHOWINTASKBAR="yes"` — タスクバーに表示
  - `SYSMENU="yes"` — システムメニュー（閉じる/最小化/最大化）表示
  - `SELECTION="yes"` — テキスト選択を許可
  - `ICON=""` — アイコン指定（空の場合はHTAデフォルト）
- `<meta http-equiv="X-UA-Compatible" content="IE=9">` でIE9モード固定

**留意:**
- IE=9を指定しない場合、IE5互換モードでレンダリングされる恐れがある
- `SINGLEINSTANCE`は同一パスのHTAファイルに対してのみ有効。ファイル名を変えると別インスタンスとして起動される

---

### M01-02: VBScript ファイルI/O層

**関数一覧:**

| VBScript関数 | 引数 | 説明 |
|-------------|------|------|
| `getAppDir()` | なし | HTAファイルの格納ディレクトリを返す。末尾に`\`付き |
| `readUTF8(filePath)` | ファイルパス | ADODB.StreamでUTF-8読み込み。結果を`_vbs_result`にセット |
| `writeUTF8(filePath, content)` | パス, 内容 | ADODB.StreamでUTF-8書き込み。成功なら`_vbs_result`に"OK"、失敗なら"ERROR:..."をセット |
| `vbsReadFile()` | なし（`_vbs_buffer`から取得） | readUTF8のラッパー。完了後`jsVbsCallback()`をJSで呼び出す |
| `vbsWriteFile()` | なし（`_vbs_filepath`/`_vbs_buffer`から取得） | writeUTF8のラッパー。完了後`jsVbsCallback()`をJSで呼び出す |
| `vbsEnsureDir(dirPath)` | ディレクトリパス | FSO.FolderExistsチェック→存在しなければFSO.CreateFolder |
| `vbsBrowseFile()` | なし | `UserAccounts.CommonDialog`でファイル選択ダイアログを表示 |
| `vbsFileExists(fp)` | ファイルパス | FSO.FileExistsの結果を返す |
| `vbsFolderExists(fp)` | フォルダパス | FSO.FolderExistsの結果を返す |

**BOM処理の詳細:**
- ADODB.StreamでUTF-8読み込み時、BOM (U+FEFF) が先頭に残る場合がある
- `AscW(Left(result, 1)) = 65279` でBOMを検出し、`Mid(result, 2)` で除去する
- この処理は `readUTF8` 内で自動実行される

**エラー処理:**
- VBScript側: `On Error Resume Next` → 処理 → `Err.Number` チェック → `Err.Clear` → `On Error GoTo 0`
- ファイルが存在しない場合: 空文字列を返す（エラーにはしない）
- 書き込み失敗時: `"ERROR:" & Err.Description` を `_vbs_result` にセット

**ADODB.Stream設定:**
- `s.Type = 2` （テキストモード）
- `s.Charset = "UTF-8"`
- 書き込み時: `s.SaveToFile filePath, 2` （2 = adSaveCreateOverWrite）

---

### M01-03: JS⇔VBS ブリッジ

**仕組み:**
1. HTML内にhidden input要素3つを配置
   - `#_vbs_buffer` — 読取対象パス or 書き込み内容のバッファ
   - `#_vbs_result` — VBScript実行結果の格納先
   - `#_vbs_filepath` — 書き込み先パスの格納先
2. JS側の呼出フロー:
   ```
   readFileSync(path):
     _vbs_buffer.value = path
     _vbs_result.value = ""
     window.execScript('vbsReadFile()', 'VBScript')
     return _vbs_result.value
   ```
   ```
   writeFileSync(path, content):
     _vbs_filepath.value = path
     _vbs_buffer.value = content
     _vbs_result.value = ""
     window.execScript('vbsWriteFile()', 'VBScript')
     return _vbs_result.value
   ```
3. `VBS_CALLBACK` グローバル変数: 非同期コールバック用（現行は同期呼出のため未使用）

**制約:**
- `window.execScript` はHTAでのみ有効。モダンブラウザでは動作しない
- 同期実行前提のため、大量ファイル操作時はUIがフリーズする可能性がある
- hidden inputの value プロパティに格納可能なテキスト量に実質的な上限あり（数MB程度）

---

### M01-04: ES5ポリフィル

**実装が必要なポリフィル一覧:**

| メソッド | 仕様 |
|---------|------|
| `Array.prototype.indexOf(v, s)` | 値`v`の最初の出現インデックスを返す。見つからなければ-1 |
| `Array.prototype.forEach(fn, ctx)` | 各要素に`fn(elem, index, array)`を実行 |
| `Array.prototype.filter(fn, ctx)` | `fn`がtrueを返す要素のみの新配列を返す |
| `Array.prototype.map(fn, ctx)` | 各要素を`fn`で変換した新配列を返す |
| `Array.isArray(a)` | `Object.prototype.toString.call(a) === '[object Array]'` で判定 |
| `String.prototype.trim()` | `this.replace(/^\s+|\s+$/g, '')` |

**追加で必要になる可能性のあるポリフィル（cockpit_v9移植時）:**
- `Array.prototype.some` — cockpit_v9のフィルタロジックで潜在的に必要
- `Array.prototype.reduce` — 集計処理で必要になる可能性
- `Object.keys` — オブジェクトのキー一覧取得
- `JSON.parse` / `JSON.stringify` — IE8以上で標準搭載。IE7以下では要ポリフィル（対象外）
- `Date.prototype.toISOString` — cockpit_v9で`nowISO()`が使用

---

### M01-05: 設定ファイル管理

**ファイル:** `app_config.json`（HTAと同階層 or `baseDir`指定パス）

**設定ロードフロー:**
1. `resolvePath('app_config.json')` でパスを解決
2. `readFileSync` でJSON読み込み
3. `safeParseJSON` でパース
4. `configVer` をチェック（現行要求バージョン: 4）
5. 設定が無い or バージョン不一致 → `getDefaultConfig()` でデフォルト生成 → 即時保存

**設定保存:**
- `JSON.stringify(CONFIG, null, 2)` でインデント付きJSON生成
- `writeFileSync` でファイルに書き込み

**バージョン管理:**
- `CONFIG.app.configVer` が `REQUIRED_CONFIG_VER`（現在4）と一致しない場合、旧版/文字化けとみなしてリセット
- バージョン番号を上げる場合、マイグレーション処理の追加を検討すること

---

### M01-06: パス解決

**関数:** `resolvePath(rel)`

**ロジック:**
1. 引数が空 → 空文字列を返す
2. 絶対パス判定: `rel.indexOf(':') > 0` → そのまま返す（例: `C:\data\file.csv`）
3. `/` を `\` に変換
4. 先頭の `.\` or `.\\` を除去
5. `APP_DIR + r` で絶対パス化

**APP_DIR取得フロー:**
1. VBScript `getAppDir()` を `window.execScript` で呼び出し
2. `document.location.pathname` から `\` に変換、先頭の `\` を除去、`%20` をスペースに変換
3. `FSO.GetParentFolderName` で親フォルダを取得
4. 末尾に `\` を付与

---

### M01-07: ディレクトリ保証

**関数:** `ensureDir(dirPath)`

**動作:**
- JS側: `window.execScript('vbsEnsureDir "' + dirPath + '"', 'VBScript')`
- VBS側: `gFSO.FolderExists(dirPath)` がFalseなら `gFSO.CreateFolder dirPath`

**制約:**
- 1階層のみ作成。再帰的な深い階層の自動作成は未対応
- 必要なら `\` 区切りでパスを分割し、ループで各階層を作成する処理を追加する

---

## Module-02: データストア

### M02-01: CSVパーサー / ジェネレーター

**パーサー (`parseCSV`)**

入力: CSV文字列
出力: `{ headers: string[], rows: object[] }`

**パースルール（RFC 4180準拠）:**
- 改行コードの正規化: `\r\n` → `\n`, `\r` → `\n`
- フィールド区切り: `,`
- クォーティング: `"` で囲まれたフィールド内では改行・カンマ・ダブルクォートを含めることが可能
- エスケープ: `""` → `"`
- 1行目をヘッダーとして扱う
- 空行（フィールド1つで空文字列のみの行）はスキップ
- 各データ行をオブジェクト `{ [header]: value }` に変換

**ジェネレーター (`generateCSV`)**

入力: `headers: string[]`, `rows: object[]`
出力: CSV文字列（改行コード: `\r\n`）

**クォーティング規則 (`csvQ`):**
- カンマ、ダブルクォート、改行を含む場合のみクォート
- ダブルクォートは `""` にエスケープ
- null/undefined は空文字列に変換

---

### M02-02: データストア管理

**グローバル変数:** `DATA_STORE = {}`

**構造:** `paneId → Array<Object>` （各ObjectはCSVの1行に対応）

**メタデータプロパティ（配列オブジェクトに直接付与）:**
- `_file` — 対応するCSVファイル名（例: "wbs_tasks.csv"）
- `_loaded` — ロード済みフラグ（true/false）
- `_headers` — カラムキーの配列

---

### M02-03: ペイン間データ共有

**`loadDataForPane(pane)` の共有ロジック:**
1. ペインのデータを読み込んだ後、全タブ・全ペインを走査
2. 同じ `dataFile` を持つ他ペインに対して `DATA_STORE[p.id] = DATA_STORE[pane.id]` で参照を共有
3. これにより、「タスク一覧」タブと「WBS」タブが同じ `wbs_tasks.csv` を参照する場合、一方の編集が他方に即座に反映される

**注意:**
- 参照共有のため、一方のソートが他方にも影響する（意図的な設計）
- 再ロード防止: `DATA_STORE[pane.id]._loaded === true` かつ `_file` が一致する場合はスキップ

---

### M02-04: ダーティフラグ管理

**グローバル変数:** `DIRTY_FLAGS = {}` （paneId → boolean）

**`markDirty(paneId)`:**
1. `DIRTY_FLAGS[paneId] = true`
2. `updateDirtyUI()` を呼び出し

**`updateDirtyUI()`:**
1. 全ペインの `DIRTY_FLAGS` をスキャン
2. いずれかがtrue → ヘッダーの「● 未保存」表示、ステータスバーに「● 未保存あり」表示（orange色）
3. すべてfalse → 非表示

---

### M02-05: 一括保存

**`saveAllData()`:**
1. 重複保存防止: `savedFiles = {}` で同一ファイル名の二重保存を防止
2. 全タブ → 全ペインを走査:
   - `dataFile` が設定されているペイン → `saveDataForPane(pane)` でCSV保存
   - `type === 'editor'` → `saveEditorContent(pane)` でテキスト保存
3. 全 `DIRTY_FLAGS` をfalseにリセット
4. `updateDirtyUI()` 呼び出し
5. トースト「すべてのデータを保存しました」表示
6. ステータスバーに「保存: HH:MM:SS」表示

**`saveDataForPane(pane)`:**
1. `DATA_STORE[pane.id]` を取得
2. `_headers` からヘッダー行生成
3. `generateCSV` でCSV文字列生成
4. `resolvePath` + `writeFileSync` でファイル書き込み
5. `DIRTY_FLAGS[pane.id] = false`

---

### M02-06: localStorage ストレージ（cockpit_v9由来）

**HTA移植時の差替仕様:**

| 元の関数 | HTA版の実装 |
|---------|-------------|
| `lsRead(key)` → `localStorage.getItem('ck_' + key)` | → `readFileSync(resolvePath(key + '.json'))` → `JSON.parse` |
| `lsWrite(key, arr)` → `localStorage.setItem(...)` | → `writeFileSync(resolvePath(key + '.json'), JSON.stringify(arr))` |
| `lsAppend(key, obj)` → read→push→write | → 同様のread→push→writeだがファイルI/O |

**データテーブルのファイルマッピング:**
- `records` → `records.json`
- `task_log` → `task_log.json`
- `intelligence_inbox` → `intelligence_inbox.json`
- `project_master` → `project_master.json`
- `client_master` → `client_master.json`

---

### M02-07: バックアップ / エクスポート

**DynamicApp版:**
- 設定JSON: `exportConfig()` → `app_config_export.json` に出力
- インポート: `importConfig()` → JSONファイル名をプロンプト入力 → 読み込み → 検証（app, tabsプロパティ存在チェック） → 適用 → `initApp()` で再初期化

**cockpit_v9版:**
- 全テーブルをJSONに統合 → クリップボードにコピー
- HTA移植時: クリップボードコピーは `document.execCommand('copy')` に変更 or ファイル出力に変更

---

## Module-03: UI フレームワーク

### M03-01: ヘッダー / リボン

**DynamicApp ヘッダー構成:**
- 高さ: 42px, 背景: #1a237e, 文字: #fff
- 左: アプリ名 `#appName`（18px bold）
- 左（アプリ名の右）: ダーティマーク `#dirtyMark`（「● 未保存」、orange色、display:noneで初期非表示）
- 右: バージョン `#appVersion`（13px, #90caf9）

**cockpit_v9 リボン構成:**
- 高さ: 44px, 背景: #1a1d28
- 左: ブランド名「COCKPIT」（18px 800weight, #e8b84b, letter-spacing:3px）
- 中: 機能ボタン群（エディタ/情報収集/議事録/全文検索/バックアップ/フォルダ）
- 右: ステータスドット（レコード数/タスク数）+ デジタル時計

**統合方針:**
- リボン形式を採用。左にアプリ名+ダーティマーク、中央にクイックアクセスボタン、右にステータス+時計

---

### M03-05: モーダルダイアログ

**タイプ別仕様:**

| タイプ | 入力 | ボタン | コールバック引数 |
|--------|------|--------|-----------------|
| `alert` | message | OK | `callback('ok')` |
| `confirm` | message, okLabel, cancelLabel, thirdLabel | キャンセル / (第3ボタン) / OK | `callback('ok'|'cancel'|'third')` |
| `prompt` | message, defaultValue | キャンセル / OK | `callback('ok'|'cancel', inputValue)` |
| `form` | fields[] | キャンセル / OK | `callback('ok'|'cancel', {key:value})` |

**formフィールドタイプ:**
- `text` → `<input type="text">`
- `select` → `<select>` + options
- `textarea` → `<textarea>`
- `radio` → `<input type="radio" name="mf_key">` × options
- `number` → `<input type="text" style="width:80px">`
- `date` → `<input type="text" style="width:160px">`
- `checkbox` → `<input type="checkbox">`

**位置:** `position:fixed; top:50%; left:50%` + `marginLeft: -w/2`, `marginTop: -120px`

**オーバーレイ:** `position:fixed; top:0; left:0; right:0; bottom:0; background:#000; opacity:0.5; z-index:1000`（IE用に `filter:alpha(opacity=50)` 併用）

**グローバル状態:**
- `window._modalCallback` — コールバック関数
- `window._modalType` — ダイアログ種別
- `window._modalOpts` — オプション

---

### M03-06: コンテキストメニュー

**`showCtxMenu(e, items)`:**
- 引数 `e`: イベントオブジェクト（`e = e || window.event` でIE対応）
- 引数 `items[]`: メニュー項目配列
  - `{ label: "テキスト", action: "function(){...}" }` — 通常項目
  - `{ separator: true }` — 区切り線
  - `{ label: "テキスト", action: "...", danger: true }` — 赤字の危険操作
- 位置: `e.clientX || e.x`, `e.clientY || e.y`
- `return false` でブラウザデフォルトの右クリックメニューを抑制

**`hideCtxMenu()`:**
- `document.onclick` に登録し、どこかクリックしたらメニューを閉じる

**action の実行方法:**
- `onclick="hideCtxMenu();(function(){...})()"` のパターン
- `action` プロパティの文字列がそのまま `onclick` 属性に埋め込まれる
- **セキュリティ注意:** 動的に生成されるaction文字列にユーザー入力が含まれる場合、XSSリスクがある

---

### M03-07: トースト通知

**`toast(type, msg)`:**
- type: `success`（緑#2e7d32）/ `error`（赤#c62828）/ `warning`（橙#e65100）/ `info`（青#1565c0）
- `#toastContainer`（position:fixed, bottom:40px, right:20px, z-index:2000）に追加
- `setTimeout` で自動削除（`CONFIG.app.toastDurationMs || 3000` ms後）
- クリックで即削除可能
- 複数トーストはスタック表示（上に積み上がる）

---

### M03-08: テーマ / スタイルシート

**DynamicApp（ライト系）:**
- 背景: #f0f2f5、サイドバー: #1a237e、アクセント: #3949ab
- テーブルヘッダー: #e8eaf6、ボーダー: #c5cae9
- `box-sizing: content-box`

**cockpit_v9（ダーク系）:**
- 背景: #12141a、サイドバー: #151820、アクセント: #e8b84b
- パネル背景: #1f2230、ボーダー: #2a2e40
- `box-sizing: border-box`

**統合時の決定事項:**
- `box-sizing` は `border-box` に統一（全要素への `* { box-sizing: border-box }` 適用）
- テーマ切替機能を設ける場合、CSS変数はIE9未対応のため、テーマごとのCSSクラスをbodyに付与する方式を採用
  - 例: `<body class="theme-dark">` → `.theme-dark #header { background: #1a1d28; }`

---

### M03-09: リアルタイム時計

**cockpit_v9 実装:**
- `setInterval(updateClock, 1000)` で1秒ごとに更新
- `updateClock()`: `new Date()` → `pad(hours) + ':' + pad(minutes) + ':' + pad(seconds)`
- 表示先: `#clock`（Consolas monospace, 16px, #e8b84b, font-weight:700, letter-spacing:2px）

---

## Module-04: タブ・ペインシステム

### M04-01: タブ描画・切替

**タブボタン描画 (`renderTabs`):**
- `CONFIG.tabs` から `visible !== false` のタブを抽出
- `order` 順にソート
- サイドバー `#tabbar` にボタン生成
- アクティブタブには `.active` クラス付与

**タブ切替 (`switchTab(tabId)`):**
1. `ACTIVE_TAB = tabId`
2. `renderTabs()` — サイドバー再描画
3. `renderMainArea()` — メインエリア全体を再描画
4. `updateToolbar()` — ツールバーの表示/非表示を更新
5. ステータスバー `#statusTab` にタブ名表示

**右クリックメニュー (`onTabCtxMenu`):**
- タブ名を変更 → `renameTab()` → prompt → `tab.label` 更新 → 保存 → 再描画
- タブを非表示 → `hideTab()` → `tab.visible = false` → 保存 → 再描画
- 上へ移動 / 下へ移動 → `moveTab()` → 隣接タブとorder値をスワップ → 保存 → 再描画

---

### M04-02: ペインレイアウト

**2ペイン水平分割（dir === 'horizontal' && panes.length === 2）:**
- 左ペイン: `left:0; right:(100-w1)%`
- スプリッター: `left:w1%; margin-left:-3px` （幅6px）
- 右ペイン: `left:w1%; right:0`
- 全ペインに `.pane` クラス → `position:absolute; top:0; bottom:0; overflow:auto`

**N ペイン（それ以外）:**
- `leftPos` を0から累積
- 各ペインに `left:leftPos%; width:w%` を設定

**垂直分割（dir === 'vertical'）:**
- `left:0; right:0; top:leftPos%; height:w%`

---

### M04-03: スプリッター

**`initSplitter(tabId, leftPane, rightPane)`:**

ドラッグ処理:
1. `splitter.onmousedown` でドラッグ開始
2. `document.onmousemove` でリアルタイム位置計算
   - `dx = ev.clientX - startX`
   - `newW = startLeft + dx`
   - `pct = Math.max(15, Math.min(85, Math.round(newW / cw * 100)))` — 15%〜85%の範囲制限
   - 左ペインの `right`、右ペインの `left`、スプリッターの `left` を更新
3. `document.onmouseup` でドラッグ終了（ハンドラ解除）
4. `e.preventDefault()` + `return false` でテキスト選択防止

**IE対応:**
- `ev.preventDefault` が存在しない場合の考慮（`if (e.preventDefault) e.preventDefault()`）

---

### M04-04: ペインリンク

**定義:** `CONFIG.paneLinks[]`

```
{ source: "pane_task_list", target: "pane_task_detail", trigger: "selectRow", action: "showDetail" }
```

**`triggerPaneLink(sourcePaneId, trigger, data)`:**
1. `CONFIG.paneLinks` をスキャン
2. `source` と `trigger` が一致するリンクを検索
3. アクション実行:
   - `showDetail` → ターゲットペインの `SELECTED_ROW` にデータインデックスをセット → `renderPaneContent(targetPane)`
   - `loadFile` → ターゲットペインの `renderPaneContent` を呼び出し（エディタがファイルを読み込む）
   - `highlightTask` → ガントチャートの行ハイライト

**デフォルトリンク定義:**
- `pane_task_list` → `pane_task_detail` (selectRow → showDetail)
- `pane_req_list` → `pane_req_detail` (selectRow → showDetail)
- `pane_risk_list` → `pane_risk_detail` (selectRow → showDetail)
- `pane_issue_list` → `pane_issue_detail` (selectRow → showDetail)
- `pane_minutes_list` → `pane_minutes_editor` (selectRow → loadFile)

---

### M04-05: ペインタイプルーター

**`renderPaneContent(pane)`:**

| type | 呼出関数 |
|------|---------|
| `list` | `renderListPane(pane, el)` |
| `detail` | `renderDetailPane(pane, el)` |
| `editor` | `renderEditorPane(pane, el)` |
| `tree` | `renderTreePane(pane, el)` |
| `wbs` | `renderWbsPane(pane, el)` |
| `gantt` | `renderGanttPane(pane, el)` |
| `menu` | `renderMenuPane(pane, el)` |
| `settings` | `renderSettingsPane(pane, el)` |
| `dashboard` | `renderDashboardPane(pane, el)` |
| default | 「未実装のペインタイプ: xxx」を表示 |

---

## Module-05: リスト（テーブル）ペイン

### M05-01: テーブル描画

**`renderListPane(pane, el)`:**

描画構造:
```
pane-header "タブ名 (N件)"
  └ div (overflow:auto, position:absolute, top:30px)
    └ table.list-table
      └ thead
        └ tr (ヘッダー行: # or 移動ボタン + 各列)
        └ tr.filter-row (フィルタ入力行)
      └ tbody
        └ tr × N (データ行)
```

**列のHTML生成ルール:**

| 列type | 表示方法 |
|--------|---------|
| `text` | `escapeHtml(val)` |
| `number` (progress系) | プログレスバー + パーセント数値 |
| `checkbox` | `<input type="checkbox">` + onclick→`toggleCheckbox()` |
| `select` | `escapeHtml(val)`（表示時はテキスト。編集時にselect化） |
| `date` | `escapeHtml(val)`（表示時はテキスト。編集時にカスタムピッカー） |

**ヘッダー行:**
- 通常モード: `#` 列（行番号表示、幅24px）
- 編集モード: `移動` 列（▲▼ボタン、幅50px）
- 各列ヘッダー: `onclick="sortList()"` + ソートアイコン（▲▼）

**データ行属性:**
- `onclick` → `selectListRow(paneId, actualIdx)`
- `ondblclick` → `editListCell(paneId, actualIdx, 1)` （2列目=最初のデータ列）
- `oncontextmenu` → `return onListRowCtxMenu(event, paneId, actualIdx)`
- ハイライトクラス: `getHighlightClass()` で条件付き書式適用

---

### M05-02: セル選択（アクティブセル）

**グローバル変数:** `ACTIVE_CELL = { paneId, row, col }`

**`setActiveCell(paneId, row, col)`:**
1. 前のアクティブセルから `.active-cell` クラスを除去
2. 新しいセルに `.active-cell` クラスを付与（outline:2px solid #1565c0）
3. `selectListRow()` を呼び出して行選択も連動

**`selectListRow(paneId, idx)`:**
1. `CELL_EDITING` 中はスキップ
2. `SELECTED_ROW[paneId] = idx`
3. DOM操作で `.selected` クラスを切替（全行再描画は行わない）
   - tbody内のtr要素をスキャンし、onclick属性内のインデックスで行を特定
4. `triggerPaneLink()` で連携ペインを更新

---

### M05-03: インライン編集

**`editListCell(paneId, rowIdx, colIdx)`:**

前提チェック:
- `col.readonly` or `col.key === 'id'` → 編集不可
- `col.type === 'checkbox'` → 編集不可（ワンクリックで切替）

**CELL_EDITING フラグ:**
- 編集開始時に `CELL_EDITING = true`
- 編集完了時に `CELL_EDITING = false`
- このフラグがtrueの間は `selectListRow()` と `setActiveCell()` を無効化（クリック誤動作防止）

**型別の編集UI:**

| 型 | 編集時のUI | 確定イベント |
|----|-----------|------------|
| `select` | `<select>` ドロップダウン | `onchange` / `onblur` → `finishEditCell()` |
| `date` | カスタム日付ピッカー（M05-04参照） | カレンダー日付クリック or OK押下 |
| `text` / `number` | `<input type="text">` | `onblur` → `finishEditCell()`、Enter → `this.blur()`、Esc → 元値復元+blur |

**`finishEditCell(paneId, rowIdx, colKey, val)`:**
1. `CELL_EDITING = false`
2. 値が変更された場合のみ `data[rowIdx][colKey] = val` + `markDirty()`
3. セル表示の部分更新（innerHTML差替、全テーブル再描画はしない）
4. 選択中の行であればペインリンクをトリガー（詳細ペインの更新）

---

### M05-04: カスタム日付ピッカー

**構成:**
- テキスト入力欄（`YYYY-MM-DD`形式）
- OKリンク
- カレンダーパネル（absolute位置、z-index:100）

**カレンダーパネル (`buildCalHTML`):**
- ヘッダー: `◀ YYYY/MM ▶` （月送りボタン）
- 曜日行: 日 月 火 水 木 金 土（日=赤、土=青）
- 日付セル: 各日をクリック → `dpSelect()` → `finishEditCell()`
- 今日ボタン: 今日の日付を即セット
- 今日の背景色: #bbdefb
- 日曜の背景: #fff5f5、土曜の背景: #f0f4ff
- ホバー効果: `onmouseover` → background:#ffe082、`onmouseout` → 元色

---

### M05-05: フィルタ行

**`onFilterInput(paneId, colKey, val)`:**
- `LIST_FILTERS[paneId][colKey] = val.trim().toLowerCase()`
- 該当ペインを再描画

**`applyFilters(pane, data)` のフィルタパイプライン:**
1. テキストフィルタ: 各列の `LIST_FILTERS` 値で部分一致フィルタ（大文字小文字無視）
2. 日付フィルタ: `pane.config.dateFilter.enabled` かつ `window._dateFilterActive` の場合、指定列の日付が `_dateFilterStart` 〜 `_dateFilterEnd` 範囲内のみ
3. フィルタビュー: `filterViewSelect` の選択値に基づくフィルタ条件を適用

---

### M05-06: ソート

**`sortList(paneId, colKey)`:**
- `LIST_SORT[paneId]` に `{ key, asc }` を保存
- 同じ列を再クリック → asc/desc切替
- 異なる列をクリック → asc固定
- ソートロジック:
  - 両方がパース可能な数値 → 数値比較
  - それ以外 → 文字列比較（`<` / `>`）
- ソート後、ペインを再描画

---

### M05-09: ハイライトルール

**`getHighlightClass(paneId, row, cols)`:**
1. `CONFIG.highlightRules.global` + `CONFIG.highlightRules.panes[paneId]` を結合
2. 各ルールについて:
   - `rule.enabled === false` → スキップ
   - `rule.condition.column` をカンマ区切りで分割し、各列名の値で条件評価
   - `evaluateCondition(val, condition)` で判定
   - `rule.statusCheck` がある場合、追加条件チェック
   - 一致 → `rule.style.className` をクラス文字列に追加

**`evaluateCondition(val, cond)` の演算子:**

| 演算子 | 説明 | 引数例 |
|--------|------|--------|
| `equals` | 完全一致 | value: "完了" |
| `notEquals` | 不一致 | value: "完了" |
| `contains` | 部分一致 | value: "田中" |
| `before` | 日付が指定値より前 | value: "TODAY" or "2025-12-31" |
| `after` | 日付が指定値より後 | value: "TODAY" |
| `within_days` | 今日からN日以内 | value: 3 |
| `isEmpty` | 空値 | - |
| `isNotEmpty` | 非空値 | - |
| `greaterThan` | 数値比較（大きい） | value: 50 |
| `lessThan` | 数値比較（小さい） | value: 50 |

**CSSクラス一覧:**

| クラス名 | 背景色 | 文字色 | 追加効果 |
|---------|-------|-------|---------|
| `hl-overdue` | #ffcdd2 | #b71c1c | font-weight:bold |
| `hl-due-soon` | #fff9c4 | - | - |
| `hl-complete` | #e8f5e9 | #9e9e9e | text-decoration:line-through |
| `hl-high-risk` | #ffe0b2 | - | font-weight:bold |
| `hl-critical` | #ffcdd2 | - | - |
| `hl-milestone` | #e3f2fd | - | - |

---

### M05-10: 行追加 / 削除 / 複製 / 挿入

**`addListRow()`:**
1. アクティブタブからlist/wbs型ペインを検索
2. ID自動採番: 既存データの最大ID番号+1 → `"ID-001"` 形式 (`padZero3`)
3. デフォルト値: date列→今日, select列→options[0], progress→"0", level→"0"
4. `data.push(row)` → `markDirty()` → 再描画 → 追加行を選択

**`deleteListRow()`:**
1. `SELECTED_ROW` から選択行インデックス取得
2. confirmダイアログ → OK → `data.splice(idx, 1)` → `markDirty()` → 再描画

**`insertRow(paneId, idx, after)`:**
- `data.splice(idx + after, 0, row)` — after=0で上に、after=1で下に挿入

**`duplicateRow(paneId, idx)`:**
- 選択行のシャローコピー（`_`プレフィックスのメタプロパティ除外）
- 新IDを `uid()` で生成
- `data.splice(idx + 1, 0, dup)`

---

## Module-06: 詳細（Detail）ペイン

### M06-01: 詳細フォーム描画

**`renderDetailPane(pane, el)`:**

表示条件:
- `linkedPane` 設定あり → `SELECTED_ROW[linkedPaneId]` で選択行を取得
- 未選択 → 「リストから行を選択してください」ガイド表示
- データなし → 「データがありません」表示

**ヘッダー表示:** `"詳細 - " + (row.task_name || row.title || row.id || "")`

**フィールド描画ルール:**

| 条件 | 描画 |
|------|------|
| `readonly` / `auto` / `key==='id'` / `key==='wbs_no'` | disabled input（灰色背景） |
| `type === 'select'` | `<select>` + onchange → `updateDetailField()` |
| `type === 'checkbox'` | `<input type="checkbox">` + onchange |
| `type === 'date'` | `<input type="text" placeholder="YYYY-MM-DD">` + onchange/onblur |
| `key` が description/notes/mitigation/resolution/response | `<textarea>` + onblur |
| その他 | `<input type="text">` + onblur |

**`updateDetailField(paneId, rowIdx, colKey, val)`:**
- データ更新 → `markDirty()`
- 対応するリストペインのセルをDOM操作で部分更新（全テーブル再描画はしない）

---

## Module-07: エディタペイン

### M07-01: テキストエリアエディタ（DynamicApp）

**`renderEditorPane(pane, el)`:**

ファイル読み込みフロー:
1. `pane.config.linkedPane` 設定あり → リストの選択行から `fileColumn`（デフォルト: "file_path"）取得
2. `pane.config.baseDir` + ファイル名 → 絶対パス解決
3. `ensureDir()` でディレクトリ保証
4. `readFileSync()` でファイル読み込み
5. ファイルが存在しない → `getMinutesTemplate()` でテンプレート生成 → ファイル書き込み

**`saveEditorContent(pane)`:**
- `EDITOR_CONTENTS[pane.id]` のテキストを `EDITOR_CONTENTS[pane.id + '_path']` に書き込み

---

### M07-02: リッチエディタ（cockpit_v9）

**`contenteditable` DIV:**
- `#editor-area` — `contenteditable="true"`
- イベント: `oninput`, `onkeydown`, `onpaste`, `ondrop`, `ondragover`
- フォント: Consolas/MS Gothic/Meiryo monospace, 16px, line-height:1.9

**DOMからrawTextへの変換 (`getEditorPlain`):**
- `collectLines(el, lines)` で再帰的にDOMノードを走査
- TABLE（`.editor-inline-table`）→ `||col1|col2|col3` 記法に変換
- DIV/P → innerTextを行として抽出。ネストDIVは再帰処理
- BR → 空行
- SPAN → テキスト抽出
- テキストノード → 空白のみでなければ行として追加

**rawTextからDOMへの変換 (`setEditorContent`):**
- `||` 始まりの行 → `<table class="editor-inline-table">` に変換（3列固定）
- それ以外 → `<div>` + `highlightLine()` でシンタックスハイライト
- `#` 行 → `.ed-tag`（#e8b84b, bold）
- `##` 行 → `.ed-field`（#5b9cf5, bold）
- `★` を含む行 → `.ed-star`（#e8b84b, bold, 背景付き）

---

### M07-03: マークダウン風パーサー（`parseText`）

**入力:** rawText文字列
**出力:**
```
{
  type: "議事録",      // #行（最初のみ）
  fields: [            // ##行とその値
    { key: "会議名", value: "定例会議" },
    { key: "日時", value: "2025-01-15" }
  ],
  tables: [            // ||行
    { fk: "Todo", cells: ["要件書修正", "田中", "1/20"] }
  ],
  freeTexts: ["..."],  // type宣言後のフィールド外テキスト
  stars: [             // ★を含む行
    { line: "★重要: フォローアップ必要", fieldKey: "Todo" }
  ]
}
```

**パースルール:**
1. `#xxx`（`##`でない）→ type設定（最初の1回のみ）
2. `##xxx` → 新フィールド開始。前フィールドのvalueをフラッシュ
3. `||xxx|yyy|zzz` → tables配列に追加（fk = 現在のフィールドキー）
4. `★`を含む行 → stars配列に追加
5. それ以外 → 現在フィールドのvalue行に追加 or freeTexts

---

### M07-04: インラインテーブル

**`||`入力検知 (`handleEditorInput`):**
1. `sel.focusNode`（テキストノード）の末尾が`||`で終わるか検知
2. `||`を削除
3. `<table class="editor-inline-table"><tbody><tr><td><br></td><td><br></td><td><br></td></tr></tbody></table>` を生成
4. 現在のDIVの直後にテーブルを挿入
5. テーブルの下にクリック用空行`<div><br></div>`を挿入
6. 最初のTDにカーソルを移動

**テーブル内キー操作 (`handleEditorKeydown`):**

| キー | 動作 |
|------|------|
| Tab | 次のセルへ移動。最終セル（最終行の最終列）なら新しい行を自動追加 |
| Shift+Tab | 前のセルへ移動 |
| Backspace（セル空+行空+テーブル全空） | テーブル全体を削除 |
| Backspace（行の全セル空） | 行のみ削除（テーブルに複数行ある場合） |
| Delete（テーブル外で直後がテーブル） | 全セル空ならテーブル削除 |

**セル移動のカーソル制御:**
- `document.createRange()` + `range.selectNodeContents(nextTd)` + `range.collapse(false)` で末尾にカーソル配置

---

### M07-05: テンプレート挿入

**テンプレート一覧（15種）:**

| カテゴリ | テンプレート名 |
|---------|--------------|
| 業務 | 議事録, タスク, ディレクション, 折衝記録 |
| フレームワーク | 3C, SCR, SWOT, MECE, 仮説思考, イシューツリー, 4P, 5Forces |
| 分析 | ギャップ分析, シナリオ分析, ピラミッド, KPIツリー, 意思決定 |

**`insertTpl(name)`:**
- 現在のエディタ内容が空でなければ末尾に `\n` + テンプレートを追加
- 空なら テンプレートのみセット
- `setEditorContent()` → `syncToView()` → `focus()` → dirty=true → autoSave予約

---

### M07-07: ペースト処理

**`handlePaste(e)`:**
1. `clipboardData.items` を走査 → 画像がある場合 → `insertMediaBlob()` で挿入
2. テキストの場合 → `e.preventDefault()` → `clipboardData.getData('text/plain')` を取得
3. HTMLは完全に排除（プレーンテキストのみ強制）
4. 改行を含む場合:
   - 1行目: テキストノードとして現在行に挿入
   - 2行目以降: 新しい`<div>`要素として追加
5. カーソルを挿入テキスト末尾に移動

**HTA移植時の注意:**
- `clipboardData.items` はIE9未対応。`window.clipboardData.getData('Text')` で代替
- 画像ペーストはIE9では非対応。VBScript経由のファイル保存に置換

---

### M07-09: 構造化ビュー（双方向同期）

**左→右同期 (`syncToView`):**
- `getEditorPlain()` → `parseText()` → `renderSection()` でHTML生成
- `#sv` にinnerHTML設定

**右→左同期 (`syncToEditor`):**
1. `#sv .sv-editable` セルから `data-field`, `data-row`, `data-col` でデータ収集
2. 元のrawTextを行単位でパースし直し、編集されたフィールド/テーブルセルだけ差替
3. 再構築テキストを `window._syncedRawText` に保持
4. 左ペインの再描画はしない（カーソル位置保持のため）
5. 次回の自動保存/手動保存時に `_syncedRawText` を優先使用

---

### M07-10: 5秒オートセーブ

**`scheduleAutoSave()`:**
- 前回のタイマーをクリア → `setTimeout(callback, 5000)` で5秒後に実行
- callback:
  1. `editorDirty === false` → 何もしない
  2. `_syncedRawText` があればそれを使用、なければ `getEditorPlain()`
  3. 既存レコード（`editId`あり）→ recordsを更新して`writeDB`
  4. 新規レコード → `appendDB` で新規作成、editIdをセット
  5. `editorDirty = false`
  6. オートセーブインジケータを2秒間表示

---

### M07-11: タスク自動起票

**DynamicApp版 (`extractTasksFromMinutes`):**
- パターン: `- [ ] 担当者: タスク名（期限）` or `- [x] ...`
- 正規表現: `/^-\s*\[[ x]\]\s*(.+)$/i`
- 担当者/タスク名/期限の分解: `/^([^:：]+)[：:]\s*(.+?)(?:[（(](.+?)[）)])?$/`
- 期限パース: `/(\d{1,2})[\/\-](\d{1,2})/` → 当年の `YYYY-MM-DD` に変換
- 重複チェック: `task_name` が既存データと一致 → スキップ
- WBS `wbs_tasks.csv` に追加。notes: "議事録から自動起票"

**cockpit_v9版 (`extractTasks`):**
- Todo/ネクストアクション/アクションアイテムフィールドの `||` テーブル行 → task_log追加
- `★`マーカー行 → task_log追加（content: "★マーカーから自動抽出"）
- 重複チェック: `title` + `source_record` の組み合わせ

---

## Module-08: WBSペイン

### M08-01: WBS番号自動採番

**`assignWbsNumbers(data)`:**
- `counters = [0]` — レベルごとのカウンター配列
- 各行について:
  - `level` に応じて counters を伸縮
  - `counters[level]++`
  - WBS番号 = counters[0].counters[1]. ... .counters[level]
  - 例: level0→"1", level1→"1.1", level0→"2"

### M08-03: インデント/アウトデント

**`wbsIndent(paneId, idx)`:**
- `data[idx].level = String(lv + 1)`
- 最初の行（idx===0）はインデント不可
- `markDirty()` → ペイン再描画

**`wbsOutdent(paneId, idx)`:**
- `data[idx].level = String(lv - 1)`
- level 0以下にはならない
- `markDirty()` → ペイン再描画

### M08-04: 子タスク/同レベル追加

**`addWbsChild(paneId, idx)`:**
- 親のlevel+1で idx+1 の位置に挿入
- デフォルト値: status="未着手", start_date=今日

**`addWbsSibling(paneId, idx)`:**
- 同じlevelで、子孫行をスキップした位置に挿入
  - `while (insertAt < data.length && level of data[insertAt] > level) insertAt++`

---

## Module-09: ガントチャートペイン

### M09-01: ガントチャート描画

**日付範囲計算:**
1. 全データからstart_date/end_dateの最小/最大を算出
2. 前後にパディング（開始-7日、終了+14日）
3. スケール別の調整: week→月曜始まり、month→1日始まり
4. 最低30日分を確保

**dayWidth:**
- day: 60px
- week: 28px
- month: 16px

**月ヘッダー:**
- 各月の開始位置・幅を計算して `position:absolute` で配置
- 表示: `"1月 2025"` 形式

**日ヘッダー:**
- 各日を `position:absolute` で配置
- dayWidth >= 28: 日付番号表示
- dayWidth < 28: 1日 or 5の倍数日のみ表示
- 土日: `.weekend` → background:#fff3e0
- 今日: `.today` → background:#bbdefb, bold

**バー描画:**
- 位置: `left = dateDiff(padStart, startDate) * dayWidth`
- 幅: `(dateDiff(startDate, endDate) + 1) * dayWidth`
- 色: `barColors[status]` — 未着手:#B0BEC5, 進行中:#42A5F5, 完了:#66BB6A, 保留:#FFA726
- 進捗表示: 0< progress <100 → 内部にoverlay（rgba(0,0,0,0.15), width:progress%）
- ツールチップ: `title="タスク名 (progress%)"` 
- マイルストーン: `milestone=true` → 三角形マーカー（border-bottom: #ff6f00）

---

## Module-10: ダッシュボード

### M10-01: DynamicApp 統計カード

**データ集約:**
- `findDataByFile('wbs_tasks.csv')` で全タスクデータ取得
- 各CSVファイルからデータを読み込み（未ロードの場合はファイルから読み込み）

**カード一覧:**

| カード名 | 計算式 |
|---------|--------|
| プロジェクト進捗 | `sum(progress) / totalTasks`（%）+ 完了/総数 |
| 期限超過タスク | status≠完了 && end_date < today |
| 未回答の依頼事項 | requests where status=依頼中 |
| アクティブリスク | risks where status∈{監視中, 対策中} |
| オープン課題 | issues where status≠完了。重大件数も表示 |

**今週タスクテーブル:**
- `getWeekStart(today)` 〜 +6日の範囲でend_dateフィルタ
- 列: タスク名, 担当者, 終了日, 進捗, 状態

**期限超過タスクテーブル:**
- status≠完了 && end_date < today
- `.hl-overdue` クラス付与

---

## Module-11: 設定ペイン

### M11-02: タブ管理

**新規タブ追加 (`addNewTab`):**

フォーム入力:
- タブ名（必須）
- ペインタイプ: list / editor / wbs / gantt / detail / menu / dashboard
- データファイル名
- レイアウト: horizontal / vertical

生成されるタブ構造:
- `id: "tab_" + uid()`
- `order: maxOrder + 1`
- `visible: true`
- 1ペイン100%
- list型の場合はデフォルト列定義を自動追加（id, title, status, notes）

### M11-03: 列定義管理

**列の追加 (`addColumn`):**
- フォーム入力: キー名, ラベル, 型（text/number/date/select/checkbox/radio）, 幅(px), 選択肢(カンマ区切り)

**列の編集 (`editColumn`):**
- フォーム入力: ラベル, 型, 幅, 選択肢, 読取専用チェック

**列の削除 (`removeColumn`):**
- confirmダイアログ → `columns.splice(colIdx, 1)` → 保存

---

## Module-13: 業務ドメイン機能（cockpit_v9）

### M13-01: 情報収集（Inbox）

**データ構造:**
```
{
  id: "II12345678",
  datetime: "2025-01-15T10:30:00.000Z",
  source_type: "URL" | "file",
  title: "記事タイトル",
  url: "https://...",
  url_summary: "要約テキスト",
  project: "P001",
  category: "業界動向" | "競合動向" | "規制・政策" | "技術トレンド" | "M&A情報" | "調査レポート" | "その他",
  tags: "AI;半導体",
  memo: "メモテキスト",
  status: "未読"
}
```

**UI構成（左右分割）:**
- 左ペイン（420px）: URL入力, ファイルドロップゾーン, 案件ID/カテゴリ/タグ/メモ入力, 保存ボタン
- 右ペイン: 保存済みアイテム一覧（検索付き）

**ファイルドロップ:**
- `ondragover` → `.over` クラス付与
- `ondrop` → `handleFiles(files)` → ファイル名をタイトルにセット

### M13-03: タスク管理

**データ構造:**
```
{
  id: "TK12345678",
  datetime: "2025-01-15T10:30:00.000Z",
  title: "タスク名",
  due: "2025-01-31",
  owner: "田中",
  project: "",
  content: "",
  status: "未着手" | "進行中" | "完了" | "中止",
  progress: "1/15 14:00:::進捗コメント|||1/16 10:00:::次のコメント",
  source_record: "RC12345678"
}
```

**進捗ログ形式:**
- `"日時:::コメント"` を `"|||"` で連結
- 表示時: `split('|||')` → 各エントリを `split(':::')` で日時とコメントに分離

**フィルタ:**
- 未完了（status≠完了 && status≠中止）
- 全て
- 完了のみ

**ソート:**
- 日付順（datetime降順）
- 期限順（due昇順、未設定は最後）

### M13-05: 案件管理

**データ構造:**
```
{
  id: "P12345678",
  name: "案件名",
  client: "顧客名",
  phase: "提案前" | ... | "完了",
  priority: "S" | "A" | "B" | "C",
  type: "戦略策定" | "業務改革" | "M&A支援" | "マーケティング" | "その他",
  fee: "1500" (万円),
  start: "2025-01-01",
  status: "進行中" | "完了"
}
```

**ボード表示:**
- フェーズごとにグルーピング（3列レイアウト）
- 優先度に応じたカード左ボーダー色: S=赤, A=金, B=青

### M13-06: 顧客管理

**データ構造:**
```
{
  id: "C12345678",
  name: "顧客名",
  division: "部署",
  key_person: "キーパーソン名",
  key_title: "役職",
  email: "email@example.com",
  tel: "03-xxxx-xxxx",
  industry: "製造" | "金融" | "商社" | "小売" | "IT" | "その他",
  relation_score: "1"〜"5",
  pain_summary: "課題テキスト",
  memo: "メモ"
}
```

**詳細表示:**
- KVテーブル: キーパーソン/役職/業種/Email/TEL/関係性/課題
- 関連案件: `project_master` から `client === c.name` でフィルタ
- 関連議事録: `records` から `type==='議事録' && rawText.indexOf(c.name) >= 0` で検索（直近5件）

### M13-08: 全文検索

**`gSearch()`:**
- `records` のrawTextを全件走査 → キーワード部分一致
- `intelligence_inbox` のtitleを全件走査 → 部分一致
- 結果をリスト表示（ソース/日時/テキスト先頭200文字）

---

## Module-14: キーボード・操作

### 全キーバインド一覧

| キー | コンテキスト | 動作 |
|------|------------|------|
| Ctrl+S | 全画面 | `saveAllData()` — 全データ一括保存 |
| Ctrl+Tab | 全画面 | 次のタブへ切替（可視タブのみ循環） |
| Ctrl+Shift+Tab | 全画面 | 前のタブへ切替 |
| Ctrl+1〜9 | 全画面 | N番目のタブに直接切替 |
| ← → ↑ ↓ | リスト/WBS | アクティブセルの移動（input/textarea/selectフォーカス中は無効） |
| Enter | リスト | アクティブセルの編集開始 |
| F2 | リスト | アクティブセルの編集開始 |
| Delete | リスト | 選択行の削除（input/textareaフォーカス中は無効） |
| Tab | WBS | 選択行のインデント |
| Shift+Tab | WBS | 選択行のアウトデント |
| Esc | 全画面 | コンテキストメニュー非表示 + モーダル閉じ |
| Ctrl+Enter | エディタ | レコード保存（cockpit_v9） |

**イベント処理のIE対応:**
- `e = e || window.event`
- `e.keyCode` でキー判定（`e.key` はIE9未対応）
- `e.preventDefault()` の存在チェック → なければ `e.returnValue = false` + `return false`
- `event.stopPropagation()` → IE8以下は `event.cancelBubble = true`

---

## デフォルトタブ構成（app_config.json）

| order | タブID | タブ名 | ペイン構成 | dataFile |
|-------|--------|--------|-----------|----------|
| 1 | tab_dashboard | ダッシュボード | dashboard×1 (100%) | - |
| 2 | tab_wbs | WBS / スケジュール | wbs (40%) + gantt (60%) | wbs_tasks.csv |
| 3 | tab_tasks | タスク一覧 | list (55%) + detail (45%) | wbs_tasks.csv |
| 4 | tab_minutes | 議事録 | list (40%) + editor (60%) | minutes.csv |
| 5 | tab_requests | 依頼事項 | list (55%) + detail (45%) | requests.csv |
| 6 | tab_risks | リスク管理 | list (55%) + detail (45%) | risks.csv |
| 7 | tab_issues | 課題管理 | list (55%) + detail (45%) | issues.csv |
| 99 | tab_settings | ⚙ 設定 | settings×1 (100%) | - |

---

## 全グローバル変数一覧

| 変数名 | 型 | 説明 |
|--------|-----|------|
| `APP_DIR` | string | HTAファイルの格納ディレクトリ（末尾\付き） |
| `CONFIG` | object | app_config.jsonの内容 |
| `DATA_STORE` | object | paneId → Array のデータ保持 |
| `DIRTY_FLAGS` | object | paneId → boolean の変更フラグ |
| `ACTIVE_TAB` | string | 現在アクティブなタブID |
| `ACTIVE_CELL` | object/null | {paneId, row, col} アクティブセル |
| `SELECTED_ROW` | object | paneId → rowIndex の選択行 |
| `DATE_SCALE` | string | "month" / "week" / "day" |
| `DATE_CURSOR` | Date | 日付ナビゲーションの現在位置 |
| `VBS_CALLBACK` | function/null | VBScript非同期コールバック（未使用） |
| `UNDO_STACKS` | object | paneId → [snapshots]（未実装） |
| `REDO_STACKS` | object | paneId → [snapshots]（未実装） |
| `EDITOR_CONTENTS` | object | paneId → テキスト内容 |
| `EDIT_MODE` | boolean | 行移動モード |
| `WBS_FOLDED` | object | paneId → {rowIdx: boolean} 折り畳み状態 |
| `SYNC_SCROLLING` | boolean | スクロール同期のループ防止フラグ |
| `CELL_EDITING` | boolean | セル編集中フラグ |
| `LIST_FILTERS` | object | paneId → {colKey: filterText} テキストフィルタ |
| `LIST_SORT` | object | paneId → {key, asc} ソート状態 |
| `editorDirty` | boolean | cockpit_v9エディタの変更フラグ |
| `autoSaveTimer` | number/null | cockpit_v9オートセーブのsetTimeout ID |
| `curMtg` | string/null | cockpit_v9 選択中の議事録ID |
| `curTk` | string/null | cockpit_v9 選択中のタスクID |
| `curDl` | string/null | cockpit_v9 選択中のディレクションID |
| `curKb` | string/null | cockpit_v9 選択中のナレッジID |
| `pendingFile` | object/null | cockpit_v9 ドロップ待ちファイル情報 |

---

## ユーティリティ関数一覧

| 関数名 | 引数 | 返値 | 説明 |
|--------|------|------|------|
| `uid()` | - | string | `"id_" + timestamp + "_" + random(0-9999)` |
| `escapeHtml(s)` | string | string | `& < > "` をエスケープ |
| `padZero(n)` | number | string | 2桁ゼロパディング |
| `padZero3(n)` | number | string | 3桁ゼロパディング |
| `formatDate(d)` | Date/string | string | `YYYY-MM-DD` 形式 |
| `parseDate(s)` | string | Date/null | `YYYY-M-D` パース |
| `todayStr()` | - | string | 今日の `YYYY-MM-DD` |
| `dateDiffDays(d1, d2)` | Date, Date | number | 日数差（d2-d1, 四捨五入） |
| `getWeekStart(d)` | Date | Date | 週の開始日（月曜） |
| `deepCopy(obj)` | object | object | JSON.stringify→parseによるディープコピー |
| `safeParseJSON(text)` | string | object/null | JSON.parse + evalフォールバック（要削除） |
| `genId(prefix)` | string | string | `prefix + timestamp下8桁` |
| `esc(s)` | string | string | cockpit_v9版escapeHtml |
| `nowISO()` | - | string | `new Date().toISOString()` |
| `todayISO()` | - | string | ISO文字列の先頭10文字 |
| `fmtDT(iso)` | string | string | `M/D HH:MM` 形式 |
| `fmtD(iso)` | string | string | `M/D` 形式 |
| `clipCopy(t)` | string | - | クリップボードにコピー |

---

## 追補: Librarian HTA 実装で発見されたIE9/HTA固有バグ一覧

### BUG-01: `<tbody>.innerHTML` への代入がエラー（★致命的）

**発生箇所:** `renderTable()` 内 `tbody.innerHTML = h;`

**エラー内容:** IE9/HTAでは `<table>`, `<thead>`, `<tbody>`, `<tfoot>`, `<tr>`, `<col>`, `<colgroup>` のinnerHTMLプロパティは**読み取り専用**。代入すると「Unknown runtime error」が発生する。

**原因:** IE独自のHTML DOMパーサー制約。テーブル関連要素は特別扱いされており、innerHTML経由のHTML挿入が禁止されている。これはIE6～IE9の全バージョンに共通する仕様。

**修正方法:**
テーブル全体（`<table>` タグごと）を包む親 `<div>` の `innerHTML` を丸ごと差し替える。`<div>` 要素のinnerHTMLは正常に動作するため、テーブルHTML文字列全体を `<div>` に書き込む方式に変更する。

```javascript
// NG: IE9でエラー
document.getElementById('tableBody').innerHTML = rowsHtml;

// OK: 親divのinnerHTMLを丸ごと差替
document.getElementById('table-wrap').innerHTML =
  '<table class="data-table"><thead>...<thead><tbody>' + rowsHtml + '</tbody></table>';
```

**影響範囲:** DynamicApp.hta の `renderListPane()`, `renderWbsPane()` も同様パターンを使用しているが、そちらはペイン全体（`el.innerHTML = h`）を差し替えているため回避されていた。今後テーブル要素のinnerHTMLに直接代入するコードを書かないこと。

---

### BUG-02: `FileReader` API 未対応

**発生箇所:** CSVドロップ処理 `csvDzDrop()` 内 `new FileReader()`

**原因:** `FileReader` はHTML5 File API（IE10以降）。IE9/HTAでは `window.FileReader` は未定義。

**修正方法:** try-catchの外側で存在チェックを行い、HTA環境ではVBScript経由のファイル読み込みに分岐する。CSVファイルドロップの代わりに、VBScriptの `UserAccounts.CommonDialog` またはパス直接入力によるファイル指定とする。

---

### BUG-03: `saveData()` のVBS文字列リテラル埋め込みが脆弱

**発生箇所:** `saveData()` 内 `jsonEscapeVBS(json)` を使ったVBSコード文字列構築

**原因:** JSONデータを `""` エスケープしてVBScript文字列リテラルに埋め込む方式は、以下のケースで破綻する:
- JSON内にシングルクォートが含まれる場合
- JSON文字列が長大（数千文字以上）な場合、VBScript文字列リテラルの行継続制限に抵触
- JSON内に改行を含むフリーテキスト（summaryフィールド等）がある場合

**修正方法:** DynamicApp.htaと同じ**hidden inputバッファパターン**を使用する:
1. JSONを `document.getElementById('_vbs_r').value` にセット
2. VBScript側で `document.getElementById("_vbs_r").value` を読み取り→ファイル書き込み
3. VBSコード文字列にデータ本体を含めない

```javascript
// NG: データをVBS文字列に埋め込み
var vbsCode = 'Call writeUTF8("path", "' + jsonEscapeVBS(json) + '")';

// OK: hidden inputバッファ経由
document.getElementById('_vbs_r').value = json;
window.execScript('Call writeUTF8("path", document.getElementById("_vbs_r").value)', 'VBScript');
```

---

### BUG-04: `e.dataTransfer.files` のHTAでの制約

**発生箇所:** ファイルドロップ処理 `dzDrop()`, `csvDzDrop()`

**原因:** HTA (IE9) のdrag-and-dropイベントでは `dataTransfer.files` プロパティが利用できない場合がある。特にExplorerからのファイルドロップ時、`getData('Text')` でファイルパスが取得できるケースと取得できないケースがある。

**修正方法:**
- ファイルドロップは「ベストエフォート」とし、失敗時にはファイルパス手入力またはVBScriptファイルダイアログにフォールバック
- CSVインポートはテキスト直接貼り付けを主要手段とし、ドロップは補助手段とする
- ドロップ処理全体をtry-catchで囲み、エラー時は「ファイル名を入力してください」ガイドを表示

---

### 総括: HTA開発での鉄則ルール

| ルール | 理由 |
|--------|------|
| テーブル要素（tbody/thead/tr等）のinnerHTMLに代入しない | IE全バージョンで読み取り専用 |
| FileReader / Blob / File API を使用しない | IE10以降のAPI |
| VBSコード文字列にデータ本体を埋め込まない | 特殊文字・長さ制限で破綻する |
| hidden inputをデータ受け渡しバッファとして使う | JS⇔VBS間の安全なデータ受渡パターン |
| ドラッグ&ドロップは補助手段として扱う | IE9のdataTransfer実装が不安定 |
| navigator.clipboard を使用しない | IE11以降のAPI。window.clipboardDataで代替 |
| CSS transition / animation を使用しない | IE9未対応 |
| `addEventListener` を使用しない | `onclick`属性 or `element.onclick = fn` で代替 |
| `<select>` のinnerHTMLに直接代入しない | IE9で選択肢が表示されない場合がある。親divごと再構築する |

---

## Module-15: Librarian（リファレンス・リンク管理アプリ）

### M15-00: 概要

ニュース記事・ガイドライン等のリンクやファイルを一元管理するライブラリアンアプリ。HTAとして単体動作する。

### M15-01: データ構造

**保存先:** `librarian_data.json`（HTAと同階層、UTF-8）

**1エントリの構造:**

| フィールド | キー | 型 | 説明 |
|-----------|------|-----|------|
| ID | `id` | string | `"L" + timestamp + "_" + random` で自動生成 |
| カテゴリ | `category` | string | `"news"` or `"guide"` |
| リンク | `link` | string | URL文字列 |
| タイトル | `title` | string | エントリの見出し |
| ソース元 | `source` | string | 情報源（例: Nikkei, Internal, Gov） |
| ファイル名 | `filename` | string | 添付ファイルの表示名 |
| ファイルパス | `filepath` | string | ファイルのフルパス（クリック時にWScript.Shell.Runで開く） |
| タグ | `tags` | string | カンマ or セミコロン区切りのタグ文字列 |
| 概要 | `summary` | string | 内容の要約・メモ |
| 作成日 | `created` | string | `YYYY-MM-DD` 形式 |

### M15-02: テーブル表示列

| # | 列名 | 幅 | 内容 |
|---|------|-----|------|
| 1 | (checkbox) | 28px | 全選択/個別選択チェックボックス |
| 2 | Type | 55px | NEWS（青バッジ）/ GUIDE（緑バッジ） |
| 3 | Title | 200px | タイトル文字列 |
| 4 | Summary | 200px | 概要（60文字で切詰め、全文はツールチップ表示） |
| 5 | Link | 170px | URL（35文字で切詰め、クリックで外部ブラウザ起動） |
| 6 | Source | 100px | ソース元 |
| 7 | Document | 150px | ファイル種別アイコン + ファイル名（クリックでファイルを開く） |
| 8 | Tags | 150px | タグをバッジ表示（複数対応） |
| 9 | Date | 80px | `M/D` 形式の作成日 |

### M15-02a: ファイル種別アイコン仕様

拡張子に基づく色分けバッジを表示し、クリックで `WScript.Shell.Run` によりファイルを開く。

| アイコン | 対象拡張子 | 色 |
|---------|-----------|-----|
| DOC | .doc, .docx, .docm | 青 (#4285f4) |
| XLS | .xls, .xlsx, .xlsm | 緑 (#34a853) |
| CSV | .csv | 緑 (#34a853) |
| PPT | .ppt, .pptx, .pptm | 橙 (#ea8600) |
| PDF | .pdf | 赤 (#ea4335) |
| IMG | .png, .jpg, .jpeg, .gif, .bmp | 紫 (#a064f0) |
| TXT | .txt, .md, .log | 灰 (#808ba0) |
| MAIL | .msg, .eml | 灰 |
| ZIP | .zip, .rar, .7z | 灰 |
| (拡張子) | その他 | 灰 (#6b7190) |

**クリック動作:** `openFile(filepath)` → `new ActiveXObject('WScript.Shell').Run('"' + filepath + '"')` で関連アプリケーション起動

### M15-03: カテゴリ分離表示

- サイドバーに動的なナビゲーションボタン: All + 設定で定義されたType数分
- 各ボタンに件数バッジ表示
- Type追加・削除時にサイドバーが自動再構築される

### M15-03a: 設定画面 — Type（カテゴリ）管理

ヘッダーの「Settings」ボタンでメインエリアに設定画面をオーバーレイ表示。

**Type定義のデータ構造:**
```
{ id: "news", label: "News", color: "#5b9cf5" }
```

**CRUD操作:**

| 操作 | 入力 | 処理 |
|------|------|------|
| 追加 | ID, Label, Color(hex) | 重複IDチェック → `config.types.push()` → 保存 → サイドバー再描画 |
| 編集 | ID, Label, Color | ID変更時は全entryのcategoryも追従変更 → 保存 |
| 削除 | confirmダイアログ | `config.types.splice()` → 保存 |

**色からバッジ生成:**
- `hexToBg(color, 0.12)` → 背景のrgba
- `hexToBg(color, 0.25)` → ボーダーのrgba
- テーブル表示時に動的に `style` 属性でバッジ色を適用

### M15-03b: 設定画面 — Column（フィールド）管理

**Column定義のデータ構造:**
```
{ key: "title", label: "Title", type: "text", width: 200, show: true, sortable: true, options: "" }
```

**対応するtype値:**

| type値 | フォーム表示 | テーブル表示 |
|--------|------------|------------|
| `text` | `<input type="text">` | テキスト表示 |
| `textarea` | `<textarea>` | 60文字切詰め+ツールチップ |
| `url` | `<input>` (placeholder: https://...) | 青リンク（クリックで外部ブラウザ） |
| `file` | `<input readonly>` + Browseボタン | ファイルアイコン + ファイル名（クリックで開く） |
| `tags` | `<input>` (カンマ区切り) | バッジ表示 |
| `select` | `<select>` (optionsからドロップダウン生成) | テキスト表示 |
| `date` | `<input type="text">` | テキスト表示 |

**CRUD操作:**

| 操作 | 入力 | 処理 |
|------|------|------|
| 追加 | Key, Label, Type, Width | 重複Keyチェック → `config.columns.push()` → 保存 → テーブル再描画 |
| 編集 | Key, Label, Type, Width, Options, Show, Sortable | Key変更時は全entryの該当フィールド名も追従変更 → 保存 |
| 削除 | confirmダイアログ | `config.columns.splice()` → 保存 |
| 並べ替え | ▲▼ボタン | 隣接カラムとswap → 保存 → テーブル列順に即時反映 |

**Show/Sortフラグ:**
- `show: false` → テーブルに列を表示しない（データは保持）
- `sortable: true` → ヘッダークリックでソート可能

**Options（select型専用）:**
- カンマ区切りの文字列（例: `"高,中,低"`）
- フォーム表示時に`<select>`のoption要素として展開

### M15-03c: 設定のデータ保存形式

`librarian_data.json` に config と entries を統合保存:

```json
{
  "config": {
    "configVer": 1,
    "types": [ ... ],
    "columns": [ ... ]
  },
  "entries": [ ... ]
}
```

設定変更は即時保存。`Reset Config` でtypes/columnsをデフォルトに戻す（entriesは保持）。

### M15-04: 入力方法

| 方法 | 動作 |
|------|------|
| フォーム手入力 | サイドパネルのフォームで全フィールド入力 |
| リンクドロップ | ブラウザからURLをウィンドウ上にドロップ → フォーム自動起動 + Link欄セット + ドメインをSource欄に自動抽出 |
| ファイルドロップ | エクスプローラからファイルをドロップゾーンにドロップ → Filename欄セット |
| ファイル参照 | VBScript `UserAccounts.CommonDialog` でファイル選択 |

### M15-05: CSV取り込み

**フロー（ファイル選択主体）:**
1. サイドバーまたは設定画面の「CSV Import」ボタンでモーダルを開く
2. **Step 1:** Type選択ドロップダウンで取り込み先カテゴリを選択（設定で定義した全Typeが表示）
3. **Step 2:** 「Browse .csv」ボタンでファイル選択ダイアログ（VBScript `UserAccounts.CommonDialog`、フィルタ: CSV/TXT/All）
4. 選択したCSVファイルをVBScript経由でUTF-8読み込み → テキストエリアにプレビュー表示
5. ファイル名・行数を表示（例: `File: data.csv (25 rows)`）
6. 「Import」ボタンで取り込み実行

**代替手段:** テキストエリアに直接CSV貼付けも可能

**Type選択のIE9対応:**
- `<select>` 要素のinnerHTMLへの代入はIE9で不安定なため、親`<div>`のinnerHTMLで`<select>`タグごと再構築する方式を採用

**CSVフォーマット:** `title,link,source,tags,summary`（5列、カラム数は可変）
- 1行目がヘッダー（"title" or "link" を含む）なら自動スキップ
- ヘッダーのカラム名を `config.columns` のkey/labelと照合して自動マッピング
- ヘッダーなしの場合はデフォルト順（title, link, source, tags, summary）で割当
- title も link も空の行はスキップ
- 取り込み完了時に件数とType名をトーストで表示（例: `12 items imported as [news]`）

### M15-06: 検索・フィルタ

| 機能 | 対象 |
|------|------|
| テキスト検索 | title, tags, source, link, summary, filename を横断部分一致 |
| カテゴリフィルタ | news / guide |
| タグフィルタ | サイドバーのタグボタンクリックでトグル（上位20タグ表示） |
| ソート | 日付新→旧 / 旧→新 / Title A-Z / Source A-Z |

### M15-07: CSV出力

- 現在のフィルタ結果をCSVファイルとしてHTAフォルダに出力
- ファイル名: `librarian_export_YYYYMMDD.csv`
- 出力列: category, title, link, source, filename, tags, summary, created

### M15-08: キーボードショートカット

| キー | 動作 |
|------|------|
| Ctrl+S | データ保存 |
| Ctrl+N | 新規エントリフォーム表示 |
| Esc | フォーム/モーダルを閉じる |
| Delete | 選択行を削除 |

