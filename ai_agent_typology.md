# AIエージェントの類型整理（OWASP 5要素ベース）

## 1. 5つの基本要素

OWASP「AI Agent Security Cheat Sheet」では、AIエージェントを次の5つの能力を持つ自律システムと定義している。

> "AI agents are autonomous systems powered by LLMs that can **reason, plan, use tools, maintain memory, and take actions** to accomplish goals."

| 略号 | 要素 | 意味 |
|------|------|------|
| **R** | Reason | 推論。問題を理解し思考する能力 |
| **P** | Plan | 計画立案。タスクを分解し手順を組む能力 |
| **T** | Tool (Function Call) | 外部APIや関数を呼び出す能力 |
| **M** | Memory | セッションや長期の状態・履歴を保持する能力 |
| **A** | Action | 実世界に副作用を出す実行能力 |

> TとAは隣接する。Tは「呼ぶ」、Aは「世界に作用する」と区別する。情報取得APIはT寄り、送信・購入・削除・ファイル変更はA寄り。

---

## 2. 判定フラグの定義

各組み合わせを2軸で判定する。

| フラグ | 意味 | 条件 |
|-------|------|------|
| **AI?** | LLMを使うAIか | Rを含む |
| **Agent?** | AIエージェントか | ⭕=標準派以上で確定 / △=緩い派で該当(R+T+A型の最小構成) / ❌=エージェントではない |

---

## 3. 32通りの組み合わせ × フラグ × 具体例

### 0個 (1通り)
| 構成 | AI? | Agent? | 具体例 |
|------|:---:|:------:|--------|
| なし | ❌ | ❌ | 電卓アプリ、静的Webサイト、ATM、自販機 |

### 1個 (5通り)
| 構成 | AI? | Agent? | 具体例 |
|------|:---:|:------:|--------|
| R | ⭕ | ❌ | GPT-3初期APIの単発QA、「フランスの首都は?」型応答 |
| P | ❌ | ❌ | Microsoft Project、ガントチャートツール、計画書テンプレ |
| T | ❌ | ❌ | cURL、Postman、固定APIラッパー |
| M | ❌ | ❌ | Notion、Evernote、Google Keep、メモアプリ |
| A | ❌ | ❌ | crontabのバックアップ、定時メール送信スクリプト |

### 2個 (10通り)
| 構成 | AI? | Agent? | 具体例 |
|------|:---:|:------:|--------|
| R+P | ⭕ | ❌ | ChatGPT素の戦略コンサル相談、論文構成提案 |
| R+T | ⭕ | ❌ | 初期Perplexity・Bing Chatの単発検索（RAG QA） |
| R+M | ⭕ | ❌ | 通常ChatGPT会話、Character.AI |
| R+A | ⭕ | △ | 単発命令で1回shellを叩くBot、反射型ロボット |
| P+T | ❌ | ❌ | テンプレ駆動の旅程作成+API連打、Expedia自動見積もり |
| P+M | ❌ | ❌ | Todoist、Trello、Asana、Jira |
| P+A | ❌ | ❌ | Zapier、IFTTT、Power Automate(ルールベース)、Airflow |
| T+M | ❌ | ❌ | キャッシュ付きAPIゲートウェイ、CDN、Redis層 |
| T+A | ❌ | ❌ | GitHub Actionsデプロイ、webhookでSlack投稿 |
| M+A | ❌ | ❌ | チェックポイント付きcron、世代管理ローテータ |

### 3個 (10通り)
| 構成 | AI? | Agent? | 具体例 |
|------|:---:|:------:|--------|
| R+P+T | ⭕ | △ | 単発リサーチエージェント、初期AutoGPT「調査」モード |
| R+P+M | ⭕ | ❌ | Notion AI、Mem.ai、ツール非使用Claude Projects（思考整理） |
| R+P+A | ⭕ | △ | ツール非使用Claude Codeでファイル書き込み |
| R+T+M | ⭕ | ❌ | Web検索付き通常ChatGPT、Claude Projects、Perplexity Pro対話 |
| R+T+A | ⭕ | △ | 単発メール送信bot、ワンショット投稿エージェント |
| R+M+A | ⭕ | △ | オフラインLLMで履歴付きファイル操作、エッジ完結型 |
| P+T+M | ❌ | ❌ | 自動ニュースレター生成パイプライン、定型レポートDB |
| P+T+A | ❌ | ❌ | Zapier/Make.comの複雑フロー、ETLバッチ |
| P+M+A | ❌ | ❌ | ステートフルなAirflow DAG、状態付きジョブキュー |
| T+M+A | ❌ | ❌ | Kafka Streams、Flink、ステートフルETLパイプライン |

### 4個 (5通り) ※1要素欠ける
| 構成 | 欠ける | AI? | Agent? | 具体例 |
|------|:-----:|:---:|:------:|--------|
| R+P+T+M | A欠 | ⭕ | ⭕ | ChatGPT Deep Research、Perplexity Spaces、Gemini Deep Research、Claude Research |
| R+P+T+A | M欠 | ⭕ | ⭕ | 単発実行Claude Code、AutoGPT/BabyAGI一回実行、CI上の自律タスクランナー |
| R+P+M+A | T欠 | ⭕ | ⭕ | オフライン環境のローカルLLMアシスタント、エッジAI |
| R+T+M+A | P欠 | ⭕ | ⭕ | ChatGPT Operator、Anthropic Computer Use、Cursor チャットモード（反応型ReAct） |
| P+T+M+A | R欠 | ❌ | ❌ | UiPath、Power Automate、Automation Anywhere（高度RPA） |

### 5個 (1通り)
| 構成 | AI? | Agent? | 具体例 |
|------|:---:|:------:|--------|
| R+P+T+M+A | ⭕ | ⭕ | Claude Code(継続セッション)、Devin、Cursor Composer/Agent、Microsoft Copilot Studio Agents、GitHub Copilot Workspace、OpenAI Operator、LangGraph/AutoGen/CrewAI製マルチエージェント、Salesforce Agentforce |

---

## 4. フラグ別の集計

| 分類 | 件数 | 概要 |
|------|:---:|------|
| ❌ 通常ソフトウェア（AI/Agent非該当） | 16 | LLM推論を持たない（電卓、RPA、ETL、ワークフロー、cron等） |
| ⭕❌ AIだがエージェントではない | 6 | R+チャット/QA/思考補助系（チャットボット、RAG、Notion AI等） |
| ⭕△ AIエージェント候補（緩い派） | 5 | R+3要素以下、R+T+A型の最小構成（単発実行bot等） |
| ⭕⭕ AIエージェント（標準派以上） | 5 | R+4要素以上（Deep Research、Operator、Cursor Agent、Claude Code等） |
| **合計** | **32** | |

---

## 5. AIエージェントの境界線（再掲）

「どこからがAIエージェントか」は業界で合意がない。3つの立場で整理。

### 厳格派（5要素必須）
- **R+P+T+M+A のみ**が「真のAIエージェント」
- Devin、継続セッションのClaude Codeなどに限定
- Anthropic研究者の一部、AutoGPT系の論者が採用

### 標準派（4要素以上、Rは必須）— **最も普及**
- 必須: **R + (T or A) + (P or M)** 系の4要素以上
- 「LLMが自律的に複数ステップを進めて世界に作用する」を基準
- OpenAI、LangChain、LlamaIndex等のドキュメントが概ねこの感覚
- ChatGPT Operator、Cursor Agent、各種ReActエージェントが該当

### 緩い派（3要素 R+T+A 等で十分）
- 「考えて、ツールを使い、世界に作用する」なら最低限エージェント
- 単発実行bot、記憶も計画もないワンショット型も含める

### 実用的な目安

> **R（推論）を必須として、そこに T+A の両方、もしくは P（計画ループ）が加わったら AIエージェント**

つまり 3要素組のうち **R+T+A、R+P+T、R+P+A** あたりが下限、4要素・5要素揃えば文句なしのエージェント。OWASPの定義も全部入りに近い水準を想定している。
