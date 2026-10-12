# grant-proposal-writer：研究計畫撰寫助手（Claude Skill）

協助研究者與研究生從研究構想一路寫到**可送件的研究計畫書（Word）**。

- **計畫類型可選**：國科會專題研究計畫、國衛院 IRG／CDG／TRG、衛福部徵求案、其他
- **年期與語言由計畫類型決定**：例如國衛院 IRG 只能選 3–5 年、CDG 固定 4 年，兩者都要求英文撰寫
- **可上傳自己的範本**；沒有範本時，自動上網查證當年度官方規定
- **可加入自訂要求**，並有明確的衝突處理順序
- **零捏造**：每筆引用都必須有工具實際查到的 PMID 或 DOI；補助規定都附來源
- **6 位模擬評審加 1 位主席**的多輪審查，以找出問題為優先
- **程式化合規檢查**：頁數、必填章節、殘留標記、文獻識別碼、引用雙向比對

## 流程

0. 計畫設定：類型、年期、範本、自訂要求、官方規定查證，產出 `proposal-config.md`
1. 研究方向與檢索策略
2. 文獻搜尋：PubMed MCP，或網路搜尋，或使用者提供的文獻
3. 文獻分析：現況、缺口、限制、展望
4. 蘇格拉底式對話，確立計畫細節與大綱
5. 分章節撰寫
6. 合規檢查與審查小組審查
7. 修改，附修改對照表
8. 重複審查，至多 5 輪
9. Word 輸出與交付清單

## 規則優先順序

1. 官方當年度硬限制（不可覆蓋）
2. 使用者自訂要求（覆蓋範本時會先提醒）
3. 使用者範本
4. 內建 profile 預設值

## 安裝

**Claude Code**
- 個人使用：把整個 `grant-proposal-writer/` 資料夾放到 `~/.claude/skills/`
- 專案共用：放到該專案的 `.claude/skills/`

**claude.ai／Claude 桌面版**
- 把資料夾壓縮成 zip，在 claude.ai 的 skills 設定中上傳並啟用

官方說明：https://code.claude.com/docs/en/skills

## 建議環境

| 功能 | 需要 | 沒有時 |
|---|---|---|
| 文獻搜尋 | PubMed MCP（建議）、Scite 等 | 改用網路搜尋；再沒有，則只建架構並標【需補文獻】 |
| 官方規定查證 | 網路搜尋 | 請使用者上傳徵求公告或範本 |
| 合規檢查 | Python 3、`python-docx`、`pypdf`、LibreOffice | 改用人工合規清單 |
| 獨立評審 | 支援子代理（如 Claude Code） | 依序扮演評審，隔離撰寫脈絡 |

合規檢查單獨使用：

```bash
pip install python-docx pypdf
python scripts/compliance_check.py 計畫書.docx --config proposal-config.md --refs refs.ris --out report.md
```

## 檔案結構

```
grant-proposal-writer/
├── SKILL.md                         # 核心流程與鐵則
├── references/
│   ├── config-template.md           # proposal-config.md 範本（含合規檢查參數）
│   ├── literature-protocol.md       # 文獻來源分層與零捏造規則
│   ├── writing-protocol.md          # 分章撰寫、多年期、Word 輸出、交付清單
│   ├── review-panel.md              # 審查小組、通過條件、輪數
│   └── profiles/
│       ├── _schema.md               # profile 格式說明（貢獻者必讀）
│       ├── nstc.md                  # 國科會專題研究計畫
│       ├── nhri-irg.md              # 國衛院 創新研究計畫
│       ├── nhri-cdg.md              # 國衛院 研究發展獎助計畫
│       ├── nhri-trg.md              # 國衛院 TRG（骨架版）
│       ├── mohw.md                  # 衛福部徵求案（流程型）
│       └── generic.md               # 其他／保底
└── scripts/
    └── compliance_check.py          # 合規檢查
```

## 各 profile 狀態

| profile | 依據文件 | 最後查證 | 狀態 |
|---|---|---|---|
| 國科會 | CM03 範本；CM03 頁數限制一覽表（2023-11-07 版） | 2026-10-12 | 多年期年數上限、各學門審查面向未收錄，執行時查證 |
| 國衛院 IRG | 116 年度申請作業手冊 | 2026-10-12 | 完整 |
| 國衛院 CDG | 116 年度申請作業手冊 | 2026-10-12 | 完整 |
| 國衛院 TRG | 111／115／116 年度手冊的徵求說明 | 2026-10-12 | 骨架版，格式細節待補 |
| 衛福部 | 各案徵求說明書（範例） | 2026-10-12 | 流程型，每案讀取該案格式 |

**profile 會過時。** 補助單位每年修訂規定。skill 執行時一定會重新查證當年度官方文件；profile 的「最後查證日期」超過一年時會提醒使用者。歡迎發 PR 更新。

## 貢獻

1. 新增或更新 profile 前，先讀 `references/profiles/_schema.md`
2. 每一條規定都要附官方來源網址與文件年度；沒有來源的不寫
3. 更新 profile 的「最後查證日期」
4. 新增計畫類型時，同步更新 SKILL.md 步驟 0 的選單與本 README 的狀態表

## 使用須知

- 本工具產出的是**草稿**。計畫書由研究者署名負責，內容須逐段確認；前導資料、團隊經歷、經費數字等標【請填入】的位置須由研究者本人填寫。
- 送件規定以補助單位當年度公告為準。
- 部分單位要求揭露 AI 使用情形，請依規定辦理。
