#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────
# transfer-to-desktop.sh
# 將成少共犯修法知識庫所有相關檔案傳到本機桌面
# 路徑：~/Desktop/2026少事法修法/
#
# 使用方式（在你本機跑、不是在沙箱跑）：
#   1. 先 cd 到你本機 policy repo 的根目錄
#   2. git checkout claude/taiwan-juvenile-adult-cooffender-EDt1Q
#   3. git pull origin claude/taiwan-juvenile-adult-cooffender-EDt1Q
#   4. bash child-rights/_drafts/transfer-to-desktop.sh
#
# 執行後會：
#   - 在桌面建立 2026少事法修法/ 目錄樹（5 個子目錄）
#   - 複製 22 個檔案，依用途分類命名
#   - 產生本機讀本 README.md
#   - 印出統計報告
# ──────────────────────────────────────────────────────────────────

set -euo pipefail

# ── 路徑設定 ──────────────────────────────────────────────────
REPO_ROOT="$(pwd)"
DEST="$HOME/Desktop/2026少事法修法"

# 偵測 macOS 與 Linux 的桌面路徑
if [[ "$(uname)" == "Darwin" ]]; then
  DEST="$HOME/Desktop/2026少事法修法"
elif [[ -d "$HOME/桌面" ]]; then
  DEST="$HOME/桌面/2026少事法修法"
fi

# ── 預檢 ────────────────────────────────────────────────────
if [[ ! -d "$REPO_ROOT/child-rights" ]]; then
  echo "❌ 錯誤：請在 policy repo 根目錄執行（找不到 child-rights/）"
  echo "   現在目錄：$REPO_ROOT"
  exit 1
fi

BRANCH=$(git rev-parse --abbrev-ref HEAD 2>/dev/null || echo "")
if [[ "$BRANCH" != "claude/taiwan-juvenile-adult-cooffender-EDt1Q" ]]; then
  echo "⚠️  警告：當前分支不是 claude/taiwan-juvenile-adult-cooffender-EDt1Q"
  echo "   當前分支：$BRANCH"
  read -rp "確定要繼續嗎？[y/N] " yn
  [[ "$yn" =~ ^[Yy]$ ]] || exit 1
fi

# ── 建立目錄樹 ──────────────────────────────────────────────
echo ""
echo "📁 建立目錄：$DEST"
mkdir -p "$DEST"/{00_索引,01_對外立場文件,02_工作筆記,03_行動文件_新聞稿與公文,04_原始證據,05_關聯背景檔案/校園安全,05_關聯背景檔案/兒少權益}

# ── 函式：複製並回報 ────────────────────────────────────────
copy_log=()
copy() {
  local src="$1" dst="$2"
  if [[ -f "$REPO_ROOT/$src" ]]; then
    cp "$REPO_ROOT/$src" "$DEST/$dst"
    copy_log+=("✅ $dst")
  else
    copy_log+=("❌ 缺：$src")
  fi
}

# ── 00 索引 ────────────────────────────────────────────────
copy "child-rights/_drafts/2026-05-24-成少共犯知識庫索引.md"          "00_索引/00-知識庫索引.md"
copy "child-rights/_drafts/2026-05-24-成少共犯知識庫索引.html"        "00_索引/00-知識庫索引.html"

# ── 01 對外立場文件 ─────────────────────────────────────────
copy "child-rights/2026-05-23-少事法成少共犯第五大訴求.md"           "01_對外立場文件/A1-第五大訴求-v1.1.md"
copy "child-rights/2026-05-23-少事法成少共犯第五大訴求.html"         "01_對外立場文件/A1-第五大訴求-v1.1.html"
copy "child-rights/2026-04-01-少事法矯正教育假釋塗銷轉銜改革.md"    "01_對外立場文件/B-既有四大訴求-基底.md"
copy "child-rights/2026-04-01-少事法矯正教育假釋塗銷轉銜改革.html"  "01_對外立場文件/B-既有四大訴求-基底.html"

# ── 02 工作筆記 ────────────────────────────────────────────
copy "child-rights/_drafts/2026-05-23-成少共犯修法工作筆記.md"       "02_工作筆記/A2-修法工作筆記-v1.1.md"
copy "child-rights/_drafts/2026-05-23-成少共犯修法工作筆記.html"     "02_工作筆記/A2-修法工作筆記-v1.1.html"
copy "child-rights/_drafts/2026-05-24-立法院1127會議紀錄分析.md"     "02_工作筆記/A3-立法院1127會議分析.md"
copy "child-rights/_drafts/2026-05-24-立法院1127會議紀錄分析.html"   "02_工作筆記/A3-立法院1127會議分析.html"

# ── 02b 深度政策研究報告（5/25 兩份） ─────────────────────
mkdir -p "$DEST/02b_深度政策研究報告"
copy "child-rights/_drafts/2026-05-25-成年犯罪組織利用少事法保護傘論述報告.md"        "02b_深度政策研究報告/F1-成年犯罪組織利用少事法保護傘論述報告.md"
copy "child-rights/_drafts/2026-05-25-成年犯罪組織利用少事法保護傘論述報告.html"      "02b_深度政策研究報告/F1-成年犯罪組織利用少事法保護傘論述報告.html"
copy "child-rights/_drafts/2026-05-25-少事法修法建議完整對照與差距分析.md"            "02b_深度政策研究報告/F2-少事法修法建議完整對照與差距分析.md"
copy "child-rights/_drafts/2026-05-25-少事法修法建議完整對照與差距分析.html"          "02b_深度政策研究報告/F2-少事法修法建議完整對照與差距分析.html"

# ── 03 行動文件 ────────────────────────────────────────────
copy "child-rights/_drafts/2026-05-24-新聞稿補位-司法院違背立委質詢.md"   "03_行動文件_新聞稿與公文/A4-新聞稿補位草稿.md"
copy "child-rights/_drafts/2026-05-24-新聞稿補位-司法院違背立委質詢.html" "03_行動文件_新聞稿與公文/A4-新聞稿補位草稿.html"
copy "child-rights/_drafts/2026-05-24-索資公文範本.md"                    "03_行動文件_新聞稿與公文/A5-索資公文範本-四份.md"
copy "child-rights/_drafts/2026-05-24-索資公文範本.html"                  "03_行動文件_新聞稿與公文/A5-索資公文範本-四份.html"

# ── 04 原始證據 ────────────────────────────────────────────
copy "child-rights/_drafts/source/2025-11-27-立法院司法及法制委員會第12次會議紀錄.txt" "04_原始證據/A6-立法院1127公報全文.txt"
copy "child-rights/_drafts/source/README.md"                                            "04_原始證據/README.md"

# ── 05 關聯背景檔案 ─────────────────────────────────────────
# 校園安全
copy "campus-safety/2025-12-16-校園菸毒與毒駕新興毒品檢測.md"        "05_關聯背景檔案/校園安全/C1-校園菸毒與毒駕.md"
copy "campus-safety/2025-12-16-校園菸毒與毒駕新興毒品檢測.html"      "05_關聯背景檔案/校園安全/C1-校園菸毒與毒駕.html"
copy "campus-safety/2026-04-02-校園唾液快篩二層檢測機制.md"          "05_關聯背景檔案/校園安全/C2-唾液快篩二層檢測.md"
copy "campus-safety/2026-04-02-校園唾液快篩二層檢測機制.html"        "05_關聯背景檔案/校園安全/C2-唾液快篩二層檢測.html"
copy "campus-safety/2026-03-23-轉介教育條例.md"                      "05_關聯背景檔案/校園安全/C3-轉介教育條例.md"
copy "campus-safety/2026-03-23-轉介教育條例.html"                    "05_關聯背景檔案/校園安全/C3-轉介教育條例.html"
copy "campus-safety/2026-03-25-永和國中霸凌案.md"                    "05_關聯背景檔案/校園安全/C4-永和國中霸凌案.md"
copy "campus-safety/2026-03-25-永和國中霸凌案.html"                  "05_關聯背景檔案/校園安全/C4-永和國中霸凌案.html"
copy "campus-safety/2026-03-23-倡議成果追蹤.md"                      "05_關聯背景檔案/校園安全/E-倡議成果追蹤.md"
copy "campus-safety/2026-03-23-倡議成果追蹤.html"                    "05_關聯背景檔案/校園安全/E-倡議成果追蹤.html"

# 兒少權益
copy "child-rights/2025-12-24-少年重大暴力案件量刑透明化.md"         "05_關聯背景檔案/兒少權益/D1-量刑透明化.md"
copy "child-rights/2025-12-24-少年重大暴力案件量刑透明化.html"       "05_關聯背景檔案/兒少權益/D1-量刑透明化.html"
copy "child-rights/2026-04-22-兒少權法修法六大缺口.md"               "05_關聯背景檔案/兒少權益/D2-兒少權法六大缺口.md"
copy "child-rights/2026-04-22-兒少權法修法六大缺口.html"             "05_關聯背景檔案/兒少權益/D2-兒少權法六大缺口.html"

# ── 產生本機 README ────────────────────────────────────────
cat > "$DEST/README.md" <<'EOF'
# 2026 少事法修法資料庫｜成少共犯議題

> 本目錄為國教行動聯盟成少共犯議題修法之完整資料庫本機版。
> 由 `transfer-to-desktop.sh` 從 GitHub repo (`naer-tw/policy` 分支 `claude/taiwan-juvenile-adult-cooffender-EDt1Q`) 自動生成。
> 最後同步時間：請以本目錄各檔案 mtime 為準。

## 目錄結構

```
2026少事法修法/
├── 00_索引/                         (從這裡開始讀)
│   ├── 00-知識庫索引.md
│   └── 00-知識庫索引.html  ★建議先開
│
├── 01_對外立場文件/                 (正式發布用)
│   ├── A1-第五大訴求-v1.1.md/.html  ★ 核心立場
│   └── B-既有四大訴求-基底.md/.html
│
├── 02_工作筆記/                     (內部戰略基底)
│   ├── A2-修法工作筆記-v1.1.md/.html
│   └── A3-立法院1127會議分析.md/.html ★ 核心證據
│
├── 02b_深度政策研究報告/            (5/25 兩份深度報告)
│   ├── F1-成年犯罪組織利用少事法保護傘論述報告.md/.html
│   └── F2-少事法修法建議完整對照與差距分析.md/.html
│
├── 03_行動文件_新聞稿與公文/        (對外行動)
│   ├── A4-新聞稿補位草稿.md/.html
│   └── A5-索資公文範本-四份.md/.html
│
├── 04_原始證據/                     (公報全文)
│   ├── A6-立法院1127公報全文.txt    ★ 191 KB、54 頁
│   └── README.md
│
└── 05_關聯背景檔案/                 (既有訴求基底)
    ├── 校園安全/
    │   ├── C1-校園菸毒與毒駕.md/.html
    │   ├── C2-唾液快篩二層檢測.md/.html
    │   ├── C3-轉介教育條例.md/.html
    │   ├── C4-永和國中霸凌案.md/.html
    │   └── E-倡議成果追蹤.md/.html
    │
    └── 兒少權益/
        ├── D1-量刑透明化.md/.html
        └── D2-兒少權法六大缺口.md/.html
```

## 推薦閱讀順序（按角色）

### 王理事長（5-10 分鐘）
1. `00_索引/00-知識庫索引.html`（速覽全貌）
2. `02_工作筆記/A3-立法院1127會議分析.html`（戰略意義 + 5 引用句）
3. `03_行動文件_新聞稿與公文/A4-新聞稿補位草稿.html`

### 政策研究組（完整版）
1. `04_原始證據/A6-立法院1127公報全文.txt`（54 頁原文，約 1 小時）
2. `02_工作筆記/A3-立法院1127會議分析.md`（對照閱讀）
3. `02_工作筆記/A2-修法工作筆記-v1.1.md`（執行清單）
4. `03_行動文件/A5-索資公文範本-四份.md`（公文發送）

### 媒體記者
1. `03_行動文件/A4-新聞稿補位草稿.html`（標題 + Q&A）
2. `01_對外立場文件/A1-第五大訴求-v1.1.html`（完整立場）

### 立委辦公室
1. `03_行動文件/A4-新聞稿補位草稿.html`（議題定位）
2. `01_對外立場文件/A1-第五大訴求-v1.1.md`（修法主張）
3. `02_工作筆記/A3-立法院1127會議分析.md`（找該立委 11/27 質詢段落）

## 同步機制

當 repo 有新更新時，重跑 `transfer-to-desktop.sh` 即可：
```bash
cd ~/policy  # 或你本機 repo 的路徑
git pull origin claude/taiwan-juvenile-adult-cooffender-EDt1Q
bash child-rights/_drafts/transfer-to-desktop.sh
```

腳本會覆蓋既有檔案，所以**請勿在本資料夾直接編輯**檔案，否則同步時會被覆蓋。如要編輯，請：
- 編輯 repo 內檔案（推送回 GitHub）
- 或先把本目錄複製為 `2026少事法修法_我的編輯版/`

## 待補資料

以下檔案尚未取得，需另行索取後手動加入本目錄：

- [ ] `04_原始證據/2026-05-22-司法院少事法修正草案.pdf`（用戶桌面待 push）
- [ ] `04_原始證據/2026-05-20-學界藥物濫用會議-全套.zip`（用戶桌面待 push）
- [ ] `04_原始證據/2025-11-監察院新北少年事件調查報告.pdf`（待向監察院索取）
- [ ] `04_原始證據/2025-04-24-司法院函1140400447號.pdf`（待向司法院索取）
- [ ] `04_原始證據/2025-04-28-司法院函1140008146號.pdf`（待向司法院索取）
- [ ] `04_原始證據/2024-12-11_2025-02-20_2025-07-09-跨院部協商會議紀錄.pdf`（待索取）

---

**維護**：國教盟政策研究組
**Repo**：https://github.com/naer-tw/policy
**分支**：claude/taiwan-juvenile-adult-cooffender-EDt1Q
EOF

# ── 報告 ────────────────────────────────────────────────────
echo ""
echo "════════════════════════════════════════════"
echo "  📋 傳檔報告"
echo "════════════════════════════════════════════"
for line in "${copy_log[@]}"; do
  echo "  $line"
done

echo ""
echo "────────────────────────────────────────────"
total=$(find "$DEST" -type f | wc -l | tr -d ' ')
size=$(du -sh "$DEST" 2>/dev/null | awk '{print $1}')
echo "  ✅ 完成！共 $total 個檔案、合計 $size"
echo "  📁 路徑：$DEST"
echo ""
echo "  下一步建議："
echo "  1. 雙擊開啟  $DEST/00_索引/00-知識庫索引.html"
echo "  2. 把  $DEST/04_原始證據/A6-立法院1127公報全文.txt  用編輯器開啟"
echo "  3. 印  $DEST/03_行動文件_新聞稿與公文/A5-索資公文範本-四份.html"
echo "      （送王理事長簽核）"
echo "════════════════════════════════════════════"
