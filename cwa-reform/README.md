# 國教行動聯盟政策知識庫(GitHub Pages)

本目錄為國教行動聯盟兒少權法修法政策知識庫的 GitHub Pages 部署檔案。

## 預設 URL 結構

部署後預期可訪問:

```
https://policy.aabe.org.tw/cwa-reform/                          (首頁)
https://policy.aabe.org.tw/cwa-reform/reform-proposals/         (6 大缺口列表)
https://policy.aabe.org.tw/cwa-reform/reform-proposals/RP-G01.html  (RP-G01 詳細)
https://policy.aabe.org.tw/cwa-reform/intl-benchmarks/          (21 國對標列表)
https://policy.aabe.org.tw/cwa-reform/intl-benchmarks/JP01.html (日本)
https://policy.aabe.org.tw/cwa-reform/policies/                  (27 篇政策列表)
https://policy.aabe.org.tw/cwa-reform/articles/                  (16 條評估列表)
```

## 部署步驟

### Option 1:獨立 repo(推薦)

```bash
# 1. 在 GitHub 上建立 naer-tw/policy repo
# 2. 將本目錄內容推送到該 repo 的 main 分支
cd docs/cwa-reform/
git init
git add .
git commit -m "Initial: 國教盟政策知識庫"
git remote add origin git@github.com:naer-tw/policy.git
git branch -M main
git push -u origin main

# 3. 在 GitHub 設定 → Pages → Source: main branch / root
# 4. 等待 1-2 分鐘即可訪問 https://policy.aabe.org.tw/cwa-reform/
```

### Option 2:整合進現有 GEO 專案

如國教盟已有 `naer-tw/policy-kb` 或類似 repo,可將本目錄複製進該 repo 的 `policy/` 子目錄。

## 內容更新

更新流程:

```bash
# 1. 在主專案(兒少權法與兒童家庭署)修改 SQLite 或 MD 政策文章
# 2. 重新跑生成腳本
cd /Users/coachyang/Documents/Claude/Projects/兒少權法與兒童家庭署/
node scripts/build_kb_pages.js

# 3. 將 docs/cwa-reform/內容推送到 GitHub Pages repo
cd docs/cwa-reform/
git add .
git commit -m "Update: 2026-MM-DD"
git push
```

## 主要文件對照

| 公開頁面 | 對應內部資料 |
|---|---|
| reform-proposals/RP-G01.html ~ G06.html | SQLite `reform_proposal` 表 6 筆 |
| intl-benchmarks/[country].html | SQLite `intl_benchmark` 表 21 筆 |
| policies/[id].html | `../兒少權監督平台/data/sources/naaes_policy/*.md` 27 篇 |
| articles/[article].html | `data/outputs/briefs/2026-05-10_165條全條評估表.md` 16 條 🟡 |

## 統計

- 首頁: 1 頁
- 6 大缺口: 1 概覽 + 6 條 = 7 頁
- 21 國對標: 1 概覽 + 21 國 = 22 頁
- 27 篇政策: 1 概覽 + 27 篇 = 28 頁
- 16 條評估: 1 概覽 + 16 條 = 17 頁
- **總計: 75 頁 + 共用 CSS**

## 聯絡

國教行動聯盟
理事長 王瀚陽 | 0983-097165
Facebook: facebook.com/twedumove
