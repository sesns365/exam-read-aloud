# 報讀考卷產生器（exam-read-aloud）

一個 Claude Code SKILL：把考卷（**PDF、Word、圖片／掃描檔**）自動做成 **版面完全不變、可點按報讀** 的網頁考卷，適合特殊需求學生的報讀考場、聽讀輔助。

> 安裝後，只要把考卷檔案丟給 Claude Code，就會自動產出「原檔名_報讀版.html」。

## 功能

- **原版面 100% 保留**：用原始考卷畫面當背景，底線、框線、字型、排版都不變；列印出來和原卷一樣。
- **三種點法**
  | 點哪裡 | 框線 | 唸什麼 |
  |---|---|---|
  | 題號（例如「1. (　)」） | 藍框 | 整題一次唸完（題目＋全部選項） |
  | 題目文字 | 橘框 | 只唸題目 |
  | 每個選項（①②③④、A–D 代號） | 橘框 | 只唸該選項 |
  | 表格格子、連連看的框、圖片標籤 | 橘框 | 只唸這一格 |
- 閱讀題組的文章依段落分開點讀。
- 右上角小圖示：**⛶ 全螢幕**、**⚙ 設定**（考卷放大縮小、整頁顯示、符合寬度、語速、聲音）。朗讀時才出現 ■ 停止鈕。
- 單一 HTML 檔，不需網路、不需安裝，雙擊就能用（建議 Chrome／Edge）。
- 用瀏覽器內建語音（繁體中文 zh-TW），不需任何 API 金鑰。

## 支援的檔案

| 檔案 | 處理方式 |
|---|---|
| PDF（有文字層，例如 Word 匯出的） | 直接讀取文字與位置，最準確 |
| Word（.docx .doc .rtf .odt） | 用 Microsoft Word 轉成 PDF 後處理（版面不變） |
| 圖片（.jpg .png …）、掃描的 PDF | Windows 內建文字辨識（OCR），再由 Claude 對照原圖校正文字 |

## 系統需求

- [Claude Code](https://claude.com/claude-code)
- Python 3.9 以上，並安裝套件：
  ```bash
  pip install pymupdf pillow
  ```
- Windows 10／11
  - Word 檔：需要 Microsoft Word（或 LibreOffice）
  - 圖片／掃描檔：使用 Windows 內建 OCR，需要「中文（繁體，台灣）」語言套件（Windows 設定 → 時間與語言 → 語言）

## 安裝

**方法一：用 skills 工具安裝（推薦）**
```bash
npx skills add sesns365/exam-read-aloud -g -y
```

**方法二：用 git 下載到 SKILL 資料夾**
```bash
git clone https://github.com/sesns365/exam-read-aloud.git "%USERPROFILE%\.claude\skills\exam-read-aloud"
```

**方法三：下載 ZIP**
1. 在本頁按綠色 **Code → Download ZIP**。
2. 解壓縮後，把資料夾改名為 `exam-read-aloud`。
3. 放到 `C:\Users\你的帳號\.claude\skills\` 底下，確認路徑是 `…\.claude\skills\exam-read-aloud\SKILL.md`。

安裝後重新開啟 Claude Code 即可使用。

## 使用方式

1. 在 Claude Code 對話框丟入考卷檔案（PDF／Word／圖片都可以，一份考卷拍成多張圖也可以一起丟）。
2. 不用多打字，Claude 會自動產出「原檔名_報讀版.html」。若沒有自動開始，補一句「做報讀版」即可。
3. 圖片或掃描檔會多一個步驟：Claude 會對照原圖校正辨識文字，確保唸出來的內容正確。

### 部署成網址（選用）

想讓學生用網址開啟時，對 Claude 說「**部署到 GitHub Pages**」。**沒有明確要求就不會上傳任何東西。**
部署時只上傳報讀版 HTML，不會上傳原始考卷檔。

> ⚠️ 公開網址任何人都看得到。建議考試當天再部署，考完刪除。

## 支援的考卷寫法

已用多份國小定期評量卷（三、五、六年級的自然、社會）測試，可辨識：

- 題號：`1. (　)`、`(　) 1.`；小題：`(　)(1)`、`(　) ①`、`(1)`
- 選擇題選項：`①②③④`、`○1 ○2`
- 代號：`A 東邊 B 西邊`、`A.紅蘿蔔 B.落地生根`、`ㄅ草本莖 ㄆ木本莖`、`Ａ.美元 Ｂ.新臺幣`
- 注音版考卷：字型內建的注音不會被唸出來；打在字後面的注音（如「上ㄕㄤˋ課」）會自動略過
- 是非題「○ ×」、角度「70°」、連連看的點線都會唸成自然的說法

- **方框與表格**：連連看的每個框、表格的每一格、圖片裡的標籤（如「ㄅ 小白菜」）都是獨立的橘框，可以單獨點讀；只有數值的格子（如「3本」）會連同該列標題一起唸
  （只認四邊都有框線的方框，字下的底線、圖中的線條不會被誤判）

圖中沒有框線的文字（例如角度標示）會併在題目中朗讀，建議搭配老師口頭說明。

## 小提醒

- 電腦語音偶爾會唸錯破音字，考前請先點幾題試聽。
- 不同裝置（平板、Chromebook）可用的聲音不同，可在 ⚙ 設定中切換。
- 格式特殊的考卷（例如選項寫成 (A)(B)(C)、題號用「（一）」），Claude 會檢查並調整辨識規則。

## 檔案結構

```
exam-read-aloud/
├── SKILL.md              # 給 Claude 的工作流程說明
├── scripts/
│   ├── build.py          # 主程式：PDF／Word／圖片 → 報讀版 HTML
│   ├── word2pdf.ps1      # Word 轉 PDF（Microsoft Word／LibreOffice）
│   └── ocr.ps1           # Windows 內建 OCR
└── assets/
    └── template.html     # 報讀版網頁樣板
```

## 授權

MIT License，歡迎自由使用、修改與分享。
