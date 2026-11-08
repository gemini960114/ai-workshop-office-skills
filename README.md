# 文件處理 — docx、pdf、pptx、xlsx

## 安裝

```bash
npx skills add anthropics/skills@docx -a antigravity -a claude-code --copy -y
npx skills add anthropics/skills@pdf -a antigravity -a claude-code --copy -y
npx skills add anthropics/skills@pptx -a antigravity -a claude-code --copy -y
npx skills add anthropics/skills@xlsx -a antigravity -a claude-code --copy -y
```

## 用途

Anthropic 官方 skill，讓 agent 能可靠地讀取與產生 Office 文件和 PDF。沒有這些 skill 時，agent 常要臨時自己寫程式讀檔，格式複雜一點就容易出錯；產出的 Excel、簡報也常只有內容、沒有版面。

| Skill | 作用 |
| --- | --- |
| docx | 讀取、建立、修改 Word 文件（標題層級、表格、目錄、追蹤修訂） |
| pdf | 擷取文字與表格、合併分割、填表單，掃描檔可做 OCR |
| pptx | 讀取、建立、修改簡報（版面、圖表、講者備註、動畫） |
| xlsx | 讀取、建立、修改試算表（公式、格式、圖表） |

這幾個 skill 不用特別點名，只要任務牽涉到這些檔案就會觸發。

## 事前準備

| 工具 | 用在哪裡 | 沒有會怎樣 |
| --- | --- | --- |
| Python 3、Node.js | 所有 skill 都靠它們執行程式 | 無法執行 |
| [LibreOffice](https://www.libreoffice.org/download/download/)（建議安裝） | 重算 Excel 公式、docx 轉 PDF、把簡報轉成圖片讓 agent 自己檢查版面 | 範例 1 的公式沒有預先算好的值、範例 4 需改用 Word 轉檔、agent 無法自己檢查版面 |
| Microsoft Office | 打開成果檢查；Windows 上沒有 LibreOffice 時，可以叫 Word 轉 PDF | 可用 LibreOffice 代替檢查，但動畫、轉場在 PowerPoint 最準 |

安裝 LibreOffice 後，請確認 `soffice` 指令能在終端機執行（Windows 預設路徑 `C:\Program Files\LibreOffice\program`，必要時加到 PATH）。

## 範例資料

```text
example/
└── 臺灣採購經理人指數(pmi及nmi).csv   ← 2012/07–2026/09 每月 PMI、NMI
```

以下五個範例是一條流水線，前一個範例的產出就是下一個範例的輸入，所有產出檔都存在 `example/` 資料夾。請依序完成：

```text
CSV ──xlsx──▶ pmi.xlsx ──pptx──▶ pmi.pptx
                 │
                 └──docx──▶ pmi-report.docx ──pdf──▶ pmi-report.pdf ──pptx──▶ pmi-lecture.pptx
```

---

## 範例 1（xlsx）：把 CSV 做成有圖表的 Excel

```text
把 example/臺灣採購經理人指數(pmi及nmi).csv 做成 Excel，存成 example/pmi.xlsx：
第一個工作表放原始資料，第二個工作表算出每年的 PMI、NMI 平均，並加上歷年走勢的折線圖。
```

**預期行為：** 產出有兩個工作表的 Excel，年度平均用公式計算（點開儲存格看得到公式，不是寫死的數字），折線圖的座標軸與圖例都有中文標示。CSV 裡 NMI 在 2014 年 8 月以前是「-」，agent 應該當成沒有資料處理，而不是當成 0。

**檢查重點：**

- 2012 年只有 7–12 月、2026 年只有 1–9 月，屬於不完整年度，最好有註記。
- 程式寫入的公式不會自帶計算結果。有 LibreOffice 時 agent 會先重算一次；沒有的話，用 Excel 打開會自動計算，但下一個範例讀檔時可能讀到空值，agent 應該改從原始資料自己算，而不是把空值當成 0。

---

## 範例 2（pptx）：根據 Excel 做簡報

```text
根據 example/pmi.xlsx 做一份 5 頁的簡報，存成 example/pmi.pptx，
說明臺灣採購經理人指數的意義、近年走勢與重點觀察，要有圖表。
```

**預期行為：** 產出 5 頁簡報，包含標題頁、指標說明、走勢圖表與重點觀察；圖表數字與 Excel 一致，文字不超出版面。用 PowerPoint 打開，檢查版面是否整齊、圖表能不能正常顯示。

---

## 範例 3（docx）：根據 Excel 寫一份 Word 報告

```text
根據 example/pmi.xlsx 寫一份 Word 分析報告，存成 example/pmi-report.docx：
要有封面、自動目錄，章節分成「指標說明」「歷年走勢」「近一年觀察」「結論」，
「歷年走勢」放年度平均的表格與折線圖。
```

**預期行為：** 標題用 Word 內建的「標題 1／標題 2」樣式，目錄是可更新的欄位（在目錄上按右鍵 →「更新功能變數」會跟著章節變化），不是手打的文字；表格數字與 Excel 的年度平均一致。

**檢查重點：**

- 剛產出的目錄可能顯示「沒有目錄項目」或空白，這是正常的：在 Word 裡更新功能變數後才會出現頁碼。範例 4 轉 PDF 前應先更新。
- Word 文件裡的折線圖通常是圖片（agent 先用程式畫好圖再插入），不是可以點開編輯資料的 Word 圖表，這是可以接受的。

---

## 範例 4（pdf）：把報告轉成上課講義 PDF

```text
把 example/pmi-report.docx 轉成 PDF，存成 example/pmi-report.pdf：
每頁下方加上頁碼，並加上淡灰色「課堂講義」浮水印。
```

**預期行為：** 中文字型正常顯示（沒有方框或亂碼），頁碼與浮水印出現在每一頁，浮水印不遮住內文、文字仍可選取複製。

**檢查重點：**

- 轉 PDF 需要 LibreOffice 或 Word。兩者都沒有時，agent 可能改用其他方式自己排版輸出 PDF，版面會和 Word 不同。
- 頁碼、浮水印用程式疊加時，必須指定中文字型（例如微軟正黑體、Noto Sans TC），否則「課堂講義」會變成方框。這是最常見的錯誤。

---

## 範例 5（pptx）：把 PDF 做成可上課用的動畫簡報

```text
把 example/pmi-report.pdf 做成 PPTX 簡報，存成 example/pmi-lecture.pptx，給大學生上課用，約 20 分鐘。
文字都要能在 PowerPoint 裡編輯（不要做成圖片）
每頁物件依序逐一出現（按一下出現一個），用淡入效果
換頁用淡入效果，慢一點
要講者備註
```

**預期行為：** agent 先用 pdf skill 擷取 PDF 的文字與表格，再重新排版成簡報，而不是把每頁 PDF 截圖貼上。20 分鐘約 10–15 頁；每頁的文字都是可編輯的文字方塊，圖表最好是 PowerPoint 原生圖表。範例 4 加的「課堂講義」浮水印和頁碼也會被擷取出來，agent 應該辨識出來並排除，不要放進投影片。

**檢查重點：** 用 PowerPoint 打開：

- 「動畫」→「動畫窗格」：每個物件是「按一下時」觸發的「淡入」，順序與閱讀順序一致
- 「轉場」：每頁都是「淡入」，期間比預設的 0.7 秒長（例如 1.5–2 秒）
- 「檢視」→「備忘稿」：每頁都有講者備註，內容是講者要說的話，不只是重複投影片文字

> 動畫與轉場是 python-pptx 等函式庫沒有直接支援的功能，agent 需要直接寫入簡報的 XML。這正是 pptx skill 派得上用場的地方——沒有 skill 時常會漏掉動畫，或產生 PowerPoint 打不開的檔案。
