# DocxAvalonia

以 **Avalonia UI** 打造的文件編輯／預覽桌面程式，支援 **`.docx` / `.doc` / `.odt` 讀寫**，可切換 **六種介面風格**：Microsoft Word、LibreOffice Writer、Google 文件、Zoho Writer、WPS Writer、FreeOffice TextMaker。

## 下載

Release：https://github.com/huang1988pioneer/DocxAvalonia/releases

| 檔案 | 說明 |
|------|------|
| **self-contained** | 免安裝 .NET，解壓後執行 `DocxAvalonia.exe`（檔案較大） |
| **framework-dependent** | 需已安裝 [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0)（檔案較小） |

self-contained ZIP 內容：`DocxAvalonia.exe`、`DocxAvalonia.pdb`、`demo.docx`

### 若出現「智慧型應用程式控制已封鎖」

本程式尚未使用商業程式碼簽章，Windows 11 **Smart App Control** 可能封鎖（錯誤常為 `0x800711C7`）。

1. **Windows 安全性** → **應用程式及瀏覽器控制** → **智慧型應用程式控制設定**
2. 改為 **關閉** 或 **評估**
3. **重新開機** 後再執行

或於已安裝 .NET 8 SDK 的環境：

```powershell
dotnet run -c Release
```

## 功能

### 檔案
- 新增、開啟、儲存、另存新檔
- 格式：**`.docx`**（原生 OOXML）、**`.odt`**（OpenDocument 讀寫）、**`.doc`**（可讀 Word 97-2003 二進位；另存為 Word 可開的 RTF 相容 `.doc`，本程式可完整往返）
- 拖放 `.docx` / `.doc` / `.odt` 開啟
- 命令列傳入路徑開啟
- **自動儲存**（可開關，每 60 秒；需已有儲存路徑）

### 編輯
- 剪下 / 複製 / 貼上
- 復原 / 重做
- 尋找、取代、全部取代

### 介面風格（可切換）
頂部 **介面風格** 下拉可即時切換六種 chrome（功能相同，外觀不同）：

| 風格 | 特色 |
|------|------|
| **Microsoft Word** | 深藍標題列 `#2B579A`、快速存取按鈕 |
| **LibreOffice Writer** | 綠標 `#18A303`、灰工具列、右側屬性 |
| **Google 文件** | 白底、藍強調 `#1A73E8`、簡潔標題與儲存狀態 |
| **Zoho Writer** | 青藍 `#00A3E0`、雙側面板、狀態膠囊 |
| **WPS Writer** | 紅標 `#E81123`、現代白底工具列 |
| **FreeOffice TextMaker** | 暖灰／橘 `#E36C0A`、經典工具列質感 |

共通：
- **分頁功能列**：首頁 / 插入 / 格式 / 工具 / 校閱 / 檢視
- **左側文件導覽**、**右側屬性面板**（可開關；各風格預設不同）
- 灰畫布 + 置中白頁；狀態列：狀態 / 選取 / 字數 / 縮放

### 字型與段落
- 字型下拉（微軟正黑體、Calibri、Arial…）
- 粗體、斜體、底線（Ctrl+B / I / U）
- 字級放大縮小、字級下拉
- 左 / 中 / 右 / 兩端對齊
- 項目符號、編號清單
- 樣式：內文、標題 1–3

### 插入
- 段落、標題、表格（3×3）、圖片、刪除區塊

### 檢視
- 縮放 50%–200%
- 狀態列字數 / 字元 / 段落統計
- 圖片預覽（EMF/WMF 等不支援時顯示提示，不閃退）

## 開發建置

```powershell
dotnet build -c Release
.\bin\Release\net8.0\DocxAvalonia.exe
```

或：

```powershell
dotnet run -c Release
```

## 技術

| 項目 | 說明 |
|------|------|
| UI | Avalonia 11 + Fluent + MVVM |
| 文件 | DocumentFormat.OpenXml（`.docx`）、自研 ODT、DocSharp（`.doc`↔DOCX/RTF） |
| 編輯模型 | 段落／表格／圖片區塊（段落級格式） |

> 此為精簡本機編輯器，外觀可仿 Word／LibreOffice／Google 文件／Zoho／WPS／FreeOffice，**非**完整商業引擎；無雲端協作、留言、追蹤修訂等。
