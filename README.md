# DocxAvalonia

以 **Avalonia UI** 打造的 Word 風格 `.docx` 編輯／預覽桌面程式。

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
- 拖放 `.docx` 開啟
- 命令列傳入路徑開啟
- **自動儲存**（可開關，每 60 秒；需已有儲存路徑）

### 編輯
- 剪下 / 複製 / 貼上
- 復原 / 重做
- 尋找、取代、全部取代

### 字型與段落
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
| 文件 | DocumentFormat.OpenXml 讀寫 `.docx` |
| 編輯模型 | 段落／表格／圖片區塊（段落級格式） |

> 此為精簡 Word 工作站，非完整 Microsoft Word 引擎；複雜版面、文字方塊、追蹤修訂等未完整支援。
