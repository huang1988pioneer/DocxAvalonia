# DocxAvalonia

以 **Avalonia UI** 打造的 Word 風格 `.docx` 編輯／預覽桌面程式。

## 常用功能

### 檔案
- 新增、開啟、儲存、另存新檔
- 拖放 `.docx` 開啟
- 命令列傳入路徑開啟

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
- 段落、標題、表格（3×3）、刪除區塊

### 檢視
- 縮放 50%–200%
- 狀態列字數 / 字元 / 段落統計

## 執行

```powershell
dotnet build -c Release
.\app-out\Release\WordViewHost.exe
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
| 編輯模型 | 段落／表格區塊（段落級格式） |

> 說明：此為精簡 Word 工作站，非完整 Microsoft Word 引擎；複雜版面、文字方塊、追蹤修訂等未支援。格式以「目前選取段落」為單位套用。
