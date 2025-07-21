# Excel 圖片提取器

此工具可從 Excel (.xlsx) 檔案中提取圖片，並建立與題號的對應關係。透過解析 Excel 的內部 XML 結構，自動辨識題號及對應的圖片位置。

## 功能特點

- 自動解壓縮 Excel 檔案並提取內嵌圖片
- 分析 XML 結構以找出圖片與題號的對應關係
- 基於章節號 (ChapterNo) 和題號 (QuestionNo) 建立唯一識別碼
- 完整支援繁體中文檔案路徑及內容
- 產生詳細的處理日誌與結果輸出
- 建立結構化的圖片對應表格

## 實作

- 使用 Python 標準函式庫的 zipfile 模組解壓縮 Excel 檔案
- 透過 ElementTree 直接解析 Excel XML 結構 (sheet1.xml, sharedStrings.xml)
- 利用 #VALUE! 錯誤標記定位圖片位置
- 採用行號與 VM 值建立智慧對應機制

## 系統需求

- Python 3.7 或更高版本
- 無需額外套件安裝 (使用 Python 標準函式庫)


## 使用方法

基本命令格式：

```bash
python extract_xlsx_images.py <xlsx_file_path> [選項]
```

**注意：`<xlsx_file_path>` 是必要參數，必須提供 Excel 檔案路徑，否則程式將無法執行。**

### 參數說明

| 參數 | 描述 |
|------|------|
| `xlsx_file` | **必要參數**。Excel 檔案的路徑，可使用絕對或相對路徑。如路徑包含空格，請用引號包圍。 |
| `--debug` | 啟用除錯模式，顯示更詳細的日誌資訊。 |
| `--keep-temp` | 保留處理過程中產生的暫存檔案，方便除錯。 |
| `--verbose`, `-v` | 顯示詳細的處理結果，包括更多題目內容資訊。 |
| `--help`, `-h` | 顯示說明訊息。 |

### 使用範例

基本用法：

```bash
python extract_xlsx_images.py "路徑/到/Excel檔案.xlsx"
```

啟用除錯模式：

```bash
python extract_xlsx_images.py "路徑/到/Excel檔案.xlsx" --debug
```

保留暫存檔案：

```bash
python extract_xlsx_images.py "路徑/到/Excel檔案.xlsx" --keep-temp
```

顯示詳細結果：

```bash
python extract_xlsx_images.py "路徑/到/Excel檔案.xlsx" --verbose
```

## 資料格式需求

Excel 檔案需包含以下欄位格式：

- **A欄**: 題號 (QuestionNo)
- **B欄**: 章節號 (ChapterNo)
- **C欄**: 題目內容

程式會根據 Excel 中圖片位置及儲存格內容建立對應關係。

## 執行結果

程式執行後會在當前目錄產生：
- `extracted_images/` 資料夾：存放所有提取出的圖片
- 詳細的執行日誌：包含題號列表和圖片對應關係
- 結構化的圖片對應表格：顯示圖片檔案名稱、章節號、題號及其他對應資訊


## 運作流程

1. **解壓縮 Excel 檔案**：使用 `zipfile` 模組將 .xlsx 檔案解壓縮到暫存目錄
2. **讀取 XML 內容**：分析 sheet1.xml 與 sharedStrings.xml 檔案
3. **題號識別**：從 A、B 欄位識別題號與章節號
4. **圖片位置分析**：尋找含有 #VALUE! 的儲存格，這些通常是圖片位置
5. **建立對應關係**：使用行號和 VM 值建立題號與圖片的對應關係
6. **複製與重命名**：將圖片複製到輸出目錄並依關聯性重命名
7. **產生報表**：輸出處理結果包含題號列表和圖片對應表格

## 注意事項

- 處理大型 Excel 檔案需要較長時間
- 預設會清除處理過程產生的暫存檔案，若需保留請使用 `--keep-temp` 參數
- 程式採用行號與 VM 值進行映射，若 Excel 結構特殊可能影響辨識準確度
- 圖片提取依賴 Excel 檔案中的內嵌圖片格式
