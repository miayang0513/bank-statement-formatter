# 銀行明細統一格式化工具

這個工具可以讀取來自4間不同銀行（Monzo、Revolut、Wise、Amex）的明細文件，並合併統一輸出為CSV格式，方便導入到Google Spreadsheet。

## 功能特點

- ✅ 支持4家銀行：Monzo、Revolut、Wise、Amex
- ✅ 自動識別和讀取不同格式的文件（CSV、XLSX）
- ✅ 統一字段格式，按日期排序
- ✅ 輸出CSV格式，可直接導入Google Spreadsheet
- ✅ 支持中文顯示（UTF-8-BOM編碼）

## 安裝

### 1. 安裝依賴

使用 `pipenv` 管理依賴：

```bash
pipenv install
```

如果沒有安裝 `pipenv`，可以先用 Homebrew 安裝：

```bash
brew install pipenv
```

或者使用傳統的 `pip`：

```bash
pip install pandas openpyxl
```

## 使用方法

### 1. 準備銀行明細文件

將各銀行的明細文件放在對應月份的目錄下，例如：

```
statements/
  └── 202510/
      ├── amex- oct.xlsx
      ├── monzo-oct.csv
      ├── revolut-oct.csv
      └── wise-oct.csv
```

### 2. 運行格式化工具

使用 `pipenv` 運行：

```bash
pipenv run python formatter.py statements/202510
```

或者如果已激活虛擬環境：

```bash
python formatter.py statements/202510
```

### 3. 輸出結果

工具會在對應月份目錄下生成 `combined_statements.csv` 文件，包含以下字段：

- `date`: 交易日期和時間（YYYY-MM-DD HH:MM:SS格式）
- `amount`: 交易金額（正數表示收入，負數表示支出）
- `currency`: 貨幣代碼
- `description`: 交易描述
- `category`: 交易類別
- `bank`: 銀行名稱
- `type`: 交易類型
- `transaction_id`: 交易ID

### 4. 導入到Google Spreadsheet

1. 打開 Google Spreadsheet
2. 點擊 `文件` > `導入`
3. 選擇 `上傳` 標籤
4. 上傳生成的 `combined_statements.csv` 文件
5. 選擇導入設置：
   - 導入位置：`替換電子表格`
   - 分隔符類型：`自動檢測` 或 `逗號`
   - 轉換文本：`是`
6. 點擊 `導入數據`

## 文件格式說明

### Monzo CSV
- 文件格式：CSV
- 必需字段：Date, Time, Amount, Currency, Name, Category

### Revolut CSV
- 文件格式：CSV
- 必需字段：Completed Date, Amount, Fee, Currency, Description

### Wise CSV
- 文件格式：CSV
- 必需字段：Status, Direction, Finished on, Source amount (after fees), Target amount (after fees), Category

### Amex XLSX
- 文件格式：Excel (.xlsx)
- 自動檢測日期、金額、描述等字段

## 注意事項

- 確保所有文件都在正確的月份目錄下
- 文件名需要包含銀行名稱（monzo、revolut、wise、amex）
- 只處理狀態為 `COMPLETED` 和 `REFUNDED` 的交易（Wise）
- 日期格式會自動識別和轉換

## 示例輸出

```
date,amount,currency,description,category,bank,type,transaction_id
2025-10-01 10:00:00,-6.29,GBP,Subway,Eating out,Monzo,Card payment,tx_0000...
2025-10-02 15:30:00,100.00,GBP,Transfer from Friend,General,Wise,TRANSFER-...
```

## 許可證

MIT License

