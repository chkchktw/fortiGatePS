# Fortigate Config Parser

## 介紹

這是一個專門解析 Fortigate 設定檔的小工具，可以將設定內容轉換成清楚的 Excel 活頁格式，方便後續管理、審核或稽核使用。

支援功能：
- 解析 `address` 設定，列出所有 Address 名稱與對應 IP
- 解析 `addrgrp` 設定，列出群組名稱、成員名稱與對應 IP
- 解析 `policy` 設定，列出策略 ID、名稱、來源介面、目的介面、來源地址、目的地址、服務、動作、狀態
- 自動產生 `.xlsx` Excel 檔案，每個項目一個活頁分頁

## 使用說明

### 環境需求
- Python 3
- openpyxl 套件 (`pip install openpyxl`)

### 執行方式

```bash
python fortips.py 設定檔.conf 輸出檔.xlsx
```

範例：
```bash
python fortips.py firewall_backup.conf result.xlsx
```

### 檔案說明
- **address** 活頁：列出所有 address 名稱與 IP
- **group** 活頁：列出所有 group 名稱、成員、IP
- **policy** 活頁：列出防火牆策略的詳細設定資訊

## 注意事項
- 活頁名稱最多 31 字元，超過會自動截斷
- 找不到 IP 的成員會標記為 `N/A`
- 設定檔需包含 `config firewall address`、`config firewall addrgrp`、`config firewall policy` 區段

## 授權
MIT License

---

如果覺得好用，歡迎點個星星 ⭐ 或提出 PR 改善功能！
