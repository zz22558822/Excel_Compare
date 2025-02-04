# Excel 檔案比對
## Excel_Compare


## 使用方法:
1. 首次運行先開啟應用後會產出 config_compare.json
2. 依照需求設定 Json
3. 運行 Excel_Compare.py


## 檔案結構
root/  
├── Excel_Compare.py → 主程序  
├── config_compare.json → 設定檔  
├── Data_1.xlsx → 比較的 Excel-1  
├── Data_2.xlsx → 比較的 Excel-2  
└── Compare_Result.txt → 對比差異的輸出紀錄(檔名可設定)  

## 設定值
``` json
{
	'Data_1': 'Data_1.xlsx',  # 比較的 Excel-1
	'Data_2': 'Data_2.xlsx',  # 比較的 Excel-2
	'Data_1_Sheet': 'Data',  # Excel-1 的分頁
	'Data_2_Sheet': 'Data', # Excel-2 的分頁
	'OUTPUT_FILE': 'Compare_Result.txt', # 對比差異的輸出紀錄
	'Scan_Method': 0  # 是否使用逐行掃描 (較慢但詳細)
}
```
