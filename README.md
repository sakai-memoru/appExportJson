# [Excel VBA] appExportJson : Excelで作成したデータをJSON形式でExportするツール

## overveiw :

- 指定の形式で作成した、Excelファイルを、JSON形式に変換し、出力する。  
    + 対象の形式  
        - Key-Value形式  
        - ListObject形式 (TBL形式)  
        - Named Range形式  

### 機能 :
- Excelにて、指定の形式で、データ表を作成し、指定フォルダ（input folder)に配置する。  
    + 指定の形式は、下記にて説明。
- メニューより、以下の機能を利用する。  
    + Export JSON (Key Value Sheet)  
    + Export JSON (ListObject:TBL)  
    + Export JSON (Named Ranges)
    
## 実行環境 :

- Microsoft Excel for Microsoft 365 MSO (16.0.13001.20142) 32-bit

## Installation :

- GitHubより、Cloneする。  
    +  https://github.com/sakai-memoru/appExportJson  

- 参照設定が必要。  

![image](https://gyazo.com/7d30f2387e7818067fd7596a82e507e9.png) 



## Usage :
- アプリは以下。
    + アプリ本体  ：appExportJson.xlsm  
        - Batch : ExportJsonMain.bas   
            + ExportJsonModule.bas 
                - GetKeyValue
                - GetTable
                - GetNames  
    + アプリconfig：config.json  

- appExportJson.xlsmを開く。  

![menu](https://gyazo.com/8b8535d17f900be082ecdffe2c6a502f.png)  


### 初期コンフィグ設定 :
   
```
{
    "BASE_FOLDER": "",
    "INPUT_FOLDER": "input",
    "OUTPUT_FOLDER": "output",
    "TEMP_FOLDER": "input/temp",
    "BACKUP_FOLDER": "input/backup",
    "FORM_FOLDER": "forms",
    "TRANSFORM_KEYVALUE": {
        "SHEET_TYPE": "KEYVALUE",
        "INPUT_LIKE": "KV*.xlsx",
        "TARGET_WORD": "no",
        "MACRO_GET_METHOD": "GetKeyValue"
    },
    "TRANSFORM_LISTOBJECT": {
        "SHEET_TYPE": "LISTOBJECT",
        "INPUT_LIKE": "TBL*.xlsx",
        "TARGET_INDEX": 1,
        "MACRO_GET_METHOD": "GetTable"
    },
    "TRANSFORM_NAMES": {
        "SHEET_TYPE": "NAMES",
        "INPUT_LIKE": "ApplicationForm*.xlsx",
        "TARGET_PARAM": "A1",
        "MACRO_GET_METHOD": "GetNames",
        "DETAIL_FIELD": "definition"
    },
    "CONTROL_PREFIX": "__",
    "SOURCE_FROM": "_source",
    "APP_NAME" : "appExportJson"
}
```  

### Environment :

![env](https://gyazo.com/de1d4b7da061302605815e150ca65658.png)

## Execution sample :

### Key Value Format
![keyval](https://gyazo.com/c1d99a5ffe5acbc253133936305c5fbc.png)

### ListObject(TBL) Format
![tbl](https://gyazo.com/331c0cf9ac4aeaf86602c13c969af11a.png)

![named](https://gyazo.com/08d3852d57dff8357eb2c48df886a55b.png)

### Named Range Format
![namedrange](https://gyazo.com/a0af01ccecd1f71f221c8fd117679d8c.png)

![named](https://gyazo.com/e321b7b8fcd4305064c529418dad3d09.png)


## application I/F :

```vb
'''' **********************************************
'' @file ExportJsonMain.bas
'' @parent appExportJson.xlsm
''

Public Sub Batch(ByVal datatype As String, Optional ByVal moveOn As Variant = False)
'''' **********************************************
'''' @function batch
'''' @param datatype {String} processing data type
''''        dictionary key in config.json
''''          "TRANSFORM_KEYVALUE" : key value formatted sheet
''''          "TRANSFORM_LISTOBJECT" : Tbl sheet
''''          "TRANSFORM_NAMES" : Named range sheet
'''' @param moveOn  {Variant<boolean>}
''''            a flag ot moving input files
''
```

## note :
- 落ち着いたら、もう少し記述を追加します。  

## reference :

- 以下の外部ライブラリを使用しています。  
  + VBA-JSON : JsonConverter.bas  
    - https://github.com/VBA-tools/VBA-JSON  
  + MiniTemplator  
    - https://www.source-code.biz/MiniTemplator/  

// --- end of README.md