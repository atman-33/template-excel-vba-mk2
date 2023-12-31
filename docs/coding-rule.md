## 命名規則

### 共通

#### プライベート変数	
- m_ + キャメルケース

e.g. :  
```txt
m_tableName
```

#### 定数	
- 大文字のスネークケース	

e.g. : 
```txt
CONFIG_SHEET
```

#### パブリック変数
- glb_ + パスカルケース

e.g. :
```text
glb_Config
```

#### モジュール、メソッド内の変数	
- キャメルケース

> VBAではデフォルトモジュールのプロパティやメソッドと同名をDimで指定した場合、
> 既に指定済みのプロパティやメソッド名も変更されてしまうため注意。
> その場合は、var を付けて別名にする。

e.g. :
```txt
tableName
varItem
```

#### メソッド	
- パスカルケース	

e.g. :
```
InitConfig
```

### 標準モジュール関連

#### モジュール名称
- Module + _ + ○○	

e.g. :
```txt
Module_Common
Module_Main
```

### クラスモジュール関連

類似した意味を持つクラスは、`_`で区切ってグルーピングする。

e.g. :  
```txt
Table_oo
Sheet_oo
```

#### DB接続クラス	
- Dao + _ + ○○	

e.g. :  
```txt
Dao_OracleOra
```

#### ベースクラス	
- oo + Base	

e.g. : 
```txt
TableBase
```
