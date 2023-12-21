## モジュール構成

```text
|- Excel Objects（Sheet）/
|   |- Sheet                : Sheet上のボタンなど、UIから操作した際の処理を格納
|
|- フォーム/
|   |- XxxForm              : Formの処理を格納
|
|- 標準モジュール/
|   |- Constants            : 共通の定数を記載
|   |- Lib                  : 共通で利用するビジネスロジックを格納
|   |- Main_Xxx             : Sheet上のボタンなど、UIから操作した際の処理を格納
|   |- Module_DataAccessDB  : DB操作を行う一般処理
|   |- Module_Shared        : パブリックなオブジェクトを格納
|   |- Tests                : ユニットテスト用コード
|   |- Utils                : 汎用メソッドを格納
|
|- クラスモジュール/
|   |- Config               : Excelシートに指定したConfigを扱うクラス
|   |- Dao_Access           : Access に接続し、CRUD処理を行うDAOクラス
|   |- Dao_OracleOra        : Oracle に接続し、CRUD処理を行うDAOクラス
|   |- Factory              : インスタンス生成Factory
|   |- IDao                 : DAOクラスのインターフェース
|   |- Repository           : DAOクラスを利用するリポジトリクラス
|   |- Table_Xxx            : Excelテーブルを操作するクラス
|   |- TableBase            : テーブル（ListObject）の基底クラスの代用（共通処理を格納）
|   |- Sheet_Xxx            : Excelシートを操作するクラス
|

```

## モジュール参照ルール

- Main_Xxx から Module_xx や Lib の参照は可能だが、逆は不可
- クラスモジュールはFactory経由でインスタンス生成すること（VBAでは引数有りのコンストラクタが使えないため）
- クラスに継承したい場合は XxxBase をクラスのコンストラクタでNewして保持させる（コンポジションの関係）