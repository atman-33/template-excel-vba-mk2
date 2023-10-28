## モジュール構成

```text
|- 標準モジュール/
|   |- Constants        : 共通の定数を記載
|   |- ModuleCommon     : Staticな共通処理を記載
|   |- Tests            : ユニットテスト用コード置き場
|   |- Utils            : 汎用メソッド置き場
|
|- クラスモジュール/
|   |- Config           : Excelシートに指定したConfigを扱うクラス
|   |- DaoAccess        : Access に接続し、CRUD処理を行うDAOクラス
|   |- DaoOracleOra     : Oracle に接続し、CRUD処理を行うDAOクラス
|   |- ExListObject     : テーブル（ListObject）の機能拡張クラス
|   |- IDao             : DAOクラスのインターフェース
|   |- Repository       : DAOクラスを利用するリポジトリクラス
|   |- TableXxx         : Excelテーブルを操作するクラス
|   |- SheetXxx         : Excelシートを操作するクラス
|

```