## Excel VBA テンプレートファイル

## 主な実装機能

- ConfigシートのデータをVBAで利用
- Oracleデータベースに接続し、Excelテーブルにデータ格納/削除する機能
- Excelテーブル操作機能（ソート、フィルター）

## モジュール構成

```text
|- 標準モジュール/
|   |- Constants            : 共通の定数を記載
|   |- Module_DataAccessDB  : DB操作を行う一般処理
|   |- Tests                : ユニットテスト用コード置き場
|   |- Utils                : 汎用メソッド置き場
|
|- クラスモジュール/
|   |- Config               : Excelシートに指定したConfigを扱うクラス
|   |- Dao_Access           : Access に接続し、CRUD処理を行うDAOクラス
|   |- Dao_OracleOra        : Oracle に接続し、CRUD処理を行うDAOクラス
|   |- IDao                 : DAOクラスのインターフェース
|   |- Repository           : DAOクラスを利用するリポジトリクラス
|   |- Table_Xxx            : Excelテーブルを操作するクラス
|   |- TableBase            : テーブル（ListObject）の基底クラスの代用（共通処理を格納）
|   |- Sheet_Xxx            : Excelシートを操作するクラス
|

```
