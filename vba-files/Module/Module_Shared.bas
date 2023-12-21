Attribute VB_Name = "Module_Shared"
Option Explicit

' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----
' 共通のパブリックオブジェクト用モジュール
' ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----

' Configデータ格納
Public glb_Config As New Config

' DAO（Oracle以外のDBに接続する時は下記のDaoクラスを変更）
Public glb_Dao As New Dao_OracleOra

' Class生成用Factory
Public glb_Factory As New Factory
