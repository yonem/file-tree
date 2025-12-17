## 動作環境

- Java21で動作

2023/9/29現在  
https://www.oracle.com/java/technologies/downloads/#java21

### 使用ライブラリ

- Apache POI

Excel2007形式（ooxml）でExcel関連の処理を行う。

```xml

<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>
```

## 操作方法

1. 起動時にディレクトリ選択のダイアログを表示
2. ツリーの起点となるディレクトリを選択する
3. 選択したディレクトリの直下に収集結果が出力される

![出力イメージ](img/output_image.png)

---

## 設計図 (System Design)

### 処理フロー

```plantuml
@startuml
title FileTreeFrame 処理フロー (onSubmit)

start
:出力ボタン押下;
if (ディレクトリパスが空?) then (yes)
  stop
else (no)
  :SwingWorker開始 (非同期処理);
  :出力ボタンを無効化;
  
  if (Excelチェックボックスが選択されている?) then (yes)
    :コンソールに "Start!!" を表示;
    partition "Excel出力処理" {
      :ExcelUtil.convertDir2Treeを実行;
    }
    :完了メッセージをダイアログとコンソールに表示;
  else (no)
    partition "コンソール出力処理 (再帰)" {
      :outputConsole(root, 0, "", false) を呼び出し;
    }
  endif

  if (例外発生?) then (yes)
    :エラーメッセージをコンソールに表示;
  endif

  :出力ボタンを有効化;
  :SwingWorker終了;
endif
stop

partition "outputConsole (再帰関数)" {
  :ファイル名を表示;
  if (インデント > 0) then (yes)
    :階層記号 (├─ / └─) を付与;
  endif
  
  if (対象はファイル?) then (yes)
  else (ディレクトリ)
    :配下のファイルリストを取得;
    :ディレクトリ優先 > 名前順でソート;
    while (要素があるか?) is (yes)
      :次の要素に対して **outputConsole** を再帰呼出;
    endwhile
  endif
}
@enduml
```

### シーケンス図

```plantuml
@startuml
title FileTreeFrame 処理シーケンス図

actor User
participant "FileTreeFrame\n(UI Thread)" as UI
participant "SwingWorker\n(Background Thread)" as Worker
participant "File" as File
participant "ExcelUtil" as Excel

User -> UI : 「出力」ボタンをクリック
activate UI

UI -> UI : onSubmit() 実行
create Worker
UI -> Worker : 生成・execute()
activate Worker

UI -> UI : btnSubmit.setEnabled(false)

Worker -> Worker : doInBackground() 実行

alt Excel出力モード
Worker -> Excel : convertDir2Tree(rootDirectory)
activate Excel
Excel --> Worker : 完了
deactivate Excel
else コンソール表示モード
Worker -> UI : outputConsole() 再帰実行
loop 全ファイル探索
Worker -> File : 情報取得
Worker -> UI : 画面更新(append)
end
end

Worker -> UI : btnSubmit.setEnabled(true)
deactivate Worker
deactivate UI
@enduml
```
