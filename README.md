# CheckFileName
フォルダ内に特定のファイル名が存在するかのチェックを行う
（同一階層のファイルしかチェックできない）

## ①定義ファイルの作成（.csv）
チェックを行うファイル名をCSVファイルに記述する。ファイル名は改行で区切る。

## ②起動する

ファイルパスには①で作成したファイルを指定する。

フォルダパスには、チェック対象のフォルダパス（FP）を指定する。

①で定義したファイル内に記載されているファイル名がFP内に存在しない場合に、ログファイルにそれを出力する。
