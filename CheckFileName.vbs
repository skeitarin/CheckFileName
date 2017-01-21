Option Explicit
 
Dim objFileSys
Dim objFolder
Dim objFile
Dim objOutputTextStream
 
'ファイルシステムを扱うオブジェクトを作成
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'ログ出力用 TextStream オブジェクトを作成
'第2引数は 1 ：読み取り、2 ：上書き、3 ：追記。
Dim logFileNm
Dim nowdate
nowdate= Year(Now())
nowdate= nowdate & Right("0" & Month(Now()) , 2)
nowdate= nowdate & Right("0" & Day(Now()) , 2)
nowdate= nowdate & Right("0" & Hour(Now()) , 2)
nowdate= nowdate & Right("0" & Minute(Now()) , 2)
nowdate= nowdate & Right("0" & Second(Now()) , 2)
logFileNm = "log_" & CStr(nowdate) & ".txt"
' logFileNm = "loga" & ".txt"
Set objOutputTextStream = objFileSys.OpenTextFile(logFileNm, 2, True)

Dim filePath
filePath = InputBox("比較対象のファイルパスの入力（選択したファイル内に記載のないファイル名をログに出力します）")
Dim objFSO, objFile2, line
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile2 = objFSO.OpenTextFile(filePath)

'フォルダのオブジェクトを取得
Dim dirPath
dirPath = InputBox("検査対象のフォルダパスの入力")
Set objFolder = objFileSys.GetFolder(dirPath)

Dim exist
 objOutputTextStream.WriteLine "定義ファイルパス：" & filePath
 objOutputTextStream.WriteLine "検査対象フォルダパス：" & dirPath
 objOutputTextStream.WriteLine "---- 存在しないファイル名のリスト ----"

Do Until objFile2.AtEndOfStream
    exist = False
    line = objFile2.ReadLine
    ' objOutputTextStream.WriteLine "log：" & line
    For Each objFile In objFolder.Files
        ' objOutputTextStream.WriteLine "_log：" & objFile.Name
        if line = objFile.Name then
            exist = True
            Exit For
        end if
    Next
    if exist = False then
        'ログファイルに出力
        objOutputTextStream.WriteLine line
    end if
Loop
 
'TextStream は Close を忘れずに
objOutputTextStream.Close
 
Set objOutputTextStream = Nothing
Set objFolder  = Nothing
Set objFileSys = Nothing 

WScript.Echo "処理が終了しました。結果がログファイルを確認してください。"