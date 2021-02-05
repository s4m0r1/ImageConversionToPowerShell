# 解像度変更ツール

## 使い方
1. 設定手順  
main.ps1をテキストエディタで開く  
`$EXPORT_PATH`の""にエクスポート先のフルパスを記入  
例: C:\Users\hoge\Documents\ExportSlide\  
必ず最後に\を入れてください

2. ショートカットを作成  
作成したショートカットのプロパティを開きショートカット先の頭に  
`powershell -NoProfile -ExecutionPolicy RemoteSigned -File`を記入  
必ず`-File`の後に半角スペースをいれること

3. 画像ファイルをショートカットにドラッグアンドドロップ

4. エクスポート先のフォルダ名を入力

5. ファイル名の入力
XXXXX_XXXXXX_XXXXXXX  
アンダースコアで分けた3つに左から順に入力  
大文字のSが入っている場合はYesを選択

6. 終わるまで触らない

7. できあがり