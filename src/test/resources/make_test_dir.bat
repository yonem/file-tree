@echo off
setlocal
mkdir test_structure
cd test_structure

:: 1. 深い階層
mkdir "01_Deep\Level2\Level3"
echo sample > "01_Deep\Level2\Level3\deep_sample.txt"

:: 2. 空のディレクトリ
mkdir "02_Empty_Directory"

:: 3. 複数のファイル
mkdir "03_Files"
echo file1 > "03_Files\file_1.log"
echo file2 > "03_Files\file_2.log"

:: 4. 今日の日付のExcelファイル (手動で日付を合わせるか、簡易的な作成)
:: ※バッチでの日付取得は環境依存が強いため、空ファイルのみ作成します
echo manual_fix > "tree_placeholder.xlsx"

:: 5. ルート直下のファイル
echo root > "root_sample.txt"

echo Done.
pause