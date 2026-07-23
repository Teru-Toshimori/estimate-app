# 古い成果物を削除
Remove-Item build -Recurse -Force -ErrorAction Ignore
Remove-Item dist -Recurse -Force -ErrorAction Ignore

# EXE作成
pyinstaller `
    --clean `
    --noconfirm `
    --windowed `
    --onefile `
    --name EstimateTool `
    main.py

# 配布フォルダ作成
New-Item -ItemType Directory -Path dist\EstimateTool -Force | Out-Null

# EXE
Move-Item dist\EstimateTool.exe dist\EstimateTool\EstimateTool.exe

# リソース
Copy-Item resources dist\EstimateTool\resources -Recurse

# 環境設定
Copy-Item .env dist\EstimateTool\.env