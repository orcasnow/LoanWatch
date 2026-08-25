含めたもの:
アプリ実行用: requests, pandas, openpyxl, playwright, beautifulsoup4, lxml, truststore など
PyInstallerビルド用: pyinstaller, pyinstaller-hooks-contrib, pefile, pywin32-ctypes, altgraph, setuptools など
確認も済みです。
python -m pip install --dry-run -r .\requirements.txt
補足として、Playwright の Chromium は requirements.txt だけでは入りません。環境再現時はインストール後に以下も実行してください。
python -m playwright install chromium
ビルドはこれです。
python -m PyInstaller --noconfirm --clean .\LoanWatch.spec