# -*- mode: python ; coding: utf-8 -*-

import os
import sys


playwright_browsers = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'ms-playwright')


def collect_tree(src_root, dest_root):
    collected = []
    for root, _dirs, files in os.walk(src_root):
        rel_root = os.path.relpath(root, src_root)
        target_dir = dest_root if rel_root == '.' else os.path.join(dest_root, rel_root)
        for filename in files:
            collected.append((os.path.join(root, filename), target_dir))
    return collected


def conda_extra_binaries():
    # Anaconda/Minicondaでは tcl86t.dll 等の標準ライブラリ拡張モジュールが
    # 依存するDLLが <prefix>\Library\bin にあり、標準レイアウト
    # (pythonw.exeと同じフォルダ)を前提にするPyInstallerの依存解決が
    # 見つけられないことがある(ビルド時の "Library not found" 警告で判明)。
    # 存在すれば明示的に同梱する。標準のpython.orgインストールでは
    # 通常見つからず、その場合は何もしない。
    # libssl/libcrypto を外すと _ssl.pyd が読めず、EXEでHTTPS通信が
    # "SSL module is not available" で全滅する。必ず入れる。
    bin_dir = os.path.join(sys.base_prefix, 'Library', 'bin')
    names = ('tcl86t.dll', 'tk86t.dll', 'ffi.dll', 'libexpat.dll',
              'libmpdec-4.dll', 'liblzma.dll', 'libbz2.dll', 'sqlite3.dll',
              'libssl-3-x64.dll', 'libcrypto-3-x64.dll')
    return [(os.path.join(bin_dir, n), '.') for n in names
            if os.path.isfile(os.path.join(bin_dir, n))]


def chromium_only_browser_dirs(browsers_root):
    # アプリは pw.chromium.launch() しか使わない(firefox/webkitは未使用)ため、
    # それらと、キャッシュに残る旧リビジョンは同梱対象から外してサイズを抑える。
    # 各プレフィックス(例: "chromium", "chromium_headless_shell")ごとに
    # リビジョン番号が最大のものだけを残す。
    if not os.path.isdir(browsers_root):
        return []
    import re
    best = {}
    for name in os.listdir(browsers_root):
        if not os.path.isdir(os.path.join(browsers_root, name)):
            continue
        m = re.match(r'^(chromium(?:_headless_shell)?)-(\d+)$', name)
        if not m:
            continue
        prefix, rev = m.group(1), int(m.group(2))
        if prefix not in best or rev > best[prefix][0]:
            best[prefix] = (rev, name)
    return [name for _rev, name in best.values()]


datas = []
if os.path.isdir(playwright_browsers):
    for dirname in chromium_only_browser_dirs(playwright_browsers):
        datas += collect_tree(os.path.join(playwright_browsers, dirname),
                               os.path.join('ms-playwright', dirname))
binaries = conda_extra_binaries()


a = Analysis(
    ['LoanWatchGui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=['LoanWatchMain', 'PlayersEditor', 'playwright.async_api', 'playwright.sync_api', 'truststore'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'torchaudio', 'tensorflow', 'tf_keras', 'keras', 'pygame', 'matplotlib', 'scipy', 'sklearn', 'IPython', 'jupyter', 'notebook', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'cv2', 'pyarrow', 'dask', 'distributed', 'xarray', 'h5py', 'bokeh', 'plotly', 'geopandas', 'shapely', 'pyproj', 'cupy', 'numba', 'llvmlite', 'altair', 'duckdb', 'pyspark'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='LoanWatch',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
