# -*- mode: python ; coding: utf-8 -*-
"""WordAgent 打包配置：python -m PyInstaller --noconfirm --clean WordAgent.spec"""
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

datas = []
# 自定义标准（agent/standards.py 运行时按相对路径加载）
datas += [("standards", "standards")]
# 品牌图标
datas += [("assets/icon.ico", "assets")]
# OCR 模型资源（rapidocr_onnxruntime 的 onnx 模型 + 配置）
datas += collect_data_files("rapidocr_onnxruntime")

hiddenimports = []
# rapidocr 及其依赖链
hiddenimports += collect_submodules("rapidocr_onnxruntime")
hiddenimports += collect_submodules("onnxruntime")
hiddenimports += ["PIL", "PIL.Image", "PIL.ImageDraw", "numpy", "cv2"]
# 文档处理
hiddenimports += ["docx", "docx.opc", "docx.oxml", "docx.oxml.ns", "fitz", "dotenv"]

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "cryptography", "IPython", "pytest", "tkinter.test"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="WordAgent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    icon="assets/icon.ico",
    version="version_info.txt",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="WordAgent",
)
