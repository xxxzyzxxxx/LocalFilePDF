# 第三方组件

本工具的转换代码和界面代码随源码提供。便携版动态加载第三方库；请保留 `_internal` 内容和随包许可。没有静态链接或修改 Qt。

| 组件 | 用途 | 上游许可 |
|---|---|---|
| Python | 运行时 | PSF License |
| PySide6 / Shiboken6 / Qt | GUI | LGPL-3.0 / GPL / 商业许可，使用开放许可版本 |
| Pillow | JPEG / PNG 解码 | HPND |
| ReportLab | 图片 PDF 生成 | BSD |
| pypdf | PDF 校验 | BSD-3-Clause |
| pywin32 | Word COM 自动化 | PSF |
| psutil | 自有子进程清理 | BSD-3-Clause |
| PyInstaller | 构建 | GPL-2.0-or-later with bootloader exception |

PyMuPDF、python-docx 和 pytest 仅用于开发验证，不属于应用运行代码。Microsoft Word 和 LibreOffice 由用户自行安装，其许可独立于本工具。

第三方完整许可文件在便携目录 `licenses` 下。允许按对应许可替换动态库；源码和构建脚本随工程交付。Qt 源代码可从 https://code.qt.io/ 获取，PySide 源代码可从 https://code.qt.io/pyside/pyside-setup.git/ 获取，版本以 requirements-lock.txt 为准。
