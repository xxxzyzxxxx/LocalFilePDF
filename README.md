# LocalFilePDF · Windows 本地文件转 PDF

将 Word 文档分别导出为 PDF，或将多张 JPG、JPEG、PNG 按指定顺序合并为一个 PDF。文件在本机处理，保留原文件，同名输出自动编号。

## 下载和使用

1. 打开本仓库的 **Releases**，下载便携版 ZIP，完整解压。
2. 双击 **LocalFilePDF.exe**，选择 **Word 转 PDF** 或 **图片转 PDF**。
3. 拖入文件。合并图片时可以调整顺序、纸张方向和边距。
4. 点击 **开始转换**。默认结果保存在首个输入文件旁的 **PDF输出** 文件夹。

便携版适用于 64 位 Windows，无需安装 Python。图片转换无需 Office；Word 转换优先使用本机 Microsoft Word，也支持检测 LibreOffice。已在 Windows 11 和 Microsoft Word 上实测，未实测 LibreOffice。

[详细使用及构建说明](使用手册.md)
