# 字体说明

模板使用两套免费字体（SIL Open Font License）：

- **Noto Serif SC**（思源宋体）：标题、数字、标签
- **Noto Sans SC**（思源黑体）：正文

## 安装

- **Windows / macOS**：从 Google Fonts 下载 Noto Serif SC 和 Noto Sans SC 并安装。较新的 Windows 11 已自带这两套字体。
- **Linux**：`sudo apt install fonts-noto-cjk`。该包里的字体族名是 Noto Serif CJK SC / Noto Sans CJK SC，
  和模板里写的名字不同。把本目录的 `60-sdu-noto.conf` 复制到 `~/.config/fontconfig/conf.d/`，
  再执行 `fc-cache -f`，LibreOffice 渲染时就会使用正确的字体。

## 换电脑演示

没装这两套字体的电脑会用别的字体顶替，排版可能走样。可以任选一种办法：

1. 在演示电脑上安装字体；
2. 导出 PDF 演示；
3. PowerPoint「文件 → 选项 → 保存 → 将字体嵌入文件」。两套字体完整嵌入后文件约 56 MB，
   勾选「仅嵌入演示文稿中使用的字符」会小很多，但在没装字体的电脑上就不能再编辑新文字。
