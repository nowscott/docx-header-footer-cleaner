# docx-header-footer-cleaner

移除 Word（docx）文档的页眉与页脚，并在页脚居中插入自动页码。支持批量、递归处理与安全备份。

## 功能
- 清除所有节的页眉与页脚（含首页/偶数页变体）。
- 在每节页脚新增居中页码（Word 字段 `PAGE`）。
- 读取配置文件，递归处理多个目录下的所有 `*.docx`。
- 处理前自动备份到指定目录，处理后原地覆写，文件名不变。

## 使用
1. 安装依赖：`pip install python-docx`
2. 准备配置文件（默认读取 `docx_config.local.txt`，不存在则读取 `docx_config.txt`）：
   - `backup=./docx_backup` 备份根目录（可改为任意路径）
   - 每行一个待处理的根目录路径，将递归处理其子目录中的 `*.docx`
3. 运行：
   - 最简：`python docx_header_footer_tool.py`
   - 指定配置：`python docx_header_footer_tool.py --config <配置路径>`
   - 临时覆盖备份目录：`python docx_header_footer_tool.py --backup <备份路径>`

## 示例配置 `docx_config.txt`
```
backup=./docx_backup
/path/to/your/word/folder
```

若需使用私有路径，请将其写入 `docx_config.local.txt`（不会提交到远端），并直接运行脚本。

## 注意
- 仅支持 `*.docx`；`*.doc` 请先在 Word 中另存为 docx。
- 打开后若页码未立即显示，在 Word 中全选并按 `F9` 更新域。

## 许可证
MIT