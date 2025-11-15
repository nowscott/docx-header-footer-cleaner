# docx-header-footer-cleaner

移除 Word（docx）文档的页眉与页脚，并在页脚居中插入自动页码。支持批量、递归处理与安全备份。

## 快速开始
- 环境要求：`Python 3.8+`
- 安装依赖：
  - 使用系统 Python：`pip install python-docx`
  - 建议虚拟环境：
    - macOS/Linux：
      - `python3 -m venv .venv && source .venv/bin/activate`
      - `pip install -U pip && pip install python-docx`
    - Windows：
      - `python -m venv .venv && .venv\\Scripts\\activate`
      - `pip install -U pip && pip install python-docx`

## 运行方式一：单文件处理
- 输入一个 `*.docx` 文件，生成一个新的已清理版本：
  - `python3 docx_header_footer_tool.py /path/to/file.docx -o /path/to/output.docx`
- 若不传 `-o`，默认在同目录生成 `file_paged.docx`。
- 处理逻辑：移除所有节的页眉/页脚（含首页/偶数页变体），并在页脚居中插入页码字段。

## 运行方式二：批量处理（推荐）
- 准备配置文件，脚本将递归处理多个根目录下的所有 `*.docx`：
 1) 创建 `docx_config.local.txt`（优先读取），若不存在则读取 `docx_config.txt`。
 2) 写入内容（每行一项）：
    - 第一类为键值：`backup=./docx_backup` 指定备份根目录（可改为绝对路径）。
    - 其余为待处理的根目录路径，脚本会递归处理其子目录中的 `*.docx`。
 3) 运行：
    - `python3 docx_header_footer_tool.py`
    - 或显式指定：`python3 docx_header_footer_tool.py --config docx_config.local.txt`
    - 临时覆盖备份目录：`python3 docx_header_footer_tool.py --backup /path/to/backup`

### 示例配置
```
backup=./docx_backup
/path/to/your/word/folder
/another/root/folder
```

## 输出与备份
- 原地覆写：每个文档会在原路径原名上被替换为“清理+加页码”的版本。
- 安全备份：处理前，会将原文件复制到备份根目录，保持原有层级结构。
- 备份位置示例：
  - 项目根为 `/project`，`backup=./docx_backup` 时，备份位于 `/project/docx_backup/<每个根目录的基本名>/...`。
- 回滚方法：将备份中的文件复制回原路径即可。

## 常见问题
- 仅支持 `*.docx`；`*.doc` 请先在 Word 中另存为 docx。
- 页码未显示：在 Word 中全选并按 `F9` 更新域。
- 文件占用导致失败：关闭正在打开的文档后重试。
- 带“冲突副本”后缀的拷贝：建议手动确认可正常打开保存后再处理。
- Windows 路径：用反斜杠或加引号，示例：`"D:\\docs\\math"`。

## 高级用法
- 指定单次备份目录（不改动配置）：`--backup /path/to/backup`
- 仅统计 `*.docx` 数量（macOS/Linux）：
  - 原目录：`find "/path/to/root" -type f -name "*.docx" | wc -l`
  - 备份目录：`find "/project/docx_backup/<root-basename>" -type f -name "*.docx" | wc -l`
- 差异核对（macOS/Linux）：
  - `find "/path/to/root" -type f -name "*.docx" | sed "s|/path/to/root/||" > /tmp/orig.list`
  - `find "/project/docx_backup/<root-basename>" -type f -name "*.docx" | sed "s|/project/docx_backup/<root-basename>/||" > /tmp/backup.list`
  - `grep -Fxv -f /tmp/backup.list /tmp/orig.list`

## 工作原理速览
- 清理页眉/页脚：`clear_hf` 在 `docx_header_footer_tool.py:9–12`
- 插入居中页码：`add_center_page_number` 在 `docx_header_footer_tool.py:14–30`
- 单文档处理：`process_document` 在 `docx_header_footer_tool.py:32–60`
- 原地覆写：`process_in_place` 在 `docx_header_footer_tool.py:96–108`
- 备份与递归遍历：`backup_file` 在 `docx_header_footer_tool.py:88–94`；`process_roots` 在 `docx_header_footer_tool.py:110–128`
- 命令行入口与配置优先级：`main` 在 `docx_header_footer_tool.py:130–166`

## 注意
- 所有节均会解除“与前一节链接”和“首页/偶数页不同”的设置后统一清理。
- 会删除页眉/页脚中的全部内容（包括文字、图片、页码、页眉线等）。
- 仅在页脚插入居中页码（Word 字段 `PAGE`）；暂不支持自定义位置或样式。

## 许可证
MIT