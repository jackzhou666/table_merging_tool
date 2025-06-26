# Excel 表格合并工具

## 功能介绍

本工具用于批量合并多个 Excel 文件（.xlsx），适用于需要将多个表格数据整合到一个文件中的场景。主要功能如下：

- 支持选择多个 Excel 文件进行合并。
- 自动识别并合并所有表格的内容，支持多表头、不同列数的情况。
- 合并结果可导出为新的 Excel 文件，方便后续处理。
- 操作简单，无需复杂配置，适合非技术用户。

## EXE 打包流程

本工具基于 Python 开发，可通过 pyinstaller 打包为 Windows 下的可执行文件（.exe），具体步骤如下：

### 1. 安装依赖

首先确保已安装 Python 3.x。推荐使用虚拟环境。

```bash
pip install pyinstaller pandas openpyxl
```

### 2. 打包命令

在命令行进入项目目录，执行以下命令：

```bash
pyinstaller --onefile merge_excel.py
```

- `--onefile` 参数表示打包成单一的 exe 文件。
- 打包完成后，`dist` 文件夹下会生成 `merge_excel.exe`。

### 3. 运行方法

将需要合并的 Excel 文件与 `merge_excel.exe` 放在同一目录下，双击运行 exe 文件，按照提示操作即可。

### 4. 注意事项

- 打包后的 exe 文件首次运行时可能会有短暂延迟，属于正常现象。
- 若合并过程中遇到格式不一致等问题，请检查源文件格式是否规范。
- 如需自定义功能，可修改 `merge_excel.py` 脚本后重新打包。

---

如需进一步帮助，请联系开发者。 