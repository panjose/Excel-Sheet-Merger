# Excel-Sheet-Merger: 批量 Excel 文件合并工具
A simple Python script using Pandas to merge multiple Excel files in a folder into a single Excel file with multiple sheets.

-----


## 🚀 项目概述 (Overview)

**Excel-Sheet-Merger** 是一个基于 **Python** 和 **Pandas** 库的小型自动化工具。

本项目提供了一个高效的脚本 `merge_excels.py`，用于解决日常工作中常见的需求：将一个文件夹内的所有 `.xlsx` 或 `.xls` 文件读取出来，并将它们各自的内容写入到一个**新的总 Excel 文件的不同 Sheet 页**中。

## ✨ 主要功能 (Features)

* **批量合并：** 自动遍历指定文件夹，识别所有 Excel 文件。
* **多表输出：** 将每个 Excel 文件的数据写入到目标文件的一个独立 Sheet 中。
* **智能命名：** **自动提取文件名**（例如：去除文件名中的扩展名）作为输出 Sheet 的名称。
* **依赖精简：** 仅依赖常用的 Pandas、OS 和 Re 库。

## ⚙️ 环境要求与安装 (Installation)

本项目运行需要 Python 3 环境，并依赖 `pandas` 和 `openpyxl` 库。

### 1. 克隆仓库

```bash
git clone https://github.com/panjose/Excel-Sheet-Merger.git
cd Excel-Sheet-Merger
````

### 2\. 安装依赖

```bash
pip install pandas openpyxl
```

## 📚 使用方法 (Usage Guide)

### 1\. 准备数据

将所有需要合并的 Excel 文件（`.xlsx` 或 `.xls`）放入一个单独的文件夹内，例如命名为 `data_to_merge`。

### 2\. 修改脚本路径

打开 `merge_excels.py` 文件，修改第 6 行的 `folder_path` 变量，将其指向您存放 Excel 文件的实际路径。

```python
# merge_excels.py (关键修改点)
# -----------------------------------------------------------------
# 1. 设置文件夹路径
# 请将 'your_folder_path' 替换为你存放 Excel 文件的实际路径
folder_path = r"C:\Users\Username\Documents\data_to_merge" 
# -----------------------------------------------------------------
```

### 3\. 运行脚本

在命令行中运行 Python 脚本：

```bash
python merge_excels.py
```

### 4\. 查看结果

脚本运行成功后，合并后的文件 `merged_file.xlsx` 将会生成在 `folder_path` **上一级目录**中。每个源 Excel 文件将对应一个 Sheet 页。

## 📝 代码说明 (Code Details)

| 行数 | 代码功能 | 说明 |
| :--- | :--- | :--- |
| 6 | `folder_path` | **用户唯一需要修改的路径变量**。 |
| 12-14 | `output_file` | 自动将输出文件放置在源文件夹的上一级目录，避免干扰源数据。 |
| 28 | `re.search` | 使用正则表达式尝试从文件名中提取一个编号或名称作为 Sheet 名。 |
| 38 | `df.to_excel` | 使用 `writer` 对象将数据写入，并设置 `index=False` 避免写入额外的索引列。 |
| 47 | `writer.close()` | 确保所有数据写入磁盘并关闭 Excel 文件。 |
