import pandas as pd
import os
import re

# 1. 设置文件夹路径
# 请将 'your_folder_path' 替换为你存放 Excel 文件的实际路径
folder_path = r"your_folder_path"

# 2. 创建一个 ExcelWriter 对象
# 这将是我们的目标文件，所有数据都将写入其中
# 请将 'merged_file.xlsx' 替换为你想要的合并后的文件名
output_path = os.path.dirname(folder_path)
output_file = os.path.join(output_path,'merged_file.xlsx')
writer = pd.ExcelWriter(output_file, engine='openpyxl')

# 3. 遍历文件夹中的所有 Excel 文件
for filename in os.listdir(folder_path):
    # 确保只处理 .xlsx 或 .xls 文件
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        # 完整的 Excel 文件路径
        file_path = os.path.join(folder_path, filename)
        
        try:
            # 4. 使用正则表达式提取编号
            # 正则表达式解释: 
            # ([\w\d\.]+) 匹配任何字母、数字或点，并捕获这一部分
            match = re.search(r'\s*([\w\d\.]+)\s*', filename)
            
            # 检查是否找到了匹配项
            if match:
                # 提取捕获的组作为新的表名
                sheet_name = match.group(1)
            else:
                # 如果没有匹配到，则使用原始文件名（不含扩展名）作为备用表名
                sheet_name = os.path.splitext(filename)[0]
            
            # 5. 读取每个 Excel 文件
            df = pd.read_excel(file_path)
            
            # 6. 将数据写入到目标文件的一个新表里
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"成功将 '{filename}' 合并到表 '{sheet_name}'。")

        except Exception as e:
            print(f"处理文件 '{filename}' 时出错：{e}")

# 7. 保存并关闭 Excel 文件
writer.close()

print("\n所有文件已成功合并到 'merged_file.xlsx' 中。")