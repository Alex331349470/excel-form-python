import pandas as pd
import os

# 设置文件路径
excel_dir = os.path.join(os.path.dirname(__file__), 'excel_files')
source_file = os.path.join(excel_dir, '云本建行.xls')
target_file = os.path.join(excel_dir, '资金日报汇总表3月7日.xlsx')

def transfer_data():
    try:
        # 读取源文件
        print(f"正在读取源文件: {source_file}")
        source_df = pd.read_excel(source_file)
        print("源文件数据预览:")
        print(source_df.head())
        
        # 读取目标文件
        print(f"正在读取目标文件: {target_file}")
        # 使用 openpyxl 引擎处理 .xlsx 文件
        target_excel = pd.ExcelFile(target_file, engine='openpyxl')
        
        # 获取目标文件的所有工作表
        sheet_names = target_excel.sheet_names
        print(f"目标文件中的工作表: {sheet_names}")
        
        # 默认选择第一个工作表，如果需要其他工作表可以调整
        sheet_to_update = sheet_names[0]
        target_df = pd.read_excel(target_file, sheet_name=sheet_to_update)
        print("目标文件数据预览:")
        print(target_df.head())
        
        # 根据实际情况，这里需要确定如何将数据写入目标文件
        # 方式1: 将源数据追加到目标工作表末尾
        # 方式2: 更新目标表中的特定列
        # 方式3: 创建新的工作表
        
        # 这里以方式1为例
        with pd.ExcelWriter(target_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # 将数据写入新的工作表，以便保留原有数据
            source_df.to_excel(writer, sheet_name='云本建行数据', index=False)
            print(f"数据已成功写入到 '{target_file}' 的 '云本建行数据' 工作表")
    
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    print("开始处理数据...")
    transfer_data()
    print("处理完成!")