import pandas as pd
import re

def process_excel_data(file_path):
    """
    直接处理Excel文件
    """
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 从规格名称中提取颜色和尺寸
    def extract_color_size(name):
        match = re.search(r'[A-Za-z]', str(name))
        if match:
            idx = match.start()
            color = name[:idx]
            size = name[idx:]
            return color, size
        return name, ''

    # 分离颜色和尺寸
    df[['颜色', '尺寸']] = df['规格名称'].apply(
        lambda x: pd.Series(extract_color_size(x))
    )

    # 分组汇总
    grouped = df.groupby(['规格编码', '颜色', '尺寸'])['规格数量'].sum().reset_index()
    grouped['尺寸数量'] = grouped['尺寸'] + '*' + grouped['规格数量'].astype(str)

    # 生成最终结果
    result_df = grouped.groupby(['规格编码', '颜色'])['尺寸数量'].apply(
        lambda x: ','.join(x)
    ).reset_index()

    result_df['结果'] = result_df['规格编码'] +'-'+ result_df['颜色'] + ' ：' + result_df['尺寸数量']

    return result_df

# 使用示例（取消注释并修改文件路径）
# file_path = "您的Excel文件路径.xlsx"
file_path = '/Users/lin/Documents/备货单2025-11-16.xlsx'
result = process_excel_data(file_path)
print(result)
result.to_excel('汇总结果.xlsx', index=False)