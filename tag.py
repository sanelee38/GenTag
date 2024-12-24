import pandas as pd
import os
import warnings
import math


def save_to_excel(df, namespace, sheet_name='tag', max_rows=30000):
    """
    将DataFrame保存到Excel文件，如果超过最大行数限制则分割成多个文件
    """
    # 先按点项名排序
    df_sorted = df.sort_values(by='点项名', ascending=True)

    # 计算需要多少个文件
    total_rows = len(df_sorted)
    num_files = math.ceil(total_rows / max_rows)

    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))

    for i in range(num_files):
        start_idx = i * max_rows
        end_idx = min((i + 1) * max_rows, total_rows)

        # 分割数据
        df_part = df_sorted.iloc[start_idx:end_idx]

        # 构建输出文件名（使用namespace作为基础名称）
        if i == 0:  # 第一个文件不带后缀
            output_file = os.path.join(current_dir, f"{namespace}.xlsx")
        else:  # 后续文件添加数字后缀
            output_file = os.path.join(current_dir, f"{namespace}_{i}.xlsx")

        # 使用xlsxwriter引擎保存
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        df_part.to_excel(writer, sheet_name=sheet_name, index=False)

        # 获取xlsxwriter工作簿和工作表对象
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # 设置列宽
        for idx, col in enumerate(df_part.columns):
            max_length = max(df_part[col].astype(str).apply(len).max(),
                             len(str(col))) + 2
            worksheet.set_column(idx, idx, max_length)

        writer.close()

        print(f"已生成文件 {output_file}，包含 {len(df_part)} 条记录")
        # 打印该文件中包含的点项名范围
        min_item = df_part['点项名'].iloc[0]
        max_item = df_part['点项名'].iloc[-1]
        print(f"点项名范围：{min_item} - {max_item}")


def transform_excel(namespace):
    # 忽略警告
    warnings.filterwarnings('ignore')

    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 构建输入文件的完整路径
    input_file = os.path.join(current_dir, 'tag.xlsx')

    # 定义需要匹配的点项名列表（分组）
    general_items = ['AV', 'DV']  # 保留AV和DV点项名
    hsscs6_items = ['L0', 'L8', 'QUA', 'SY', 'MO', 'MC', 'V1', 'V2', 'ZT']  # 这些点项名只适用于HSSCS6点类型

    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"错误：输入文件 {input_file} 不存在！")
        return

    try:
        # 读取输入Excel文件的两个sheet
        df_tag = pd.read_excel(input_file, sheet_name='tag')
        df_tagtype = pd.read_excel(input_file, sheet_name='tagtype')

        # 过滤掉以SYS和FIO开头的点名
        df_tag = df_tag[
            (~df_tag['点名'].str.startswith('SYS', na=False)) &
            (~df_tag['点名'].str.startswith('FIO', na=False))
            ]

        # 创建一个空的结果列表来存储所有行
        result_rows = []

        # 处理一般点项名（AV, DV）
        df_tagtype_general = df_tagtype[df_tagtype['点项名'].isin(general_items)]
        general_types = df_tagtype_general['点类型'].unique()
        df_tag_general = df_tag[df_tag['点类型'].isin(general_types)]

        # 处理HSSCS6特定的点项名
        df_tag_hsscs6 = df_tag[df_tag['点类型'] == 'HSSCS6']
        df_tagtype_hsscs6 = df_tagtype[
            (df_tagtype['点类型'] == 'HSSCS6') &
            (df_tagtype['点项名'].isin(hsscs6_items))
            ]

        # 处理一般点项名的数据
        for _, tag_row in df_tag_general.iterrows():
            type_items = df_tagtype_general[
                (df_tagtype_general['点类型'] == tag_row['点类型']) &
                (df_tagtype_general['点项名'].isin(general_items))
                ]

            for _, type_item in type_items.iterrows():
                # 如果点项类型不是boolean或float，则改为boolean
                point_type = type_item['点项类型']
                if point_type.lower() not in ['boolean', 'float']:
                    point_type = 'boolean'

                # 对于AV和DV，统一设置点项描述为"当前值"
                point_item_desc = "当前值" if type_item['点项名'] in ['AV', 'DV'] else type_item['点项描述']

                new_row = {
                    '点名': tag_row['点名'],
                    '点描述': tag_row['点描述'],
                    '点项名': type_item['点项名'],
                    '点项描述': point_item_desc,
                    '点项类型': point_type,
                    '源名称空间': namespace,
                    '源点名': tag_row['点名'],
                    '源点项名': type_item['点项名'],
                    '是否周期(0否,1是)': 1
                }
                result_rows.append(new_row)

        # 处理HSSCS6特定点项名的数据
        for _, tag_row in df_tag_hsscs6.iterrows():
            for _, type_item in df_tagtype_hsscs6.iterrows():
                # 如果点项类型不是boolean或float，则改为boolean
                point_type = type_item['点项类型']
                if point_type.lower() not in ['boolean', 'float']:
                    point_type = 'boolean'

                # 对于AV和DV，统一设置点项描述为"当前值"
                point_item_desc = "当前值" if type_item['点项名'] in ['AV', 'DV'] else type_item['点项描述']

                new_row = {
                    '点名': tag_row['点名'],
                    '点描述': tag_row['点描述'],
                    '点项名': type_item['点项名'],
                    '点项描述': point_item_desc,
                    '点项类型': point_type,
                    '源名称空间': namespace,
                    '源点名': tag_row['点名'],
                    '源点项名': type_item['点项名'],
                    '是否周期(0否,1是)': 1
                }
                result_rows.append(new_row)

        if len(result_rows) == 0:
            print("警告：没有找到匹配的数据！")
            return

        # 将结果转换为DataFrame
        result_df = pd.DataFrame(result_rows)

        # 设置列的顺序
        columns = ['点名', '点描述', '点项名', '点项描述', '点项类型',
                   '源名称空间', '源点名', '源点项名', '是否周期(0否,1是)']
        result_df = result_df[columns]

        # 保存结果，如果超过30000行则分割成多个文件
        save_to_excel(result_df, namespace, max_rows=30000)

        print(f"转换完成！共处理 {len(df_tag_general) + len(df_tag_hsscs6)} 个点名，生成 {len(result_rows)} 条记录")

    except Exception as e:
        print(f"处理过程中出现错误：{str(e)}")


# 使用示例
if __name__ == "__main__":
    namespace = "unit01"  # 设置源名称空间的值
    transform_excel(namespace)