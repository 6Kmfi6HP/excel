import pandas as pd
import os
from pathlib import Path
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import re

def clean_phone_number(phone: str) -> str:
    """清理并格式化电话号码"""
    phone = str(phone)
    # 移除特殊字符
    for char in ['-', ' ', '(', ')', ':']:
        phone = phone.replace(char, '')
    # 添加+号前缀
    return "+" + phone if not phone.startswith('+') else phone

def extract_price(product_info: str) -> float:
    """从产品信息中提取价格"""
    try:
        # 查找价格模式：$数字.数字
        price_match = re.search(r'\$(\d+\.?\d*)', product_info)
        if price_match:
            return float(price_match.group(1))
        return 0.0
    except Exception:
        return 0.0

def process_excel_file(file_path: str, min_price: float = None, max_price: float = None) -> pd.DataFrame:
    """处理单个Excel文件"""
    print(f"\n处理文件: {file_path}")
    try:
        df = pd.read_excel(file_path)
        processed_data = []

        for index, row in df.iterrows():
            try:
                # 提取价格信息
                product_info = str(row.get('产品信息', ''))
                price = extract_price(product_info)
                
                # 如果设置了价格过滤条件，检查是否符合条件
                if min_price is not None and price < min_price:
                    continue
                if max_price is not None and price > max_price:
                    continue
                
                # 提取并处理数据
                order_data = {
                    '订单号': row['订单号'],
                    '收货人名称': str(row['收货人名称']),
                    '联系电话': clean_phone_number(row['联系电话']),
                    '国家': str(row['国家']),
                    '收货地址': str(row['收货地址']),
                    '价格': price,
                    '来源文件': os.path.basename(file_path)
                }
                
                processed_data.append(order_data)
                
                # 打印处理信息
                print(f"订单号: {order_data['订单号']}")
                print(f"收货人名称: {order_data['收货人名称']}")
                print(f"联系电话: {order_data['联系电话']}")
                print(f"国家: {order_data['国家']}")
                print(f"价格: ${order_data['价格']:.2f}")
                print("-" * 50)
                
            except Exception as e:
                print(f"处理第 {index} 行时出错: {str(e)}")
                continue
                
        return pd.DataFrame(processed_data)
        
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {str(e)}")
        return pd.DataFrame()  # 返回空DataFrame

def find_excel_files(directory: str) -> list:
    """递归查找目录下所有的xls文件，排除processed_orders开头的文件"""
    excel_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            # 排除processed_orders开头的文件
            if file.endswith('.xls') and not file.startswith('processed_orders'):
                excel_files.append(os.path.join(root, file))
    return excel_files

def save_to_excel(df: pd.DataFrame, output_file: str):
    """保存数据到Excel文件，设置单元格格式"""
    # 将DataFrame保存为Excel文件
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='订单数据')
    
    # 获取工作簿对象
    workbook = writer.book
    worksheet = workbook.active
    
    # 设置列宽和对齐方式
    column_widths = {
        'A': 20,  # 订单号
        'B': 15,  # 收货人名称
        'C': 20,  # 联系电话
        'D': 10,  # 国家
        'E': 40,  # 收货地址
        'F': 10,  # 价格
        'G': 20,  # 来源文件
    }
    
    # 设置每列的宽度和对齐方式
    for col_letter, width in column_widths.items():
        worksheet.column_dimensions[col_letter].width = width
        # 设置居中对齐
        for cell in worksheet[col_letter]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # 保存文件
    writer.close()

def main():
    # 获取当前目录
    current_dir = os.getcwd()
    
    # 查找所有Excel文件
    excel_files = find_excel_files(current_dir)
    
    if not excel_files:
        print("未找到需要处理的.xls文件！")
        return
        
    print(f"找到 {len(excel_files)} 个.xls文件")
    
    # 设置价格过滤范围（可以根据需要修改）
    min_price = None  # 最低价格，设置为None表示不限制
    max_price = None  # 最高价格，设置为None表示不限制
    
    # 处理所有文件并合并结果
    all_data = []
    for file_path in excel_files:
        df = process_excel_file(file_path, min_price, max_price)
        if not df.empty:
            all_data.append(df)
    
    if not all_data:
        print("没有成功处理任何文件！")
        return
        
    # 合并所有数据
    result_df = pd.concat(all_data, ignore_index=True)
    
    # 根据联系电话去重，保留最后一次出现的记录
    result_df = result_df.drop_duplicates(subset=['联系电话'], keep='last')
    
    # 按价格降序排序
    result_df = result_df.sort_values(by='价格', ascending=False)
    
    # 生成带时间戳的输出文件名
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'processed_orders_{timestamp}.xlsx'
    
    # 保存到Excel文件
    save_to_excel(result_df, output_file)
    
    print(f"\n处理完成！结果已保存到 {output_file}")
    print(f"处理的文件数: {len(excel_files)}")
    print(f"总记录数: {len(pd.concat(all_data, ignore_index=True))}")
    print(f"去重后记录数: {len(result_df)}")
    
    # 打印价格统计信息
    if not result_df.empty:
        print("\n价格统计信息:")
        print(f"最低价格: ${result_df['价格'].min():.2f}")
        print(f"最高价格: ${result_df['价格'].max():.2f}")
        print(f"平均价格: ${result_df['价格'].mean():.2f}")

if __name__ == '__main__':
    main()