# Excel 订单处理工具

这是一个用于处理 Excel 订单文件的工具，可以自动处理和合并多个订单文件，并进行电话号码格式化和去重。

## 功能特点

- 自动遍历目录下所有 .xls 文件
- 格式化电话号码（添加 + 号，移除特殊字符）
- 根据电话号码去重
- 保存为格式化的 Excel 文件
- 支持 Windows、macOS 和 Linux

## 使用方法

1. 下载对应您操作系统的可执行文件：
   - Windows: `excel_processor_windows.exe`
   - macOS: `excel_processor_macos`
   - Linux: `excel_processor_linux`

2. 将可执行文件放在包含 Excel 文件的目录中

3. 运行程序：
   - Windows: 双击 `excel_processor_windows.exe`
   - macOS/Linux: 
     ```bash
     chmod +x excel_processor_macos  # 或 excel_processor_linux
     ./excel_processor_macos  # 或 ./excel_processor_linux
     ```

4. 程序会自动：
   - 处理当前目录及子目录中的所有 .xls 文件
   - 格式化电话号码
   - 去除重复记录
   - 生成带时间戳的结果文件（格式：processed_orders_YYYYMMDD_HHMMSS.xlsx）

## 注意事项

- 程序会自动跳过文件名以 "processed_orders" 开头的文件
- 去重时保留最后一次出现的记录
- 输出文件包含所有原始字段，并添加了"来源文件"列
- 所有数据在输出文件中居中对齐，并设置了适当的列宽

## 开发环境

- Python 3.9
- pandas
- openpyxl
- xlrd
