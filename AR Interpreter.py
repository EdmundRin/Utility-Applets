import pandas as pd
from tkinter import Tk, filedialog, messagebox

def select_file(title):
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx")]
    )

def select_save_path():
    root = Tk()
    root.withdraw()
    return filedialog.asksaveasfilename(
        title="保存Excel文件",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

def validate_columns(df, required_cols, df_name):
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"{df_name} 缺少必要列: {', '.join(missing)}")

def generate_pivot(raw_file: str, customer_file: str, output_file: str):
    try:
        # 读取原始数据
        raw_df = pd.read_excel(raw_file, sheet_name='Raw')
        default_sales_df = pd.read_excel(customer_file, sheet_name='export')
        default_sales_df = default_sales_df[['Number', 'Name', 'Salesman']]

        # 校验必需字段
        validate_columns(raw_df, ['Customer ID'] + ['1-30', '31-60', '61-90', '91-365', '366-730', '731+'], 'Raw sheet')
        validate_columns(default_sales_df, ['Number', 'Salesman'], 'Customer sheet (export)')

        # 匹配 Salesman
        sales_lookup = default_sales_df[['Number', 'Salesman']].drop_duplicates()
        merged_raw_df = raw_df.merge(
            sales_lookup,
            how='left',
            left_on='Customer ID',
            right_on='Number'
        )
        merged_raw_df.drop(columns=['Number'], inplace=True)

        if 'Salesman_y' in merged_raw_df.columns:
            merged_raw_df.drop(columns=['Salesman'], errors='ignore', inplace=True)
            merged_raw_df.rename(columns={'Salesman_y': 'Salesman'}, inplace=True)

        # 创建 Pivot Table
        pivot_columns = ['1-30', '31-60', '61-90', '91-365', '366-730', '731+']
        pivot_table = pd.pivot_table(
            merged_raw_df,
            index='Salesman',
            values=pivot_columns,
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        pivot_table['Total'] = pivot_table[pivot_columns].sum(axis=1)

        # 写入 Excel 文件
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pivot_table.to_excel(writer, index=False, sheet_name='Pivot')
            default_sales_df.to_excel(writer, index=False, sheet_name='Default Sales')
            merged_raw_df.to_excel(writer, index=False, sheet_name='Raw')

        messagebox.showinfo("完成", f"文件已成功生成：\n{output_file}")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误：\n{str(e).replace(',', ',\n')}")

def main():
    raw_file = select_file("选择包含 Raw sheet 的 Excel 文件")
    if not raw_file:
        messagebox.showwarning("取消", "未选择 Raw 文件，程序已退出。")
        return

    customer_file = select_file("选择 Customer 文件")
    if not customer_file:
        messagebox.showwarning("取消", "未选择客户文件，程序已退出。")
        return

    output_file = select_save_path()
    if not output_file:
        messagebox.showwarning("取消", "未选择保存路径，程序已退出。")
        return

    generate_pivot(raw_file, customer_file, output_file)

if __name__ == "__main__":
    main()