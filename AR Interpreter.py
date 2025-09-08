import pandas as pd
from tkinter import Tk, filedialog, messagebox

PIVOT_BUCKETS = ['Current','1-30','31-60','61-90','91-365','366-730','731+']

def select_file(root, title):
    return filedialog.askopenfilename(
        parent=root,
        title=title,
        filetypes=[("Excel files", "*.xlsx")]
    )

def select_save_path(root):
    return filedialog.asksaveasfilename(
        parent=root,
        title="保存Excel文件",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

def validate_columns(df, required_cols, df_name):
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        # 逗号保留，便于阅读列名
        raise ValueError(f"{df_name} 缺少必要列: {', '.join(missing)}")

def generate_pivot(raw_file: str, customer_file: str, output_file: str):
    raw_df = pd.read_excel(raw_file, sheet_name='Raw')
    default_sales_df = pd.read_excel(customer_file, sheet_name='export')

    # 仅保留需要字段，并去重
    default_sales_df = default_sales_df[['Number', 'Name', 'Salesman']].drop_duplicates(subset=['Number'])
    validate_columns(raw_df, ['Customer ID'] + PIVOT_BUCKETS, 'Raw sheet')
    validate_columns(default_sales_df, ['Number', 'Salesman'], 'Customer sheet (export)')

    # 匹配 Salesman
    merged_raw_df = raw_df.merge(
        default_sales_df[['Number','Salesman']],
        how='left',
        left_on='Customer ID',
        right_on='Number'
    ).drop(columns=['Number'])

    merged_raw_df['Salesman'] = merged_raw_df['Salesman'].fillna('Unassigned')

    # 透视/汇总
    pivot_table = (
        merged_raw_df
        .groupby('Salesman', dropna=False)[PIVOT_BUCKETS]
        .sum(numeric_only=True)
        .reset_index()
    )
    pivot_table['Total'] = pivot_table[PIVOT_BUCKETS].sum(axis=1)

    # 输出
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        pivot_table.to_excel(writer, index=False, sheet_name='Pivot')
        default_sales_df.to_excel(writer, index=False, sheet_name='Default Sales')
        merged_raw_df.to_excel(writer, index=False, sheet_name='Raw')

def main():
    root = Tk()
    root.withdraw()

    raw_file = select_file(root, "选择包含 Raw sheet 的 Excel 文件")
    if not raw_file:
        messagebox.showwarning("取消", "未选择 Raw 文件，程序已退出。", parent=root)
        return

    customer_file = select_file(root, "选择 Customer 文件")
    if not customer_file:
        messagebox.showwarning("取消", "未选择客户文件，程序已退出。", parent=root)
        return

    output_file = select_save_path(root)
    if not output_file:
        messagebox.showwarning("取消", "未选择保存路径，程序已退出。", parent=root)
        return

    try:
        generate_pivot(raw_file, customer_file, output_file)
        messagebox.showinfo("完成", f"文件已成功生成：\n{output_file}", parent=root)
    except Exception as e:
        messagebox.showerror("错误", f"发生错误：\n{e}", parent=root)

if __name__ == "__main__":
    main()