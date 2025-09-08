import warnings
warnings.filterwarnings("ignore",
                        category=UserWarning,
                        module="openpyxl.styles.stylesheet")

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import ImageFont
from tkinter import (
    Tk, filedialog, messagebox,
    Toplevel, Frame, Label, Entry, Listbox, Scrollbar, Button,
    StringVar, END, SINGLE, BOTH, RIGHT, LEFT, Y, X
)

PIVOT_BUCKETS = ['Current','1-30','31-60','61-90','91-365','366-730','731+']

# ---------- 带搜索框的选择器 ----------
def choose_company_dialog(root, companies, title="选择公司"):
    win = Toplevel(root)
    win.title(title)
    # 不立即 grab_set, 避免初始化渲染阻塞
    win.geometry("560x520")
    win.minsize(460, 380)

    # 归一 & 去重 & 排序
    uniq_companies = sorted({str(c).strip() for c in companies if str(c).strip()})
    companies_lower = [c.casefold() for c in uniq_companies]

    # 懒加载阈值
    LAZY_THRESHOLD = 1000
    lazy_mode = len(uniq_companies) > LAZY_THRESHOLD

    # 顶部: 搜索框
    top = Frame(win); top.pack(fill=X, padx=10, pady=(12, 6))
    Label(top, text="搜索公司: ").pack(side=LEFT, padx=(0, 6))
    qvar = StringVar()
    ent = Entry(top, textvariable=qvar); ent.pack(side=LEFT, fill=X, expand=True)

    # 提示/计数
    hint = StringVar()
    if lazy_mode:
        hint.set(f"共 {len(uniq_companies)} 家公司。为提速, 请先输入 ≥2 个字符再显示结果。")
    else:
        hint.set(f"共 {len(uniq_companies)} 家公司。可直接滚动或输入过滤。")
    Label(win, textvariable=hint, anchor="w").pack(fill=X, padx=12)

    # 中部: 列表 + 滚动条
    mid = Frame(win); mid.pack(fill=BOTH, expand=True, padx=10, pady=6)
    sb = Scrollbar(mid); sb.pack(side=RIGHT, fill=Y)
    lb = Listbox(mid, selectmode=SINGLE)
    lb.pack(side=LEFT, fill=BOTH, expand=True)
    lb.config(yscrollcommand=sb.set); sb.config(command=lb.yview)

    # 状态
    filtered_idx = list(range(len(uniq_companies)))  # 当前展示的索引列表
    last_query = {"text": None}
    pending_after = {"id": None}

    def build_items_from_idx(idxs):
        # 从索引构建原始字符串列表
        return [uniq_companies[i] for i in idxs]

    def listbox_fill(items):
        lb.delete(0, END)
        if items:
            # 批量插入
            lb.insert(END, *items)
            # 选中第一项, 便于键盘回车快速确认
            lb.selection_set(0); lb.see(0)

    def do_filter():
        # 读取并归一查询
        q = qvar.get().strip()
        q_norm = q.casefold()
        if lazy_mode and len(q_norm) < 2:
            # 懒加载: 要求至少 2 字符
            listbox_fill([])
            hint.set(f"共 {len(uniq_companies)} 家公司。请输入 ≥2 个字符进行搜索。")
            return

        # 若与上次相同, 不必重算
        if q_norm == last_query["text"]:
            return
        last_query["text"] = q_norm

        if not q_norm:
            # 空查询: 展示全量(若非懒加载)
            if lazy_mode:
                listbox_fill([])
                hint.set(f"共 {len(uniq_companies)} 家公司。请输入 ≥2 个字符进行搜索。")
                return
            filtered = list(range(len(uniq_companies)))
        else:
            # 先前缀匹配, 再子串匹配(去重)
            starts = [i for i, s in enumerate(companies_lower) if s.startswith(q_norm)]
            if len(starts) < 5000:  # 经验: 命中不太多时再做包含匹配
                contains = [i for i, s in enumerate(companies_lower)
                            if q_norm in s and not s.startswith(q_norm)]
                filtered = starts + contains
            else:
                filtered = starts  # 减少额外包含匹配的成本

        items = build_items_from_idx(filtered)
        listbox_fill(items)
        hint.set(f"匹配 {len(items)} / {len(uniq_companies)}")

    def schedule_filter():
        # 防抖: 取消上一次计划
        if pending_after["id"] is not None:
            try:
                win.after_cancel(pending_after["id"])
            except Exception:
                pass
            pending_after["id"] = None
        # 120ms 后执行过滤
        pending_after["id"] = win.after(120, do_filter)

    def confirm_selection():
        sel = lb.curselection()
        if not sel:
            # 如果没有选中但输入是全量匹配, 也允许直接确认
            q = qvar.get().strip()
            if q and q in uniq_companies:
                choice = q
            else:
                messagebox.showwarning("未选择", "请从列表中选中一个公司。", parent=win)
                return
        else:
            choice = lb.get(sel[0])
        win.grab_release()
        win.destroy()
        selected["value"] = choice

    def cancel():
        win.grab_release()
        win.destroy()
        selected["value"] = None

    # 底部按钮
    bot = Frame(win); bot.pack(fill=X, padx=10, pady=(6, 10))
    Button(bot, text="确定", command=confirm_selection).pack(side=LEFT, padx=(0, 6))
    Button(bot, text="取消", command=cancel).pack(side=LEFT)

    # 事件绑定
    ent.bind("<KeyRelease>", lambda e: schedule_filter())
    ent.bind("<Return>",     lambda e: confirm_selection())
    lb.bind("<Return>",      lambda e: confirm_selection())
    lb.bind("<Double-1>",    lambda e: confirm_selection())
    win.bind("<Escape>",     lambda e: cancel())

    # 初始渲染(非懒加载时才显示全量)
    if not lazy_mode:
        listbox_fill(uniq_companies)
    else:
        listbox_fill([])

    # 渲染完再设为模态, 避免“白窗等渲染”的卡顿体感
    win.update_idletasks()
    win.grab_set()
    ent.focus_set()

    selected = {"value": None}
    win.wait_window()
    return selected["value"]

# ---------- 文件对话框 ----------
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
        raise ValueError(f"{df_name} 缺少必要列: {', '.join(missing)}")

# ---------- 自动探测表头并读取总表 ----------
def _detect_header_row(df_no_header: pd.DataFrame) -> int:
    for i, row in df_no_header.iterrows():
        s = row.astype(str)
        if s.str.contains("Customer", case=False, na=False).any() and \
           s.str.contains("Current", case=False, na=False).any():
            return i
    return 0

def read_master_raw(master_file: str, sheet_name: str | int | None = None) -> pd.DataFrame:
    probe = pd.read_excel(master_file, sheet_name=sheet_name or 0, header=None)
    header_idx = _detect_header_row(probe)
    df = pd.read_excel(master_file, sheet_name=sheet_name or 0, header=header_idx)
    return df

# ---------- 从“数据行”提取公司列表(排除说明区 & All Companies) ----------
def extract_companies_for_choice(df: pd.DataFrame, company_col: str) -> list[str]:
    # aging 桶转数值, 标记至少一个桶非空/数字
    aging_num = df[PIVOT_BUCKETS].apply(pd.to_numeric, errors="coerce")
    mask_data_row = df['Customer ID'].notna() & aging_num.notna().any(axis=1)
    companies = (
        df.loc[mask_data_row, company_col]
          .dropna()
          .astype(str).str.strip()
    )
    companies = companies[~companies.str.fullmatch(r'(?i)all\s+companies')]  # 去掉 All Companies
    return sorted(companies.unique().tolist())

# ---------- 自动列宽 ----------
def autosize_columns_xlsx(path, padding=2):
    """智能调整列宽"""
    font = ImageFont.load_default()
    wb = load_workbook(path, data_only=True)

    def _display_text(cell):
        v = cell.value
        if v is None:
            return ""
        if cell.is_date:
            return v.strftime("%Y-%m-%d")
        if isinstance(v, (float, int)):
            # round 到小数点后 5 位
            return str(round(v, 5))
        return str(v)

    for ws in wb.worksheets:
        for col_idx, col_cells in enumerate(ws.columns, 1):
            max_width = 0
            for cell in col_cells:
                val = _display_text(cell)
                try:
                    width_px = font.getlength(val)
                except Exception:
                    width_px = len(val) * 11 * 0.6
                if width_px > max_width:
                    max_width = width_px
            # 转换为 Excel 列宽单位
            adjusted = (max_width / 7) + padding
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = adjusted

    wb.save(path)

# ---------- 主逻辑 ----------
def generate_pivot(master_file: str, customer_file: str, output_file: str, root: Tk):
    master_df = read_master_raw(master_file)

    # 找到 Company 列
    company_col = None
    for c in master_df.columns:
        if str(c).strip().lower() == 'company':
            company_col = c; break
    if company_col is None:
        candidates = [c for c in master_df.columns if 'company' in str(c).lower()]
        if len(candidates) == 1:
            company_col = candidates[0]
    if company_col is None:
        raise ValueError("未找到 Company 列, 请确认总表包含公司列。")

    # 校验必要列(在“真正数据列”层面)
    validate_columns(master_df, [company_col, 'Customer ID'] + PIVOT_BUCKETS, 'Master sheet')

    # 读取客户-业务员映射
    default_sales_df = pd.read_excel(customer_file, sheet_name='export')
    validate_columns(default_sales_df, ['Number', 'Salesman'], 'Customer sheet (export)')
    default_sales_df = default_sales_df[['Number', 'Name', 'Salesman']].drop_duplicates(subset=['Number'])

    # 仅从“数据行”提取公司列表
    companies = extract_companies_for_choice(master_df, company_col)
    if not companies:
        raise ValueError("未识别到可用公司。请确认总表的数据区(非说明区)中存在公司名称。")

    # 一律弹窗让用户确认(即使只有 1 家也弹)
    chosen_company = choose_company_dialog(root, companies)
    if not chosen_company:
        messagebox.showwarning("取消", "未选择公司, 程序已退出。", parent=root)
        return False

    # 过滤出所选公司的数据行
    aging_num = master_df[PIVOT_BUCKETS].apply(pd.to_numeric, errors="coerce")
    mask_data_row = master_df['Customer ID'].notna() & aging_num.notna().any(axis=1)
    raw_df = master_df.loc[
        mask_data_row & (master_df[company_col].astype(str).str.strip() == chosen_company),
    ].copy()

    if raw_df.empty:
        raise ValueError(f"公司 {chosen_company} 在数据区没有记录。")

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
    autosize_columns_xlsx(output_file)

    return True

def main():
    root = Tk(); root.withdraw()

    master_file = select_file(root, "选择总表(包含所有公司的明细)")
    if not master_file:
        messagebox.showwarning("取消", "未选择总表, 程序已退出。", parent=root); return

    customer_file = select_file(root, "选择 Customer 文件(export: Number/Name/Salesman)")
    if not customer_file:
        messagebox.showwarning("取消", "未选择客户文件, 程序已退出。", parent=root); return

    output_file = select_save_path(root)
    if not output_file:
        messagebox.showwarning("取消", "未选择保存路径, 程序已退出。", parent=root); return

    try:
        ok = generate_pivot(master_file, customer_file, output_file, root)
        if ok:
            messagebox.showinfo("完成", f"文件已成功生成：\n{output_file}", parent=root)
        else:
            pass
    except Exception as e:
        messagebox.showerror("错误", f"发生错误: \n{e}", parent=root)

if __name__ == "__main__":
    main()