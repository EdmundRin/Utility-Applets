import warnings
warnings.filterwarnings("ignore",
                        category=UserWarning,
                        module="openpyxl.styles.stylesheet")

import pandas as pd
from openpyxl.utils import get_column_letter
from tkinter import (
    Tk, filedialog, messagebox,
    Toplevel, Frame, Label, Entry, Listbox, Scrollbar, Button,
    StringVar, END, SINGLE, BOTH, RIGHT, LEFT, Y, X
)

from decimal import Decimal, ROUND_HALF_UP, getcontext
getcontext().prec = 28  # 足够高的计算精度

PIVOT_BUCKETS = ['Current','1-30','31-60','61-90','91-365','366-730','731+']
DEC5 = Decimal("0.00001")  # 5位小数

# ---------- Decimal 工具 ----------
def to_dec5(x):
    """把任何输入转为保留5位小数的 Decimal；NaN/None -> 0"""
    if pd.isna(x):
        return Decimal("0").quantize(DEC5, rounding=ROUND_HALF_UP)
    return Decimal(str(x)).quantize(DEC5, rounding=ROUND_HALF_UP)

def df_to_dec5(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """将 df 的指定列转为 Decimal(5位) 对象列"""
    for c in cols:
        if c in df.columns:
            df[c] = df[c].apply(to_dec5)
    return df

def sum_dec(series: pd.Series) -> Decimal:
    """对一列 Decimal 求和（Python 层求和，避免浮点）"""
    total = Decimal("0")
    for v in series:
        if isinstance(v, Decimal):
            total += v
        elif pd.isna(v):
            continue
        else:
            total += to_dec5(v)
    return total.quantize(DEC5, rounding=ROUND_HALF_UP)

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
    # aging 桶转数值, 标记至少一个桶非空/数字（仅用于判定数据行，不影响后续 Decimal 计算）
    aging_num = df[PIVOT_BUCKETS].apply(pd.to_numeric, errors="coerce")
    mask_data_row = df['Customer ID'].notna() & aging_num.notna().any(axis=1)
    companies = (
        df.loc[mask_data_row, company_col]
          .dropna()
          .astype(str).str.strip()
    )
    companies = companies[~companies.str.fullmatch(r'(?i)all\s+companies')]  # 去掉 All Companies
    return sorted(companies.unique().tolist())

# ---------- 导出 ----------
def export_xlsx_multi(sheets: dict[str, pd.DataFrame], path: str):
    try:
        with pd.ExcelWriter(path, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss") as xw:
            for sheet_name, out in sheets.items():
                if out is None or out.empty:
                    pd.DataFrame().to_excel(xw, index=False, sheet_name=sheet_name)
                    ws = xw.sheets[sheet_name]
                    ws.freeze_panes = "A2"
                    continue

                out.to_excel(xw, index=False, sheet_name=sheet_name)
                ws = xw.sheets[sheet_name]

                # 冻结首行
                ws.freeze_panes = "A2"
                # 自动调整列宽
                for j, col in enumerate(out.columns, start=1):
                    values = out[col].tolist()[:1000]
                    maxlen = 0
                    for v in values:
                        if v is None:
                            s = ""
                        else:
                            s = str(v)
                        l = len(s)
                        if l > maxlen:
                            maxlen = l
                    maxlen = max(maxlen, len(str(col))) + 2
                    ws.column_dimensions[get_column_letter(j)].width = min(maxlen, 80)
    except Exception as ex:
        raise RuntimeError(f"Failed to save Excel: {ex}") from ex

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
    # （这里只是识别哪些是有效数据行，不改变原值）
    companies = extract_companies_for_choice(master_df, company_col)
    if not companies:
        raise ValueError("未识别到可用公司。请确认总表的数据区(非说明区)中存在公司名称。")

    # 一律弹窗让用户确认(即使只有 1 家也弹)
    chosen_company = choose_company_dialog(root, companies)
    if not chosen_company:
        messagebox.showwarning("取消", "未选择公司, 程序已退出。", parent=root)
        return False

    # 过滤出所选公司的数据行
    # 注意：这里不要把桶列转 numeric，不然会提前进 float
    aging_num_probe = master_df[PIVOT_BUCKETS].apply(pd.to_numeric, errors="coerce")
    mask_data_row = master_df['Customer ID'].notna() & aging_num_probe.notna().any(axis=1)
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

    # ======= 关键：把金额桶列转为 Decimal(5位) 再参与运算 =======
    merged_raw_df = df_to_dec5(merged_raw_df, PIVOT_BUCKETS)

    # 透视/汇总（对每个桶列做 Decimal 求和）
    gb = merged_raw_df.groupby('Salesman', dropna=False)
    pivot_data = {}
    for bucket in PIVOT_BUCKETS:
        pivot_data[bucket] = gb[bucket].apply(sum_dec)

    pivot_table = pd.DataFrame(pivot_data).reset_index()

    # Total = 各桶列逐行 Decimal 求和（保持 5 位）
    pivot_table['Total'] = pivot_table[PIVOT_BUCKETS].apply(
        lambda row: sum((v if isinstance(v, Decimal) else to_dec5(v)) for v in row).quantize(DEC5, rounding=ROUND_HALF_UP),
        axis=1
    )

    # ======= 导出前：将 Decimal 转成 float（Excel 里可继续运算；显示位数交给 Excel）=======
    def dec_to_float_df(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
        for c in cols:
            if c in df.columns:
                df[c] = df[c].apply(lambda v: float(v) if isinstance(v, Decimal) else v)
        return df

    pivot_table = dec_to_float_df(pivot_table, PIVOT_BUCKETS + ['Total'])
    merged_raw_df = dec_to_float_df(merged_raw_df, PIVOT_BUCKETS)

    # 输出
    try:
        export_xlsx_multi(
            {
                'Pivot': pivot_table,
                'Default Sales': default_sales_df,
                'Raw': merged_raw_df
            },
            output_file
        )
    except Exception as e:
        messagebox.showerror("错误", f"保存Excel失败：\n{e}", parent=root)
        return False

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
