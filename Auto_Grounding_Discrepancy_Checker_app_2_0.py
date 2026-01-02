import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
import os
import subprocess
import platform
import sys
import datetime
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_exe_dir() -> str:
    return os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__)

# Open the folder containing the given file.
def open_file_location(path: str):
    folder = os.path.dirname(path)
    if platform.system() == "Windows":
        os.startfile(folder)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


language = "en"
TEXT = {
    "en": {
        "title": "Grounding Discrepancy Checker 2.0",
        "grounding": "Please choose Grounding.xlsx (Optional)",
        "inbound": "Please choose InboundRecordFile.xlsx (Optional)",
        "export": "Please choose ExportCustomRecordsFile.xlsx (Optional)",
        "cancel": "Please choose Cancel Order.xlsx (Optional)",
        "transfer": "Please choose Transfer Order.xlsx (Optional)",
        "history": "Please choose historical.xlsx (Optional cumulative file)",
        "browse": "📂 Browse",
        "not_selected": "Not Selected",
        "run_button": "🚀 Check Discrepancy",
        "refresh_button": "🔄 Refresh Auto Match",
        "error_missing_all": "Please upload at least one file (Grounding / Inbound / Export / Cancel / Transfer) before running.",
        "success": "✅ Analysis complete. Report generated.",
        "error_title": "Error",
        "success_title": "Done",
        "menu_language": "Language",
        "menu_zh": "Chinese",
        "menu_en": "English",
        "tab_home": "Home",
        "tab_checker": "Statistics",
        "home_title": "Daily Discrepancy Statistics",
        "home_date": "Date (YYYY-MM-DD)",
        "home_source": "Inbound Source",
        "home_count": "Mismatch Count",
        "home_total_upc": "Total UPC Count",
        "home_add": "➕ Add Record",
        "home_export": "💾 Export to Excel",
        "home_chart": "📊 Generate Chart",
        "home_start_date": "Start Date (YYYY-MM-DD)",
        "home_end_date": "End Date (YYYY-MM-DD)",
        "home_chart_type": "Chart Type",
        "home_chart_bar": "Bar",
        "home_chart_line": "Line",
        "stats_file_label": "Stats history file:",
        "stats_auto_hint": "Tip: Total UPC is auto-filled from today's files when available.",
        "stats_no_file": "none",
        "stats_load_error": "Error loading stats history file: ",
    },
    "zh": {
        "title": "地面入库差异检查工具 2.0",
        "grounding": "请选择 Grounding.xlsx（可选）",
        "inbound": "请选择 InboundRecordFile.xlsx（可选）",
        "export": "请选择 ExportCustomRecordsFile.xlsx（可选）",
        "cancel": "请选择 Cancel Order.xlsx（可选）",
        "transfer": "请选择 Transfer Order.xlsx（可选）",
        "history": "请选择 historical.xlsx（历史累计表，可选）",
        "browse": "📂 浏览",
        "not_selected": "未选择",
        "run_button": "🚀 执行差异核查",
        "refresh_button": "🔄 重新自动匹配",
        "error_missing_all": "请至少选择 Grounding / Inbound / Export / Cancel / Transfer 之一再执行核查。",
        "success": "✅ 分析完成，已生成报告",
        "error_title": "错误",
        "success_title": "完成",
        "menu_language": "语言",
        "menu_zh": "中文",
        "menu_en": "English",
        "tab_home": "主页",
        "tab_checker": "统计",
        "home_title": "每日差异数量统计",
        "home_date": "日期（YYYY-MM-DD）",
        "home_source": "入库来源",
        "home_count": "不匹配数量",
        "home_total_upc": "当日入库 UPC 数量",
        "home_add": "➕ 添加记录",
        "home_export": "💾 导出 Excel",
        "home_chart": "📊 生成图表",
        "home_start_date": "开始日期（YYYY-MM-DD）",
        "home_end_date": "结束日期（YYYY-MM-DD）",
        "home_chart_type": "图表类型",
        "home_chart_bar": "柱状图",
        "home_chart_line": "折线图",
        "stats_file_label": "统计历史文件：",
        "stats_auto_hint": "小提示：当日入库 UPC 数量会在有数据时自动填入。",
        "stats_no_file": "无",
        "stats_load_error": "读取统计历史文件出错：",
    },
}

file_paths: dict[str, str] = {}
status_labels: dict[str, tk.Label] = {}
label_refs: dict[str, tk.Label] = {}

STATS_HISTORY_FILENAME = "discrepancy_stats.xlsx"

stats_df = pd.DataFrame(columns = ["Date", "Source", "MismatchCount", "TotalUPC", "Rate"])
chart_canvas: FigureCanvasTkAgg | None = None

today_source_upc_counts = {
    "Inbound": 0,
    "Value-Add": 0,
    "Cancel": 0,
    "Transfer": 0,
}

# Source filter options for chart
source_filter_vars = {}


def select_file(file_type: str):
    path = filedialog.askopenfilename(filetypes = [("Excel files", "*.xlsx")])
    if path:
        file_paths[file_type] = path
        if file_type in status_labels:
            status_labels[file_type].config(
                text = os.path.basename(path),
                fg = "green",
                font = ("Arial", 9, "bold")
            )


def auto_load_files():
    file_paths.clear()
    for key, lbl in status_labels.items():
        lbl.config(
            text = TEXT[language]["not_selected"],
            fg = "gray",
            font = ("Arial", 9)
        )

    default_files = {
        "Grounding": "Grounding.xlsx",
        "Inbound": "InboundRecordFile.xlsx",
        "Export": "ExportCustomRecordsFile.xlsx",
        "Cancel": "Cancel Order.xlsx",
        "Transfer": "Transfer Order.xlsx",
        "History": "historical.xlsx",
    }

    exe_dir = get_exe_dir()
    for key, filename in default_files.items():
        path = os.path.join(exe_dir, filename)
        if os.path.exists(path):
            file_paths[key] = path
            if key in status_labels:
                status_labels[key].config(
                    text = filename,
                    fg = "green",
                    font = ("Arial", 9, "bold")
                )


def save_stats_history():
    if stats_df is None:
        return
    exe_dir = get_exe_dir()
    path = os.path.join(exe_dir, STATS_HISTORY_FILENAME)
    try:
        df_to_save = stats_df.copy()
        if not df_to_save.empty:
            df_to_save["Date"] = df_to_save["Date"].astype(str)
        df_to_save.to_excel(path, index=False)
        stats_file_label.config(
            text = f"{TEXT[language]['stats_file_label']} {STATS_HISTORY_FILENAME}"
        )
    except Exception as e:
        messagebox.showerror(TEXT[language]["error_title"], str(e))


def load_stats_history():
    global stats_df
    exe_dir = get_exe_dir()
    path = os.path.join(exe_dir, STATS_HISTORY_FILENAME)

    if not os.path.exists(path):
        stats_df = pd.DataFrame(columns = ["Date", "Source", "MismatchCount", "TotalUPC", "Rate"])
        stats_file_label.config(
            text = f"{TEXT[language]['stats_file_label']} {TEXT[language]['stats_no_file']}"
        )
        return

    try:
        df = pd.read_excel(path)
        expected_cols = ["Date", "Source", "MismatchCount", "TotalUPC", "Rate"]
        for c in expected_cols:
            if c not in df.columns:
                if c == "Date":
                    df[c] = datetime.date.today()
                else:
                    df[c] = 0
        df = df[expected_cols]

        try:
            df["Date"] = pd.to_datetime(df["Date"]).dt.date
        except Exception:
            df["Date"] = datetime.date.today()

        stats_df = df

        for _, row in stats_df.iterrows():
            dt = row["Date"]
            if isinstance(dt, str):
                try:
                    dt = datetime.datetime.strptime(dt, "%Y-%m-%d").date()
                except Exception:
                    dt = datetime.date.today()
            rate_val = row["Rate"] if pd.notna(row["Rate"]) else 0.0
            rate_str = f"{rate_val * 100:.1f}%" if rate_val != 0 else "0.0%"
            stats_tree.insert(
                "",
                "end",
                values = (
                    dt.strftime("%Y-%m-%d"),
                    row["Source"],
                    int(row["MismatchCount"]),
                    int(row["TotalUPC"]),
                    rate_str,
                ),
            )

        stats_file_label.config(
            text=f"{TEXT[language]['stats_file_label']} {STATS_HISTORY_FILENAME}"
        )

    except Exception as e:
        messagebox.showerror(
            TEXT[language]["error_title"],
            TEXT[language]["stats_load_error"] + str(e),
        )
        stats_df = pd.DataFrame(columns=["Date", "Source", "MismatchCount", "TotalUPC", "Rate"])
        stats_file_label.config(
            text=f"{TEXT[language]['stats_file_label']} {TEXT[language]['stats_no_file']}"
        )


def run_check():
    global today_source_upc_counts

    if not any(k in file_paths for k in ["Grounding", "Inbound", "Export", "Cancel", "Transfer"]):
        messagebox.showerror(TEXT[language]["error_title"], TEXT[language]["error_missing_all"])
        return

    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
    if not save_path:
        return

    try:
        grounding_exists = "Grounding" in file_paths
        inbound_exists = "Inbound" in file_paths
        export_exists = "Export" in file_paths
        cancel_exists = "Cancel" in file_paths
        transfer_exists = "Transfer" in file_paths
        history_path = file_paths.get("History", None)

        if grounding_exists:
            grounding_df = pd.read_excel(file_paths["Grounding"])
            records = []
            for i in range(0, len(grounding_df), 3):
                group = grounding_df.iloc[i:i + 3]
                if len(group) == 3:
                    upc = str(group.iloc[0, 0])
                    product_name = str(group.iloc[1, 0])
                    grounding_num = group.iloc[2, 1]
                    try:
                        grounding_num = int(float(grounding_num))
                    except Exception:
                        grounding_num = 0
                    records.append({
                        "UPC": upc,
                        "Product Name": product_name,
                        "Grounding Qty": grounding_num
                    })
            df_final_grounding = pd.DataFrame(records)
            grounding_summary = df_final_grounding.groupby("UPC", as_index=False)["Grounding Qty"].sum()
        else:
            df_final_grounding = pd.DataFrame(columns=["UPC", "Product Name", "Grounding Qty"])
            grounding_summary = pd.DataFrame(columns=["UPC", "Grounding Qty"])

        if inbound_exists:
            inbound_df = pd.read_excel(file_paths["Inbound"])
            inbound_summary = inbound_df.groupby("UPC", as_index=False)["Inbound Qty"].sum()
        else:
            inbound_summary = pd.DataFrame(columns=["UPC", "Inbound Qty"])

        if export_exists:
            export_df_raw = pd.read_excel(file_paths["Export"])
            valueadd_df = export_df_raw[export_df_raw["Outbound SN"].isna()]
            valueadd_summary = valueadd_df.groupby("UPC after value-added", as_index=False)["Value-added Num."].sum()
            valueadd_summary.rename(
                columns={"UPC after value-added": "UPC", "Value-added Num.": "Value-add Qty"},
                inplace=True
            )
        else:
            valueadd_summary = pd.DataFrame(columns=["UPC", "Value-add Qty"])

        if cancel_exists:
            cancel_df_raw = pd.read_excel(file_paths["Cancel"], header=None)
            cancel_df_raw = cancel_df_raw.iloc[1:].reset_index(drop=True)
            col0 = cancel_df_raw[0].tolist()
            cancel_records = []
            i = 0
            while i < len(col0):
                upc = str(col0[i]).strip()
                if i + 1 >= len(col0):
                    break
                product_name = str(col0[i + 1]).strip()
                if i + 2 >= len(col0):
                    break
                next_line = str(col0[i + 2]).strip()
                if next_line.isnumeric():
                    shelf_code = None
                    canceled_qty = int(float(next_line))
                    i += 3
                else:
                    shelf_code = next_line
                    if i + 3 < len(col0):
                        canceled_qty_str = str(col0[i + 3]).strip()
                        try:
                            canceled_qty = int(float(canceled_qty_str))
                        except Exception:
                            canceled_qty = 0
                    else:
                        canceled_qty = 0
                    i += 4
                cancel_records.append({
                    "UPC": upc,
                    "Product Name": product_name,
                    "Shelf Code": shelf_code,
                    "Canceled Qty": canceled_qty
                })
            cancel_df = pd.DataFrame(cancel_records)
            cancel_df = cancel_df[(cancel_df["Shelf Code"].isna()) | (cancel_df["Shelf Code"] == "CUSTOM0")]
            cancel_summary = cancel_df.groupby("UPC", as_index=False)["Canceled Qty"].sum()
        else:
            cancel_summary = pd.DataFrame(columns=["UPC", "Canceled Qty"])

        if transfer_exists:
            transfer_df = pd.read_excel(file_paths["Transfer"])
            transfer_summary = transfer_df.groupby("UPC", as_index=False)["transfer Num."].sum()
            transfer_summary.rename(columns={"transfer Num.": "Transfer Qty"}, inplace=True)
        else:
            transfer_summary = pd.DataFrame(columns=["UPC", "Transfer Qty"])

        for df in [grounding_summary, inbound_summary, valueadd_summary, cancel_summary, transfer_summary]:
            if not df.empty:
                df["UPC"] = df["UPC"].astype(str)

        today_source_upc_counts = {
            "Inbound": int((inbound_summary["Inbound Qty"] > 0).sum()) if not inbound_summary.empty else 0,
            "Value-Add": int((valueadd_summary["Value-add Qty"] > 0).sum()) if not valueadd_summary.empty else 0,
            "Cancel": int((cancel_summary["Canceled Qty"] > 0).sum()) if not cancel_summary.empty else 0,
            "Transfer": int((transfer_summary["Transfer Qty"] > 0).sum()) if not transfer_summary.empty else 0,
        }

        all_upcs = set()
        for df in [grounding_summary, inbound_summary, valueadd_summary, cancel_summary, transfer_summary]:
            if not df.empty:
                all_upcs.update(df["UPC"].unique())

        if not all_upcs:
            messagebox.showerror(TEXT[language]["error_title"], "No data rows found in selected files.")
            return

        merged = pd.DataFrame({"UPC": list(all_upcs)})
        merged = merged.merge(grounding_summary, on="UPC", how="left")
        merged = merged.merge(inbound_summary, on="UPC", how="left")
        merged = merged.merge(valueadd_summary, on="UPC", how="left")
        merged = merged.merge(cancel_summary, on="UPC", how="left")
        merged = merged.merge(transfer_summary, on="UPC", how="left")

        for col in ["Grounding Qty", "Inbound Qty", "Value-add Qty", "Canceled Qty", "Transfer Qty"]:
            if col in merged.columns:
                merged[col] = merged[col].fillna(0).astype(int)
            else:
                merged[col] = 0

        merged["Total Inbound Qty"] = (
            merged["Inbound Qty"]
            + merged["Value-add Qty"]
            + merged["Canceled Qty"]
            + merged["Transfer Qty"]
        )

        if not df_final_grounding.empty:
            product_map = df_final_grounding[["UPC", "Product Name"]].drop_duplicates("UPC")
        else:
            product_map = pd.DataFrame(columns=["UPC", "Product Name"])
        merged = merged.merge(product_map, on="UPC", how="left")

        def check_issue(row):
            if row["Total Inbound Qty"] == 0 and row["Grounding Qty"] > 0:
                return "Missing Inbound"
            if row["Grounding Qty"] == 0 and row["Total Inbound Qty"] > 0:
                return "Missing Grounding"
            if row["Total Inbound Qty"] != row["Grounding Qty"]:
                return "Quantity Mismatch"
            return None

        merged["Issue"] = merged.apply(check_issue, axis=1)
        today_result = merged[merged["Issue"].notna()][[
            "UPC",
            "Product Name",
            "Inbound Qty",
            "Value-add Qty",
            "Canceled Qty",
            "Transfer Qty",
            "Total Inbound Qty",
            "Grounding Qty",
            "Issue"
        ]]

        export_df = today_result.copy()

        if history_path is not None:
            try:
                if os.path.exists(history_path):
                    hist_df = pd.read_excel(history_path)
                else:
                    hist_df = pd.DataFrame(columns=today_result.columns)

                cols = [
                    "UPC",
                    "Product Name",
                    "Inbound Qty",
                    "Value-add Qty",
                    "Canceled Qty",
                    "Transfer Qty",
                    "Total Inbound Qty",
                    "Grounding Qty",
                    "Issue"
                ]
                for c in cols:
                    if c not in hist_df.columns:
                        hist_df[c] = pd.NA
                hist_df = hist_df[cols]

                numeric_cols = [
                    "Inbound Qty",
                    "Value-add Qty",
                    "Canceled Qty",
                    "Transfer Qty",
                    "Total Inbound Qty",
                    "Grounding Qty"
                ]

                for df_tmp in (hist_df, today_result):
                    if not df_tmp.empty:
                        df_tmp["UPC"] = df_tmp["UPC"].astype(str)
                        for col in numeric_cols:
                            df_tmp[col] = pd.to_numeric(df_tmp[col], errors="coerce").fillna(0).astype(int)

                combined = pd.merge(
                    hist_df,
                    today_result,
                    on="UPC",
                    how="outer",
                    suffixes=("_hist", "_today")
                )

                combined["Product Name"] = combined["Product Name_today"].combine_first(combined["Product Name_hist"])

                for col in numeric_cols:
                    col_hist = f"{col}_hist"
                    col_today = f"{col}_today"
                    if col_hist not in combined.columns:
                        combined[col_hist] = 0
                    if col_today not in combined.columns:
                        combined[col_today] = 0
                    combined[col] = (
                        combined[col_hist].fillna(0).astype(int)
                        + combined[col_today].fillna(0).astype(int)
                    )

                def check_issue_cumulative(row):
                    if row["Total Inbound Qty"] == 0 and row["Grounding Qty"] > 0:
                        return "Missing Inbound"
                    if row["Grounding Qty"] == 0 and row["Total Inbound Qty"] > 0:
                        return "Missing Grounding"
                    if row["Total Inbound Qty"] != row["Grounding Qty"]:
                        return "Quantity Mismatch"
                    return None

                combined["Issue"] = combined.apply(check_issue_cumulative, axis=1)

                cumulative_unresolved = combined[combined["Issue"].notna()].copy()

                history_to_save = cumulative_unresolved[cols]

                export_df = history_to_save.copy()


            except Exception as e:
                messagebox.showerror(TEXT[language]["error_title"], f"Error updating historical.xlsx: {e}")
                export_df = today_result.copy()

        if export_df.empty:
            messagebox.showinfo(TEXT[language]["success_title"], "No discrepancy found. Export file may be empty.")
        if save_path.endswith(".csv"):
            export_df.to_csv(save_path, index=False, encoding="utf-8-sig")
        else:
            export_df.to_excel(save_path, index=False)

        msg = TEXT[language]["success"] + f"\n\n{save_path}"
        msg += "\n\nToday UPC counts by source:\n"
        msg += f"Inbound: {today_source_upc_counts['Inbound']}, "
        msg += f"Value-Add: {today_source_upc_counts['Value-Add']}, "
        msg += f"Cancel: {today_source_upc_counts['Cancel']}, "
        msg += f"Transfer: {today_source_upc_counts['Transfer']}"

        messagebox.showinfo(TEXT[language]["success_title"], msg)
        open_file_location(save_path)

        on_source_change()

    except Exception as e:
        messagebox.showerror(TEXT[language]["error_title"], str(e))


def switch_language(lang: str):
    global language
    language = lang
    refresh_ui()


def refresh_ui():
    root.title(TEXT[language]["title"])
    menubar.entryconfigure(0, label=TEXT[language]["menu_language"])
    lang_menu.entryconfigure(0, label=TEXT["zh"]["menu_zh"])
    lang_menu.entryconfigure(1, label=TEXT["en"]["menu_en"])

    for key, label in label_refs.items():
        label.config(text=TEXT[language].get(key, label.cget("text")))
    run_button.config(text=TEXT[language]["run_button"])
    refresh_button.config(text=TEXT[language]["refresh_button"])

    notebook.tab(checker_frame, text=TEXT[language]["tab_home"])
    notebook.tab(home_frame, text=TEXT[language]["tab_checker"])

    checker_title_label.config(text=TEXT[language]["title"])
    home_title_label.config(text=TEXT[language]["home_title"])
    date_label.config(text=TEXT[language]["home_date"])
    source_label.config(text=TEXT[language]["home_source"])
    count_label.config(text=TEXT[language]["home_count"])
    total_upc_label.config(text=TEXT[language]["home_total_upc"])
    add_record_button.config(text=TEXT[language]["home_add"])
    export_stats_button.config(text=TEXT[language]["home_export"])
    chart_button.config(text=TEXT[language]["home_chart"])
    start_date_label.config(text=TEXT[language]["home_start_date"])
    end_date_label.config(text=TEXT[language]["home_end_date"])
    chart_type_label.config(text=TEXT[language]["home_chart_type"])
    bar_radio.config(text=TEXT[language]["home_chart_bar"])
    line_radio.config(text=TEXT[language]["home_chart_line"])
    stats_file_label.config(text=TEXT[language]["stats_file_label"])
    stats_hint_label.config(text=TEXT[language]["stats_auto_hint"])


def create_file_selector(master, label_key: str, file_type: str, row: int):
    label = tk.Label(master, text=TEXT[language][label_key], font=("Arial", 10, "bold"))
    label.grid(row=row, column=0, padx=20, pady=8, sticky="w")
    tk.Button(master, text=TEXT[language]["browse"], command=lambda: select_file(file_type)).grid(
        row=row, column=1, padx=10
    )
    status = tk.Label(master, text=TEXT[language]["not_selected"], fg="gray", font=("Arial", 9))
    status.grid(row=row, column=2, padx=10, sticky="w")
    status_labels[file_type] = status
    label_refs[label_key] = label


def add_stats_record():
    global stats_df
    date_str = date_entry.get().strip()
    if not date_str:
        date_val = datetime.date.today()
    else:
        try:
            date_val = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror(TEXT[language]["error_title"], "Date format should be YYYY-MM-DD.")
            return

    source_val = source_combo.get().strip() or "Inbound"

    mismatch_str = count_entry.get().strip()
    total_upc_str = total_upc_entry.get().strip()

    try:
        mismatch_val = int(mismatch_str)
    except ValueError:
        messagebox.showerror(TEXT[language]["error_title"], "Mismatch count should be an integer.")
        return

    if not total_upc_str:
        total_val = int(today_source_upc_counts.get(source_val, 0))
    else:
        try:
            total_val = int(total_upc_str)
        except ValueError:
            messagebox.showerror(TEXT[language]["error_title"], "Total UPC count should be an integer.")
            return

    rate_val = mismatch_val / total_val if total_val > 0 else 0.0

    stats_df.loc[len(stats_df)] = [date_val, source_val, mismatch_val, total_val, rate_val]

    rate_str = f"{rate_val * 100:.1f}%"
    stats_tree.insert("", "end", values=(
        date_val.strftime("%Y-%m-%d"),
        source_val,
        mismatch_val,
        total_val,
        rate_str
    ))

    save_stats_history()

    count_entry.delete(0, tk.END)
    total_upc_entry.delete(0, tk.END)


def export_stats():
    if stats_df.empty:
        messagebox.showerror(TEXT[language]["error_title"], "No statistics data to export.")
        return
    filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
    if not save_path:
        return
    try:
        df_to_save = stats_df.copy()
        df_to_save["Date"] = df_to_save["Date"].astype(str)
        if save_path.endswith(".csv"):
            df_to_save.to_csv(save_path, index=False, encoding="utf-8-sig")
        else:
            df_to_save.to_excel(save_path, index=False)
        messagebox.showinfo(TEXT[language]["success_title"], TEXT[language]["success"] + f"\n\n{save_path}")
        open_file_location(save_path)
    except Exception as e:
        messagebox.showerror(TEXT[language]["error_title"], str(e))



def generate_chart():
    global chart_canvas

    if stats_df.empty:
        messagebox.showerror(TEXT[language]["error_title"], "No statistics data to plot.")
        return

    # Copy and apply date range filter
    df_plot = stats_df.copy()
    start_str = start_date_entry.get().strip()
    end_str = end_date_entry.get().strip()
    try:
        if start_str:
            start_date = datetime.datetime.strptime(start_str, "%Y-%m-%d").date()
            df_plot = df_plot[df_plot["Date"] >= start_date]
        if end_str:
            end_date = datetime.datetime.strptime(end_str, "%Y-%m-%d").date()
            df_plot = df_plot[df_plot["Date"] <= end_date]
    except ValueError:
        messagebox.showerror(TEXT[language]["error_title"], "Date format should be YYYY-MM-DD.")
        return

    if df_plot.empty:
        messagebox.showerror(TEXT[language]["error_title"], "No data in the selected date range.")
        return

    # Apply source filter (if user checked any source)
    if source_filter_vars:
        selected_sources = [src for src, var in source_filter_vars.items() if var.get()]
    else:
        selected_sources = sorted(df_plot["Source"].unique().tolist())

    if not selected_sources:
        # If nothing is selected, fall back to all sources in the data
        selected_sources = sorted(df_plot["Source"].unique().tolist())
        for src in source_filter_vars:
            source_filter_vars[src].set(True)

    df_plot = df_plot[df_plot["Source"].isin(selected_sources)]

    if df_plot.empty:
        messagebox.showerror(TEXT[language]["error_title"], "No data for the selected source(s) and date range.")
        return

    # Aggregate by Date + Source, then compute daily mismatch rate per source
    grouped = df_plot.groupby(["Date", "Source"], as_index=False).agg({
        "MismatchCount": "sum",
        "TotalUPC": "sum"
    })
    grouped["Rate"] = grouped.apply(
        lambda r: (r["MismatchCount"] / r["TotalUPC"]) if r["TotalUPC"] > 0 else 0.0,
        axis=1
    )

    if grouped.empty:
        messagebox.showerror(TEXT[language]["error_title"], "No aggregated data to plot.")
        return

    # Build complete date list to align different sources
    dates = sorted(grouped["Date"].unique())
    if not dates:
        messagebox.showerror(TEXT[language]["error_title"], "No dates available to plot.")
        return

    # Prepare figure
    if chart_canvas is not None:
        chart_canvas.get_tk_widget().destroy()
        chart_canvas = None

    fig = Figure(figsize=(6, 3), dpi=100)
    ax = fig.add_subplot(111)

    chart_mode = chart_type_var.get()

    if chart_mode == "line":
        # One line per source
        for src in selected_sources:
            df_src = grouped[grouped["Source"] == src]
            y_vals = []
            for d in dates:
                row = df_src[df_src["Date"] == d]
                if not row.empty:
                    y_vals.append(float(row["Rate"].iloc[0]))
                else:
                    y_vals.append(0.0)
            ax.plot(dates, y_vals, marker="o", label=src)
    else:
        # Grouped bar chart by date and source
        x_idx = list(range(len(dates)))
        n_src = max(len(selected_sources), 1)
        width = 0.8 / n_src  # total width around each date

        for i, src in enumerate(selected_sources):
            df_src = grouped[grouped["Source"] == src]
            y_vals = []
            for d in dates:
                row = df_src[df_src["Date"] == d]
                if not row.empty:
                    y_vals.append(float(row["Rate"].iloc[0]))
                else:
                    y_vals.append(0.0)
            # shift each source slightly so they appear side by side
            offsets = [x + (i - (n_src - 1) / 2) * width for x in x_idx]
            ax.bar(offsets, y_vals, width=width, label=src)

        # Use date strings as x-axis tick labels
        ax.set_xticks(x_idx)
        ax.set_xticklabels([d.strftime("%Y-%m-%d") for d in dates], rotation=45, ha="right")

    ax.set_xlabel("Date")
    ax.set_ylabel("Mismatch Rate")
    ax.set_title(TEXT[language]["home_title"])
    ax.legend()
    fig.autofmt_xdate()

    chart_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    chart_canvas.draw()
    chart_canvas.get_tk_widget().pack(fill="both", expand=True)
def on_source_change(event=None):
    src = source_combo.get().strip() or "Inbound"
    val = int(today_source_upc_counts.get(src, 0))
    total_upc_entry.delete(0, tk.END)
    if val > 0:
        total_upc_entry.insert(0, str(val))


root = tk.Tk()
root.title(TEXT[language]["title"])
root.geometry("900x740")

try:
    root.iconbitmap(resource_path("logo_on_shelf_bold.ico"))
except Exception:
    pass

menubar = tk.Menu(root)
lang_menu = tk.Menu(menubar, tearoff=0)
lang_menu.add_command(label=TEXT["zh"]["menu_zh"], command=lambda: switch_language("zh"))
lang_menu.add_command(label=TEXT["en"]["menu_en"], command=lambda: switch_language("en"))
menubar.add_cascade(label=TEXT[language]["menu_language"], menu=lang_menu)
root.config(menu=menubar)

try:
    img = Image.open(resource_path("logo_on_shelf_bold.jfif"))
    img = img.resize((80, 80), Image.Resampling.LANCZOS)
    logo_img = ImageTk.PhotoImage(img)
    img_label = tk.Label(root, image=logo_img)
    img_label.pack(pady=5)
except Exception:
    logo_img = None

notebook = ttk.Notebook(root)
checker_frame = tk.Frame(notebook)
home_frame = tk.Frame(notebook)
notebook.add(checker_frame, text=TEXT[language]["tab_home"])
notebook.add(home_frame, text=TEXT[language]["tab_checker"])
notebook.pack(expand=True, fill="both", pady=10)

checker_title_label = tk.Label(checker_frame, text=TEXT[language]["title"], font=("Arial", 14, "bold"))
checker_title_label.pack(pady=10)

file_frame = tk.Frame(checker_frame)
file_frame.pack(pady=10)

create_file_selector(file_frame, "grounding", "Grounding", 0)
create_file_selector(file_frame, "inbound", "Inbound", 1)
create_file_selector(file_frame, "export", "Export", 2)
create_file_selector(file_frame, "cancel", "Cancel", 3)
create_file_selector(file_frame, "transfer", "Transfer", 4)
create_file_selector(file_frame, "history", "History", 5)

refresh_button = tk.Button(
    checker_frame,
    text=TEXT[language]["refresh_button"],
    font=("Arial", 10, "bold"),
    bg="#007ACC",
    fg="white",
    width=25,
    command=auto_load_files
)
refresh_button.pack(pady=5)

auto_load_files()

run_button = tk.Button(
    checker_frame,
    text=TEXT[language]["run_button"],
    font=("Arial", 12, "bold"),
    bg="green",
    fg="white",
    width=25,
    command=run_check
)
run_button.pack(pady=10)

home_title_label = tk.Label(home_frame, text=TEXT[language]["home_title"], font=("Arial", 14, "bold"))
home_title_label.pack(pady=10)

home_top_frame = tk.Frame(home_frame)
home_top_frame.pack(pady=5, fill="x", padx=20)

date_label = tk.Label(home_top_frame, text=TEXT[language]["home_date"])
date_label.grid(row=0, column=0, sticky="w", padx=5, pady=2)
date_entry = tk.Entry(home_top_frame, width=15)
date_entry.grid(row=0, column=1, padx=5, pady=2)
date_entry.insert(0, datetime.date.today().strftime("%Y-%m-%d"))

source_label = tk.Label(home_top_frame, text=TEXT[language]["home_source"])
source_label.grid(row=0, column=2, sticky="w", padx=5, pady=2)
source_combo = ttk.Combobox(
    home_top_frame,
    values=["Inbound", "Value-Add", "Transfer", "Cancel"],
    state="readonly",
    width=15
)
source_combo.grid(row=0, column=3, padx=5, pady=2)
source_combo.set("Inbound")
source_combo.bind("<<ComboboxSelected>>", on_source_change)

count_label = tk.Label(home_top_frame, text=TEXT[language]["home_count"])
count_label.grid(row=0, column=4, sticky="w", padx=5, pady=2)
count_entry = tk.Entry(home_top_frame, width=10)
count_entry.grid(row=0, column=5, padx=5, pady=2)

total_upc_label = tk.Label(home_top_frame, text=TEXT[language]["home_total_upc"])
total_upc_label.grid(row=0, column=6, sticky="w", padx=5, pady=2)
total_upc_entry = tk.Entry(home_top_frame, width=10)
total_upc_entry.grid(row=0, column=7, padx=5, pady=2)

add_record_button = tk.Button(home_top_frame, text=TEXT[language]["home_add"], command=add_stats_record)
add_record_button.grid(row=0, column=8, padx=10, pady=2)

stats_hint_label = tk.Label(home_frame, text=TEXT[language]["stats_auto_hint"], fg="gray", font=("Arial", 9))
stats_hint_label.pack(anchor="w", padx=20, pady=2)

stats_tree = ttk.Treeview(
    home_frame,
    columns=("date", "source", "mismatch", "total", "rate"),
    show="headings",
    height=8
)
stats_tree.heading("date", text=TEXT[language]["home_date"])
stats_tree.heading("source", text=TEXT[language]["home_source"])
stats_tree.heading("mismatch", text=TEXT[language]["home_count"])
stats_tree.heading("total", text=TEXT[language]["home_total_upc"])
stats_tree.heading("rate", text="Mismatch Rate")
stats_tree.column("date", width=110)
stats_tree.column("source", width=110)
stats_tree.column("mismatch", width=110)
stats_tree.column("total", width=110)
stats_tree.column("rate", width=120)
stats_tree.pack(padx=20, pady=10, fill="x")

home_bottom_frame = tk.Frame(home_frame)
home_bottom_frame.pack(pady=5, fill="x", padx=20)

export_stats_button = tk.Button(home_bottom_frame, text=TEXT[language]["home_export"], command=export_stats)
export_stats_button.grid(row=0, column=0, padx=5, pady=2)

start_date_label = tk.Label(home_bottom_frame, text=TEXT[language]["home_start_date"])
start_date_label.grid(row=1, column=0, sticky="w", padx=5, pady=2)
start_date_entry = tk.Entry(home_bottom_frame, width=15)
start_date_entry.grid(row=1, column=1, padx=5, pady=2)

end_date_label = tk.Label(home_bottom_frame, text=TEXT[language]["home_end_date"])
end_date_label.grid(row=1, column=2, sticky="w", padx=5, pady=2)
end_date_entry = tk.Entry(home_bottom_frame, width=15)
end_date_entry.grid(row=1, column=3, padx=5, pady=2)

chart_type_label = tk.Label(home_bottom_frame, text=TEXT[language]["home_chart_type"])
chart_type_label.grid(row=1, column=4, sticky="w", padx=5, pady=2)
chart_type_var = tk.StringVar(value="bar")
bar_radio = tk.Radiobutton(home_bottom_frame, text=TEXT[language]["home_chart_bar"],
                           variable=chart_type_var, value="bar")
bar_radio.grid(row=1, column=5, padx=5, pady=2)
line_radio = tk.Radiobutton(home_bottom_frame, text=TEXT[language]["home_chart_line"],
                            variable=chart_type_var, value="line")
line_radio.grid(row=1, column=6, padx=5, pady=2)

chart_button = tk.Button(home_bottom_frame, text=TEXT[language]["home_chart"], command=generate_chart)
chart_button.grid(row=1, column=7, padx=10, pady=2)

# Source filter checkbuttons for chart (optional)
source_filter_label = tk.Label(home_bottom_frame, text="Sources:")
source_filter_label.grid(row=2, column=0, sticky="w", padx=5, pady=2)

source_filter_vars["Inbound"] = tk.BooleanVar(value=True)
source_filter_vars["Value-Add"] = tk.BooleanVar(value=True)
source_filter_vars["Transfer"] = tk.BooleanVar(value=True)
source_filter_vars["Cancel"] = tk.BooleanVar(value=True)

inbound_cb = tk.Checkbutton(home_bottom_frame, text="Inbound", variable=source_filter_vars["Inbound"])
inbound_cb.grid(row=2, column=1, padx=5, pady=2, sticky="w")

valueadd_cb = tk.Checkbutton(home_bottom_frame, text="Value-Add", variable=source_filter_vars["Value-Add"])
valueadd_cb.grid(row=2, column=2, padx=5, pady=2, sticky="w")

transfer_cb = tk.Checkbutton(home_bottom_frame, text="Transfer", variable=source_filter_vars["Transfer"])
transfer_cb.grid(row=2, column=3, padx=5, pady=2, sticky="w")

cancel_cb = tk.Checkbutton(home_bottom_frame, text="Cancel", variable=source_filter_vars["Cancel"])
cancel_cb.grid(row=2, column=4, padx=5, pady=2, sticky="w")

stats_file_label = tk.Label(home_frame, text=f"{TEXT[language]['stats_file_label']} {TEXT[language]['stats_no_file']}")
stats_file_label.pack(anchor="w", padx=20, pady=5)

chart_frame = tk.Frame(home_frame, height=250)
chart_frame.pack(fill="both", expand=True, padx=20, pady=5)

load_stats_history()

if __name__ == "__main__":
    root.mainloop()
