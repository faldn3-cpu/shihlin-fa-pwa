# -*- coding: utf-8 -*-
"""
士林電機 FA 業務系統 - 加密 JSON 資料保險箱桌面檢視與密碼管理工具 (json_vault_viewer.py)

功能說明：
1. 在電腦記憶體內即時解密本機或從 Google 雲端硬碟下載的加密 .json 檔案（絕不落地產生明文暫存檔）。
2. 以中文表格呈現客戶資料、標準產品庫、個人特殊品、歷史報價單及產品明細，支援關鍵字即時篩選。
3. 提供「查看／修改主密碼」功能，修改後可一鍵重新加密本機檔案、網頁資料包並同步至 Google 雲端硬碟。
4. 嚴格遵守 Windows UTF-8 編碼、Zero Emoji 與 .env 機密隔離規範。
"""

import csv
import json
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any, Dict, List, Optional

if sys.platform.startswith("win"):
    try:
        reconfig_out = getattr(sys.stdout, "reconfigure", None)
        if callable(reconfig_out):
            reconfig_out(encoding="utf-8")
        reconfig_err = getattr(sys.stderr, "reconfigure", None)
        if callable(reconfig_err):
            reconfig_err(encoding="utf-8")
    except Exception:
        pass

import data_access


class JsonVaultViewerApp:
    """加密 JSON 資料保險箱桌面視窗應用程式"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("士林電機 FA 業務系統 - 加密資料保險箱檢視與密碼管理工具")
        self.root.geometry("1280x780")
        self.root.minsize(960, 620)

        self.current_file_path: Optional[Path] = None
        self.current_decrypted_data: Any = None
        self.current_rows: List[Dict[str, Any]] = []
        self.filtered_rows: List[Dict[str, Any]] = []
        self.current_schema: str = "generic"

        self.var_password = tk.StringVar()
        self.var_show_pw = tk.BooleanVar(value=False)
        self.var_search = tk.StringVar()
        self.var_status = tk.StringVar(value="狀態：請從左側選擇檔案，或點擊「選擇檔案」開啟從雲端下載的 .json 檔")

        self._load_initial_password_from_env()
        self._build_ui()
        self.refresh_file_tree()

    def _load_initial_password_from_env(self) -> None:
        """從本機 .env 自動讀取目前設定的解密密碼（預設以星號遮罩顯示）"""
        try:
            pw = data_access.get_vault_password()
            self.var_password.set(pw)
        except Exception:
            self.var_password.set("")

    def _build_ui(self) -> None:
        """建立視窗主版面、頂部工具列、左側檔案清單與右側中文表格檢視區"""
        # 1. 頂部密碼鎖與控制工具列
        top_frame = ttk.Frame(self.root, padding=(10, 8))
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="解密密碼:").pack(side=tk.LEFT, padx=(0, 4))
        self.entry_pw = ttk.Entry(top_frame, textvariable=self.var_password, show="*", width=24)
        self.entry_pw.pack(side=tk.LEFT, padx=(0, 6))
        self.entry_pw.bind("<Return>", lambda _e: self.reload_current_file())

        chk_show = ttk.Checkbutton(
            top_frame,
            text="顯示密碼",
            variable=self.var_show_pw,
            command=self._toggle_password_visibility,
        )
        chk_show.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(top_frame, text="解鎖並重新載入", command=self.reload_current_file).pack(side=tk.LEFT, padx=3)
        ttk.Button(top_frame, text="開啟任意 JSON 檔...", command=self.open_external_json_file).pack(side=tk.LEFT, padx=3)
        ttk.Button(top_frame, text="匯出目前表格 (CSV)", command=self.export_current_table_csv).pack(side=tk.LEFT, padx=3)

        ttk.Separator(top_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Button(
            top_frame,
            text="變更密碼並重新加密全系統與雲端...",
            command=self.change_vault_password_and_resync,
        ).pack(side=tk.LEFT, padx=3)

        # 2. 狀態與關鍵字搜尋列
        bar_frame = ttk.Frame(self.root, padding=(10, 0, 10, 6))
        bar_frame.pack(fill=tk.X)

        lbl_status = ttk.Label(bar_frame, textvariable=self.var_status, foreground="#0f4c81")
        lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(bar_frame, text="關鍵字搜尋:").pack(side=tk.LEFT, padx=(8, 4))
        entry_search = ttk.Entry(bar_frame, textvariable=self.var_search, width=26)
        entry_search.pack(side=tk.LEFT)
        self.var_search.trace_add("write", lambda *_args: self._apply_search_filter())

        ttk.Button(bar_frame, text="重新掃描清單", command=self.refresh_file_tree).pack(side=tk.LEFT, padx=(8, 0))

        # 3. 主分割視窗（左側檔案樹、右側分頁資料檢視區）
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # 左側導覽樹
        left_frame = ttk.LabelFrame(main_pane, text="本機與雲端資料分類清單", padding=6)
        main_pane.add(left_frame, weight=1)

        self.tree_files = ttk.Treeview(left_frame, show="tree")
        sb_tree = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree_files.yview)
        self.tree_files.configure(yscrollcommand=sb_tree.set)
        self.tree_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_tree.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_files.bind("<<TreeviewSelect>>", self._on_tree_file_selected)

        # 右側檢視區
        right_notebook = ttk.Notebook(main_pane)
        main_pane.add(right_notebook, weight=4)

        # 分頁 1：中文表格檢視（含上方主表與下方明細表）
        tab_table = ttk.Frame(right_notebook, padding=4)
        right_notebook.add(tab_table, text="中文解密內容表格")

        table_split = ttk.PanedWindow(tab_table, orient=tk.VERTICAL)
        table_split.pack(fill=tk.BOTH, expand=True)

        # 上半部：主表格
        top_tbl_frame = ttk.LabelFrame(table_split, text="資料清單總覽", padding=4)
        table_split.add(top_tbl_frame, weight=3)

        self.tree_main = ttk.Treeview(top_tbl_frame, show="headings", selectmode="browse")
        sb_main_y = ttk.Scrollbar(top_tbl_frame, orient=tk.VERTICAL, command=self.tree_main.yview)
        sb_main_x = ttk.Scrollbar(top_tbl_frame, orient=tk.HORIZONTAL, command=self.tree_main.xview)
        self.tree_main.configure(yscrollcommand=sb_main_y.set, xscrollcommand=sb_main_x.set)
        sb_main_y.pack(side=tk.RIGHT, fill=tk.Y)
        sb_main_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_main.pack(fill=tk.BOTH, expand=True)
        self.tree_main.bind("<<TreeviewSelect>>", self._on_main_row_selected)

        # 下半部：明細展開區
        self.detail_frame = ttk.LabelFrame(table_split, text="點選上方報價單／客戶／草稿列，可於此檢視明細內容", padding=4)
        table_split.add(self.detail_frame, weight=2)

        self.tree_detail = ttk.Treeview(self.detail_frame, show="headings", selectmode="browse")
        sb_det_y = ttk.Scrollbar(self.detail_frame, orient=tk.VERTICAL, command=self.tree_detail.yview)
        sb_det_x = ttk.Scrollbar(self.detail_frame, orient=tk.HORIZONTAL, command=self.tree_detail.xview)
        self.tree_detail.configure(yscrollcommand=sb_det_y.set, xscrollcommand=sb_det_x.set)
        sb_det_y.pack(side=tk.RIGHT, fill=tk.Y)
        sb_det_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_detail.pack(fill=tk.BOTH, expand=True)

        # 分頁 2：原始檔案加密狀態檢視（讓使用者直接確認磁碟檔案為亂碼）
        tab_raw = ttk.Frame(right_notebook, padding=6)
        right_notebook.add(tab_raw, text="磁碟原始加密檔檢視")

        ttk.Label(
            tab_raw,
            text="說明：以下顯示的是儲存在硬碟與 Google 雲端硬碟上的真實原始內容。只要經過加密保護，外人打開只會看到亂碼。",
            foreground="#475569",
        ).pack(anchor=tk.W, pady=(0, 6))

        self.txt_raw_preview = tk.Text(tab_raw, wrap=tk.CHAR, font=("Consolas", 10))
        sb_raw_y = ttk.Scrollbar(tab_raw, orient=tk.VERTICAL, command=self.txt_raw_preview.yview)
        self.txt_raw_preview.configure(yscrollcommand=sb_raw_y.set)
        sb_raw_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_raw_preview.pack(fill=tk.BOTH, expand=True)

    def _toggle_password_visibility(self) -> None:
        self.entry_pw.configure(show="" if self.var_show_pw.get() else "*")

    def refresh_file_tree(self) -> None:
        """掃描本機集中管理資料夾與網頁加密資料包並填入左側清單"""
        for item in self.tree_files.get_children():
            self.tree_files.delete(item)

        store_dir = data_access.CENTRALIZED_LOCAL_STORE_DIR
        root_store = self.tree_files.insert("", tk.END, text="集中管理資料庫 (Centralized)", open=True)

        shared_dir = store_dir / "00_全體共用區"
        shared_node = self.tree_files.insert(root_store, tk.END, text="00_全體共用區", open=True)
        if shared_dir.exists():
            for fp in sorted(shared_dir.glob("*.json")):
                size_str = self._fmt_size(fp.stat().st_size)
                self.tree_files.insert(shared_node, tk.END, text=f"{fp.name} ({size_str})", values=(str(fp),))

        personal_dir = store_dir / "01_業務個人資料區"
        personal_node = self.tree_files.insert(root_store, tk.END, text="01_業務個人資料區", open=True)
        if personal_dir.exists():
            for s_dir in sorted(personal_dir.iterdir()):
                if not s_dir.is_dir():
                    continue
                s_node = self.tree_files.insert(personal_node, tk.END, text=s_dir.name, open=(s_dir.name == "曾維崧"))
                for fp in sorted(s_dir.glob("*.json")):
                    size_str = self._fmt_size(fp.stat().st_size)
                    self.tree_files.insert(s_node, tk.END, text=f"{fp.name} ({size_str})", values=(str(fp),))

        web_seed = data_access.DEFAULT_BUILTIN_SEED_PATH
        web_node = self.tree_files.insert("", tk.END, text="GitHub 網頁加密資料包", open=True)
        if web_seed.exists():
            size_str = self._fmt_size(web_seed.stat().st_size)
            self.tree_files.insert(web_node, tk.END, text=f"{web_seed.name} ({size_str})", values=(str(web_seed),))

    @staticmethod
    def _fmt_size(num_bytes: int) -> str:
        if num_bytes < 1024:
            return f"{num_bytes} B"
        if num_bytes < 1024 * 1024:
            return f"{num_bytes / 1024:.1f} KB"
        return f"{num_bytes / (1024 * 1024):.2f} MB"

    def _on_tree_file_selected(self, _event: Any = None) -> None:
        sel = self.tree_files.selection()
        if not sel:
            return
        vals = self.tree_files.item(sel[0], "values")
        if vals and vals[0]:
            self.load_json_file(Path(vals[0]))

    def open_external_json_file(self) -> None:
        chosen = filedialog.askopenfilename(
            title="選擇要檢視的加密或一般 JSON 檔案",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if chosen:
            self.load_json_file(Path(chosen))

    def reload_current_file(self) -> None:
        if self.current_file_path and self.current_file_path.exists():
            self.load_json_file(self.current_file_path)
        else:
            self.refresh_file_tree()

    def load_json_file(self, file_path: Path) -> None:
        """讀取磁碟上的 JSON 檔案、於記憶體解密並渲染成中文表格"""
        self.current_file_path = file_path
        if not file_path.exists():
            messagebox.showwarning("檔案不存在", f"找不到檔案：{file_path}")
            return

        try:
            raw_text = file_path.read_text(encoding="utf-8")
            self.txt_raw_preview.delete("1.0", tk.END)
            self.txt_raw_preview.insert("1.0", raw_text[:6000] + ("\n...（後續亂碼已省略顯示）" if len(raw_text) > 6000 else ""))
            envelope = json.loads(raw_text)
        except Exception as exc:
            self.var_status.set(f"狀態：檔案讀取失敗 - {exc}")
            messagebox.showerror("讀取失敗", f"無法解析檔案為 JSON 格式：\n{exc}")
            return

        is_encrypted = isinstance(envelope, dict) and envelope.get("__encrypted_vault__") is True
        pw = self.var_password.get().strip()
        try:
            decrypted = data_access.decrypt_json_payload(envelope, password=pw or None)
        except Exception as exc:
            self.current_decrypted_data = None
            self.current_rows = []
            self._clear_treeview(self.tree_main)
            self._clear_treeview(self.tree_detail)
            self.var_status.set(f"[鎖定中] {file_path.name} 已加密保護，請確認上方「解密密碼」是否正確。")
            messagebox.showerror("解密失敗", str(exc))
            return

        self.current_decrypted_data = decrypted
        self._detect_and_load_schema(file_path.name, decrypted)
        if is_encrypted and envelope.get("key_scope") == "web_auto":
            lock_label = "[網頁自動解鎖資料包（亂碼存放） -> 管理員驗證通過於記憶體解密]"
        elif is_encrypted:
            lock_label = "[管理員私人保險箱 -> 成功於記憶體解密]"
        else:
            lock_label = "[警告：此為舊版未加密檔案]"
        self.var_status.set(
            f"{lock_label} 檔案：{file_path.name} ｜ 大小：{self._fmt_size(file_path.stat().st_size)} ｜ 筆數：{len(self.current_rows)} 筆"
        )
        self._apply_search_filter()

    @staticmethod
    def _clear_treeview(tv: ttk.Treeview) -> None:
        for item in tv.get_children():
            tv.delete(item)
        tv["columns"] = ()

    def _detect_and_load_schema(self, filename: str, data: Any) -> None:
        """依據檔案類型自動辨識中文欄位結構"""
        self._clear_treeview(self.tree_detail)
        fname = filename.lower()

        if isinstance(data, dict) and "products" in data and "customers_by_sales" in data:
            self.current_schema = "vault_bundle"
            summary_rows: List[Dict[str, Any]] = []
            summary_rows.append({
                "分類": "00_全體共用區",
                "項目": "標準產品主檔",
                "業務員": "全體共用",
                "筆數": len(data.get("products") or []),
                "備註": f"建立時間：{data.get('generated_at', '')}",
            })
            for s_name in data.get("salespersons") or []:
                c_cnt = len((data.get("customers_by_sales") or {}).get(s_name, []))
                q_cnt = len((data.get("quotations_by_sales") or {}).get(s_name, []))
                d_cnt = len((data.get("drafts_by_sales") or {}).get(s_name, []))
                sp_cnt = len((data.get("special_products_by_sales") or {}).get(s_name, []))
                summary_rows.append({
                    "分類": "業務個人區",
                    "項目": f"{s_name} 個人完整資料",
                    "業務員": s_name,
                    "筆數": c_cnt + q_cnt + d_cnt + sp_cnt,
                    "備註": f"客戶 {c_cnt} 家 / 報價單 {q_cnt} 筆 / 草稿 {d_cnt} 筆 / 特殊品 {sp_cnt} 筆",
                })
            self.current_rows = summary_rows
            self._configure_tree_columns(self.tree_main, [
                ("分類", 120), ("項目", 200), ("業務員", 120), ("筆數", 90), ("備註", 480)
            ])
            return

        if isinstance(data, dict) and "settings" in data:
            self.current_schema = "settings"
            rows = [{"設定鍵": k, "設定值": str(v)} for k, v in (data.get("settings") or {}).items()]
            self.current_rows = rows
            self._configure_tree_columns(self.tree_main, [("設定鍵", 260), ("設定值", 600)])
            return

        if not isinstance(data, list):
            self.current_schema = "generic"
            self.current_rows = [{"內容預覽": json.dumps(data, ensure_ascii=False)}]
            self._configure_tree_columns(self.tree_main, [("內容預覽", 800)])
            return

        if fname.startswith("customers_"):
            self.current_schema = "customers"
            self.current_rows = [
                {
                    "_raw": c,
                    "公司名稱": c.get("name", ""),
                    "負責業務": c.get("owner_sales", ""),
                    "統一編號": c.get("tax_id", ""),
                    "預設聯絡人": c.get("default_contact", ""),
                    "電話": c.get("phone", ""),
                    "付款方式": c.get("payment_terms", ""),
                    "常用": "是" if c.get("is_favorite") else "",
                }
                for c in data
            ]
            self._configure_tree_columns(self.tree_main, [
                ("公司名稱", 260), ("負責業務", 100), ("統一編號", 110),
                ("預設聯絡人", 120), ("電話", 140), ("付款方式", 140), ("常用", 70)
            ])
        elif fname.startswith("quotations_"):
            self.current_schema = "quotations"
            self.current_rows = [
                {
                    "_raw": q,
                    "報價單號": q.get("quote_no", ""),
                    "版次": q.get("revision", 0),
                    "報價日期": q.get("quote_date", ""),
                    "客戶名稱": q.get("customer_name", ""),
                    "專案名稱": q.get("project_name", ""),
                    "未稅金額": f"{float(q.get('subtotal') or 0):,.0f}",
                    "狀態": q.get("status", ""),
                    "品項數": len(q.get("items") or []),
                }
                for q in data
            ]
            self._configure_tree_columns(self.tree_main, [
                ("報價單號", 160), ("版次", 60), ("報價日期", 110), ("客戶名稱", 220),
                ("專案名稱", 200), ("未稅金額", 110), ("狀態", 90), ("品項數", 70)
            ])
        elif "standard_products" in fname or fname.startswith("special_products_"):
            self.current_schema = "products"
            self.current_rows = [
                {
                    "_raw": p,
                    "型號": p.get("model_no", ""),
                    "規格說明": p.get("spec", ""),
                    "經銷價": f"{float(p.get('dealer_price') or 0):,.0f}",
                    "牌價": f"{float(p.get('list_price') or 0):,.0f}",
                    "庫存量": p.get("stock_qty", 0),
                    "料號": p.get("sku", ""),
                    "訂購品": "是" if p.get("is_order_item") else "",
                }
                for p in data
            ]
            self._configure_tree_columns(self.tree_main, [
                ("型號", 200), ("規格說明", 300), ("經銷價", 100),
                ("牌價", 100), ("庫存量", 80), ("料號", 160), ("訂購品", 70)
            ])
        else:
            self.current_schema = "generic"
            if len(data) > 0 and isinstance(data[0], dict):
                keys = [k for k in data[0].keys() if not isinstance(data[0][k], (list, dict))][:8]
                self.current_rows = [{k: str(row.get(k, "")) for k in keys} | {"_raw": row} for row in data]
                self._configure_tree_columns(self.tree_main, [(k, 140) for k in keys])
            else:
                self.current_rows = [{"資料": str(x)} for x in data]
                self._configure_tree_columns(self.tree_main, [("資料", 700)])

    @staticmethod
    def _configure_tree_columns(tv: ttk.Treeview, col_specs: List[tuple]) -> None:
        cols = [c[0] for c in col_specs]
        tv["columns"] = cols
        for col_name, col_w in col_specs:
            tv.heading(col_name, text=col_name)
            tv.column(col_name, width=col_w, minwidth=50)

    def _apply_search_filter(self) -> None:
        kw = self.var_search.get().strip().lower()
        for item in self.tree_main.get_children():
            self.tree_main.delete(item)

        if not kw:
            self.filtered_rows = list(self.current_rows)
        else:
            self.filtered_rows = [
                r for r in self.current_rows
                if any(kw in str(v).lower() for k, v in r.items() if k != "_raw")
            ]

        cols = self.tree_main["columns"]
        for idx, row in enumerate(self.filtered_rows[:2000]):
            vals = [row.get(c, "") for c in cols]
            self.tree_main.insert("", tk.END, iid=str(idx), values=vals)

    def _on_main_row_selected(self, _event: Any = None) -> None:
        sel = self.tree_main.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx < 0 or idx >= len(self.filtered_rows):
            return
        row_obj = self.filtered_rows[idx]
        raw = row_obj.get("_raw")
        self._clear_treeview(self.tree_detail)
        if not isinstance(raw, dict):
            return

        if self.current_schema == "quotations":
            items = raw.get("items") or []
            self._configure_tree_columns(self.tree_detail, [
                ("項次", 50), ("型號", 180), ("規格", 260), ("數量", 70),
                ("經銷價", 90), ("折數%", 80), ("最終單價", 100), ("小計", 110), ("備註", 160)
            ])
            for i, itm in enumerate(items, 1):
                self.tree_detail.insert("", tk.END, values=(
                    i,
                    itm.get("model_no", ""),
                    itm.get("spec", ""),
                    itm.get("quantity", 0),
                    f"{float(itm.get('dealer_price') or 0):,.0f}",
                    itm.get("discount_pct", ""),
                    f"{float(itm.get('final_unit_price') or 0):,.0f}",
                    f"{float(itm.get('row_subtotal') or 0):,.0f}",
                    itm.get("note_lead_time", ""),
                ))
        elif self.current_schema == "customers":
            contacts = raw.get("contacts") or []
            self._configure_tree_columns(self.tree_detail, [
                ("聯絡人姓名", 160), ("手機", 150), ("Email", 240)
            ])
            for ct in contacts:
                self.tree_detail.insert("", tk.END, values=(
                    ct.get("name", ""), ct.get("mobile", ""), ct.get("email", "")
                ))

    def export_current_table_csv(self) -> None:
        """將目前解密後表格安全匯出為 Excel 相容之 UTF-8-SIG CSV 檔"""
        if not self.filtered_rows:
            messagebox.showinfo("提示", "目前沒有可匯出的資料列。")
            return
        save_path = filedialog.asksaveasfilename(
            title="匯出解密表格為 CSV",
            defaultextension=".csv",
            filetypes=[("CSV 檔案 (Excel 開啟)", "*.csv")],
        )
        if not save_path:
            return
        cols = list(self.tree_main["columns"])
        with open(save_path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            for r in self.filtered_rows:
                writer.writerow([r.get(c, "") for c in cols])
        messagebox.showinfo("匯出成功", f"已匯出 {len(self.filtered_rows)} 筆資料至：\n{save_path}")

    def change_vault_password_and_resync(self) -> None:
        """修改主密碼，並以新密碼重新加密本機資料夾、網頁種子檔與同步 Google 雲端硬碟"""
        new_pw1 = simpledialog.askstring("設定新密碼", "請輸入新的資料加密主密碼（至少 4 碼）：", show="*", parent=self.root)
        if new_pw1 is None:
            return
        new_pw1 = new_pw1.strip()
        if len(new_pw1) < 4:
            messagebox.showwarning("格式不符", "密碼長度至少需要 4 個字元。")
            return

        new_pw2 = simpledialog.askstring("再次確認密碼", "請再輸入一次新的加密密碼：", show="*", parent=self.root)
        if new_pw2 is None or new_pw2.strip() != new_pw1:
            messagebox.showerror("密碼不一致", "兩次輸入的密碼不同，已取消變更。")
            return

        sync_cloud = messagebox.askyesno(
            "確認重新加密",
            "即將把新密碼寫入本機 .env，並以新密碼重新加密本機與雲端後台檔案\n"
            "（網頁端採用獨立自動解鎖機制，管理員修改私人密碼不影響業務正常免密碼登入）。\n\n"
            "是否同時重新加密並上傳覆蓋 Google 雲端硬碟「PWA業務系統」資料夾？",
        )
        try:
            data_access.save_vault_password_to_env(new_pw1)
            self.var_password.set(new_pw1)
            if sync_cloud:
                res = data_access.sync_to_gdrive_pwa_folder("PWA業務系統", password=new_pw1)
                msg = (
                    f"已完成管理員私人密碼更新與後台資料重新加密！\n"
                    f"雲端同步狀態：{res.get('status')}（共上傳 {res.get('uploaded_files', 0)} 個管理員專屬加密檔）\n"
                    f"（網頁端業務登入維持免輸密碼自動解鎖，不受影響）"
                )
            else:
                res_loc = data_access.export_centralized_sales_folders(encrypt=True, password=new_pw1)
                res_web = data_access.export_builtin_web_seed_json(encrypt=True)
                msg = (
                    f"已完成本機後台資料夾（{res_loc.get('files_written', 0)} 個管理員專屬加密檔）重新加密！\n"
                    f"網頁加密資料包（{res_web.get('product_count', 0)} 筆產品）維持自動解鎖機制，業務端無須變更密碼。"
                )
            self.refresh_file_tree()
            self.reload_current_file()
            messagebox.showinfo("更新完成", msg)
        except Exception as exc:
            messagebox.showerror("執行失敗", f"重新加密過程中發生錯誤：\n{exc}")


def main() -> None:
    root = tk.Tk()
    app = JsonVaultViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

