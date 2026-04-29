def update_report_view_ui(self):
    if self.report_only_mode:
        self.btn_report_only.config(text="Sair do relatório")
        try:
            self._report_left_sash = self.center_pane.sashpos(0)
        except Exception:
            pass

        try:
            self.left_panel.pack_forget()
        except Exception:
            pass

        try:
            self.metrics_wrap.pack_forget()
        except Exception:
            pass

        try:
            self.top_header.pack_forget()
        except Exception:
            pass

        try:
            self.progress_wrap.pack_forget()
        except Exception:
            pass

        if not self.inline_progress_host.winfo_manager():
            self.inline_progress_host.pack(side="left", fill="x", expand=True, padx=(12, 0))

        try:
            self.editor_box.pack_forget()
        except Exception:
            pass

        try:
            self.xsb.pack_forget()
        except Exception:
            pass

        self.table_header.pack_configure(fill="x", pady=(0, 4))
        self.table_wrap.pack_configure(fill="both", expand=True)
        self.table_title_label.configure(font=("Segoe UI", 13, "bold"))
    else:
        self.btn_report_only.config(text="Somente relatório")

        try:
            self.inline_progress_host.pack_forget()
        except Exception:
            pass

        if not self.top_header.winfo_manager():
            self.top_header.pack(fill="x", pady=(0, 2), before=self.progress_wrap)

        if not self.progress_wrap.winfo_manager():
            self.progress_wrap.pack(fill="x", pady=(2, 4))

        if not self.metrics_wrap.winfo_manager():
            self.metrics_wrap.pack(fill="x", pady=METRICS_WRAP_PADY, after=self.progress_wrap)

        if not self.left_panel.winfo_manager():
            try:
                self.center_pane.insert(0, self.left_panel, weight=1)
            except Exception:
                try:
                    self.center_pane.add(self.left_panel, weight=1)
                except Exception:
                    pass

        if not self.editor_box.winfo_manager():
            self.editor_box.pack(fill="x", pady=(8, 0))

        if not self.xsb.winfo_manager():
            self.xsb.pack(fill="x")

        self.table_title_label.configure(font=("Segoe UI", 12, "bold"))

    self.update_fullscreen_ui()
    self.update_idletasks()

    if self.report_only_mode:
        try:
            total_w = self.center_pane.winfo_width()
            if total_w > 100:
                self.center_pane.sashpos(0, 1)
        except Exception:
            pass
    else:
        try:
            if self._report_left_sash is not None:
                self.center_pane.sashpos(0, self._report_left_sash)
        except Exception:
            pass