from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QLabel, QTabWidget, QPushButton,
    QSpacerItem, QSizePolicy, QDockWidget, QListWidget, QListWidgetItem,
    QDateEdit, QSpinBox, QFormLayout
)
from PySide6.QtCore import Qt, QDate
from matplotlib.figure import Figure
from PySide6.QtGui import QAction, QFont, QIcon
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from analytics import (
    load_and_validate, kpi, monthly_trends, quarterly_trends,
    regional_breakdown, top_bottom_products, get_filter_options, apply_filters,
    product_month_pivot_profit_filtered
)
from charts import (
    revenue_trend_with_fit, regional_pie, quarterly_trend_chart, margin_hist,
    heatmap_product_month
)


import pandas as pd
from pathlib import Path

from analytics import (
    load_and_validate, kpi, monthly_trends, quarterly_trends,
    regional_breakdown, top_bottom_products,
    by_customer_type, product_month_pivot_profit,
    region_month_pivot_sales, margins_describe, data_dictionary,
    monthly_growth_table
)
from charts import (
    revenue_trend_with_fit, regional_pie, quarterly_trend_chart, margin_hist
)
from export import export_pdf, export_excel_full, export_pngs

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sales Analytics (Stage D — Full Excel)")
        self.resize(1240, 820)

        self.df: pd.DataFrame | None = None
        self.figures = []

        # Menu
        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        export_menu = menubar.addMenu("Export")
        help_menu = menubar.addMenu("Help")

        self.act_load = QAction("Load Excel...", self)
        self.act_load.triggered.connect(self.on_load_excel)
        file_menu.addAction(self.act_load)

        self.act_export_pdf = QAction("Export PDF...", self); self.act_export_pdf.setEnabled(False)
        self.act_export_pdf.triggered.connect(self.on_export_pdf)
        export_menu.addAction(self.act_export_pdf)

        self.act_export_excel = QAction("Export Excel Summary (Full)...", self); self.act_export_excel.setEnabled(False)
        self.act_export_excel.triggered.connect(self.on_export_excel)
        export_menu.addAction(self.act_export_excel)

        self.act_export_png = QAction("Export PNG Charts...", self); self.act_export_png.setEnabled(False)
        self.act_export_png.triggered.connect(self.on_export_png)
        export_menu.addAction(self.act_export_png)

        self.act_about = QAction("User Guide", self)
        self.act_about.triggered.connect(self.on_show_help)
        help_menu.addAction(self.act_about)

        # Central widget with tabs
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        # ---- Filters Dock (left) ----
        self.filters = {
            "date_from": None,
            "date_to": None,
            "regions": [],
            "customer_types": [],
            "top_n": 5,
        }

        self.dock = QDockWidget("Filters", self)
        self.dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.LeftDockWidgetArea, self.dock)

        dock_widget = QWidget()
        self.dock.setWidget(dock_widget)

        form = QFormLayout(dock_widget)

        # Date range
        self.dt_from = QDateEdit(calendarPopup=True)
        self.dt_to = QDateEdit(calendarPopup=True)
        self.dt_from.setDisplayFormat("yyyy-MM-dd")
        self.dt_to.setDisplayFormat("yyyy-MM-dd")

        # Multi-select lists
        self.lst_regions = QListWidget(); self.lst_regions.setSelectionMode(QListWidget.MultiSelection)
        self.lst_ctypes = QListWidget();  self.lst_ctypes.setSelectionMode(QListWidget.MultiSelection)

        # Top-N
        self.spin_topn = QSpinBox(); self.spin_topn.setRange(1, 100); self.spin_topn.setValue(5)

        # Buttons
        self.btn_apply = QPushButton("Apply")
        self.btn_reset = QPushButton("Reset")

        form.addRow(QLabel("<b>Date from</b>"), self.dt_from)
        form.addRow(QLabel("<b>Date to</b>"), self.dt_to)
        form.addRow(QLabel("<b>Regions</b>"), self.lst_regions)
        form.addRow(QLabel("<b>Customer types</b>"), self.lst_ctypes)
        form.addRow(QLabel("<b>Top N products</b>"), self.spin_topn)
        form.addRow(self.btn_apply, self.btn_reset)

        self.btn_apply.clicked.connect(self.on_apply_filters)
        self.btn_reset.clicked.connect(self.on_reset_filters)

        # путь к папке assets рядом с ui.py
        assets_dir = Path(__file__).with_name("assets")
        icon_overview = QIcon(str(assets_dir / "overview.svg"))
        icon_charts = QIcon(str(assets_dir / "charts.svg"))
        icon_heatmap = QIcon(str(assets_dir / "heatmap.svg"))
        icon_more = QIcon(str(assets_dir / "more.svg"))

        # Overview tab
        self.tab_overview = QWidget()
        (self.tab_overview, icon_overview, "Overview")

        ov_layout = QVBoxLayout(self.tab_overview)
        self.kpi_label = QLabel("Load an Excel file to see KPIs.")
        self.kpi_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.kpi_label.setWordWrap(True)
        font = QFont()
        font.setPointSize(10)
        self.kpi_label.setFont(font)
        ov_layout.addWidget(self.kpi_label)

        self.top_bottom_label = QLabel("")
        self.top_bottom_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        self.top_bottom_label.setWordWrap(True)
        mono = QFont("Consolas")
        mono.setPointSize(9)
        self.top_bottom_label.setFont(mono)
        ov_layout.addWidget(self.top_bottom_label)

        # Charts tab
        self.tab_charts = QWidget()
        self.tabs.addTab(self.tab_charts, icon_charts, "Charts")

        ch_layout = QVBoxLayout(self.tab_charts)
        self.canvas_rev = FigureCanvas(revenue_trend_with_fit(pd.DataFrame({"month": [], "sales_amount": []})))
        self.canvas_reg = FigureCanvas(regional_pie(pd.DataFrame({"region": [], "sales_amount": []})))
        ch_layout.addWidget(self.canvas_rev)
        ch_layout.addWidget(self.canvas_reg)

        if not hasattr(self, "tab_heatmap"):
            from matplotlib.figure import Figure
            self.tab_heatmap = QWidget()
            hm_layout = QVBoxLayout(self.tab_heatmap)
            self.canvas_heatmap = FigureCanvas(Figure())
            hm_layout.addWidget(self.canvas_heatmap)

        # Heatmap tab  ← ДОЛЖНО БЫТЬ в __init__ ДО вызовов refresh_all
        self.tab_heatmap = QWidget()
        self.tabs.addTab(self.tab_heatmap, icon_heatmap, "Heatmap")
        hm_layout = QVBoxLayout(self.tab_heatmap)

        self.canvas_heatmap = FigureCanvas(Figure())  # пустая заглушка
        hm_layout.addWidget(self.canvas_heatmap)

        # More Charts tab
        self.tab_more = QWidget()
        self.tabs.addTab(self.tab_more, icon_more, "More Charts")
        mc_layout = QVBoxLayout(self.tab_more)
        self.canvas_quarter = FigureCanvas(quarterly_trend_chart(pd.DataFrame({"quarter": [], "sales_amount": []})))
        self.canvas_margin = FigureCanvas(margin_hist(pd.DataFrame({"margin": []})))
        mc_layout.addWidget(self.canvas_quarter)
        mc_layout.addWidget(self.canvas_margin)

        # Bottom bar with quick export buttons
        bar = QHBoxLayout()
        self.btn_export_pdf = QPushButton("Export PDF"); self.btn_export_pdf.setEnabled(False)
        self.btn_export_pdf.clicked.connect(self.on_export_pdf)
        self.btn_export_excel = QPushButton("Export Excel (Full)"); self.btn_export_excel.setEnabled(False)
        self.btn_export_excel.clicked.connect(self.on_export_excel)
        self.btn_export_png = QPushButton("Export PNGs"); self.btn_export_png.setEnabled(False)
        self.btn_export_png.clicked.connect(self.on_export_png)
        bar.addWidget(self.btn_export_pdf)
        bar.addWidget(self.btn_export_excel)
        bar.addWidget(self.btn_export_png)
        bar.addItem(QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        ch_layout.addLayout(bar)

        # (опционально) иконка окна
        app_icon_path = assets_dir / "app.ico"
        if app_icon_path.exists():
            self.setWindowIcon(QIcon(str(app_icon_path)))

    def on_show_help(self):
        QMessageBox.information(self, "User Guide",
            "1) File -> Load Excel... (columns: date, product, region, sales_amount, cost, customer_type)\n"
            "2) Overview: KPIs + Top/Bottom products\n"
            "3) Charts: Revenue trend (with trend line) + Regional pie\n"
            "4) More Charts: Quarterly revenue + Margin histogram\n"
            "5) Export: PDF/PNGs and Full Excel Summary (many sheets with formatting)"
        )

    def on_load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = load_and_validate(path)
        except Exception as e:
            QMessageBox.critical(self, "Validation error", str(e))
            return

        self.df = df
        # populate filters
        opts = get_filter_options(self.df)
        # даты
        if opts["date_min"] and opts["date_max"]:
            self.dt_from.setDate(QDate(opts["date_min"].year, opts["date_min"].month, opts["date_min"].day))
            self.dt_to.setDate(QDate(opts["date_max"].year, opts["date_max"].month, opts["date_max"].day))

        # списки регионов
        self.lst_regions.clear()
        for r in opts["regions"]:
            item = QListWidgetItem(str(r)); item.setSelected(False)
            self.lst_regions.addItem(item)

        # списки customer types
        self.lst_ctypes.clear()
        for ct in opts["customer_types"]:
            item = QListWidgetItem(str(ct)); item.setSelected(False)
            self.lst_ctypes.addItem(item)

        self.refresh_all()

        for act in (self.act_export_pdf, self.act_export_excel, self.act_export_png):
            act.setEnabled(True)
        for btn in (self.btn_export_pdf, self.btn_export_excel, self.btn_export_png):
            btn.setEnabled(True)

    def current_filters(self):
        # собираем выбранные значения
        date_from = self.dt_from.date().toPython() if self.dt_from.date().isValid() else None
        date_to   = self.dt_to.date().toPython() if self.dt_to.date().isValid() else None

        regions = [i.text() for i in self.lst_regions.selectedItems()]
        ctypes  = [i.text() for i in self.lst_ctypes.selectedItems()]
        top_n   = self.spin_topn.value()

        return {
            "date_from": date_from,
            "date_to": date_to,
            "regions": regions,
            "customer_types": ctypes,
            "top_n": top_n
        }

    def on_apply_filters(self):
        self.filters = self.current_filters()
        self.refresh_all()

    def on_reset_filters(self):
        # сброс выделений и дат на весь диапазон
        for i in range(self.lst_regions.count()):
            self.lst_regions.item(i).setSelected(False)
        for i in range(self.lst_ctypes.count()):
            self.lst_ctypes.item(i).setSelected(False)
        opts = get_filter_options(self.df) if self.df is not None else None
        if opts and opts["date_min"] and opts["date_max"]:
            self.dt_from.setDate(QDate(opts["date_min"].year, opts["date_min"].month, opts["date_min"].day))
            self.dt_to.setDate(QDate(opts["date_max"].year, opts["date_max"].month, opts["date_max"].day))
        self.spin_topn.setValue(5)
        self.filters = self.current_filters()
        self.refresh_all()

    def refresh_all(self):
        if self.df is None:
            return

        # 1) применяем фильтры
        f = self.filters
        df_filtered = apply_filters(
            self.df,
            date_from=f.get("date_from"),
            date_to=f.get("date_to"),
            regions=f.get("regions") or None,
            customer_types=f.get("customer_types") or None,
        )

        # 2) KPI
        k = kpi(df_filtered if not df_filtered.empty else self.df)
        lines = [
            f"Total Revenue: {k['total_revenue']:.2f}",
            f"Average Revenue: {k['avg_revenue']:.2f}",
            f"Total Profit: {k['total_profit']:.2f}",
        ]
        if k["avg_margin"] is not None:
            lines.append(f"Average Margin: {k['avg_margin'] * 100:.2f}%")
        if k["growth_mom"] is not None:
            lines.append(f"Growth MoM: {k['growth_mom'] * 100:.2f}%")
        self.kpi_label.setText("\n".join(lines))

        # 3) Top/Bottom по отфильтрованным
        top_n = max(1, int(f.get("top_n") or 5))
        top, bottom = top_bottom_products(df_filtered if not df_filtered.empty else self.df, n=top_n)

        def fmt(df):
            if df.empty:
                return "—"
            return "\n".join(f"{row.product:<12}  profit={row.profit:.2f}  sales={row.sales_amount:.2f}" for row in
                             df.itertuples(index=False))

        tb_txt = f"Top {top_n} products by profit:\n{fmt(top)}\n\nBottom {top_n} products by profit:\n{fmt(bottom)}"
        self.top_bottom_label.setText(tb_txt)

        # 4) Графики по отфильтрованным
        m = monthly_trends(df_filtered) if not df_filtered.empty else monthly_trends(self.df)
        q = quarterly_trends(df_filtered) if not df_filtered.empty else quarterly_trends(self.df)
        r = regional_breakdown(df_filtered) if not df_filtered.empty else regional_breakdown(self.df)

        new_rev = FigureCanvas(revenue_trend_with_fit(m))
        new_reg = FigureCanvas(regional_pie(r))
        new_qua = FigureCanvas(quarterly_trend_chart(q))
        new_mrg = FigureCanvas(margin_hist(df_filtered if not df_filtered.empty else self.df))

        charts_layout: QVBoxLayout = self.tab_charts.layout()
        charts_layout.replaceWidget(self.canvas_rev, new_rev)
        charts_layout.replaceWidget(self.canvas_reg, new_reg)
        self.canvas_rev.setParent(None);
        self.canvas_reg.setParent(None)
        self.canvas_rev = new_rev;
        self.canvas_reg = new_reg

        more_layout: QVBoxLayout = self.tab_more.layout()
        more_layout.replaceWidget(self.canvas_quarter, new_qua)
        more_layout.replaceWidget(self.canvas_margin, new_mrg)
        self.canvas_quarter.setParent(None);
        self.canvas_margin.setParent(None)
        self.canvas_quarter = new_qua;
        self.canvas_margin = new_mrg

        # 5) Heatmap из отфильтрованных данных
        pivot = product_month_pivot_profit_filtered(df_filtered if not df_filtered.empty else self.df)
        new_hm = FigureCanvas(heatmap_product_month(pivot))
        hm_layout: QVBoxLayout = self.tab_heatmap.layout()
        hm_layout.replaceWidget(self.canvas_heatmap, new_hm)
        self.canvas_heatmap.setParent(None)
        self.canvas_heatmap = new_hm

        # 6) список фигур для экспорта (добавим и heatmap)
        self.figures = [
            self.canvas_rev.figure,
            self.canvas_reg.figure,
            self.canvas_quarter.figure,
            self.canvas_margin.figure,
            self.canvas_heatmap.figure,
        ]

    def on_export_pdf(self):
        if not self.figures:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export PDF", "report.pdf", "PDF (*.pdf)")
        if not path:
            return
        try:
            export_pdf(self.figures, path, {"title": "Sales Analytics Report"})
            QMessageBox.information(self, "PDF Export", "PDF report saved.")
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))

    def on_export_excel(self):
        if self.df is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Export Excel Summary (Full)", "summary_full.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            compute_funcs = {
                "ByMonth": monthly_trends,
                "ByQuarter": quarterly_trends,
                "ByRegion": regional_breakdown,
                "ByCustomerType": by_customer_type,
                "TopProducts": lambda d: top_bottom_products(d, n=20)[0],
                "BottomProducts": lambda d: top_bottom_products(d, n=20)[1],
                "Product×Month_Profit": product_month_pivot_profit,
                "Region×Month_Sales": region_month_pivot_sales,
                "MarginsStats": margins_describe,
                "MonthlyGrowth": monthly_growth_table,
                "DataDictionary": data_dictionary,
            }
            export_excel_full(self.df, path, compute_funcs, kpi)
            QMessageBox.information(self, "Excel Export", "Full Excel summary saved.")
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))

    def on_export_png(self):
        if not self.figures:
            return
        dir_path = QFileDialog.getExistingDirectory(self, "Select output folder")
        if not dir_path:
            return
        try:
            export_pngs(self.figures, dir_path)
            QMessageBox.information(self, "PNG Export", "PNG charts saved.")
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))