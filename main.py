from pathlib import Path
import sys

from PySide6 import QtCore, QtWidgets, QtGui

from parser_core import parse_xml_file
from exporters import (
    to_excel_multisheet,
    to_excel_with_photo_dropdown,
)

APP_DIR = Path(__file__).parent.resolve()
PHOTO_DROPDOWN_FIELDS = ["photoname", "photolat", "photolon"]


def _short(v, n=6):
    try:
        f = float(v)
        return f"{f:.{n}f}"
    except Exception:
        return str(v) if v is not None else ""


class MainWin(QtWidgets.QMainWindow):
    """
    Simple app: Load XML -> Preview/edit table -> Export to Excel (single or multi-sheet).
    Map code removed entirely.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("XML Survey Extractor – Preview & Export")
        self.resize(1100, 700)

        self.rows = []       # raw rows (with 'photos' list)
        self.flat_rows = []  # table rows (with 'photo' value)
        self.visible_cols = []

        # ---- Central UI: just the table + controls ----
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        v = QtWidgets.QVBoxLayout(central)
        v.setContentsMargins(8, 8, 8, 8)

        # Top controls
        top = QtWidgets.QHBoxLayout()
        self.btn_load = QtWidgets.QPushButton("Load XML…")
        self.btn_load.clicked.connect(self.load_xml)
        top.addWidget(self.btn_load)

        self.btn_choose_cols = QtWidgets.QPushButton("Columns…")
        self.btn_choose_cols.clicked.connect(self.choose_columns)
        self.btn_choose_cols.setEnabled(False)
        top.addWidget(self.btn_choose_cols)

        top.addStretch(1)

        self.chk_only_selected = QtWidgets.QCheckBox("Export only selected rows")
        top.addWidget(self.chk_only_selected)

        v.addLayout(top)

        # Table
        self.table = QtWidgets.QTableWidget(0, 0)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.setSortingEnabled(True)
        v.addWidget(self.table)

        tip = QtWidgets.QLabel("Tip: Drag columns to reorder. Double-click cells to edit. Ctrl/Shift for multi-select.")
        tip.setStyleSheet("color: gray;")
        v.addWidget(tip)

        # Toolbar with exports
        tb = QtWidgets.QToolBar("Main")
        self.addToolBar(tb)

        self.act_export_excel = QtGui.QAction("Export Excel…", self)
        self.act_export_excel.setEnabled(False)
        self.act_export_excel.triggered.connect(self.export_excel_from_table)
        tb.addAction(self.act_export_excel)

        self.act_export_multi = QtGui.QAction("Export Excel (multi-sheet)…", self)
        self.act_export_multi.setEnabled(False)
        self.act_export_multi.triggered.connect(self.export_excel_multisheet)
        tb.addAction(self.act_export_multi)

        self.status = self.statusBar()

    # ---------- Load & prepare ----------
    def load_xml(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Select XML files", str(APP_DIR), "XML files (*.xml)"
        )
        if not files:
            return

        all_rows, all_errors = [], []
        for f in files:
            rows, errs = parse_xml_file(f)
            all_rows.extend(rows)
            all_errors.extend(errs)

        self.rows = all_rows
        if not self.rows:
            QtWidgets.QMessageBox.information(self, "Parse Result", "No rows found.")
            self.act_export_excel.setEnabled(False)
            self.act_export_multi.setEnabled(False)
            self.btn_choose_cols.setEnabled(False)
            self._load_table([])
            return

        self.flat_rows = self._build_flat_rows_with_photo_value(self.rows)

        # Column order (no 'photos' key exposed)
        all_cols = list(self.flat_rows[0].keys())
        preferred = [
            "source_file", "seqno",
            "featurecoords_raw", "lat", "lon", "altitude_m",
            "event_timestamp", "gps_type", "gps_accuracy", "gps_speed",
            "history", "event_date_reported", "event_time_reported",
            "district", "state",
            "length_m", "breadth_m", "height_m",
            "type_landslide", "material", "occurrence", "structure",
            "trigger", "causes", "landslide_category", "remedial",
            "photo",
        ]
        self.visible_cols = [c for c in preferred if c in all_cols] + [
            c for c in all_cols if c not in preferred and c != "photos"
        ]

        self._load_table(self.flat_rows)
        self.act_export_excel.setEnabled(True)
        self.act_export_multi.setEnabled(True)
        self.btn_choose_cols.setEnabled(True)

        msg = f"Parsed {len(self.rows)} row(s)."
        if all_errors:
            msg += "\n\nNotes:\n" + "\n".join(all_errors[:8])
        QtWidgets.QMessageBox.information(self, "Parse Result", msg)
        self._show_basic_validation()

    def _build_flat_rows_with_photo_value(self, base_rows):
        out = []
        for r in base_rows:
            row = dict(r)
            photos = row.get("photos", [])
            default = ""
            for p in photos:
                for key in ("photolon", "photolat", "photoname"):
                    val = p.get(key, "")
                    if val:
                        default = val
                        break
                if default:
                    break
            row["photo"] = default
            out.append(row)
        return out

    # ---------- Table ----------
    def _load_table(self, rows):
        self.table.clear()
        if not rows:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        cols = [c for c in (self.visible_cols or list(rows[0].keys())) if c != "photos"]
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        self.table.setRowCount(len(rows))

        for r, row in enumerate(rows):
            for c, col in enumerate(cols):
                if col == "photo":
                    combo = QtWidgets.QComboBox()
                    combo.setEditable(False)
                    options = self._build_photo_options_for_row(row)
                    if not options:
                        combo.addItem("(no photo values)", userData="")
                    else:
                        for label, value in options:
                            combo.addItem(label, userData=value)
                        want = row.get("photo", "")
                        idx = 0
                        for i in range(combo.count()):
                            if str(combo.itemData(i)) == str(want):
                                idx = i
                                break
                        combo.setCurrentIndex(idx)

                    def on_change(i, rr=r, cc=c, cmb=combo):
                        val = cmb.itemData(i)
                        self.table.setItem(rr, cc, QtWidgets.QTableWidgetItem("" if val is None else str(val)))

                    combo.currentIndexChanged.connect(on_change)
                    self.table.setItem(r, c, QtWidgets.QTableWidgetItem(row.get("photo", "")))
                    self.table.setCellWidget(r, c, combo)
                else:
                    val = row.get(col, "")
                    item = QtWidgets.QTableWidgetItem(str(val))
                    if col in ("lat", "lon", "altitude_m"):
                        try:
                            item.setData(QtCore.Qt.ItemDataRole.UserRole, float(val))
                        except Exception:
                            pass
                    self.table.setItem(r, c, item)

    def _build_photo_options_for_row(self, row):
        opts = []
        photos = row.get("photos", [])
        for i, p in enumerate(photos, start=1):
            if "photoname" in PHOTO_DROPDOWN_FIELDS and p.get("photoname"):
                opts.append((f"photo{i} name", p.get("photoname", "")))
            if "photolat" in PHOTO_DROPDOWN_FIELDS and p.get("photolat"):
                opts.append((f"photo{i} lat", p.get("photolat", "")))
            if "photolon" in PHOTO_DROPDOWN_FIELDS and p.get("photolon"):
                opts.append((f"photo{i} lon", p.get("photolon", "")))
        return opts

    def choose_columns(self):
        if not self.flat_rows:
            QtWidgets.QMessageBox.information(self, "Columns", "Load some XML rows first.")
            return
        all_cols = [c for c in self.flat_rows[0].keys() if c != "photos"]
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Choose Columns")
        dlg.resize(360, 440)
        v = QtWidgets.QVBoxLayout(dlg)
        listw = QtWidgets.QListWidget()
        listw.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        for col in all_cols:
            it = QtWidgets.QListWidgetItem(col)
            it.setFlags(it.flags() | QtCore.Qt.ItemIsUserCheckable)
            it.setCheckState(QtCore.Qt.CheckState.Checked if col in self.visible_cols else QtCore.Qt.CheckState.Unchecked)
            listw.addItem(it)
        v.addWidget(listw)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        v.addWidget(bb)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        if dlg.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            sel = []
            for i in range(listw.count()):
                it = listw.item(i)
                if it.checkState() == QtCore.Qt.CheckState.Checked:
                    sel.append(it.text())
            if not sel:
                QtWidgets.QMessageBox.warning(self, "Columns", "At least one column must be selected.")
                return
            self.visible_cols = [c for c in sel if c != "photos"]
            self._load_table(self.flat_rows)

    # ---------- Excel export ----------
    def export_excel_from_table(self):
        if not self.rows:
            QtWidgets.QMessageBox.warning(self, "Export", "Nothing to export.")
            return
        selected_rows = sorted({i.row() for i in self.table.selectedIndexes()})
        use_selected = self.chk_only_selected.isChecked() and selected_rows
        selected_indices = selected_rows if use_selected else None
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Excel", str(APP_DIR / "export.xlsx"), "Excel (*.xlsx)"
        )
        if not out:
            return
        path = to_excel_with_photo_dropdown(self.rows, selected_indices, out)
        QtWidgets.QMessageBox.information(self, "Export", f"Saved Excel → {path}")
        self.status.showMessage(f"Saved: {path}", 5000)

    def export_excel_multisheet(self):
        if not self.rows:
            QtWidgets.QMessageBox.warning(self, "Export", "Nothing to export.")
            return
        out, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save Excel (multi-sheet)", str(APP_DIR / "export_multi.xlsx"), "Excel (*.xlsx)"
        )
        if not out:
            return
        path = to_excel_multisheet(self.rows, out)
        QtWidgets.QMessageBox.information(self, "Export", f"Saved Excel (multi-sheet) → {path}")
        self.status.showMessage(f"Saved: {path}", 4000)

    # ---------- Validation ----------
    def _show_basic_validation(self):
        valid = 0
        for r in self.rows:
            try:
                lat = float(r.get("lat", ""))
                lon = float(r.get("lon", ""))
                if -90 <= lat <= 90 and -180 <= lon <= 180:
                    valid += 1
            except Exception:
                pass
        self.status.showMessage(f"{valid}/{len(self.rows)} rows have numeric lat/lon.", 4000)


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWin()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
