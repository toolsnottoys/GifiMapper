// ──────────────────────────────────────────────────────────────────────────────
//  MainForm.cs  —  GIFI Mapper v1.1  (Windows Forms .NET 8)
// ──────────────────────────────────────────────────────────────────────────────
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace GifiMapper
{
    public partial class MainForm : Form
    {
        // ── State ─────────────────────────────────────────────────────────────
        private DataTable _gifiTable      = new();
        private DataTable _tbTable        = new();
        private DataTable _mergedTable    = new();
        private DataTable _unmatchedTable = new();

        private DataView _gifiView      = new();
        private DataView _tbView        = new();
        private DataView _mergedView    = new();
        private DataView _unmatchedView = new();

        private bool _gifiLoaded = false;
        private bool _tbLoaded   = false;
        private bool _merged     = false;

        // ── Constructor ───────────────────────────────────────────────────────
        public MainForm()
        {
            InitializeComponent();
            WireEvents();
            TryAutoLoadGifi();
        }

        // ── Event wiring ──────────────────────────────────────────────────────
        private void WireEvents()
        {
            btnBrowseGifi.Click      += (_, _) => BrowseAndLoad(isGifi: true);
            btnLoadGifi.Click        += (_, _) => LoadGifi(txtGifiPath.Text);
            btnBrowseTB.Click        += (_, _) => BrowseAndLoad(isGifi: false);
            btnLoadTB.Click          += (_, _) => LoadTrialBalance(txtTBPath.Text);

            btnMerge.Click           += (_, _) => DoMerge();
            btnSaveMerged.Click      += (_, _) => SaveCsv(_mergedTable,    "Save Merged File");
            btnSaveUnmatched.Click   += (_, _) => SaveCsv(_unmatchedTable, "Save Unmatched File");
            btnExportExcel.Click     += (_, _) => ExportExcel();
            btnExportFutureTax.Click += (_, _) => ExportFutureTax();
            btnClear.Click           += (_, _) => ClearAll();

            menuFileLoadGifi.Click        += (_, _) => BrowseAndLoad(isGifi: true);
            menuFileLoadTB.Click          += (_, _) => BrowseAndLoad(isGifi: false);
            menuFileSaveMerged.Click      += (_, _) => SaveCsv(_mergedTable,    "Save Merged File");
            menuFileSaveUnmatched.Click   += (_, _) => SaveCsv(_unmatchedTable, "Save Unmatched File");
            menuFileExportExcel.Click     += (_, _) => ExportExcel();
            menuFileExportFutureTax.Click += (_, _) => ExportFutureTax();
            menuFileExit.Click            += (_, _) => Application.Exit();
            menuHelpAbout.Click           += (_, _) => ShowAbout();

            menuFileLoadGifi.ShortcutKeys        = Keys.Control | Keys.G;
            menuFileLoadTB.ShortcutKeys          = Keys.Control | Keys.T;
            menuFileSaveMerged.ShortcutKeys      = Keys.Control | Keys.S;
            menuFileExportExcel.ShortcutKeys     = Keys.Control | Keys.E;
            menuFileExportFutureTax.ShortcutKeys = Keys.Control | Keys.F;

            txtSearch.TextChanged              += (_, _) => ApplySearch();
            btnSearchClear.Click               += (_, _) => { txtSearch.Clear(); ApplySearch(); };
            tabControl.SelectedIndexChanged    += (_, _) => ApplySearch();
            dgvUnmatched.CellDoubleClick       += DgvUnmatched_CellDoubleClick;

            foreach (var dgv in AllGrids())
                dgv.KeyDown += Dgv_KeyDown;
        }

        // ── Auto-load GIFI from exe folder ───────────────────────────────────
        private void TryAutoLoadGifi()
        {
            var candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IG_gifi_codes.csv");
            if (File.Exists(candidate))
            {
                txtGifiPath.Text = candidate;
                LoadGifi(candidate);
            }
        }

        // ── Browse ────────────────────────────────────────────────────────────
        private void BrowseAndLoad(bool isGifi)
        {
            using var dlg = new OpenFileDialog
            {
                Title  = isGifi ? "Select GIFI Mapping CSV" : "Select Trial Balance CSV",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            if (isGifi) { txtGifiPath.Text = dlg.FileName; LoadGifi(dlg.FileName); }
            else        { txtTBPath.Text   = dlg.FileName; LoadTrialBalance(dlg.FileName); }
        }

        // ── Load GIFI mapping ─────────────────────────────────────────────────
        private void LoadGifi(string path)
        {
            if (!ValidatePath(path, "GIFI mapping")) return;
            try
            {
                SetStatus("Loading GIFI mapping…", true);
                _gifiTable = ReadCsvRaw(path, hasHeader: true,
                    forceHeaders: new[] { "Accounting Code", "GIFI Code", "Description" });
                _gifiView = _gifiTable.DefaultView;
                dgvGifi.DataSource = _gifiView;
                StyleGifiGrid();
                _gifiLoaded = true;
                lblGifiCount.Text = $"GIFI: {_gifiTable.Rows.Count} rows";
                tabControl.SelectedTab = tabGifi;
                SetStatus($"GIFI mapping loaded — {_gifiTable.Rows.Count} codes.", false);
            }
            catch (Exception ex) { ShowError("Failed to load GIFI file", ex); }
        }

        // ── Load trial balance ────────────────────────────────────────────────
        private void LoadTrialBalance(string path)
        {
            if (!ValidatePath(path, "trial balance")) return;
            try
            {
                SetStatus("Loading trial balance…", true);
                _tbTable = ReadCsvRaw(path, hasHeader: false);
                if (_tbTable.Columns.Count > 0)
                    _tbTable.Columns[0].ColumnName = "Accounting Code";
                for (int i = 1; i < _tbTable.Columns.Count; i++)
                    _tbTable.Columns[i].ColumnName = $"Column {i + 1}";
                _tbView = _tbTable.DefaultView;
                dgvTB.DataSource = _tbView;
                StyleGrid(dgvTB, "Accounting Code");
                _tbLoaded = true;
                lblTBCount.Text = $"TB: {_tbTable.Rows.Count} rows";
                tabControl.SelectedTab = tabTB;
                SetStatus($"Trial balance loaded — {_tbTable.Rows.Count} rows.", false);
                UpdateMergeButton();
            }
            catch (Exception ex) { ShowError("Failed to load trial balance", ex); }
        }

        // ── Merge ─────────────────────────────────────────────────────────────
        private void DoMerge()
        {
            if (!_gifiLoaded) { Warn("Please load the GIFI mapping file first."); return; }
            if (!_tbLoaded)   { Warn("Please load the trial balance file first."); return; }
            try
            {
                SetStatus("Merging…", true);

                var lookup = new Dictionary<string, (string Gifi, string Desc)>(StringComparer.OrdinalIgnoreCase);
                foreach (DataRow r in _gifiTable.Rows)
                {
                    var key  = r["Accounting Code"]?.ToString()?.Trim() ?? "";
                    var gifi = r["GIFI Code"]?.ToString()?.Trim() ?? "";
                    var desc = _gifiTable.Columns.Count > 2 ? r["Description"]?.ToString()?.Trim() ?? "" : "";
                    if (!string.IsNullOrEmpty(key)) lookup[key] = (gifi, desc);
                }

                _mergedTable = _tbTable.Clone();
                _mergedTable.Columns.Add("GIFI Code",        typeof(string));
                _mergedTable.Columns.Add("GIFI Description", typeof(string));
                _unmatchedTable = _mergedTable.Clone();

                int matched = 0, unmatched = 0;
                foreach (DataRow tbRow in _tbTable.Rows)
                {
                    var acct   = tbRow["Accounting Code"]?.ToString()?.Trim() ?? "";
                    var newRow = _mergedTable.NewRow();
                    foreach (DataColumn col in _tbTable.Columns)
                        newRow[col.ColumnName] = tbRow[col];

                    if (lookup.TryGetValue(acct, out var hit))
                    {
                        newRow["GIFI Code"]        = hit.Gifi;
                        newRow["GIFI Description"] = hit.Desc;
                        matched++;
                    }
                    else
                    {
                        newRow["GIFI Code"]        = DBNull.Value;
                        newRow["GIFI Description"] = DBNull.Value;
                        unmatched++;
                        _unmatchedTable.ImportRow(newRow);
                    }
                    _mergedTable.Rows.Add(newRow);
                }

                _mergedView    = _mergedTable.DefaultView;
                _unmatchedView = _unmatchedTable.DefaultView;
                dgvMerged.DataSource    = _mergedView;
                dgvUnmatched.DataSource = _unmatchedView;
                StyleMergedGrid(dgvMerged);
                StyleMergedGrid(dgvUnmatched);
                ColorUnmatchedRows(dgvMerged);

                _merged = true;
                lblMatchedCount.Text   = $"Matched: {matched}";
                lblUnmatchedCount.Text = $"Unmatched: {unmatched}";

                SetSaveButtonsEnabled(true);

                if (unmatched > 0)
                {
                    tabUnmatched.Text = $"⚠  Unmatched ({unmatched})";
                    tabControl.SelectedTab = tabUnmatched;
                    tabControl.SelectedTab = tabMerged;
                }
                else
                {
                    tabUnmatched.Text = "⚠  Unmatched";
                    tabControl.SelectedTab = tabMerged;
                }

                SetStatus(unmatched == 0
                    ? $"Merge complete — all {matched} rows matched."
                    : $"Merge complete — {matched} matched, {unmatched} unmatched (see ⚠ tab).", false);
            }
            catch (Exception ex) { ShowError("Merge failed", ex); }
        }

        // ── Save CSV ──────────────────────────────────────────────────────────
        private void SaveCsv(DataTable table, string title)
        {
            if (table.Rows.Count == 0) { MessageBox.Show("Nothing to save.", title, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            using var dlg = new SaveFileDialog
            {
                Title = title, Filter = "CSV Files (*.csv)|*.csv", DefaultExt = "csv",
                FileName = title.Contains("Unmatched") ? "unmatched.csv" : "merged.csv"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                WriteCsv(table, dlg.FileName);
                SetStatus($"Saved: {dlg.FileName}", false);
                if (MessageBox.Show("File saved. Open it now?", "Saved", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    OpenFile(dlg.FileName);
            }
            catch (Exception ex) { ShowError("Save failed", ex); }
        }

        // ── Export Excel ──────────────────────────────────────────────────────
        private void ExportExcel()
        {
            if (!_merged) { Warn("Please run Merge first."); return; }
            using var dlg = new SaveFileDialog
            {
                Title = "Export to Excel", Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx", FileName = "gifi_merged.xlsx"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                SetStatus("Exporting to Excel…", true);
                using var wb = new XLWorkbook();
                WriteSheet(wb, "Merged",    _mergedTable,    highlightUnmatched: true);
                WriteSheet(wb, "Unmatched", _unmatchedTable, highlightUnmatched: false);
                WriteSheet(wb, "GIFI Map",  _gifiTable,      highlightUnmatched: false);
                wb.SaveAs(dlg.FileName);
                SetStatus($"Excel exported: {dlg.FileName}", false);
                if (MessageBox.Show("Excel file saved. Open it now?", "Exported", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    OpenFile(dlg.FileName);
            }
            catch (Exception ex) { ShowError("Excel export failed", ex); }
        }

        // ── Export for FutureTax ──────────────────────────────────────────────
        // Column 3 = Debit  (positive),  Column 4 = Credit (negative).
        // Net per GIFI = sum(Debits) - sum(Credits).
        // Amount written to file = Math.Abs(Net).
        private void ExportFutureTax()
        {
            if (!_merged) { Warn("Please run Merge first."); return; }

            // Confirm Column 3 and Column 4 exist
            const string debitCol  = "Column 3";
            const string creditCol = "Column 4";
            bool hasDebit  = _mergedTable.Columns.Contains(debitCol);
            bool hasCredit = _mergedTable.Columns.Contains(creditCol);

            if (!hasDebit || !hasCredit)
            {
                MessageBox.Show(
                    $"Expected columns \"{debitCol}\" (Debit) and \"{creditCol}\" (Credit) " +
                    $"in the trial balance.\n\n" +
                    $"Found: {string.Join(", ", _mergedTable.Columns.Cast<DataColumn>().Select(c => c.ColumnName))}",
                    "Missing Columns", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var dlg = new SaveFileDialog
            {
                Title = "Save FutureTax Import File",
                Filter = "CSV Files (*.csv)|*.csv|GFI Files (*.gfi)|*.gfi",
                DefaultExt = "csv", FileName = "futuretax_import.csv"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                SetStatus("Building FutureTax file…", true);

                // Accumulate debit and credit totals per GIFI code separately
                var debits  = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
                var credits = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

                foreach (DataRow row in _mergedTable.Rows)
                {
                    var gifi = row["GIFI Code"]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(gifi)) continue;

                    debits[gifi]  = debits.GetValueOrDefault(gifi)  + ParseAmount(row[debitCol]);
                    credits[gifi] = credits.GetValueOrDefault(gifi) + ParseAmount(row[creditCol]);
                }

                if (debits.Count == 0 && credits.Count == 0)
                {
                    MessageBox.Show("No numeric values found in Column 3 or Column 4.\nCheck the trial balance file.",
                        "Nothing to Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    SetStatus("FutureTax export cancelled.", false);
                    return;
                }

                // Build result: net = debit - credit,  exported amount = Abs(net)
                var allGifi = debits.Keys.Union(credits.Keys, StringComparer.OrdinalIgnoreCase)
                                         .OrderBy(g => g)
                                         .ToList();

                var results = allGifi.Select(gifi =>
                {
                    decimal debit  = debits.GetValueOrDefault(gifi);
                    decimal credit = credits.GetValueOrDefault(gifi);
                    decimal net    = debit - credit;
                    decimal absAmt = Math.Abs(net);
                    return (Gifi: gifi, Debit: debit, Credit: credit, Net: net, AbsAmt: absAmt);
                }).ToList();

                // Write: Code,AbsoluteAmount  (no header row — FutureTax format)
                using (var w = new StreamWriter(dlg.FileName, false, Encoding.UTF8))
                    foreach (var r in results)
                        w.WriteLine($"{r.Gifi},{(long)Math.Round(r.AbsAmt, MidpointRounding.AwayFromZero)}");

                SetStatus($"FutureTax file saved: {dlg.FileName} ({results.Count} GIFI codes)", false);
                ShowFutureTaxPreview(results, dlg.FileName);
            }
            catch (Exception ex) { ShowError("FutureTax export failed", ex); }
        }

        // Parse an amount cell: handles blanks, commas, $, and (1234) bracket negatives
        private static decimal ParseAmount(object? cell)
        {
            var raw = cell?.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(raw)) return 0m;
            bool neg = raw.StartsWith('(') && raw.EndsWith(')');
            raw = raw.Replace("(", "").Replace(")", "")
                     .Replace(",", "").Replace("$", "").Trim();
            if (!decimal.TryParse(raw,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out decimal amt)) return 0m;
            return neg ? -amt : amt;
        }

        private void ShowFutureTaxPreview(
            List<(string Gifi, decimal Debit, decimal Credit, decimal Net, decimal AbsAmt)> results,
            string savedPath)
        {
            var frm = new Form
            {
                Text = "FutureTax Export — Preview", Size = new Size(590, 460),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false
            };
            var lbl = new Label
            {
                Text = $"✅  {results.Count} GIFI codes written.\nFile: {Path.GetFileName(savedPath)}",
                Left = 10, Top = 8, Width = 560, Height = 36, AutoSize = false
            };
            var dgv = MakeGrid();
            dgv.Left = 10; dgv.Top = 50; dgv.Width = 560; dgv.Height = 310;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            var dt = new DataTable();
            dt.Columns.Add("GIFI Code");
            dt.Columns.Add("Debit Total");
            dt.Columns.Add("Credit Total");
            dt.Columns.Add("Net (D−C)");
            dt.Columns.Add("Exported Amount");
            foreach (var r in results)
                dt.Rows.Add(r.Gifi,
                    r.Debit.ToString("0.##"),
                    r.Credit.ToString("0.##"),
                    r.Net.ToString("0.##"),
                    ((long)Math.Round(r.AbsAmt, MidpointRounding.AwayFromZero)).ToString());
            dgv.DataSource = dt;

            var btnOpen  = new Button { Text = "Open File", Left = 200, Top = 372, Width = 110, Height = 28 };
            var btnClose = new Button { Text = "Close",      Left = 322, Top = 372, Width = 80,  Height = 28, DialogResult = DialogResult.OK };
            btnOpen.Click += (_, _) => OpenFile(savedPath);
            frm.Controls.AddRange(new Control[] { lbl, dgv, btnOpen, btnClose });
            frm.AcceptButton = btnClose;
            frm.ShowDialog(this);
        }

        // ── Excel helper ──────────────────────────────────────────────────────
        private static void WriteSheet(XLWorkbook wb, string name, DataTable dt, bool highlightUnmatched)
        {
            var ws = wb.Worksheets.Add(name);

            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = dt.Columns[c].ColumnName;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#33558B");
                cell.Style.Font.FontColor       = XLColor.White;
                cell.Style.Font.Bold            = true;
            }

            int gifiIdx = dt.Columns.IndexOf("GIFI Code");
            for (int r = 0; r < dt.Rows.Count; r++)
            {
                bool amber = highlightUnmatched && gifiIdx >= 0
                             && dt.Rows[r][gifiIdx] == DBNull.Value;
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    var cell = ws.Cell(r + 2, c + 1);
                    var v    = dt.Rows[r][c];
                    cell.Value = (v == DBNull.Value) ? "" : v.ToString() ?? "";
                    if (amber)
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3CD");
                }
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
            if (dt.Rows.Count > 0)
                ws.RangeUsed()?.SetAutoFilter();
        }

        // ── Live search ───────────────────────────────────────────────────────
        private void ApplySearch()
        {
            var term = txtSearch.Text.Trim();
            DataView? view = null;
            if      (tabControl.SelectedTab == tabGifi)      view = _gifiView;
            else if (tabControl.SelectedTab == tabTB)        view = _tbView;
            else if (tabControl.SelectedTab == tabMerged)    view = _mergedView;
            else if (tabControl.SelectedTab == tabUnmatched) view = _unmatchedView;
            if (view == null) return;

            if (string.IsNullOrEmpty(term)) { view.RowFilter = ""; return; }

            var clauses = view.Table!.Columns
                .Cast<DataColumn>()
                .Where(c => c.DataType == typeof(string))
                .Select(c => $"CONVERT([{c.ColumnName}], System.String) LIKE '%{EscFilter(term)}%'");
            view.RowFilter = string.Join(" OR ", clauses);
        }

        private static string EscFilter(string s) =>
            s.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]");

        // ── Clear all ─────────────────────────────────────────────────────────
        private void ClearAll()
        {
            if (MessageBox.Show("Clear all loaded data?", "Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            _gifiTable.Clear(); _tbTable.Clear();
            _mergedTable.Clear(); _unmatchedTable.Clear();
            dgvGifi.DataSource = dgvTB.DataSource =
            dgvMerged.DataSource = dgvUnmatched.DataSource = null;

            txtGifiPath.Clear(); txtTBPath.Clear(); txtSearch.Clear();
            _gifiLoaded = _tbLoaded = _merged = false;
            lblGifiCount.Text = "GIFI: —"; lblTBCount.Text = "TB: —";
            lblMatchedCount.Text = "Matched: —"; lblUnmatchedCount.Text = "Unmatched: —";
            tabUnmatched.Text = "⚠  Unmatched";
            SetSaveButtonsEnabled(false);
            UpdateMergeButton();
            SetStatus("Cleared.", false);
        }

        // ── Grid styling ──────────────────────────────────────────────────────
        private void StyleGifiGrid()
        {
            StyleGrid(dgvGifi, "Accounting Code");
            if (dgvGifi.Columns.Contains("GIFI Code"))
                dgvGifi.Columns["GIFI Code"]!.DefaultCellStyle.ForeColor = Color.DarkBlue;
        }

        private static void StyleGrid(DataGridView dgv, string? boldCol = null)
        {
            if (boldCol != null && dgv.Columns.Contains(boldCol))
                dgv.Columns[boldCol]!.DefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        }

        private static void StyleMergedGrid(DataGridView dgv)
        {
            StyleGrid(dgv, "Accounting Code");
            if (dgv.Columns.Contains("GIFI Code"))
            {
                dgv.Columns["GIFI Code"]!.DefaultCellStyle.ForeColor = Color.DarkGreen;
                dgv.Columns["GIFI Code"]!.DefaultCellStyle.Font      = new Font("Segoe UI", 9F, FontStyle.Bold);
            }
        }

        private static void ColorUnmatchedRows(DataGridView dgv)
        {
            if (!dgv.Columns.Contains("GIFI Code")) return;
            dgv.CellFormatting += (sender, e) =>
            {
                if (e.RowIndex < 0) return;
                var val = ((DataGridView)sender!).Rows[e.RowIndex].Cells["GIFI Code"].Value;
                if (val == null || val == DBNull.Value || string.IsNullOrWhiteSpace(val.ToString()))
                    e.CellStyle.BackColor = Color.FromArgb(255, 243, 205);
            };
        }

        // ── CSV helpers ───────────────────────────────────────────────────────
        private static DataTable ReadCsvRaw(string path, bool hasHeader, string[]? forceHeaders = null)
        {
            var dt   = new DataTable();
            int cols = 0;

            using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            bool firstLine = true;

            while (reader.ReadLine() is string line)
            {
                var fields = SplitCsvLine(line);
                if (cols == 0)
                {
                    cols = fields.Count;
                    if (forceHeaders != null)
                    {
                        for (int i = 0; i < cols; i++)
                            dt.Columns.Add(i < forceHeaders.Length ? forceHeaders[i] : $"Column {i + 1}");
                    }
                    else if (hasHeader)
                    {
                        foreach (var f in fields)
                            dt.Columns.Add(string.IsNullOrWhiteSpace(f) ? $"Col{dt.Columns.Count + 1}" : f.Trim());
                        firstLine = false;
                        continue;
                    }
                    else
                    {
                        for (int i = 0; i < cols; i++)
                            dt.Columns.Add($"col{i}");
                    }
                }

                if (firstLine && hasHeader && forceHeaders == null) { firstLine = false; continue; }
                firstLine = false;

                if (fields.All(f => string.IsNullOrWhiteSpace(f))) continue;

                var row = dt.NewRow();
                for (int i = 0; i < Math.Min(fields.Count, cols); i++)
                    row[i] = fields[i].Trim();
                dt.Rows.Add(row);
            }
            return dt;
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            var sb     = new StringBuilder();
            bool inQ   = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQ && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else inQ = !inQ;
                }
                else if (c == ',' && !inQ) { result.Add(sb.ToString()); sb.Clear(); }
                else sb.Append(c);
            }
            result.Add(sb.ToString());
            return result;
        }

        private static void WriteCsv(DataTable dt, string path)
        {
            using var w = new StreamWriter(path, false, Encoding.UTF8);
            w.WriteLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => CsvQ(c.ColumnName))));
            foreach (DataRow row in dt.Rows)
                w.WriteLine(string.Join(",", row.ItemArray.Select(v => CsvQ(v?.ToString() ?? ""))));
        }

        private static string CsvQ(string s) =>
            (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
            ? $"\"{s.Replace("\"", "\"\"")}\"" : s;

        // ── Misc helpers ──────────────────────────────────────────────────────
        private bool ValidatePath(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
                { Warn($"Please select a {label} file."); return false; }
            if (!File.Exists(path))
                { MessageBox.Show($"File not found:\n{path}", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
            return true;
        }

        private void SetStatus(string msg, bool working)
        {
            lblStatus.Text      = msg;
            progressBar.Visible = working;
            progressBar.Style   = ProgressBarStyle.Marquee;
            Application.DoEvents();
        }

        private void SetSaveButtonsEnabled(bool on)
        {
            btnSaveMerged.Enabled = btnSaveUnmatched.Enabled =
            btnExportExcel.Enabled = btnExportFutureTax.Enabled = on;
            menuFileSaveMerged.Enabled = menuFileSaveUnmatched.Enabled =
            menuFileExportExcel.Enabled = menuFileExportFutureTax.Enabled = on;
        }

        private void UpdateMergeButton() => btnMerge.Enabled = _gifiLoaded && _tbLoaded;

        private static void Warn(string msg) =>
            MessageBox.Show(msg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private static void ShowError(string title, Exception ex) =>
            MessageBox.Show($"{ex.Message}\n\nDetails:\n{ex}", title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        private static void OpenFile(string path)
        {
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true }); }
            catch { /* ignore */ }
        }

        private DataGridView[] AllGrids() => new[] { dgvGifi, dgvTB, dgvMerged, dgvUnmatched };

        // ── Ctrl+C from grid ──────────────────────────────────────────────────
        private void Dgv_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C && sender is DataGridView dgv)
            {
                var sb = new StringBuilder();
                foreach (DataGridViewRow row in dgv.SelectedRows)
                    sb.AppendLine(string.Join("\t", row.Cells.Cast<DataGridViewCell>().Select(c => c.Value?.ToString() ?? "")));
                if (sb.Length > 0) Clipboard.SetText(sb.ToString());
                e.Handled = true;
            }
        }

        // ── Double-click unmatched row ────────────────────────────────────────
        private void DgvUnmatched_CellDoubleClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var code = dgvUnmatched.Rows[e.RowIndex].Cells["Accounting Code"].Value?.ToString() ?? "";
            MessageBox.Show(
                $"Accounting Code \"{code}\" has no matching GIFI code.\n\nAdd it to your GIFI mapping file and reload.",
                "Unmatched Code", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── About ─────────────────────────────────────────────────────────────
        private static void ShowAbout() =>
            MessageBox.Show(
                "GIFI Mapper  v1.1\n\n" +
                "Loads a GIFI mapping and a trial-balance CSV,\n" +
                "merges on Accounting Code, and exports as:\n" +
                "  • Merged CSV  • Unmatched CSV\n" +
                "  • Excel (.xlsx) with 3 sheets\n" +
                "  • FutureTax import (Code + Amount, summed)\n\n" +
                "Shortcuts:\n" +
                "  Ctrl+G  Load GIFI      Ctrl+T  Load Trial Balance\n" +
                "  Ctrl+S  Save Merged    Ctrl+E  Export Excel\n" +
                "  Ctrl+F  FutureTax      Ctrl+C  Copy selected rows",
                "About GIFI Mapper", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
