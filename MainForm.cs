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
        private DataTable _summaryTable   = new();

        private DataView _gifiView      = new();
        private DataView _tbView        = new();
        private DataView _mergedView    = new();
        private DataView _unmatchedView = new();
        private DataView _summaryView   = new();

        private bool _gifiLoaded = false;
        private bool _tbLoaded   = false;
        private bool _merged     = false;

        private AppSettings _settings = new();
        private FutureTaxSortConfig _sortConfig = new();

        // ── Constructor ───────────────────────────────────────────────────────
        public MainForm()
        {
            InitializeComponent();
            WireEvents();
            _settings = AppSettings.Load();
            _sortConfig = FutureTaxSortConfig.Load();
            TryAutoLoadGifi();
        }

        // ── Event wiring ──────────────────────────────────────────────────────
        private void WireEvents()
        {
            btnBrowseGifi.Click      += (_, _) => BrowseAndLoad(isGifi: true);
            btnLoadGifi.Click        += (_, _) => LoadGifi(txtGifiPath.Text);
            btnEditGifi.Click        += (_, _) => EditGifiFile();
            btnReloadGifi.Click      += (_, _) => ReloadGifiFile();
            btnBrowseTB.Click        += (_, _) => BrowseAndLoad(isGifi: false);
            btnLoadTB.Click          += (_, _) => LoadTrialBalance(txtTBPath.Text);
            btnEditTB.Click          += (_, _) => EditTBFile();
            btnReloadTB.Click        += (_, _) => ReloadTBFile();

            btnMerge.Click           += (_, _) => DoMerge();
            btnSaveMerged.Click      += (_, _) => SaveCsv(_mergedTable,    "Save Merged File");
            btnSaveUnmatched.Click   += (_, _) => SaveCsv(_unmatchedTable, "Save Unmatched File");
            btnExportSummary.Click   += (_, _) => SaveCsv(_summaryTable,   "Save GIFI Summary");
            btnExportExcel.Click     += (_, _) => ExportExcel();
            btnExportFutureTax.Click += (_, _) => ExportFutureTax();
            btnViewFutureTax.Click   += (_, _) => ViewFutureTax();
            btnClear.Click           += (_, _) => ClearAll();

            menuFileLoadGifi.Click        += (_, _) => BrowseAndLoad(isGifi: true);
            menuFileEditGifi.Click        += (_, _) => EditGifiFile();
            menuFileReloadGifi.Click      += (_, _) => ReloadGifiFile();
            menuFileLoadTB.Click          += (_, _) => BrowseAndLoad(isGifi: false);
            menuFileEditTB.Click          += (_, _) => EditTBFile();
            menuFileReloadTB.Click        += (_, _) => ReloadTBFile();
            menuFileSaveMerged.Click      += (_, _) => SaveCsv(_mergedTable,    "Save Merged File");
            menuFileSaveUnmatched.Click   += (_, _) => SaveCsv(_unmatchedTable, "Save Unmatched File");
            menuFileExportSummary.Click   += (_, _) => SaveCsv(_summaryTable,   "Save GIFI Summary");
            menuFileExportExcel.Click     += (_, _) => ExportExcel();
            menuFileExportFutureTax.Click += (_, _) => ExportFutureTax();
            menuFileViewFutureTax.Click   += (_, _) => ViewFutureTax();
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
            dgvMerged.CellEndEdit              += Dgv_GifiCodeCellEndEdit;
            dgvUnmatched.CellEndEdit           += Dgv_GifiCodeCellEndEdit;

            foreach (var dgv in AllGrids())
                dgv.KeyDown += Dgv_KeyDown;
        }

        // ── Auto-load GIFI from exe folder ───────────────────────────────────
        private void TryAutoLoadGifi()
        {
            // Prefer the last file the user explicitly loaded, if it still exists.
            if (!string.IsNullOrEmpty(_settings.GifiMappingPath) && File.Exists(_settings.GifiMappingPath))
            {
                txtGifiPath.Text = _settings.GifiMappingPath;
                LoadGifi(_settings.GifiMappingPath);
                return;
            }

            var candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IG_gifi_codes.csv");
            // Fall back to the Excel version if no CSV is present next to the exe.
            if (!File.Exists(candidate))
                candidate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "IG_gifi_codes.xlsx");
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
                Filter = "All Supported Files (*.csv;*.xlsx;*.xlsm)|*.csv;*.xlsx;*.xlsm|CSV Files (*.csv)|*.csv|Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            if (isGifi) { txtGifiPath.Text = dlg.FileName; LoadGifi(dlg.FileName); }
            else        { txtTBPath.Text   = dlg.FileName; LoadTrialBalance(dlg.FileName); }
        }

        // ── Edit / reload GIFI mapping ───────────────────────────────────────
        private void EditGifiFile()
        {
            if (string.IsNullOrEmpty(txtGifiPath.Text) || !File.Exists(txtGifiPath.Text))
            {
                Warn("No GIFI mapping file loaded yet.");
                return;
            }
            OpenFile(txtGifiPath.Text);
        }

        private void ReloadGifiFile()
        {
            if (string.IsNullOrEmpty(txtGifiPath.Text))
            {
                Warn("No GIFI mapping file loaded yet.");
                return;
            }
            LoadGifi(txtGifiPath.Text);
        }

        // ── Edit / reload trial balance ──────────────────────────────────────
        private void EditTBFile()
        {
            if (string.IsNullOrEmpty(txtTBPath.Text) || !File.Exists(txtTBPath.Text))
            {
                Warn("No trial balance file loaded yet.");
                return;
            }
            OpenFile(txtTBPath.Text);
        }

        private void ReloadTBFile()
        {
            if (string.IsNullOrEmpty(txtTBPath.Text))
            {
                Warn("No trial balance file loaded yet.");
                return;
            }
            LoadTrialBalance(txtTBPath.Text);
        }

        // ── Load GIFI mapping ─────────────────────────────────────────────────
        private void LoadGifi(string path)
        {
            if (!ValidatePath(path, "GIFI mapping")) return;
            if (IsLegacyExcelFile(path))
            {
                MessageBox.Show(
                    ".xls files are not supported — please open the file in Excel and save it as .xlsx, then try again.",
                    "Unsupported File Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                SetStatus("Loading GIFI mapping…", true);
                var forceHeaders = new[] { "Accounting Code", "GIFI Code", "Description" };
                _gifiTable = IsExcelFile(path)
                    ? ReadExcelRaw(path, hasHeader: true, forceHeaders: forceHeaders)
                    : ReadCsvRaw(path, hasHeader: true, forceHeaders: forceHeaders);
                _gifiView = _gifiTable.DefaultView;
                dgvGifi.DataSource = _gifiView;
                StyleGifiGrid();
                _gifiLoaded = true;
                lblGifiCount.Text = $"GIFI: {_gifiTable.Rows.Count} rows";
                tabControl.SelectedTab = tabGifi;
                SetStatus($"GIFI mapping loaded — {_gifiTable.Rows.Count} codes.", false);

                _settings.GifiMappingPath = path;
                AppSettings.Save(_settings);
            }
            catch (Exception ex) { ShowError("Failed to load GIFI file", ex); }
        }

        // ── Load trial balance ────────────────────────────────────────────────
        private void LoadTrialBalance(string path)
        {
            if (!ValidatePath(path, "trial balance")) return;
            if (IsLegacyExcelFile(path))
            {
                MessageBox.Show(
                    ".xls files are not supported — please open the file in Excel and save it as .xlsx, then try again.",
                    "Unsupported File Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                SetStatus("Loading trial balance…", true);
                _tbTable = IsExcelFile(path)
                    ? ReadExcelRaw(path, hasHeader: false)
                    : ReadCsvRaw(path, hasHeader: false);
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
                MakeGifiCodeEditable(dgvMerged);
                MakeGifiCodeEditable(dgvUnmatched);

                DebugLog($"[DoMerge] unmatched counter={unmatched}, _unmatchedTable.Rows.Count={_unmatchedTable.Rows.Count}, _unmatchedView.Count={_unmatchedView.Count}, _unmatchedView.RowFilter='{_unmatchedView.RowFilter}', dgvUnmatched.RowCount={dgvUnmatched.RowCount}, dgvUnmatched.ColumnCount={dgvUnmatched.ColumnCount}, dgvUnmatched.Visible={dgvUnmatched.Visible}, ReferenceEquals(dgvUnmatched.DataSource,_unmatchedView)={ReferenceEquals(dgvUnmatched.DataSource, _unmatchedView)}");

                _merged = true;
                lblMatchedCount.Text   = $"Matched: {matched}";
                lblUnmatchedCount.Text = $"Unmatched: {unmatched}";
                BuildGifiSummary();

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
                WriteSheet(wb, "Merged",       _mergedTable,    highlightUnmatched: true);
                WriteSheet(wb, "GIFI Summary", _summaryTable,   highlightUnmatched: false);
                WriteSheet(wb, "Unmatched",    _unmatchedTable, highlightUnmatched: false);
                WriteSheet(wb, "GIFI Map",     _gifiTable,      highlightUnmatched: false);
                wb.SaveAs(dlg.FileName);
                SetStatus($"Excel exported: {dlg.FileName}", false);
                if (MessageBox.Show("Excel file saved. Open it now?", "Exported", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    OpenFile(dlg.FileName);
            }
            catch (Exception ex) { ShowError("Excel export failed", ex); }
        }

        // ── GIFI summary aggregation ─────────────────────────────────────────
        // Shared by the GIFI Summary tab and the FutureTax export: collapses
        // mergedTable rows sharing the same GIFI Code into one totals row.
        private static List<(string Gifi, string Desc, decimal Debit, decimal Credit, decimal Net, decimal AbsAmt, int AccountCount)>
            ComputeGifiSummary(DataTable mergedTable, DataTable gifiTable)
        {
            const string debitCol  = "Column 3";
            const string creditCol = "Column 4";

            // First non-blank description per GIFI code, pulled from the mapping
            // table rather than mergedTable rows (different accounts can share a
            // GIFI code but carry different source descriptions).
            var descByGifi = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (gifiTable.Columns.Contains("GIFI Code") && gifiTable.Columns.Contains("Description"))
            {
                foreach (DataRow r in gifiTable.Rows)
                {
                    var code = r["GIFI Code"]?.ToString()?.Trim() ?? "";
                    var desc = r["Description"]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(desc)) continue;
                    if (!descByGifi.ContainsKey(code)) descByGifi[code] = desc;
                }
            }

            bool hasDebitCol  = mergedTable.Columns.Contains(debitCol);
            bool hasCreditCol = mergedTable.Columns.Contains(creditCol);

            var totals = new Dictionary<string, (decimal Debit, decimal Credit, int Count)>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in mergedTable.Rows)
            {
                var gifi = row["GIFI Code"]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(gifi)) continue;

                var debit  = hasDebitCol  ? ParseAmount(row[debitCol])  : 0m;
                var credit = hasCreditCol ? ParseAmount(row[creditCol]) : 0m;
                var cur    = totals.GetValueOrDefault(gifi);
                totals[gifi] = (cur.Debit + debit, cur.Credit + credit, cur.Count + 1);
            }

            return totals
                .OrderBy(kv => kv.Key)
                .Select(kv =>
                {
                    decimal net    = kv.Value.Debit - kv.Value.Credit;
                    decimal absAmt = Math.Abs(net);
                    return (Gifi: kv.Key, Desc: descByGifi.GetValueOrDefault(kv.Key, ""),
                            Debit: kv.Value.Debit, Credit: kv.Value.Credit,
                            Net: net, AbsAmt: absAmt, AccountCount: kv.Value.Count);
                })
                .ToList();
        }

        // ── GIFI Summary tab ──────────────────────────────────────────────────
        private void BuildGifiSummary()
        {
            var summary = ComputeGifiSummary(_mergedTable, _gifiTable);

            _summaryTable = new DataTable();
            _summaryTable.Columns.Add("GIFI Code",        typeof(string));
            _summaryTable.Columns.Add("GIFI Description", typeof(string));
            _summaryTable.Columns.Add("Debit Total",      typeof(string));
            _summaryTable.Columns.Add("Credit Total",     typeof(string));
            _summaryTable.Columns.Add("Net (D-C)",        typeof(string));
            _summaryTable.Columns.Add("Accounts",         typeof(int));

            foreach (var s in summary)
                _summaryTable.Rows.Add(s.Gifi, s.Desc,
                    s.Debit.ToString("0.##"), s.Credit.ToString("0.##"), s.Net.ToString("0.##"),
                    s.AccountCount);

            _summaryView = _summaryTable.DefaultView;
            dgvSummary.DataSource = _summaryView;
            StyleMergedGrid(dgvSummary);
            lblSummaryCount.Text = $"GIFI codes: {_summaryTable.Rows.Count}";
        }

        // ── Export for FutureTax ──────────────────────────────────────────────
        // Column 3 = Debit  (positive),  Column 4 = Credit (negative).
        // Net per GIFI = sum(Debits) - sum(Credits).
        // Amount written to file = Math.Abs(Net).
        private void ExportFutureTax()
        {
            if (!_merged) { Warn("Please run Merge first."); return; }

            var results = ComputeFutureTaxResults();
            if (results == null) return;

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
                WriteFutureTaxFile(dlg.FileName, results);
                SetStatus($"FutureTax file saved: {dlg.FileName} ({results.Count} GIFI codes)", false);
                ShowFutureTaxPreview(results, dlg.FileName);
            }
            catch (Exception ex) { ShowError("FutureTax export failed", ex); }
        }

        // ── View FutureTax (no save required) ────────────────────────────────
        private void ViewFutureTax()
        {
            if (!_merged) { Warn("Please run Merge first."); return; }

            var results = ComputeFutureTaxResults();
            if (results == null) return;

            ShowFutureTaxPreview(results, null);
        }

        // Validates the merged table has the expected columns, aggregates per-GIFI
        // totals, and returns the FutureTax result rows — or null if it already
        // showed the user a message explaining why there's nothing to compute.
        private List<(string Gifi, decimal Debit, decimal Credit, decimal Net, decimal AbsAmt)>? ComputeFutureTaxResults()
        {
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
                return null;
            }

            // Per-GIFI debit/credit totals, shared with the GIFI Summary tab.
            var summary = ComputeGifiSummary(_mergedTable, _gifiTable);

            if (summary.Count == 0)
            {
                MessageBox.Show("No numeric values found in Column 3 or Column 4.\nCheck the trial balance file.",
                    "Nothing to Export", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("FutureTax export cancelled.", false);
                return null;
            }

            // Build result: net = debit - credit,  exported amount = Abs(net)
            return summary
                .Select(s => (Gifi: s.Gifi, Debit: s.Debit, Credit: s.Credit, Net: s.Net, AbsAmt: s.AbsAmt))
                .ToList();
        }

        // Write: Code,AbsoluteAmount  (no header row — FutureTax format)
        private static void WriteFutureTaxFile(string path, List<(string Gifi, decimal Debit, decimal Credit, decimal Net, decimal AbsAmt)> results)
        {
            using var w = new StreamWriter(path, false, Encoding.UTF8);
            foreach (var r in results)
                w.WriteLine($"{r.Gifi},{(long)Math.Round(r.AbsAmt, MidpointRounding.AwayFromZero)}");
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
            string? savedPath)
        {
            var frm = new Form
            {
                Text = "FutureTax Export — Preview", Size = new Size(780, 620),
                MinimumSize = new Size(560, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MaximizeBox = true, MinimizeBox = true,
                Padding = Padding.Empty
            };
            var lbl = new Label
            {
                Text = savedPath != null
                    ? $"✅  {results.Count} GIFI codes written.\nFile: {Path.GetFileName(savedPath)}"
                    : $"{results.Count} GIFI codes computed.",
                Dock = DockStyle.Top, Height = 44, AutoSize = false,
                Padding = new Padding(10, 8, 10, 0)
            };

            var dgv = MakeGrid();
            dgv.Dock = DockStyle.Fill;
            dgv.Margin = Padding.Empty;

            // Best-effort scroll-to-top; FirstDisplayedScrollingRowIndex can throw
            // in some grid states (e.g. before layout has fully settled), so this
            // is always wrapped and never allowed to crash the preview.
            void ScrollToTop()
            {
                try
                {
                    if (dgv.Rows.Count > 0)
                        dgv.FirstDisplayedScrollingRowIndex = 0;
                }
                catch { /* ignore — purely a cosmetic scroll reset */ }
            }

            // GIFI codes are normally numeric — type the column as int (for correct
            // numeric sort) unless any result's code fails to parse, in which case
            // fall back to string typing for all rows rather than throwing.
            bool allNumericGifi = results.All(r => int.TryParse(r.Gifi, out _));

            // User-assigned Sort values come from the persisted config; codes with
            // no saved value get the next free slot so they land at the end,
            // in a stable, collision-free order. Fresh defaults are NOT saved here —
            // only an explicit edit (below) writes to the config file.
            int nextSortValue = (_sortConfig.Order.Count > 0 ? _sortConfig.Order.Values.Max() : 0) + 1;

            var dt = new DataTable();
            dt.Columns.Add("Sort",            typeof(int));
            dt.Columns.Add("GIFI Code",       allNumericGifi ? typeof(int) : typeof(string));
            dt.Columns.Add("Exported Amount", typeof(long));
            dt.Columns.Add("Debit Total",     typeof(decimal));
            dt.Columns.Add("Credit Total",    typeof(decimal));
            dt.Columns.Add("Net (D-C)",       typeof(decimal));
            foreach (var r in results)
            {
                object gifiValue = allNumericGifi ? int.Parse(r.Gifi) : r.Gifi;
                int sortValue = _sortConfig.Order.TryGetValue(r.Gifi, out int saved) ? saved : nextSortValue++;
                dt.Rows.Add(sortValue, gifiValue,
                    (long)Math.Round(r.AbsAmt, MidpointRounding.AwayFromZero),
                    r.Debit,
                    r.Credit,
                    r.Net);
            }

            var view = new DataView(dt) { Sort = "Sort ASC" };
            dgv.DataSource = view;
            ScrollToTop();

            dgv.CellEndEdit += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var sortCol = dgv.Columns["Sort"];
                if (sortCol == null || e.ColumnIndex != sortCol.Index) return;

                var row      = dgv.Rows[e.RowIndex];
                var gifiCode = row.Cells["GIFI Code"].Value?.ToString() ?? "";
                var cellVal  = row.Cells["Sort"].Value;

                if (cellVal == null || cellVal == DBNull.Value || !int.TryParse(cellVal.ToString(), out int newSortValue))
                {
                    // Non-numeric/invalid input — revert without touching the saved config.
                    row.Cells["Sort"].Value = _sortConfig.Order.TryGetValue(gifiCode, out int prev) ? prev : cellVal;
                    return;
                }
                if (string.IsNullOrEmpty(gifiCode)) return;

                _sortConfig.Order[gifiCode] = newSortValue;
                FutureTaxSortConfig.Save(_sortConfig);

                if (dgv.DataSource is DataView dv) dv.Sort = "Sort ASC";
                ScrollToTop();
            };

            foreach (DataGridViewColumn col in dgv.Columns)
                col.SortMode = DataGridViewColumnSortMode.Automatic;

            // Grid-level ReadOnly overrides per-column settings (see the earlier
            // Merged/Unmatched edit fix) — lower it first, then lock every column
            // except "Sort".
            dgv.ReadOnly = false;
            foreach (DataGridViewColumn col in dgv.Columns)
                col.ReadOnly = col.Name != "Sort";

            // Fixed, compact widths so Sort / GIFI Code / Exported Amount sit close
            // together on the left; only the last column absorbs extra window width.
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            if (dgv.Columns["Sort"] is DataGridViewColumn colSort)
            {
                colSort.Width = 50;
                colSort.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            if (dgv.Columns["GIFI Code"] is DataGridViewColumn colGifi)
            {
                colGifi.Width = 90;
                if (allNumericGifi)
                    colGifi.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
            if (dgv.Columns["Exported Amount"] is DataGridViewColumn colAmt)
            {
                colAmt.Width = 110;
                colAmt.DefaultCellStyle.Format    = "#,##0";
                colAmt.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            foreach (var colName in new[] { "Debit Total", "Credit Total" })
            {
                if (dgv.Columns[colName] is DataGridViewColumn colDec)
                {
                    colDec.Width = 110;
                    colDec.DefaultCellStyle.Format    = "#,##0.##";
                    colDec.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
            }
            if (dgv.Columns["Net (D-C)"] is DataGridViewColumn colNet)
            {
                colNet.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                colNet.DefaultCellStyle.Format    = "#,##0.##";
                colNet.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            var buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 48 };
            var buttonFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(10)
            };

            var btnSaveAs = new Button { Text = "Save As…", Width = 100, Height = 28, Margin = new Padding(6, 10, 0, 6) };
            var btnOpen   = new Button { Text = "Open File", Width = 110, Height = 28, Enabled = savedPath != null, Margin = new Padding(6, 10, 0, 6) };
            var btnClose  = new Button { Text = "Close",      Width = 80,  Height = 28, DialogResult = DialogResult.OK, Margin = new Padding(6, 10, 0, 6) };

            btnOpen.Click += (_, _) => { if (savedPath != null) OpenFile(savedPath); };

            btnSaveAs.Click += (_, _) =>
            {
                using var saveDlg = new SaveFileDialog
                {
                    Title = "Save FutureTax Import File",
                    Filter = "CSV Files (*.csv)|*.csv|GFI Files (*.gfi)|*.gfi",
                    DefaultExt = "csv", FileName = "futuretax_import.csv"
                };
                if (saveDlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    WriteFutureTaxFile(saveDlg.FileName, results);
                    savedPath = saveDlg.FileName;
                    lbl.Text = $"✅  {results.Count} GIFI codes written.\nFile: {Path.GetFileName(savedPath)}";
                    btnOpen.Enabled = true;
                    SetStatus($"FutureTax file saved: {saveDlg.FileName} ({results.Count} GIFI codes)", false);
                }
                catch (Exception ex) { ShowError("FutureTax export failed", ex); }
            };

            // Added right-to-left so the visual order reads Save As, Open File, Close.
            buttonFlow.Controls.AddRange(new Control[] { btnClose, btnOpen, btnSaveAs });
            buttonPanel.Controls.Add(buttonFlow);

            frm.Controls.AddRange(new Control[] { lbl, buttonPanel, dgv });
            frm.AcceptButton = btnClose;
            frm.Shown += (_, _) => ScrollToTop();
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
            DebugLog($"[ApplySearch] selectedTab={tabControl.SelectedTab?.Name}, txtSearch.Text='{txtSearch.Text}'");

            var term = txtSearch.Text.Trim();
            DataView? view = null;
            if      (tabControl.SelectedTab == tabGifi)      view = _gifiView;
            else if (tabControl.SelectedTab == tabTB)        view = _tbView;
            else if (tabControl.SelectedTab == tabMerged)    view = _mergedView;
            else if (tabControl.SelectedTab == tabSummary)   view = _summaryView;
            else if (tabControl.SelectedTab == tabUnmatched) view = _unmatchedView;
            if (view == null) return;

            if (string.IsNullOrEmpty(term))
            {
                DebugLog($"[ApplySearch] about to set RowFilter on view for {tabControl.SelectedTab?.Name} to: (empty-clear)");
                view.RowFilter = "";
                DebugLog($"[ApplySearch] after filter: view.Count={view.Count}, view.RowFilter='{view.RowFilter}', view.Table.Rows.Count={view.Table?.Rows.Count}");
                return;
            }

            var clauses = view.Table!.Columns
                .Cast<DataColumn>()
                .Where(c => c.DataType == typeof(string))
                .Select(c => $"CONVERT([{c.ColumnName}], System.String) LIKE '%{EscFilter(term)}%'");
            var filterExpr = string.Join(" OR ", clauses);
            DebugLog($"[ApplySearch] about to set RowFilter on view for {tabControl.SelectedTab?.Name} to: '{filterExpr}'");
            view.RowFilter = filterExpr;
            DebugLog($"[ApplySearch] after filter: view.Count={view.Count}, view.RowFilter='{view.RowFilter}', view.Table.Rows.Count={view.Table?.Rows.Count}");

            // A search term carried over from another tab can filter this tab's
            // grid down to zero rows even though it has data — flag that so it
            // doesn't look like a data/merge bug.
            if (view.Count == 0 && view.Table.Rows.Count > 0)
                SetStatus($"No rows match \"{term}\" in this tab — clear Search to see all rows.", false);
        }

        private static string EscFilter(string s) =>
            s.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]");

        // ── Clear all ─────────────────────────────────────────────────────────
        private void ClearAll()
        {
            if (MessageBox.Show("Clear all loaded data?", "Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            _gifiTable.Clear(); _tbTable.Clear();
            _mergedTable.Clear(); _unmatchedTable.Clear(); _summaryTable.Clear();
            dgvGifi.DataSource = dgvTB.DataSource =
            dgvMerged.DataSource = dgvUnmatched.DataSource = dgvSummary.DataSource = null;

            txtGifiPath.Clear(); txtTBPath.Clear(); txtSearch.Clear();
            _gifiLoaded = _tbLoaded = _merged = false;
            lblGifiCount.Text = "GIFI: —"; lblTBCount.Text = "TB: —";
            lblMatchedCount.Text = "Matched: —"; lblUnmatchedCount.Text = "Unmatched: —";
            lblSummaryCount.Text = "GIFI codes: —";
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

        // ── Inline GIFI Code editing (Merged / Unmatched grids) ──────────────
        private static void MakeGifiCodeEditable(DataGridView dgv)
        {
            // Grid-level ReadOnly overrides per-column settings — MakeGrid() creates
            // every grid with ReadOnly = true, so it must be lowered here before
            // locking columns individually, or the per-column settings below are ignored.
            dgv.ReadOnly = false;
            foreach (DataGridViewColumn col in dgv.Columns)
                col.ReadOnly = col.Name != "GIFI Code";
        }

        private void Dgv_GifiCodeCellEndEdit(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (sender is not DataGridView dgv) return;

            var gifiCol = dgv.Columns["GIFI Code"];
            if (gifiCol == null || e.ColumnIndex != gifiCol.Index) return;

            var row            = dgv.Rows[e.RowIndex];
            var accountingCode = row.Cells["Accounting Code"].Value?.ToString() ?? "";
            var newGifiCode    = row.Cells["GIFI Code"].Value?.ToString() ?? "";
            UpdateGifiMapping(accountingCode, newGifiCode);
        }

        private void UpdateGifiMapping(string accountingCode, string newGifiCode)
        {
            accountingCode = accountingCode.Trim();
            newGifiCode    = newGifiCode.Trim();
            if (string.IsNullOrEmpty(accountingCode)) return;
            if (string.IsNullOrEmpty(newGifiCode)) return;

            DataRow? match = null;
            foreach (DataRow r in _gifiTable.Rows)
            {
                var key = r["Accounting Code"]?.ToString()?.Trim() ?? "";
                if (string.Equals(key, accountingCode, StringComparison.OrdinalIgnoreCase))
                {
                    match = r;
                    break;
                }
            }

            if (match != null)
            {
                match["GIFI Code"] = newGifiCode;
            }
            else
            {
                var newRow = _gifiTable.NewRow();
                newRow["Accounting Code"] = accountingCode;
                newRow["GIFI Code"]       = newGifiCode;
                _gifiTable.Rows.Add(newRow);
            }

            lblGifiCount.Text = $"GIFI: {_gifiTable.Rows.Count} rows";

            try
            {
                WriteGifiMappingFile(txtGifiPath.Text);
                SetStatusQuiet($"GIFI mapping updated: {accountingCode} → {newGifiCode}. Saved to mapping file.");
            }
            catch (Exception ex) { ShowError("Failed to save GIFI mapping file", ex); }
        }

        private void WriteGifiMappingFile(string path)
        {
            if (IsExcelFile(path))
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("GIFI Map");
                for (int c = 0; c < _gifiTable.Columns.Count; c++)
                {
                    var cell = ws.Cell(1, c + 1);
                    cell.Value = _gifiTable.Columns[c].ColumnName;
                    cell.Style.Font.Bold = true;
                }
                for (int r = 0; r < _gifiTable.Rows.Count; r++)
                    for (int c = 0; c < _gifiTable.Columns.Count; c++)
                    {
                        var v = _gifiTable.Rows[r][c];
                        ws.Cell(r + 2, c + 1).Value = (v == DBNull.Value) ? "" : v.ToString() ?? "";
                    }
                wb.SaveAs(path);
            }
            else
            {
                WriteCsv(_gifiTable, path);
            }
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

        // ── Excel import helpers ──────────────────────────────────────────────
        private static bool IsExcelFile(string path)
        {
            var ext = Path.GetExtension(path);
            return string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase)
                || string.Equals(ext, ".xlsm", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsLegacyExcelFile(string path) =>
            string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase);

        private static DataTable ReadExcelRaw(string path, bool hasHeader, string[]? forceHeaders = null)
        {
            var dt = new DataTable();

            using var wb = new XLWorkbook(path);
            var worksheet = wb.Worksheet(1);
            var used = worksheet.RangeUsed();
            if (used == null) return dt;

            int rowCount = used.RowCount();
            int colCount = used.ColumnCount();
            int cols = 0;
            bool firstRow = true;

            for (int r = 1; r <= rowCount; r++)
            {
                var fields = new List<string>(colCount);
                for (int c = 1; c <= colCount; c++)
                    fields.Add(used.Cell(r, c).GetString().Trim());

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
                            dt.Columns.Add(string.IsNullOrWhiteSpace(f) ? $"Col{dt.Columns.Count + 1}" : f);
                        firstRow = false;
                        continue;
                    }
                    else
                    {
                        for (int i = 0; i < cols; i++)
                            dt.Columns.Add($"col{i}");
                    }
                }

                if (firstRow && hasHeader && forceHeaders == null) { firstRow = false; continue; }
                firstRow = false;

                if (fields.All(f => string.IsNullOrWhiteSpace(f))) continue;

                var newRow = dt.NewRow();
                for (int i = 0; i < Math.Min(fields.Count, cols); i++)
                    newRow[i] = fields[i];
                dt.Rows.Add(newRow);
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

        // Status update safe to call from inside a DataGridView cell-edit callback
        // (e.g. CellEndEdit) — must NOT call Application.DoEvents(), which can pump
        // a queued message back into the grid mid-commit and trigger a reentrant
        // SetCurrentCellAddressCore crash.
        private void SetStatusQuiet(string msg)
        {
            lblStatus.Text      = msg;
            progressBar.Visible = false;
        }

        private void SetSaveButtonsEnabled(bool on)
        {
            btnSaveMerged.Enabled = btnSaveUnmatched.Enabled = btnExportSummary.Enabled =
            btnExportExcel.Enabled = btnExportFutureTax.Enabled = btnViewFutureTax.Enabled = on;
            menuFileSaveMerged.Enabled = menuFileSaveUnmatched.Enabled = menuFileExportSummary.Enabled =
            menuFileExportExcel.Enabled = menuFileExportFutureTax.Enabled = menuFileViewFutureTax.Enabled = on;
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

        // TEMPORARY diagnostic logging for the Unmatched-tab empty-grid investigation.
        // Remove once the root cause is identified.
        private static void DebugLog(string msg)
        {
            try
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gifimapper_debug.log");
                File.AppendAllText(path, $"{DateTime.Now:HH:mm:ss.fff}  {msg}{Environment.NewLine}");
            }
            catch { /* ignore logging failures */ }
        }

        private DataGridView[] AllGrids() => new[] { dgvGifi, dgvTB, dgvMerged, dgvUnmatched, dgvSummary };

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
                "  • GIFI Summary tab/CSV — totals grouped by GIFI code\n" +
                "  • Excel (.xlsx) with 4 sheets\n" +
                "  • FutureTax import (Code + Amount, summed)\n\n" +
                "The last-used GIFI mapping file is remembered across launches,\n" +
                "and can be edited or reloaded from the File menu.\n" +
                "The trial balance file can also be edited or reloaded from the File menu.\n" +
                "GIFI Code cells in the Merged/Unmatched grids can be edited directly —\n" +
                "changes update and save the GIFI mapping file automatically.\n" +
                "View FutureTax previews the codes/amounts on demand without saving first,\n" +
                "and the preview window can save the file directly via Save As.\n\n" +
                "Shortcuts:\n" +
                "  Ctrl+G  Load GIFI      Ctrl+T  Load Trial Balance\n" +
                "  Ctrl+S  Save Merged    Ctrl+E  Export Excel\n" +
                "  Ctrl+F  FutureTax      Ctrl+C  Copy selected rows",
                "About GIFI Mapper", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
