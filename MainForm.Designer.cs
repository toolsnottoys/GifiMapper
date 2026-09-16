namespace GifiMapper
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // Controls
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem menuFile;
        private System.Windows.Forms.ToolStripMenuItem menuFileLoadGifi;
        private System.Windows.Forms.ToolStripMenuItem menuFileLoadTB;
        private System.Windows.Forms.ToolStripSeparator menuFileSep1;
        private System.Windows.Forms.ToolStripMenuItem menuFileSaveMerged;
        private System.Windows.Forms.ToolStripMenuItem menuFileSaveUnmatched;
        private System.Windows.Forms.ToolStripMenuItem menuFileExportExcel;
        private System.Windows.Forms.ToolStripMenuItem menuFileExportFutureTax;
        private System.Windows.Forms.ToolStripSeparator menuFileSep2;
        private System.Windows.Forms.ToolStripMenuItem menuFileExit;
        private System.Windows.Forms.ToolStripMenuItem menuHelp;
        private System.Windows.Forms.ToolStripMenuItem menuHelpAbout;

        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.GroupBox grpGifi;
        private System.Windows.Forms.TextBox txtGifiPath;
        private System.Windows.Forms.Button btnBrowseGifi;
        private System.Windows.Forms.Button btnLoadGifi;

        private System.Windows.Forms.GroupBox grpTB;
        private System.Windows.Forms.TextBox txtTBPath;
        private System.Windows.Forms.Button btnBrowseTB;
        private System.Windows.Forms.Button btnLoadTB;

        private System.Windows.Forms.Panel panelActions;
        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.Button btnSaveMerged;
        private System.Windows.Forms.Button btnSaveUnmatched;
        private System.Windows.Forms.Button btnExportExcel;
        private System.Windows.Forms.Button btnExportFutureTax;
        private System.Windows.Forms.Button btnClear;

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabGifi;
        private System.Windows.Forms.TabPage tabTB;
        private System.Windows.Forms.TabPage tabMerged;
        private System.Windows.Forms.TabPage tabUnmatched;

        private System.Windows.Forms.DataGridView dgvGifi;
        private System.Windows.Forms.DataGridView dgvTB;
        private System.Windows.Forms.DataGridView dgvMerged;
        private System.Windows.Forms.DataGridView dgvUnmatched;

        private System.Windows.Forms.Panel panelSearch;
        private System.Windows.Forms.Label lblSearch;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnSearchClear;

        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;
        private System.Windows.Forms.ToolStripStatusLabel lblGifiCount;
        private System.Windows.Forms.ToolStripStatusLabel lblTBCount;
        private System.Windows.Forms.ToolStripStatusLabel lblMatchedCount;
        private System.Windows.Forms.ToolStripStatusLabel lblUnmatchedCount;
        private System.Windows.Forms.ToolStripProgressBar progressBar;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            this.SuspendLayout();

            // Form
            this.Text = "GIFI Mapper — Trial Balance Mapper";
            this.Size = new System.Drawing.Size(1150, 720);
            this.MinimumSize = new System.Drawing.Size(950, 600);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);

            // Menu strip
            menuStrip = new System.Windows.Forms.MenuStrip();
            menuFile = new System.Windows.Forms.ToolStripMenuItem("&File");

            menuFileLoadGifi      = new System.Windows.Forms.ToolStripMenuItem("Load &GIFI Mapping File…");
            menuFileLoadTB        = new System.Windows.Forms.ToolStripMenuItem("Load &Trial Balance File…");
            menuFileSep1          = new System.Windows.Forms.ToolStripSeparator();
            menuFileSaveMerged    = new System.Windows.Forms.ToolStripMenuItem("Save &Merged CSV…")    { Enabled = false };
            menuFileSaveUnmatched = new System.Windows.Forms.ToolStripMenuItem("Save &Unmatched CSV…") { Enabled = false };
            menuFileExportExcel   = new System.Windows.Forms.ToolStripMenuItem("&Export to Excel (.xlsx)…") { Enabled = false };
            menuFileExportFutureTax = new System.Windows.Forms.ToolStripMenuItem("Export for &FutureTax (.csv)…") { Enabled = false };
            menuFileSep2          = new System.Windows.Forms.ToolStripSeparator();
            menuFileExit          = new System.Windows.Forms.ToolStripMenuItem("E&xit");

            menuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
                menuFileLoadGifi, menuFileLoadTB, menuFileSep1,
                menuFileSaveMerged, menuFileSaveUnmatched,
                menuFileExportExcel, menuFileExportFutureTax,
                menuFileSep2, menuFileExit });

            menuHelp = new System.Windows.Forms.ToolStripMenuItem("&Help");
            menuHelpAbout = new System.Windows.Forms.ToolStripMenuItem("&About…");
            menuHelp.DropDownItems.Add(menuHelpAbout);

            menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { menuFile, menuHelp });
            menuStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.System;

            // Top panel
            panelTop = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 130,
                Padding = new System.Windows.Forms.Padding(6, 4, 6, 4)
            };

            grpGifi = new System.Windows.Forms.GroupBox { Text = "GIFI Mapping File",  Left = 8, Top = 8,  Width = 510, Height = 54 };
            txtGifiPath  = new System.Windows.Forms.TextBox  { Left = 8,   Top = 22, Width = 350, Height = 22, ReadOnly = true, BackColor = System.Drawing.SystemColors.Window, PlaceholderText = "No file selected…" };
            btnBrowseGifi = new System.Windows.Forms.Button  { Left = 364, Top = 21, Width = 65,  Height = 24, Text = "Browse…" };
            btnLoadGifi   = new System.Windows.Forms.Button  { Left = 434, Top = 21, Width = 65,  Height = 24, Text = "Load" };
            grpGifi.Controls.AddRange(new System.Windows.Forms.Control[] { txtGifiPath, btnBrowseGifi, btnLoadGifi });

            grpTB = new System.Windows.Forms.GroupBox { Text = "Trial Balance File", Left = 8, Top = 68, Width = 510, Height = 54 };
            txtTBPath    = new System.Windows.Forms.TextBox  { Left = 8,   Top = 22, Width = 350, Height = 22, ReadOnly = true, BackColor = System.Drawing.SystemColors.Window, PlaceholderText = "No file selected…" };
            btnBrowseTB  = new System.Windows.Forms.Button  { Left = 364, Top = 21, Width = 65,  Height = 24, Text = "Browse…" };
            btnLoadTB    = new System.Windows.Forms.Button  { Left = 434, Top = 21, Width = 65,  Height = 24, Text = "Load" };
            grpTB.Controls.AddRange(new System.Windows.Forms.Control[] { txtTBPath, btnBrowseTB, btnLoadTB });

            panelTop.Controls.AddRange(new System.Windows.Forms.Control[] { grpGifi, grpTB });

            // Action button bar
            panelActions = new System.Windows.Forms.Panel
            {
                Dock = System.Windows.Forms.DockStyle.Top,
                Height = 42,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
            };

            int bx = 8;
            btnMerge           = Btn("▶  Merge",           bx, 100, true);  bx += 108;
            btnSaveMerged      = Btn("💾  Save Merged",     bx, 130, false); bx += 138;
            btnSaveUnmatched   = Btn("📋  Save Unmatched",  bx, 140, false); bx += 148;
            btnExportExcel     = Btn("📊  Export Excel",    bx, 130, false); bx += 138;
            btnExportFutureTax = Btn("🍁  FutureTax",       bx, 110, false); bx += 118;
            btnClear           = Btn("✖  Clear All",        bx, 100, true);

            panelActions.Controls.AddRange(new System.Windows.Forms.Control[]
                { btnMerge, btnSaveMerged, btnSaveUnmatched, btnExportExcel, btnExportFutureTax, btnClear });

            // Search bar
            panelSearch = new System.Windows.Forms.Panel { Dock = System.Windows.Forms.DockStyle.Top, Height = 34 };
            lblSearch   = new System.Windows.Forms.Label  { Left = 8,  Top = 8, Width = 50,  Text = "Search:", TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            txtSearch   = new System.Windows.Forms.TextBox{ Left = 60, Top = 6, Width = 260, Height = 22, PlaceholderText = "Filter rows in active tab…" };
            btnSearchClear = new System.Windows.Forms.Button { Left = 326, Top = 5, Width = 60, Height = 24, Text = "Clear" };
            panelSearch.Controls.AddRange(new System.Windows.Forms.Control[] { lblSearch, txtSearch, btnSearchClear });

            // Tab control
            tabControl = new System.Windows.Forms.TabControl { Dock = System.Windows.Forms.DockStyle.Fill };

            tabGifi      = new System.Windows.Forms.TabPage("📂  GIFI Mapping");
            tabTB        = new System.Windows.Forms.TabPage("📄  Trial Balance");
            tabMerged    = new System.Windows.Forms.TabPage("🔗  Merged Results");
            tabUnmatched = new System.Windows.Forms.TabPage("⚠  Unmatched");

            dgvGifi      = MakeGrid(); tabGifi.Controls.Add(dgvGifi);
            dgvTB        = MakeGrid(); tabTB.Controls.Add(dgvTB);
            dgvMerged    = MakeGrid(); tabMerged.Controls.Add(dgvMerged);
            dgvUnmatched = MakeGrid(); tabUnmatched.Controls.Add(dgvUnmatched);

            tabControl.TabPages.AddRange(new System.Windows.Forms.TabPage[]
                { tabGifi, tabTB, tabMerged, tabUnmatched });

            // Status strip
            statusStrip       = new System.Windows.Forms.StatusStrip();
            lblStatus         = new System.Windows.Forms.ToolStripStatusLabel("Ready") { Spring = true, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            lblGifiCount      = SLabel("GIFI: —");
            lblTBCount        = SLabel("TB: —");
            lblMatchedCount   = SLabel("Matched: —");
            lblUnmatchedCount = SLabel("Unmatched: —");
            progressBar       = new System.Windows.Forms.ToolStripProgressBar { Visible = false, Width = 120 };
            statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[]
                { lblStatus, lblGifiCount, lblTBCount, lblMatchedCount, lblUnmatchedCount, progressBar });

            // Compose form
            this.Controls.Add(tabControl);
            this.Controls.Add(panelSearch);
            this.Controls.Add(panelActions);
            this.Controls.Add(panelTop);
            this.Controls.Add(menuStrip);
            this.Controls.Add(statusStrip);
            this.MainMenuStrip = menuStrip;

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private static System.Windows.Forms.Button Btn(string text, int x, int w, bool enabled) =>
            new System.Windows.Forms.Button
            {
                Text = text, Left = x, Top = 7, Width = w, Height = 28,
                Enabled = enabled,
                FlatStyle = System.Windows.Forms.FlatStyle.System
            };

        internal static System.Windows.Forms.DataGridView MakeGrid() =>
            new System.Windows.Forms.DataGridView
            {
                Dock = System.Windows.Forms.DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true,
                RowHeadersVisible = false,
                BackgroundColor = System.Drawing.SystemColors.Window,
                BorderStyle = System.Windows.Forms.BorderStyle.None,
                AlternatingRowsDefaultCellStyle = new System.Windows.Forms.DataGridViewCellStyle
                    { BackColor = System.Drawing.Color.FromArgb(245, 247, 250) },
                ColumnHeadersDefaultCellStyle = new System.Windows.Forms.DataGridViewCellStyle
                {
                    BackColor = System.Drawing.Color.FromArgb(51, 85, 139),
                    ForeColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold)
                },
                EnableHeadersVisualStyles = false,
                GridColor = System.Drawing.Color.FromArgb(220, 225, 235),
                CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal
            };

        private static System.Windows.Forms.ToolStripStatusLabel SLabel(string text) =>
            new System.Windows.Forms.ToolStripStatusLabel(text)
            {
                BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left,
                BorderStyle = System.Windows.Forms.Border3DStyle.Etched,
                Padding = new System.Windows.Forms.Padding(8, 0, 8, 0)
            };
    }
}
