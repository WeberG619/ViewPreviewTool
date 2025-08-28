using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;
using System.Collections.Generic;

public class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MultiVersionInstallerForm());
    }
}

public class MultiVersionInstallerForm : Form
{
    private PictureBox logoBox;
    private Label titleLabel;
    private Label descLabel;
    private CheckBox chk2024;
    private CheckBox chk2025;
    private CheckBox chk2026;
    private ProgressBar progressBar;
    private Button installButton;
    private Button cancelButton;
    private Label statusLabel;
    private Panel headerPanel;
    private Panel contentPanel;

    public MultiVersionInstallerForm()
    {
        InitializeUI();
    }

    private void InitializeUI()
    {
        // Form settings
        this.Text = "ViewPreviewTool v1.0 Setup - BIM Ops Studio";
        this.Size = new Size(700, 650);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.Icon = SystemIcons.Application;
        this.BackColor = Color.White;

        // Header panel
        headerPanel = new Panel();
        headerPanel.Dock = DockStyle.Top;
        headerPanel.Height = 140;
        headerPanel.BackColor = Color.White;
        headerPanel.Paint += (s, e) => {
            using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1))
            {
                e.Graphics.DrawLine(pen, 0, headerPanel.Height - 1, headerPanel.Width, headerPanel.Height - 1);
            }
        };

        // Logo
        logoBox = new PictureBox();
        logoBox.Size = new Size(100, 100);
        logoBox.Location = new Point(30, 20);
        logoBox.SizeMode = PictureBoxSizeMode.Zoom;
        logoBox.BackColor = Color.White;
        
        // Try to load high-res logo for UI
        try
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream("BIMOpsStudio_Logo_HD.png");
            if (stream != null)
            {
                logoBox.Image = Image.FromStream(stream);
            }
            else
            {
                logoBox.Image = CreateLogo();
            }
        }
        catch
        {
            logoBox.Image = CreateLogo();
        }

        // Title
        titleLabel = new Label();
        titleLabel.Text = "ViewPreviewTool for Revit";
        titleLabel.Font = new Font("Segoe UI", 24, FontStyle.Regular);
        titleLabel.ForeColor = Color.FromArgb(30, 50, 90);
        titleLabel.Location = new Point(150, 30);
        titleLabel.Size = new Size(500, 40);

        // Subtitle
        var subtitleLabel = new Label();
        subtitleLabel.Text = "BIM Ops Studio - Professional Revit Add-in";
        subtitleLabel.Font = new Font("Segoe UI", 12);
        subtitleLabel.ForeColor = Color.FromArgb(100, 100, 100);
        subtitleLabel.Location = new Point(150, 75);
        subtitleLabel.Size = new Size(500, 25);

        headerPanel.Controls.AddRange(new Control[] { logoBox, titleLabel, subtitleLabel });

        // Content panel
        contentPanel = new Panel();
        contentPanel.Location = new Point(0, 140);
        contentPanel.Size = new Size(700, 510);
        contentPanel.BackColor = Color.FromArgb(248, 248, 248);

        // Description
        descLabel = new Label();
        descLabel.Text = "This installer will add the View Preview Tool to your selected Revit installations.\n\n" +
                        "Features (v2.5):\n" +
                        "• Real-time view previews when selecting views\n" +
                        "• Auto-fit to window when view loads (2025/2026)\n" +
                        "• Zoom with mouse wheel / Pan by dragging\n" +
                        "• Double-click to fit view to window\n" +
                        "• BIM Ops Studio branding\n" +
                        "• Fixed: Window no longer stays on top\n" +
                        "• Fixed: Title text stays on single line\n\n" +
                        "Select which Revit versions to install:";
        descLabel.Font = new Font("Segoe UI", 11);
        descLabel.Location = new Point(50, 20);
        descLabel.Size = new Size(600, 200);
        descLabel.ForeColor = Color.FromArgb(60, 60, 60);

        // Version selection group
        var versionGroup = new GroupBox();
        versionGroup.Text = "Select Revit Versions";
        versionGroup.Font = new Font("Segoe UI", 10);
        versionGroup.Location = new Point(50, 230);
        versionGroup.Size = new Size(600, 180);
        versionGroup.ForeColor = Color.FromArgb(60, 60, 60);

        // Checkboxes for versions
        chk2024 = new CheckBox();
        chk2024.Text = "Revit 2024";
        chk2024.Font = new Font("Segoe UI", 11);
        chk2024.Location = new Point(30, 40);
        chk2024.Size = new Size(500, 30);
        chk2024.Checked = true;

        chk2025 = new CheckBox();
        chk2025.Text = "Revit 2025";
        chk2025.Font = new Font("Segoe UI", 11);
        chk2025.Location = new Point(30, 80);
        chk2025.Size = new Size(500, 30);
        chk2025.Checked = true;

        chk2026 = new CheckBox();
        chk2026.Text = "Revit 2026";
        chk2026.Font = new Font("Segoe UI", 11);
        chk2026.Location = new Point(30, 120);
        chk2026.Size = new Size(500, 30);
        chk2026.Checked = true;

        versionGroup.Controls.AddRange(new Control[] { chk2024, chk2025, chk2026 });

        // Features text
        var featuresLabel = new Label();
        featuresLabel.Text = "Features:\n" +
                            "• Real-time view previews in the Project Browser\n" +
                            "• Interactive zoom and pan controls\n" +
                            "• Family type information display\n" +
                            "• Professional schedule notifications\n" +
                            "• Optimized performance";
        featuresLabel.Font = new Font("Segoe UI", 10);
        featuresLabel.Location = new Point(50, 285);
        featuresLabel.Size = new Size(600, 120);
        featuresLabel.ForeColor = Color.FromArgb(60, 60, 60);

        // Status label  
        statusLabel = new Label();
        statusLabel.Location = new Point(50, 415);
        statusLabel.Size = new Size(350, 25);
        statusLabel.Font = new Font("Segoe UI", 9);
        statusLabel.ForeColor = Color.FromArgb(30, 50, 90);
        statusLabel.Text = "Select versions and click Install";

        // Cancel button
        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Font = new Font("Segoe UI", 10);
        cancelButton.Location = new Point(420, 410);
        cancelButton.Size = new Size(100, 32);
        cancelButton.BackColor = Color.White;
        cancelButton.ForeColor = Color.FromArgb(60, 60, 60);
        cancelButton.FlatStyle = FlatStyle.Flat;
        cancelButton.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
        cancelButton.Cursor = Cursors.Hand;
        cancelButton.Click += (s, e) => this.Close();

        // Install button
        installButton = new Button();
        installButton.Text = "Install";
        installButton.Font = new Font("Segoe UI", 10);
        installButton.Location = new Point(530, 410);
        installButton.Size = new Size(100, 32);
        installButton.BackColor = Color.FromArgb(30, 50, 90);
        installButton.ForeColor = Color.White;
        installButton.FlatStyle = FlatStyle.Flat;
        installButton.FlatAppearance.BorderSize = 0;
        installButton.Cursor = Cursors.Hand;
        installButton.Click += InstallButton_Click;

        // Progress bar
        progressBar = new ProgressBar();
        progressBar.Location = new Point(50, 455);
        progressBar.Size = new Size(600, 23);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Visible = false;

        contentPanel.Controls.AddRange(new Control[] { 
            descLabel, versionGroup, featuresLabel, progressBar, statusLabel, installButton, cancelButton 
        });

        this.Controls.Add(contentPanel);
        this.Controls.Add(headerPanel);
    }

    private Bitmap CreateLogo()
    {
        try
        {
            // Try to load the BIM Ops Studio logo
            string logoPath = @"D:\BIM_Ops_Studio\BIM_Ops_Studio_logo_all_sizes\BIM_Ops_Studio_logo_1500x1500.png";
            if (File.Exists(logoPath))
            {
                using (var originalImage = Image.FromFile(logoPath))
                {
                    // Resize to 80x80
                    var logo = new Bitmap(80, 80);
                    using (var g = Graphics.FromImage(logo))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalImage, 0, 0, 80, 80);
                    }
                    return logo;
                }
            }
        }
        catch { }
        
        // Fallback to simple logo if file not found
        var fallbackLogo = new Bitmap(80, 80);
        using (var g = Graphics.FromImage(fallbackLogo))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);

            using (var brush = new LinearGradientBrush(
                new Rectangle(5, 5, 70, 70),
                Color.FromArgb(0, 150, 255),
                Color.FromArgb(0, 100, 200),
                45F))
            {
                g.FillEllipse(brush, 5, 5, 70, 70);
            }

            using (var font = new Font("Segoe UI", 20, FontStyle.Bold))
            {
                var sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                g.DrawString("BIM", font, Brushes.White, new RectangleF(0, 0, 80, 80), sf);
            }
        }
        return fallbackLogo;
    }

    private void InstallButton_Click(object sender, EventArgs e)
    {
        // Check if at least one version is selected
        if (!chk2024.Checked && !chk2025.Checked && !chk2026.Checked)
        {
            MessageBox.Show("Please select at least one Revit version to install.", 
                "No Version Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        installButton.Enabled = false;
        cancelButton.Enabled = false;
        progressBar.Visible = true;
        progressBar.Value = 0;

        var selectedVersions = new List<string>();
        if (chk2024.Checked) selectedVersions.Add("2024");
        if (chk2025.Checked) selectedVersions.Add("2025");
        if (chk2026.Checked) selectedVersions.Add("2026");

        // Store for access in completed event
        var versions = selectedVersions;

        var worker = new System.ComponentModel.BackgroundWorker();
        worker.WorkerReportsProgress = true;
        
        worker.DoWork += (s, args) =>
        {
            try
            {
                int progressStep = 100 / (selectedVersions.Count * 2 + 1);
                int currentProgress = 0;

                worker.ReportProgress(currentProgress, "Starting installation...");
                Thread.Sleep(300);

                foreach (string version in selectedVersions)
                {
                    currentProgress += progressStep;
                    worker.ReportProgress(currentProgress, "Installing for Revit " + version + "...");
                    
                    string targetDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        @"Autodesk\Revit\Addins\" + version);

                    // Show the exact path being used
                    worker.ReportProgress(currentProgress, "Target: " + targetDir);
                    Thread.Sleep(500); // Let user see the path

                    if (!Directory.Exists(targetDir))
                        Directory.CreateDirectory(targetDir);

                    // Extract embedded resources
                    var assembly = Assembly.GetExecutingAssembly();
                    var resources = assembly.GetManifestResourceNames();
                    
                    // Extract DLL
                    bool foundDll = false;
                    foreach (string resourceName in resources)
                    {
                        if (resourceName.EndsWith(".dll"))
                        {
                            string dllPath = Path.Combine(targetDir, "ViewPreviewTool.dll");
                            ExtractResource(assembly, resourceName, dllPath);
                            foundDll = true;
                            worker.ReportProgress(currentProgress, "Extracted DLL to: " + dllPath);
                            Thread.Sleep(300);
                            break;
                        }
                    }
                    if (!foundDll)
                    {
                        throw new Exception("DLL resource not found in installer!");
                    }

                    currentProgress += progressStep;
                    worker.ReportProgress(currentProgress, "Configuring for Revit " + version + "...");
                    
                    // Extract and modify addin file for this version
                    bool foundAddin = false;
                    foreach (string resourceName in resources)
                    {
                        if (resourceName.EndsWith(".addin"))
                        {
                            string addinPath = Path.Combine(targetDir, "ViewPreviewTool.addin");
                            ExtractResource(assembly, resourceName, addinPath);
                            foundAddin = true;
                            worker.ReportProgress(currentProgress, "Extracted ADDIN to: " + addinPath);
                            Thread.Sleep(300);
                            break;
                        }
                    }
                    if (!foundAddin)
                    {
                        throw new Exception("ADDIN resource not found in installer!");
                    }

                    Thread.Sleep(300);
                }

                worker.ReportProgress(100, "Installation completed successfully!");
            }
            catch (Exception ex)
            {
                args.Result = ex;
            }
        };

        worker.ProgressChanged += (s, args) =>
        {
            progressBar.Value = args.ProgressPercentage;
            statusLabel.Text = args.UserState.ToString();
        };

        worker.RunWorkerCompleted += (s, args) =>
        {
            if (args.Result is Exception)
            {
                Exception ex = (Exception)args.Result;
                statusLabel.Text = "Installation failed!";
                statusLabel.ForeColor = Color.Red;
                MessageBox.Show(
                    "Installation failed:\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                installButton.Enabled = true;
                cancelButton.Enabled = true;
            }
            else
            {
                statusLabel.ForeColor = Color.Green;
                
                string versionsText = string.Join(", ", versions);
                MessageBox.Show(
                    "ViewPreviewTool v1.0 has been installed successfully!\n\n" +
                    "Installed for: Revit " + versionsText + "\n\n" +
                    "Please restart Revit to activate the add-in.\n" +
                    "Look for 'Toggle View Preview' in the Add-Ins tab.\n\n" +
                    "Developed by BIM Ops Studio",
                    "Installation Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                this.Close();
            }
        };

        worker.RunWorkerAsync();
    }

    private void ExtractResource(Assembly assembly, string resourceName, string fileName)
    {
        using (Stream stream = assembly.GetManifestResourceStream(resourceName))
        {
            if (stream == null) return;
            
            using (FileStream fileStream = File.Create(fileName))
            {
                stream.CopyTo(fileStream);
            }
        }
    }
}