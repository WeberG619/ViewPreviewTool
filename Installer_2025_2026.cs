using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Linq;

public class ViewPreviewToolInstaller : Form
{
    private PictureBox logoPicBox;
    private Label titleLabel;
    private Label descLabel;
    private CheckBox chk2025;
    private CheckBox chk2026;
    private Button installButton;
    private Button cancelButton;
    private ProgressBar progressBar;
    private Label statusLabel;

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new ViewPreviewToolInstaller());
    }

    public ViewPreviewToolInstaller()
    {
        InitializeUI();
    }

    private void InitializeUI()
    {
        // Form settings
        this.Text = "ViewPreviewTool v1.0 Setup - Revit 2025/2026 - BIM Ops Studio";
        this.Size = new Size(700, 650);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.BackColor = Color.White;

        try
        {
            // Try to set the form icon
            string iconPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "BIMOpsStudio.ico");
            if (File.Exists(iconPath))
            {
                this.Icon = new Icon(iconPath);
            }
        }
        catch { }

        // Header panel with gradient
        var headerPanel = new Panel();
        headerPanel.Dock = DockStyle.Top;
        headerPanel.Height = 100;
        headerPanel.Paint += (s, e) =>
        {
            using (var brush = new LinearGradientBrush(
                headerPanel.ClientRectangle,
                Color.FromArgb(0, 122, 204),
                Color.FromArgb(0, 88, 156),
                LinearGradientMode.Horizontal))
            {
                e.Graphics.FillRectangle(brush, headerPanel.ClientRectangle);
            }
        };

        // Logo
        logoPicBox = new PictureBox();
        logoPicBox.Location = new Point(20, 15);
        logoPicBox.Size = new Size(70, 70);
        logoPicBox.SizeMode = PictureBoxSizeMode.Zoom;
        logoPicBox.Image = CreateLogo();
        logoPicBox.BackColor = Color.Transparent;
        headerPanel.Controls.Add(logoPicBox);

        // Title
        titleLabel = new Label();
        titleLabel.Text = "ViewPreviewTool v1.0 for Revit 2025/2026";
        titleLabel.Font = new Font("Segoe UI", 18, FontStyle.Bold);
        titleLabel.ForeColor = Color.White;
        titleLabel.BackColor = Color.Transparent;
        titleLabel.AutoSize = true;
        titleLabel.Location = new Point(105, 25);
        headerPanel.Controls.Add(titleLabel);

        // Subtitle
        var subtitleLabel = new Label();
        subtitleLabel.Text = "BIM Ops Studio - Professional View Preview Solution";
        subtitleLabel.Font = new Font("Segoe UI", 11);
        subtitleLabel.ForeColor = Color.FromArgb(220, 220, 220);
        subtitleLabel.BackColor = Color.Transparent;
        subtitleLabel.AutoSize = true;
        subtitleLabel.Location = new Point(105, 55);
        headerPanel.Controls.Add(subtitleLabel);

        // Content panel
        var contentPanel = new Panel();
        contentPanel.Location = new Point(0, 100);
        contentPanel.Size = new Size(700, 500);

        // Description
        descLabel = new Label();
        descLabel.Text = "This installer will add the View Preview Tool to your Revit 2025 and 2026 installations.\n\n" +
                        "Features:\n" +
                        "• Real-time view previews when selecting views\n" +
                        "• Zoom and pan functionality\n" +
                        "• Family information display\n" +
                        "• Single-line titles with ellipsis for long names\n\n" +
                        "Select Revit versions to install:";
        descLabel.Location = new Point(50, 20);
        descLabel.Size = new Size(600, 150);
        descLabel.Font = new Font("Segoe UI", 10);

        // Version selection
        var versionGroup = new GroupBox();
        versionGroup.Text = "Revit Versions";
        versionGroup.Location = new Point(50, 180);
        versionGroup.Size = new Size(600, 100);
        versionGroup.Font = new Font("Segoe UI", 10);

        chk2025 = new CheckBox();
        chk2025.Text = "Revit 2025";
        chk2025.Location = new Point(30, 30);
        chk2025.Size = new Size(250, 25);
        chk2025.Checked = true;
        chk2025.Font = new Font("Segoe UI", 10);

        chk2026 = new CheckBox();
        chk2026.Text = "Revit 2026";
        chk2026.Location = new Point(300, 30);
        chk2026.Size = new Size(250, 25);
        chk2026.Checked = true;
        chk2026.Font = new Font("Segoe UI", 10);

        versionGroup.Controls.Add(chk2025);
        versionGroup.Controls.Add(chk2026);

        // Status label
        statusLabel = new Label();
        statusLabel.Text = "Ready to install";
        statusLabel.Location = new Point(50, 370);
        statusLabel.Size = new Size(600, 25);
        statusLabel.Font = new Font("Segoe UI", 9);
        statusLabel.ForeColor = Color.FromArgb(60, 60, 60);

        // Progress bar
        progressBar = new ProgressBar();
        progressBar.Location = new Point(50, 400);
        progressBar.Size = new Size(600, 23);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Visible = false;

        // Buttons
        installButton = new Button();
        installButton.Text = "Install";
        installButton.Location = new Point(450, 420);
        installButton.Size = new Size(90, 30);
        installButton.Font = new Font("Segoe UI", 10);
        installButton.BackColor = Color.FromArgb(0, 122, 204);
        installButton.ForeColor = Color.White;
        installButton.FlatStyle = FlatStyle.Flat;
        installButton.FlatAppearance.BorderSize = 0;
        installButton.Click += InstallButton_Click;

        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(560, 420);
        cancelButton.Size = new Size(90, 30);
        cancelButton.Font = new Font("Segoe UI", 10);
        cancelButton.Click += (s, e) => this.Close();

        contentPanel.Controls.AddRange(new Control[] { 
            descLabel, versionGroup, statusLabel, progressBar, installButton, cancelButton 
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
                    var logo = new Bitmap(70, 70);
                    using (var g = Graphics.FromImage(logo))
                    {
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.Clear(Color.Transparent);
                        g.DrawImage(originalImage, 0, 0, 70, 70);
                    }
                    return logo;
                }
            }
        }
        catch { }
        
        // Fallback logo
        var fallbackLogo = new Bitmap(70, 70);
        using (var g = Graphics.FromImage(fallbackLogo))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using (var brush = new LinearGradientBrush(
                new Rectangle(5, 5, 60, 60),
                Color.FromArgb(0, 150, 255),
                Color.FromArgb(0, 100, 200),
                45F))
            {
                g.FillEllipse(brush, 5, 5, 60, 60);
            }
        }
        return fallbackLogo;
    }

    private void InstallButton_Click(object sender, EventArgs e)
    {
        if (!chk2025.Checked && !chk2026.Checked)
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
        if (chk2025.Checked) selectedVersions.Add("2025");
        if (chk2026.Checked) selectedVersions.Add("2026");

        var worker = new BackgroundWorker();
        worker.WorkerReportsProgress = true;
        
        worker.DoWork += (s, args) =>
        {
            try
            {
                int progressStep = 100 / (selectedVersions.Count * 2);
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
                            break;
                        }
                    }
                    
                    if (!foundDll)
                    {
                        throw new Exception("DLL resource not found in installer!");
                    }

                    currentProgress += progressStep;
                    worker.ReportProgress(currentProgress, "Configuring for Revit " + version + "...");
                    
                    // Extract addin file
                    bool foundAddin = false;
                    foreach (string resourceName in resources)
                    {
                        if (resourceName.EndsWith(".addin"))
                        {
                            string addinPath = Path.Combine(targetDir, "ViewPreviewTool.addin");
                            ExtractResource(assembly, resourceName, addinPath);
                            foundAddin = true;
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
                MessageBox.Show(
                    "Installation completed successfully!\n\nPlease restart Revit to use the View Preview Tool.",
                    "Success",
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