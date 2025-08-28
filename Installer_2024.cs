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

public class ViewPreviewTool2024Installer : Form
{
    private PictureBox logoPicBox;
    private Label titleLabel;
    private Label descLabel;
    private CheckBox chk2024;
    private Button installButton;
    private Button cancelButton;
    private ProgressBar progressBar;
    private Label statusLabel;

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new ViewPreviewTool2024Installer());
    }

    public ViewPreviewTool2024Installer()
    {
        InitializeUI();
    }

    private void InitializeUI()
    {
        // Form settings
        this.Text = "ViewPreviewTool v1.0 Setup - Revit 2024 - BIM Ops Studio";
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
                Color.FromArgb(0, 122, 204),  // Same blue theme
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
        titleLabel.Text = "ViewPreviewTool v2.5 for Revit 2024";
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
        contentPanel.Size = new Size(700, 550);

        // Description
        descLabel = new Label();
        descLabel.Text = "This installer will add the View Preview Tool to your Revit 2024 installation.\n\n" +
                        "IMPORTANT: This is a special build for Revit 2024 using .NET Framework 4.8.\n\n" +
                        "Features:\n" +
                        "• Real-time view previews when selecting views\n" +
                        "• Auto-fit to window when view loads\n" +
                        "• Zoom with mouse wheel / Pan by dragging\n" +
                        "• Double-click to re-fit view to window\n" +
                        "• BIM Ops Studio branding\n" +
                        "• Fixed: Zoom label stays on single line\n" +
                        "• Fixed: Window no longer stays on top";
        descLabel.Location = new Point(50, 20);
        descLabel.Size = new Size(600, 240);
        descLabel.Font = new Font("Segoe UI", 10);

        // Version selection (just 2024)
        chk2024 = new CheckBox();
        chk2024.Text = "Install for Revit 2024";
        chk2024.Location = new Point(50, 230);
        chk2024.Size = new Size(600, 30);
        chk2024.Checked = true;
        chk2024.Enabled = false; // Always checked
        chk2024.Font = new Font("Segoe UI", 11);

        // Info message
        var infoLabel = new Label();
        infoLabel.Text = "ℹ This version is specifically built for Revit 2024 (.NET Framework 4.8)";
        infoLabel.Location = new Point(50, 270);
        infoLabel.Size = new Size(600, 25);
        infoLabel.Font = new Font("Segoe UI", 9, FontStyle.Italic);
        infoLabel.ForeColor = Color.FromArgb(0, 122, 204);

        // Status label
        statusLabel = new Label();
        statusLabel.Text = "Ready to install";
        statusLabel.Location = new Point(50, 380);
        statusLabel.Size = new Size(600, 25);
        statusLabel.Font = new Font("Segoe UI", 9);
        statusLabel.ForeColor = Color.FromArgb(60, 60, 60);

        // Progress bar
        progressBar = new ProgressBar();
        progressBar.Location = new Point(50, 410);
        progressBar.Size = new Size(600, 23);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.Visible = false;

        // Buttons
        installButton = new Button();
        installButton.Text = "Install";
        installButton.Location = new Point(450, 460);
        installButton.Size = new Size(90, 30);
        installButton.Font = new Font("Segoe UI", 10);
        installButton.BackColor = Color.FromArgb(0, 122, 204);
        installButton.ForeColor = Color.White;
        installButton.FlatStyle = FlatStyle.Flat;
        installButton.FlatAppearance.BorderSize = 0;
        installButton.Click += InstallButton_Click;

        cancelButton = new Button();
        cancelButton.Text = "Cancel";
        cancelButton.Location = new Point(560, 460);
        cancelButton.Size = new Size(90, 30);
        cancelButton.Font = new Font("Segoe UI", 10);
        cancelButton.Click += (s, e) => this.Close();

        contentPanel.Controls.AddRange(new Control[] { 
            descLabel, chk2024, infoLabel, statusLabel, progressBar, installButton, cancelButton 
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
                Color.FromArgb(0, 122, 204),
                Color.FromArgb(0, 88, 156),
                45F))
            {
                g.FillEllipse(brush, 5, 5, 60, 60);
            }
        }
        return fallbackLogo;
    }

    private void InstallButton_Click(object sender, EventArgs e)
    {
        installButton.Enabled = false;
        cancelButton.Enabled = false;
        progressBar.Visible = true;
        progressBar.Value = 0;

        var worker = new BackgroundWorker();
        worker.WorkerReportsProgress = true;
        
        worker.DoWork += (s, args) =>
        {
            try
            {
                worker.ReportProgress(10, "Starting installation for Revit 2024...");
                Thread.Sleep(500);

                string targetDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    @"Autodesk\Revit\Addins\2024");

                worker.ReportProgress(20, "Creating directory: " + targetDir);

                if (!Directory.Exists(targetDir))
                    Directory.CreateDirectory(targetDir);

                Thread.Sleep(300);

                // Extract embedded resources
                var assembly = Assembly.GetExecutingAssembly();
                var resources = assembly.GetManifestResourceNames();
                
                worker.ReportProgress(40, "Installing ViewPreviewTool.dll...");
                
                // Extract DLL
                bool foundDll = false;
                foreach (string resourceName in resources)
                {
                    if (resourceName.EndsWith(".dll"))
                    {
                        string dllPath = Path.Combine(targetDir, "ViewPreviewTool.dll");
                        ExtractResource(assembly, resourceName, dllPath);
                        foundDll = true;
                        worker.ReportProgress(60, "Installed DLL successfully");
                        Thread.Sleep(300);
                        break;
                    }
                }
                
                if (!foundDll)
                {
                    throw new Exception("DLL resource not found in installer!");
                }

                worker.ReportProgress(70, "Installing ViewPreviewTool.addin...");
                
                // Extract addin file
                bool foundAddin = false;
                foreach (string resourceName in resources)
                {
                    if (resourceName.EndsWith(".addin"))
                    {
                        string addinPath = Path.Combine(targetDir, "ViewPreviewTool.addin");
                        ExtractResource(assembly, resourceName, addinPath);
                        foundAddin = true;
                        worker.ReportProgress(90, "Installed ADDIN successfully");
                        Thread.Sleep(300);
                        break;
                    }
                }
                
                if (!foundAddin)
                {
                    throw new Exception("ADDIN resource not found in installer!");
                }

                Thread.Sleep(500);
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
                    "Installation failed:\n" + ex.Message + 
                    "\n\nPlease ensure:\n" +
                    "• Revit 2024 is closed\n" +
                    "• You have write permissions to the Addins folder\n" +
                    "• No antivirus is blocking the installation",
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
                    "Installation completed successfully!\n\n" +
                    "Please restart Revit 2024 to use the View Preview Tool.\n\n" +
                    "Note: If the tool doesn't appear, check the Revit journal file for any loading errors.",
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