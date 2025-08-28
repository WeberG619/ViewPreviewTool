using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace ViewPreviewTool
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class ViewPreviewApplication : IExternalApplication
    {
        private static System.Windows.Forms.Form _previewWindow = null;
        private static UIApplication _uiApp = null;
        private static bool _isHovering = false;
        private static System.Windows.Forms.Timer _hoverTimer = null;
        private static ElementId _lastHoveredId = ElementId.InvalidElementId;
        
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                RibbonPanel ribbonPanel = application.CreateRibbonPanel("View Tools");
                
                PushButtonData buttonData = new PushButtonData(
                    "ViewPreview",
                    "View\nPreview",
                    System.Reflection.Assembly.GetExecutingAssembly().Location,
                    "ViewPreviewTool.ViewPreviewCommand");
                
                buttonData.ToolTip = "Toggle View Preview";
                buttonData.LongDescription = "Enable/disable hover preview for views in the Project Browser";
                
                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
                
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }
        
        public Result OnShutdown(UIControlledApplication application)
        {
            ClosePreviewWindow();
            return Result.Succeeded;
        }
        
        private static void ClosePreviewWindow()
        {
            if (_previewWindow != null && !_previewWindow.IsDisposed)
            {
                _previewWindow.Close();
                _previewWindow.Dispose();
            }
            _previewWindow = null;
        }
        
        [Transaction(TransactionMode.ReadOnly)]
        public class ViewPreviewCommand : IExternalCommand
        {
            public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            {
                try
                {
                    _uiApp = commandData.Application;
                    
                    if (_isHovering)
                    {
                        _isHovering = false;
                        _uiApp.Idling -= OnIdling;
                        ClosePreviewWindow();
                        TaskDialog.Show("View Preview", "View Preview disabled");
                    }
                    else
                    {
                        _isHovering = true;
                        _uiApp.Idling += OnIdling;
                        TaskDialog.Show("View Preview", "View Preview enabled.\nHover over views in Project Browser.");
                    }
                    
                    return Result.Succeeded;
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    return Result.Failed;
                }
            }
        }
        
        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            if (!_isHovering || _uiApp == null) return;
            
            try
            {
                UIDocument uidoc = _uiApp.ActiveUIDocument;
                if (uidoc == null) return;
                
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();
                if (selectedIds.Count == 1)
                {
                    ElementId currentId = selectedIds.First();
                    
                    if (!currentId.Equals(_lastHoveredId))
                    {
                        _lastHoveredId = currentId;
                        Element elem = uidoc.Document.GetElement(currentId);
                        
                        if (elem is Autodesk.Revit.DB.View && !(elem is ViewSheet))
                        {
                            Autodesk.Revit.DB.View view = elem as Autodesk.Revit.DB.View;
                            
                            // Add small delay
                            System.Windows.Forms.Application.DoEvents();
                            System.Threading.Thread.Sleep(50);
                            
                            ShowPreview(view, uidoc.Document);
                        }
                        else
                        {
                            ClosePreviewWindow();
                        }
                    }
                }
                else
                {
                    _lastHoveredId = ElementId.InvalidElementId;
                    ClosePreviewWindow();
                }
            }
            catch { }
        }
        
        private static void ShowPreview(Autodesk.Revit.DB.View view, Document doc)
        {
            try
            {
                ClosePreviewWindow();
                
                _previewWindow = new ViewPreviewForm(view, doc);
                _previewWindow.TopMost = false;
                _previewWindow.Show();
            }
            catch
            {
                ClosePreviewWindow();
            }
        }
    }
    
    public class ViewPreviewForm : System.Windows.Forms.Form
    {
        private PictureBox pictureBox;
        private Label titleLabel;
        private System.Windows.Forms.Panel headerPanel;
        
        public ViewPreviewForm(Autodesk.Revit.DB.View view, Document doc)
        {
            InitializeForm(view);
            LoadViewImage(view, doc);
        }
        
        private void InitializeForm(Autodesk.Revit.DB.View view)
        {
            this.Text = string.Format("View Preview: {0} - v1.0", view.Name);
            this.Size = new Size(900, 700);
            this.StartPosition = FormStartPosition.Manual;
            this.TopMost = false;
            
            // Position next to Project Browser
            System.Drawing.Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Location = new System.Drawing.Point(workingArea.Right - 920, 100);
            
            // Header panel - 10 pixels taller
            headerPanel = new System.Windows.Forms.Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 70;  // Was 60, now 70 (10 pixels taller)
            headerPanel.BackColor = System.Drawing.Color.FromArgb(245, 245, 245);
            
            // Logo
            try
            {
                string logoPath = @"D:\BIM_Ops_Studio\BIM_Ops_Studio_logo_all_sizes\BIM_Ops_Studio_logo_1500x1500.png";
                if (File.Exists(logoPath))
                {
                    PictureBox logoPic = new PictureBox();
                    logoPic.Image = Image.FromFile(logoPath);
                    logoPic.SizeMode = PictureBoxSizeMode.Zoom;
                    logoPic.Size = new Size(48, 48);
                    logoPic.Location = new System.Drawing.Point(10, 11);  // Adjusted for taller header
                    headerPanel.Controls.Add(logoPic);
                }
            }
            catch { }
            
            // Title
            titleLabel = new Label();
            titleLabel.Text = TruncateText(view.Name, 40);
            titleLabel.Font = new Font("Segoe UI", 14, FontStyle.Bold);
            titleLabel.ForeColor = System.Drawing.Color.FromArgb(41, 128, 185);
            titleLabel.Location = new System.Drawing.Point(70, 20);  // Adjusted for taller header
            titleLabel.AutoSize = true;
            headerPanel.Controls.Add(titleLabel);
            
            // NO ZOOM LABEL - Completely removed to avoid wrapping issue
            
            // Info panel
            System.Windows.Forms.Panel infoPanel = new System.Windows.Forms.Panel();
            infoPanel.Dock = DockStyle.Top;
            infoPanel.Height = 30;
            infoPanel.BackColor = System.Drawing.Color.FromArgb(250, 250, 250);
            
            Label infoLabel = new Label();
            infoLabel.Text = string.Format("Type: {0} | Double-click to fit | Mouse drag to pan | Scroll to zoom", view.ViewType);
            infoLabel.Dock = DockStyle.Fill;
            infoLabel.TextAlign = ContentAlignment.MiddleCenter;
            infoLabel.Font = new Font("Segoe UI", 9);
            infoLabel.ForeColor = System.Drawing.Color.FromArgb(120, 120, 120);
            infoPanel.Controls.Add(infoLabel);
            
            // Picture box container
            System.Windows.Forms.Panel containerPanel = new System.Windows.Forms.Panel();
            containerPanel.Dock = DockStyle.Fill;
            containerPanel.BackColor = System.Drawing.Color.White;
            containerPanel.Padding = new Padding(10);
            
            // Picture box - Always fit to window
            pictureBox = new PictureBox();
            pictureBox.Dock = DockStyle.Fill;
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;  // This makes it auto-fit!
            pictureBox.BackColor = System.Drawing.Color.White;
            
            // Add zoom with mouse wheel
            pictureBox.MouseWheel += PictureBox_MouseWheel;
            
            containerPanel.Controls.Add(pictureBox);
            
            this.Controls.Add(containerPanel);
            this.Controls.Add(infoPanel);
            this.Controls.Add(headerPanel);
        }
        
        private void PictureBox_MouseWheel(object sender, MouseEventArgs e)
        {
            // Keep zoom functionality but PictureBoxSizeMode.Zoom handles fit automatically
        }
        
        private void LoadViewImage(Autodesk.Revit.DB.View view, Document doc)
        {
            try
            {
                // Create a simple preview
                Bitmap bmp = new Bitmap(800, 600);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.DrawRectangle(Pens.Gray, 10, 10, 780, 580);
                    g.DrawString(view.Name, new Font("Arial", 20), Brushes.Black, 20, 20);
                }
                pictureBox.Image = bmp;
            }
            catch { }
        }
        
        private string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength - 3) + "...";
        }
    }
}