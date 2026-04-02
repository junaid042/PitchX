using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PitchX
{
    [ComVisible(true)]
    public partial class TLBRRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public TLBRRibbon() { }

        public Bitmap getImage(Office.IRibbonControl control)
        {
            System.Diagnostics.Debug.WriteLine("getImage called for: " + control.Id);
            switch (control.Id)
            {
                case "TextBoxPlain":
                    System.Diagnostics.Debug.WriteLine("Loading TextBoxPlain");
                    return Properties.Resources.Text;
                case "TextBoxFill":
                    System.Diagnostics.Debug.WriteLine("Loading TextBoxFill");
                    return Properties.Resources.Text2;
                case "Widescreen":
                    System.Diagnostics.Debug.WriteLine("Loading Widescreen");
                    return Properties.Resources.Widescreen;
                case "A4":
                    System.Diagnostics.Debug.WriteLine("Loading A4");
                    return Properties.Resources.A4;
                case "UST":
                    System.Diagnostics.Debug.WriteLine("Loading UST");
                    return Properties.Resources.UST;
                case "Margin0":
                    System.Diagnostics.Debug.WriteLine("Loading Margin0");
                    return Properties.Resources.Margin0;
                case "Margin1":
                    System.Diagnostics.Debug.WriteLine("Loading Margin1");
                    return Properties.Resources.Margin1;
                case "Margin2":
                    System.Diagnostics.Debug.WriteLine("Loading Margin2");
                    return Properties.Resources.Margin2;
                case "LS1pt":
                    System.Diagnostics.Debug.WriteLine("Loading LS_1pt");
                    return Properties.Resources.LS_1pt; 
                case "LS6pt":
                    System.Diagnostics.Debug.WriteLine("Loading LS6_pt");
                    return Properties.Resources.LS6_pt;
                case "HighlightCells":
                    System.Diagnostics.Debug.WriteLine("Loading HighlightCells");
                    return Properties.Resources.HighCells;
                case "HorizontalDivider":
                    System.Diagnostics.Debug.WriteLine("Loading HorizontalDivider");
                    return Properties.Resources.HorDiv;
                case "VerticalDivider":
                    System.Diagnostics.Debug.WriteLine("Loading VerticalDivider");
                    return Properties.Resources.VerDiv;
                case "TotalRow":
                    System.Diagnostics.Debug.WriteLine("Loading TotalRow");
                    return Properties.Resources.Total2;
                case "Margin15":
                    System.Diagnostics.Debug.WriteLine("Loading Margin15");
                    return Properties.Resources.Margin15;
                case "guide1":
                    System.Diagnostics.Debug.WriteLine("Loading guide1");
                    return Properties.Resources.WidescreenO;
                case "guide2":
                    System.Diagnostics.Debug.WriteLine("Loading guide2");
                    return Properties.Resources.A4O;
                case "guide3":
                    System.Diagnostics.Debug.WriteLine("Loading guide3");
                    return Properties.Resources.USO;
                case "Bracket":
                    System.Diagnostics.Debug.WriteLine("Loading Bracket");
                    return Properties.Resources.Bracket;
                case "Chevron":
                    System.Diagnostics.Debug.WriteLine("Loading Chevron");
                    return Properties.Resources.Chevron;
                case "BigStat":
                    System.Diagnostics.Debug.WriteLine("Loading BigStat");
                    return Properties.Resources.Bigstat;
                case "LitStat":
                    System.Diagnostics.Debug.WriteLine("Loading LitStat");
                    return Properties.Resources.Litstat;
                case "Quote":
                    System.Diagnostics.Debug.WriteLine("Loading Quote");
                    return Properties.Resources.Quote;
                case "Ready":
                    System.Diagnostics.Debug.WriteLine("Loading Ready");
                    return Properties.Resources.Red;
                case "Hold":
                    System.Diagnostics.Debug.WriteLine("Loading Hold");
                    return Properties.Resources.Amber;
                case "Done":
                    System.Diagnostics.Debug.WriteLine("Loading Done");
                    return Properties.Resources.Green;
                case "BottomAdj":
                    System.Diagnostics.Debug.WriteLine("Loading BottomAdj");
                    return Properties.Resources.BottomAdj;
                case "NoFill":
                    System.Diagnostics.Debug.WriteLine("Loading NoFill");
                    return Properties.Resources.NoFill;
                case "NoOutline":
                    System.Diagnostics.Debug.WriteLine("Loading NoOutline");
                    return Properties.Resources.NoOutline;
                case "NoteRed":
                    System.Diagnostics.Debug.WriteLine("Loading NoteRed");
                    return Properties.Resources.NoteRed;
                case "NoteYellow":
                    System.Diagnostics.Debug.WriteLine("Loading NoteYellow");
                    return Properties.Resources.NoteYellow;
                case "RightAdj":
                    System.Diagnostics.Debug.WriteLine("Loading RightAdj");
                    return Properties.Resources.RightAdj;
                case "SelectFill":
                    System.Diagnostics.Debug.WriteLine("Loading SelectFill");
                    return Properties.Resources.SelectFill;
                case "SelectWidth":
                    System.Diagnostics.Debug.WriteLine("Loading SelectWidth");
                    return Properties.Resources.SelectWidth;
                case "SelectHeight":
                    System.Diagnostics.Debug.WriteLine("Loading SelectHeight");
                    return Properties.Resources.SelectHeight;
                case "SelectOutline":
                    System.Diagnostics.Debug.WriteLine("Loading SelectOutline");
                    return Properties.Resources.SelectOutline;
                case "UK":
                    System.Diagnostics.Debug.WriteLine("Loading UK");
                    return Properties.Resources.UK;
                case "US":
                    System.Diagnostics.Debug.WriteLine("Loading US");
                    return Properties.Resources.US;
                case "btnLoginToTLBR":
                    System.Diagnostics.Debug.WriteLine("Loading Margin2 for btnLoginToTLBR");
                    return Properties.Resources.Margin2;
                default:
                    System.Diagnostics.Debug.WriteLine("No image found for: " + control.Id);
                    return null;
            }
        }

        // Load the embedded PitchXRibbon.xml (make sure Build Action = Embedded Resource)
        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceName = asm.GetManifestResourceNames()
                                .FirstOrDefault(n => n.EndsWith("PitchXRibbon.xml", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrEmpty(resourceName))
            {
                // Fallback to a more specific search if the exact name isn't found
                resourceName = asm.GetManifestResourceNames()
                                .FirstOrDefault(n => n.IndexOf("PitchXRibbon", StringComparison.OrdinalIgnoreCase) >= 0 && n.EndsWith(".xml"));
            }

            if (string.IsNullOrEmpty(resourceName))
            {
                throw new InvalidOperationException("Could not find embedded PitchXRibbon.xml resource. Ensure the file is included with Build Action set to Embedded Resource.");
            }

            try
            {
                using (var stream = asm.GetManifestResourceStream(resourceName))
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error loading PitchXRibbon.xml: {ex.Message}", ex);
            }
        }

        public void Margin1Macro(Office.IRibbonControl control) { }
        public void Margin2Macro(Office.IRibbonControl control) { }
        public void Margin15Macro(Office.IRibbonControl control) { }
        public void LineSpace1Macro(Office.IRibbonControl control) { }
        public void LineSpace6Macro(Office.IRibbonControl control) { }
        public void MergeTextBoxesMacro(Office.IRibbonControl control) { }
        public void RightSpaceAdjustmentMacro(Office.IRibbonControl control) { }
        public void NoFillMacro(Office.IRibbonControl control) { }
        public void NoOutlineMacro(Office.IRibbonControl control) { }
        public void InsertNoteMacro(Office.IRibbonControl control) { }
        public void SetUKLanguageMacro(Office.IRibbonControl control) { }
        public void SetUSLanguageMacro(Office.IRibbonControl control) { }
        public void OnOpenLoginToPitchX(Office.IRibbonControl control) { }
        public void CustomMarginMacro(Office.IRibbonControl control) { }
        public void TableMarginMacro(Office.IRibbonControl control) { }
        public void SelectSameFillMacro(Office.IRibbonControl control) { }
        public void SelectSameHeightMacro(Office.IRibbonControl control) { }
        public void SelectSameOutlineMacro(Office.IRibbonControl control) { }

        // onLoad callback
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        // getImage callback — returns stdole.IPictureDisp for the button image
        public stdole.IPictureDisp GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                Image img = null;
                var asm = Assembly.GetExecutingAssembly();

                // Option A: try project resources class: PitchX.Properties.Resources.InsertIcon
                var resourcesType = asm.GetType("PitchX.Properties.Resources");
                if (resourcesType != null)
                {
                    var prop = resourcesType.GetProperty("InsertIcon", BindingFlags.Public | BindingFlags.Static);
                    if (prop != null)
                        img = prop.GetValue(null, null) as Image;
                }

                // Option B: fallback to an embedded file named InsertIcon.png (embed it as resource)
                if (img == null)
                {
                    var imgName = asm.GetManifestResourceNames()
                                     .FirstOrDefault(n => n.ToLower().EndsWith("inserticon.png") || n.ToLower().EndsWith("inserticon.jpg") || n.ToLower().EndsWith("inserticon.ico"));
                    if (imgName != null)
                    {
                        using (var s = asm.GetManifestResourceStream(imgName))
                        {
                            img = Image.FromStream(s);
                        }
                    }
                }

                if (img == null)
                    return null; // no image found — Office will show label only

                return PictureConverter.ImageToPictureDisp(img);
            }
            catch
            {
                return null;
            }
        }

        public void Invalidate()
        {
            ribbon?.Invalidate();
        }

        public void InvalidateControl(string id)
        {
            ribbon?.InvalidateControl(id);
        }

        public void WSMacro(Office.IRibbonControl control)
        {
            try
            {
                // Check if email/password are saved
                string savedEmail = Properties.Settings.Default.Email;
                string savedPassword = Properties.Settings.Default.Password;

                if (string.IsNullOrWhiteSpace(savedEmail) || string.IsNullOrWhiteSpace(savedPassword))
                {
                    // Stylish professional message before redirecting to login
                    DialogResult result = MessageBox.Show(
                        "You need to log in before accessing the widescreen PowerPoint template.\n\n" +
                        "Click OK to securely sign in now.",
                        "Authentication Required",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Information);

                    if (result == DialogResult.OK)
                    {
                        // Open login dialog
                        OnOpenLoginToTLBR(control);
                    }
                    else
                    {
                        return; 
                    }
                }
                else
                {
                    // User already logged in — continue to open template
                    using (TempleateOpening dialog = new TempleateOpening())
                    {
                        dialog.Show();
                        System.Windows.Forms.Application.DoEvents();

                        ServicePointManager.Expect100Continue = true;
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                        string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/Widescreen-template.potx";
                        string tempPath = Path.Combine(Path.GetTempPath(), "Widescreen-template.potx");

                        using (WebClient client = new WebClient())
                        {
                            try
                            {
                                client.DownloadFile(url, tempPath);
                            }
                            catch (WebException ex)
                            {
                                MessageBox.Show($"Failed to download the template: {ex.Message}\nPlease check your internet connection and try again.",
                                    "WSMacro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                dialog.Close();
                                return;
                            }
                        }

                        PowerPoint.Application pptApp;
                        try
                        {
                            pptApp = (PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
                        }
                        catch
                        {
                            pptApp = new PowerPoint.Application();
                        }

                        pptApp.Visible = Office.MsoTriState.msoTrue;

                        // Open as a new editable presentation
                        pptApp.Presentations.Open(tempPath, WithWindow: Office.MsoTriState.msoTrue);

                        dialog.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}\nPlease try again.", "WSMacro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void A4Macro(Office.IRibbonControl control)
        {
            try
            {
                // Show the dialog as the first task
                using (TempleateOpening dialog = new TempleateOpening())
                {
                    dialog.Show();
                    System.Windows.Forms.Application.DoEvents(); 

                    // Force TLS 1.2/1.3 for modern HTTPS
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                    string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/A4-template.potx";
                    string tempPath = Path.Combine(Path.GetTempPath(), "A4-template.potx");

                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(url, tempPath);
                    }

                    PowerPoint.Application pptApp;
                    try
                    {
                        pptApp = (PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        pptApp = new PowerPoint.Application();
                    }

                    pptApp.Visible = Office.MsoTriState.msoTrue;

                    // Open as a new untitled editable presentation
                    PowerPoint.Presentation presentation = pptApp.Presentations.Open(
                        tempPath,
                        WithWindow: Office.MsoTriState.msoTrue
                    );

                    // Close the dialog after the presentation opens
                    dialog.Close();

                    presentation = null;
                    pptApp = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}\nPlease try again.", "A4Macro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void USLMacro(Office.IRibbonControl control)
        {
            try
            {
                // Show the dialog as the first task
                using (TempleateOpening dialog = new TempleateOpening())
                {
                    dialog.Show();
                    System.Windows.Forms.Application.DoEvents(); 

                    // Force TLS 1.2/1.3 for modern HTTPS
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                    string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/US-Letter-template.potx";
                    string tempPath = Path.Combine(Path.GetTempPath(), "US-Letter-template.potx");

                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(url, tempPath);
                    }

                    PowerPoint.Application pptApp;
                    try
                    {
                        pptApp = (PowerPoint.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        pptApp = new PowerPoint.Application();
                    }

                    pptApp.Visible = Office.MsoTriState.msoTrue;

                    // Open as a new untitled editable presentation
                    PowerPoint.Presentation presentation = pptApp.Presentations.Open(
                        tempPath,
                        WithWindow: Office.MsoTriState.msoTrue
                    );

                    // Close the dialog after the presentation opens
                    dialog.Close();

                    presentation = null;
                    pptApp = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}\nPlease try again.", "USLMacro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnOpenLoginToTLBR(Office.IRibbonControl control)
        {
            using (var dialog = new LoginToTLBR())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var info = dialog.GetLoginData();
                    if (info != null)
                    {
                        MessageBox.Show($"Welcome! Email: {info.Email}, Password: {info.Password}", "PitchX");
                    }
                    else
                    {
                        MessageBox.Show("Login data not available.", "Toolbar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        public void Margin0(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                    return;

                // Check selection type: shapes, text, or text range
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;
                    float marginValue = 0;

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shp = shapeRange[i];

                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame textFrame = shp.TextFrame;

                            // Set all margins to 0 points
                            textFrame.MarginLeft = marginValue;
                            textFrame.MarginRight = marginValue;
                            textFrame.MarginTop = marginValue;
                            textFrame.MarginBottom = marginValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Margin00Macro");
            }
        }

        public void Margin1(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                    return;

                // Check if shapes or text are selected
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    // 0.1 cm = 2.834645669 points
                    float marginValue = 2.834645669f;

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shp = shapeRange[i];

                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame textFrame = shp.TextFrame;

                            // Set margins to 0.1 cm
                            textFrame.MarginLeft = marginValue;
                            textFrame.MarginRight = marginValue;
                            textFrame.MarginTop = marginValue;
                            textFrame.MarginBottom = marginValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Margin01Macro");
            }
        }

        public void Margin2(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                    return;

                // Check if shapes or text are selected
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    // 0.2 cm = 5.669291338 points
                    float marginValue = 5.669291338f;

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shp = shapeRange[i];

                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame textFrame = shp.TextFrame;

                            // Set margins to 0.2 cm
                            textFrame.MarginLeft = marginValue;
                            textFrame.MarginRight = marginValue;
                            textFrame.MarginTop = marginValue;
                            textFrame.MarginBottom = marginValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Margin02Macro");
            }
        }

        public void CustomMargin(Office.IRibbonControl control)
        {
            try
            {
                using (var dialog = new CustomMarginMacro())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // Convert cm to points (1 cm = 28.3464567 points)
                        float marginPoints = dialog.MarginValueCm * 28.3464567f;

                        bool applyLeft = dialog.ApplyLeft;
                        bool applyRight = dialog.ApplyRight;
                        bool applyTop = dialog.ApplyTop;
                        bool applyBottom = dialog.ApplyBottom;

                        PowerPoint.Application app = Globals.ThisAddIn.Application;
                        PowerPoint.Selection sel = app.ActiveWindow.Selection;

                        if (sel == null)
                            return;

                        // Apply to selected shapes or text
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                            sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                            for (int i = 1; i <= shapeRange.Count; i++)
                            {
                                PowerPoint.Shape shp = shapeRange[i];

                                if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                    PowerPoint.TextFrame textFrame = shp.TextFrame;

                                    // Apply only selected sides
                                    if (applyLeft)
                                        textFrame.MarginLeft = marginPoints;

                                    if (applyRight)
                                        textFrame.MarginRight = marginPoints;

                                    if (applyTop)
                                        textFrame.MarginTop = marginPoints;

                                    if (applyBottom)
                                        textFrame.MarginBottom = marginPoints;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "CustomMargin");
            }
        }

        public void LineSpace6(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("Please select a text box or a table to apply line spacing.", "LineSpace6");
                    return;
                }

                // Handle shapes or text selection
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shp = shapeRange[i];

                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue &&
                            shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange txtRange = shp.TextFrame.TextRange;

                            // Loop through each paragraph in the text
                            for (int j = 1; j <= txtRange.Paragraphs().Count; j++)
                            {
                                PowerPoint.TextRange paragraph = txtRange.Paragraphs(j);
                                PowerPoint.ParagraphFormat pf = paragraph.ParagraphFormat;

                                // Apply spacing
                                pf.SpaceBefore = 0f;   
                                pf.SpaceAfter = 6f;  
                                pf.SpaceWithin = 1f; 
                            }
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    MessageBox.Show("Please select a text box or table with text.", "LineSpace6");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "LineSpace6");
            }
        }

        public void LineSpace1(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("Please select a text box or a table to adjust line spacing.", "LineSpace1");
                    return;
                }

                // Process selected shapes or text
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    for (int i = 1; i <= shapeRange.Count; i++)
                    {
                        PowerPoint.Shape shp = shapeRange[i];

                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue &&
                            shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextRange txtRange = shp.TextFrame.TextRange;

                            // Loop through each paragraph
                            for (int j = 1; j <= txtRange.Paragraphs().Count; j++)
                            {
                                PowerPoint.TextRange paragraph = txtRange.Paragraphs(j);
                                PowerPoint.ParagraphFormat pf = paragraph.ParagraphFormat;

                                float currentAfter = pf.SpaceAfter;

                                // Decrease by 1 point, not below 0
                                if (currentAfter > 0f)
                                {
                                    float newAfter = currentAfter - 1f;
                                    if (newAfter < 0f) newAfter = 0f;
                                    pf.SpaceAfter = newAfter;
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a text box or table with text.", "LineSpace1");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "LineSpace1");
            }
        }

        public void MergeTextBoxes(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                // Ensure shapes are selected
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select text boxes to merge.", "Merge Text Boxes");
                    return;
                }

                PowerPoint.ShapeRange shapeRange = sel.ShapeRange;
                StringBuilder mergedText = new StringBuilder();

                // Collect text from all selected shapes
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    PowerPoint.Shape shp = shapeRange[i];

                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue &&
                        shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        mergedText.AppendLine(shp.TextFrame.TextRange.Text.Trim());
                    }
                }

                string finalText = mergedText.ToString().Trim();

                if (string.IsNullOrEmpty(finalText))
                {
                    MessageBox.Show("No text found in selected text boxes.", "Merge Text Boxes");
                    return;
                }

                // Add new textbox on current slide
                PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide;
                PowerPoint.Shape newTextBox = activeSlide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    100, 100, 400, 100);

                newTextBox.TextFrame.TextRange.Text = finalText;

                // Optionally, delete original shapes after merging
                // for (int i = shapeRange.Count; i >= 1; i--)
                // {
                //     shapeRange[i].Delete();
                // }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Merge Text Boxes");
            }
        }

        public void SplitTextBoxes(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("Please select a text box to split.", "SplitTextBoxes");
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.ShapeRange shapeRange = sel.ShapeRange;

                    // iterate backwards so deleting shapes doesn't break the loop
                    for (int idx = shapeRange.Count; idx >= 1; idx--)
                    {
                        PowerPoint.Shape oShape = shapeRange[idx];

                        if (oShape.HasTextFrame == Office.MsoTriState.msoTrue &&
                            oShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame2 textFrame2 = oShape.TextFrame2;
                            string fullText = Convert.ToString(textFrame2.TextRange.Text) ?? string.Empty;

                            // split by any newline variant (VBA paragraph behavior)
                            string[] paragraphs = Regex.Split(fullText, @"\r\n|\r|\n");

                            float tTop = oShape.Top;
                            float tLeft = oShape.Left;
                            float tWidth = oShape.Width;
                            float tHeight = oShape.Height;

                            // get slide to add shapes to (use view slide)
                            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                            foreach (string p in paragraphs)
                            {
                                string paraText = p?.Trim();
                                if (string.IsNullOrEmpty(paraText))
                                    continue;

                                PowerPoint.Shape newShape = slide.Shapes.AddTextbox(
                                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                    tLeft, tTop, tWidth, tHeight);

                                newShape.TextFrame2.TextRange.Text = paraText;
                                newShape.TextFrame2.WordWrap = Office.MsoTriState.msoTrue;
                                newShape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeShapeToFitText;

                                // move down for the next textbox
                                tTop += newShape.Height + 5f;
                            }

                            // delete the original combined shape
                            oShape.Delete();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a text box with text.", "SplitTextBoxes");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SplitTextBoxes");
            }
        }

        public void PlainText(Office.IRibbonControl control)
        {
            PowerPoint.Presentation sourcePresentation = null;

            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation destPresentation = app.ActivePresentation;
                PowerPoint.Slide destSlide = app.ActiveWindow.View.Slide;

                string tempFilePath = Path.GetTempPath();
                string tempFileName = "PPT-elements.pptx";
                string tempFullPath = Path.Combine(tempFilePath, tempFileName);
                string fileUrl = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx";

                // Force TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // Download file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(fileUrl, tempFullPath);
                }

                // Open downloaded presentation (hidden)
                sourcePresentation = app.Presentations.Open(tempFullPath, WithWindow: Office.MsoTriState.msoFalse);
                PowerPoint.Slide sourceSlide = sourcePresentation.Slides[1];

                PowerPoint.Shape sourceShape = null;

                // Find the first text box with text
                foreach (PowerPoint.Shape shp in sourceSlide.Shapes)
                {
                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue &&
                        shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        sourceShape = shp;
                        shp.Copy();
                        break;
                    }
                }

                if (sourceShape == null)
                {
                    MessageBox.Show("No text box found on the first slide.", "Text Box Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Paste to current slide
                PowerPoint.ShapeRange pastedRange = destSlide.Shapes.Paste();
                PowerPoint.Shape pastedShape = pastedRange[1];

                // Set fixed size and position
                float cmToPoints = 28.3465f;
                float shapeWidth = 6 * cmToPoints;
                float shapeHeight = 4 * cmToPoints;
                float leftPosition = destSlide.Parent.PageSetup.SlideWidth + (1.2f * cmToPoints);
                float topPosition = (destSlide.Parent.PageSetup.SlideHeight - shapeHeight) / 2;

                pastedShape.Width = shapeWidth;
                pastedShape.Height = shapeHeight;
                pastedShape.Left = leftPosition;
                pastedShape.Top = topPosition;

                // Close hidden source
                sourcePresentation.Close();
            }
            catch (WebException ex)
            {
                MessageBox.Show($"Failed to download the file.\n\n{ex.Message}", "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "PlainText Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sourcePresentation?.Close();
            }
        }

        public void Strapline1(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Convert slide size from points to cm
                double slideWidth = slide.Master.Width / 28.3465;
                double slideHeight = slide.Master.Height / 28.3465;

                PowerPoint.Shape selectedShape = null;
                bool shapeSelected = false;

                // Check if any shape is selected
                if (app.ActiveWindow.Selection != null &&
                    app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    selectedShape = app.ActiveWindow.Selection.ShapeRange[1];
                    shapeSelected = true;
                }

                // Define parameters
                double boxWidth = 0;
                double boxHeight = 0;
                double boxHorizontal = 0;
                double boxVertical = 0;
                double marginLeft = 0;
                double marginRight = 0;
                double marginTop = 0;
                double marginBottom = 0;

                // Determine template type and set parameters
                if (slideWidth >= 29.6 && slideWidth <= 29.8 && slideHeight >= 20.9 && slideHeight <= 21.1)
                {
                    boxWidth = 29.704;
                    boxHeight = 1;
                    boxHorizontal = 0;
                    boxVertical = 18.2;
                    marginLeft = 2.8;
                    marginRight = 1.9;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else if (slideWidth >= 33.7 && slideWidth <= 33.9 && slideHeight >= 18.9 && slideHeight <= 19.1)
                {
                    boxWidth = 33.876;
                    boxHeight = 1;
                    boxHorizontal = 0;
                    boxVertical = 15.99;
                    marginLeft = 3.2;
                    marginRight = 1.06;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else if (slideWidth >= 27.8 && slideWidth <= 28 && slideHeight >= 21.4 && slideHeight <= 21.7)
                {
                    boxWidth = 27.94;
                    boxHeight = 1;
                    boxHorizontal = 0;
                    boxVertical = 18.52;
                    marginLeft = 2.6;
                    marginRight = 1.8;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else
                {
                    MessageBox.Show("The template in use is not supported by this toolbar.", "Strapline1");
                    return;
                }

                PowerPoint.Shape shape;
                string shapeText = "";

                // Add or replace rectangle
                if (!shapeSelected)
                {
                    // Add a new rectangle
                    shape = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        (float)(boxHorizontal * 28.3465),
                        (float)(boxVertical * 28.3465),
                        (float)(boxWidth * 28.3465),
                        (float)(boxHeight * 28.3465)
                    );
                    shape.TextFrame.TextRange.Text = "[Key message – optional]";
                }
                else
                {
                    // Copy text, delete old shape, and add a new one
                    shapeText = selectedShape.TextFrame.TextRange.Text;
                    selectedShape.Delete();

                    shape = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        (float)(boxHorizontal * 28.3465),
                        (float)(boxVertical * 28.3465),
                        (float)(boxWidth * 28.3465),
                        (float)(boxHeight * 28.3465)
                    );
                    shape.TextFrame.TextRange.Text = shapeText;
                }

                // Apply formatting
                shape.Width = (float)(boxWidth * 28.3465);
                shape.Height = (float)(boxHeight * 28.3465);
                shape.Left = (float)(boxHorizontal * 28.3465);
                shape.Top = (float)(boxVertical * 28.3465);

                shape.TextFrame.MarginLeft = (float)(marginLeft * 28.3465);
                shape.TextFrame.MarginRight = (float)(marginRight * 28.3465);
                shape.TextFrame.MarginTop = (float)(marginTop * 28.3465);
                shape.TextFrame.MarginBottom = (float)(marginBottom * 28.3465);

                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

                var font = shape.TextFrame.TextRange.Font;
                font.Name = "Helvetica Now Text";
                font.Size = 14;
                font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 117, 133));

                shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(229, 239, 240));
                shape.Fill.Transparency = 0f;
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                var para = shape.TextFrame.TextRange.ParagraphFormat;
                para.SpaceBefore = 0f;
                para.SpaceAfter = 6f;
                para.SpaceWithin = 1f;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Strapline1Macro Error");
            }
        }

        public void Strapline2(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Convert slide size from points to cm
                double slideWidth = slide.Master.Width / 28.3465;
                double slideHeight = slide.Master.Height / 28.3465;

                PowerPoint.Shape selectedShape = null;
                bool shapeSelected = false;

                // Check if any shape is selected
                if (app.ActiveWindow.Selection != null &&
                    app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    selectedShape = app.ActiveWindow.Selection.ShapeRange[1];
                    shapeSelected = true;
                }

                // Define parameters
                double boxWidth = 0;
                double boxHeight = 0;
                double boxHorizontal = 0;
                double boxVertical = 0;
                double marginLeft = 0;
                double marginRight = 0;
                double marginTop = 0;
                double marginBottom = 0;

                // Determine slide size and set parameters
                if (slideWidth >= 29.6 && slideWidth <= 29.8 && slideHeight >= 20.9 && slideHeight <= 21.1)
                {
                    boxWidth = 29.704;
                    boxHeight = 1.4;
                    boxHorizontal = 0;
                    boxVertical = 17.79;
                    marginLeft = 2.8;
                    marginRight = 1.9;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else if (slideWidth >= 33.7 && slideWidth <= 33.9 && slideHeight >= 18.9 && slideHeight <= 19.1)
                {
                    boxWidth = 33.876;
                    boxHeight = 1.4;
                    boxHorizontal = 0;
                    boxVertical = 15.6;
                    marginLeft = 3.2;
                    marginRight = 1.06;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else if (slideWidth >= 27.8 && slideWidth <= 28 && slideHeight >= 21.4 && slideHeight <= 21.7)
                {
                    boxWidth = 27.94;
                    boxHeight = 1.4;
                    boxHorizontal = 0;
                    boxVertical = 18.1;
                    marginLeft = 2.6;
                    marginRight = 1.8;
                    marginTop = 0.1;
                    marginBottom = 0.1;
                }
                else
                {
                    MessageBox.Show("The template in use is not supported by this toolbar.", "Strapline 2");
                    return;
                }

                PowerPoint.Shape shape;
                string shapeText = "";

                // Add or replace rectangle
                if (!shapeSelected)
                {
                    // Add a new rectangle
                    shape = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        (float)(boxHorizontal * 28.3465),
                        (float)(boxVertical * 28.3465),
                        (float)(boxWidth * 28.3465),
                        (float)(boxHeight * 28.3465)
                    );
                    shape.TextFrame.TextRange.Text = "[Key message – optional]";
                }
                else
                {
                    // Copy text, delete old shape, and add a new one
                    shapeText = selectedShape.TextFrame.TextRange.Text;
                    selectedShape.Delete();

                    shape = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        (float)(boxHorizontal * 28.3465),
                        (float)(boxVertical * 28.3465),
                        (float)(boxWidth * 28.3465),
                        (float)(boxHeight * 28.3465)
                    );
                    shape.TextFrame.TextRange.Text = shapeText;
                }

                // Apply formatting
                shape.Width = (float)(boxWidth * 28.3465);
                shape.Height = (float)(boxHeight * 28.3465);
                shape.Left = (float)(boxHorizontal * 28.3465);
                shape.Top = (float)(boxVertical * 28.3465);

                shape.TextFrame.MarginLeft = (float)(marginLeft * 28.3465);
                shape.TextFrame.MarginRight = (float)(marginRight * 28.3465);
                shape.TextFrame.MarginTop = (float)(marginTop * 28.3465);
                shape.TextFrame.MarginBottom = (float)(marginBottom * 28.3465);

                shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

                var font = shape.TextFrame.TextRange.Font;
                font.Name = "Helvetica Now Text";
                font.Size = 14;
                font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 117, 133));

                shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(229, 239, 240));
                shape.Fill.Transparency = 0f;
                shape.Line.Visible = Office.MsoTriState.msoFalse;

                var para = shape.TextFrame.TextRange.ParagraphFormat;
                para.SpaceBefore = 0f;
                para.SpaceAfter = 6f;
                para.SpaceWithin = 1f;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Strapline2Macro Error");
            }
        }

        public void ToggleWrap(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    MessageBox.Show("Please select a text box to toggle WordWrap.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame tf = shp.TextFrame;
                            if (tf.HasText == Office.MsoTriState.msoTrue)
                            {
                                // Toggle WordWrap
                                if (tf.WordWrap == Office.MsoTriState.msoTrue)
                                    tf.WordWrap = Office.MsoTriState.msoFalse;
                                else
                                    tf.WordWrap = Office.MsoTriState.msoTrue;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a text box to toggle WordWrap.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "ToggleWrap Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ToggleResize(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    MessageBox.Show("Please select a text box to toggle AutoSize.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.TextFrame tf = shp.TextFrame;
                            if (tf.HasText == Office.MsoTriState.msoTrue)
                            {
                                // Toggle AutoSize between none and fit-to-text
                                if (tf.AutoSize == PowerPoint.PpAutoSize.ppAutoSizeNone)
                                    tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                                else
                                    tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select a text box to toggle AutoSize.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "ToggleResize Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FillText(Office.IRibbonControl control)
        {
            PowerPoint.Presentation sourcePresentation = null;

            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation destPresentation = app.ActivePresentation;
                PowerPoint.Slide destSlide = app.ActiveWindow.View.Slide;

                string tempFilePath = Path.GetTempPath();
                string tempFileName = "PPT-elements.pptx";
                string tempFullPath = Path.Combine(tempFilePath, tempFileName);
                string fileUrl = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx";

                // Force TLS 1.2
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // Download file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(fileUrl, tempFullPath);
                }

                // Open downloaded presentation (hidden)
                sourcePresentation = app.Presentations.Open(tempFullPath, WithWindow: Office.MsoTriState.msoFalse);
                PowerPoint.Slide sourceSlide = sourcePresentation.Slides[2]; 

                PowerPoint.Shape sourceShape = null;

                // Find the first text box with text
                foreach (PowerPoint.Shape shp in sourceSlide.Shapes)
                {
                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue &&
                        shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        sourceShape = shp;
                        shp.Copy();
                        break;
                    }
                }

                if (sourceShape == null)
                {
                    MessageBox.Show("No text box found on the second slide.", "Text Box Not Found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Paste to current slide
                PowerPoint.ShapeRange pastedRange = destSlide.Shapes.Paste();
                PowerPoint.Shape pastedShape = pastedRange[1];

                // Same base as PlainText, but shifted downward
                float cmToPoints = 28.3465f;
                float shapeWidth = 6 * cmToPoints;
                float shapeHeight = 4 * cmToPoints;
                float leftPosition = destSlide.Parent.PageSetup.SlideWidth + (1.2f * cmToPoints);
                float topPosition = (destSlide.Parent.PageSetup.SlideHeight - shapeHeight) / 2 + shapeHeight + (0.5f * cmToPoints);

                pastedShape.Width = shapeWidth;
                pastedShape.Height = shapeHeight;
                pastedShape.Left = leftPosition;
                pastedShape.Top = topPosition;

                // Custom fill color (light blue)
                pastedShape.Fill.Visible = Office.MsoTriState.msoTrue;
                pastedShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(0xEE, 0xF3, 0xF6));
                pastedShape.Fill.Solid();

                // Black text
                if (pastedShape.TextFrame?.TextRange != null)
                {
                    pastedShape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
                }

                sourcePresentation.Close();
            }
            catch (WebException ex)
            {
                MessageBox.Show($"Failed to download the file.\n\n{ex.Message}", "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "FillText Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sourcePresentation?.Close();
            }
        }

        public void Insertkey(Office.IRibbonControl control)
        {
            try
            {
                // Force TLS 1.2/1.3 for modern HTTPS
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                string tempFilePath = Path.Combine(Path.GetTempPath(), "downloadedPresentation.pptx");

                // Download the PPT file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", tempFilePath);
                }

                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation sourcePresentation = app.Presentations.Open(tempFilePath, WithWindow: MsoTriState.msoFalse);
                PowerPoint.Slide sourceSlide = sourcePresentation.Slides[3];

                PowerPoint.Shape textBox = null;
                foreach (PowerPoint.Shape shape in sourceSlide.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        textBox = shape;
                        break;
                    }
                }

                if (textBox == null)
                {
                    MessageBox.Show("Text box not found on page 3.", "Insertkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    sourcePresentation.Close();
                    return;
                }

                PowerPoint.Presentation destPresentation = app.ActivePresentation;
                PowerPoint.Slide destSlide = app.ActiveWindow.View.Slide;

                float pageWidth = destSlide.Master.Width;
                float pageHeight = destSlide.Master.Height;

                bool shouldPaste = false;

                // Determine if paste should occur
                if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    shouldPaste = true;
                }
                else if ((pageWidth >= 29.6 * 28.3465 && pageWidth <= 29.8 * 28.3465 &&
                          pageHeight >= 20.9 * 28.3465 && pageHeight <= 21.1 * 28.3465) ||
                         (pageWidth >= 33.7 * 28.3465 && pageWidth <= 33.9 * 28.3465 &&
                          pageHeight >= 18.9 * 28.3465 && pageHeight <= 19.1 * 28.3465) ||
                         (pageWidth >= 27.8 * 28.3465 && pageWidth <= 28 * 28.3465 &&
                          pageHeight >= 21.4 * 28.3465 && pageHeight <= 21.7 * 28.3465))
                {
                    shouldPaste = true;
                }

                if (shouldPaste)
                {
                    textBox.Copy();

                    PowerPoint.Shape pastedShape = null;

                    if (app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.ShapeRange shapeRange = app.ActiveWindow.Selection.ShapeRange;
                        for (int i = 1; i <= shapeRange.Count; i++)
                        {
                            PowerPoint.Shape selectedShape = shapeRange[i];
                            destSlide.Shapes.Paste();
                            pastedShape = destSlide.Shapes[destSlide.Shapes.Count];

                            pastedShape.Width = selectedShape.Width;
                            pastedShape.Left = selectedShape.Left;
                            pastedShape.Top = selectedShape.Top + selectedShape.Height + (float)(0.1 * 28.3465);
                        }
                    }
                    else
                    {
                        destSlide.Shapes.Paste();
                        pastedShape = destSlide.Shapes[destSlide.Shapes.Count];

                        if (pageWidth >= 29.6 * 28.3465 && pageWidth <= 29.8 * 28.3465 &&
                            pageHeight >= 20.9 * 28.3465 && pageHeight <= 21.1 * 28.3465)
                        {
                            pastedShape.Width = (float)(26 * 28.3465);
                            pastedShape.Height = (float)(1.37 * 28.3465);
                            pastedShape.Left = (float)(2.78 * 28.3465);
                            pastedShape.Top = (float)(16.65 * 28.3465);
                        }
                        else if (pageWidth >= 33.7 * 28.3465 && pageWidth <= 33.9 * 28.3465 &&
                                 pageHeight >= 18.9 * 28.3465 && pageHeight <= 19.1 * 28.3465)
                        {
                            pastedShape.Width = (float)(29.64 * 28.3465);
                            pastedShape.Height = (float)(1.37 * 28.3465);
                            pastedShape.Left = (float)(3.18 * 28.3465);
                            pastedShape.Top = (float)(14.4 * 28.3465);
                        }
                        else if (pageWidth >= 27.8 * 28.3465 && pageWidth <= 28 * 28.3465 &&
                                 pageHeight >= 21.4 * 28.3465 && pageHeight <= 21.7 * 28.3465)
                        {
                            pastedShape.Width = (float)(24.45 * 28.3465);
                            pastedShape.Height = (float)(1.37 * 28.3465);
                            pastedShape.Left = (float)(2.62 * 28.3465);
                            pastedShape.Top = (float)(16.98 * 28.3465);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("The template in use is not supported by this toolbar.",
                        "Insertkey", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                sourcePresentation.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}", "Insertkey Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void fullStop(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;
                PowerPoint.Shape shp;
                PowerPoint.TextRange txtRange;
                PowerPoint.Table tbl;
                int iRow, iCol;

                // Create regex equivalent of VBScript.RegExp
                Regex regEx = new Regex(@"\s+([.!?:;])", RegexOptions.Compiled);

                if (selection == null ||
                   (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a shape, text box, or table cell before running the TLBR.",
                                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    txtRange = selection.TextRange;
                    ProcessTextRange(txtRange, regEx);
                }

                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    shp = shape;

                    if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        if (shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            txtRange = shp.TextFrame.TextRange;
                            ProcessTextRange(txtRange, regEx);
                        }
                    }

                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        tbl = shp.Table;
                        for (iRow = 1; iRow <= tbl.Rows.Count; iRow++)
                        {
                            for (iCol = 1; iCol <= tbl.Columns.Count; iCol++)
                            {
                                PowerPoint.Cell cell = tbl.Cell(iRow, iCol);
                                if (cell.Selected)
                                {
                                    txtRange = cell.Shape.TextFrame.TextRange;
                                    ProcessTextRange(txtRange, regEx);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred:\n{ex.Message}",
                                "fullStop Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProcessTextRange(PowerPoint.TextRange txtRange, Regex regEx)
        {
            // Remove extra spaces before punctuation while preserving formatting
            string text = txtRange.Text;
            if (!string.IsNullOrEmpty(text))
            {
                string cleanedText = regEx.Replace(text, "$1");
                if (cleanedText != text)
                {
                    txtRange.Text = cleanedText;
                }
            }

            // Process each paragraph to add a full stop if needed
            for (int i = 1; i <= txtRange.Paragraphs().Count; i++)
            {
                PowerPoint.TextRange paragraph = txtRange.Paragraphs(i, 1);
                string pText = paragraph.Text.TrimEnd('\r', '\n');

                if (!string.IsNullOrEmpty(pText) &&
                    !pText.EndsWith(".") && !pText.EndsWith("!") &&
                    !pText.EndsWith("?") && !pText.EndsWith(";") && !pText.EndsWith(":"))
                {
                    // Insert a full stop at the end of the paragraph
                    PowerPoint.TextRange endOfParagraph = paragraph.Characters(paragraph.Text.Length, 0);
                    endOfParagraph.InsertAfter(".");
                }
            }
        }

        public void NoFill(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("Please select a shape first!", "No Fill");
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        shp.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                }
                else
                {
                    MessageBox.Show("Please select a shape first!", "No Fill");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "NoFill Error");
            }
        }

        public void NoOutline(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("Please select a shape first!", "No Outline");
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        shp.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                }
                else
                {
                    MessageBox.Show("Please select a shape first!", "No Outline");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "NoOutline Error");
            }
        }

        public void SelectSameFill(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a single shape first.", "Select Same Fill");
                    return;
                }

                // Get the selected shape
                PowerPoint.Shape selectedShape = sel.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Ensure the shape has a visible fill
                if (selectedShape.Fill.Visible != Office.MsoTriState.msoTrue)
                {
                    MessageBox.Show("The selected shape has no visible fill.", "Select Same Fill");
                    return;
                }

                // Reference fill properties
                int fillColor = selectedShape.Fill.ForeColor.RGB;
                float fillTransparency = selectedShape.Fill.Transparency;

                // Collect matching shapes
                List<PowerPoint.Shape> matchingShapes = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shp in slide.Shapes)
                {
                    try
                    {
                        if (shp.Fill.Visible == Office.MsoTriState.msoTrue &&
                            shp.Fill.ForeColor.RGB == fillColor &&
                            Math.Abs(shp.Fill.Transparency - fillTransparency) < 0.001)
                        {
                            matchingShapes.Add(shp);
                        }
                    }
                    catch
                    {
                        // Skip shapes that don't support fill
                        continue;
                    }
                }

                if (matchingShapes.Count > 0)
                {
                    // Select all matching shapes together
                    PowerPoint.ShapeRange range = slide.Shapes.Range(matchingShapes.Select(s => s.Name).ToArray());
                    range.Select();
                }
                else
                {
                    MessageBox.Show("No shapes found with the same fill color.", "Select Same Fill");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SelectSameFillMacro Error");
            }
        }

        public void SelectSameWidth(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a single shape.", "Select Same Width");
                    return;
                }

                PowerPoint.Shape selectedShape = sel.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float selectedWidth = selectedShape.Width;

                foreach (PowerPoint.Shape shp in slide.Shapes)
                {
                    try
                    {
                        if (Math.Abs(shp.Width - selectedWidth) < 0.001)
                        {
                            shp.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SelectSameWidth Error");
            }
        }

        public void SelectSameHeight(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a single shape.", "Select Same Height");
                    return;
                }

                PowerPoint.Shape selectedShape = sel.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float selectedHeight = selectedShape.Height;

                foreach (PowerPoint.Shape shp in slide.Shapes)
                {
                    try
                    {
                        if (Math.Abs(shp.Height - selectedHeight) < 0.001)
                        {
                            shp.Select(Office.MsoTriState.msoFalse);
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SelectSameHeight Error");
            }
        }

        public void SwapPositions(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count != 2)
                {
                    MessageBox.Show("Select exactly two shapes to swap positions.", "Swap Positions");
                    return;
                }

                PowerPoint.Shape shp1 = sel.ShapeRange[1];
                PowerPoint.Shape shp2 = sel.ShapeRange[2];

                float tempLeft = shp1.Left;
                float tempTop = shp1.Top;

                shp1.Left = shp2.Left;
                shp1.Top = shp2.Top;

                shp2.Left = tempLeft;
                shp2.Top = tempTop;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SwapPositions Error");
            }
        }

        public void CopyPosition(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a shape.", "Copy Position");
                    return;
                }

                if (sel.ShapeRange.Count != 1)
                {
                    MessageBox.Show("Please select only one shape to copy its position.", "Copy Position");
                    return;
                }

                PowerPoint.Shape selectedShape = sel.ShapeRange[1];

                float copiedLeft = selectedShape.Left;
                float copiedTop = selectedShape.Top;

                // Store values in presentation tags for later use
                app.ActivePresentation.Tags.Add("CopiedLeft", copiedLeft.ToString());
                app.ActivePresentation.Tags.Add("CopiedTop", copiedTop.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "CopyPositionMacro Error");
            }
        }

        public void PastePosition(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                // Retrieve stored position values
                string leftTag = pres.Tags["CopiedLeft"];
                string topTag = pres.Tags["CopiedTop"];

                if (string.IsNullOrEmpty(leftTag) || string.IsNullOrEmpty(topTag))
                {
                    MessageBox.Show("No position has been copied. Use CopyPositionMacro first.", "Paste Position");
                    return;
                }

                float copiedLeft = float.Parse(leftTag);
                float copiedTop = float.Parse(topTag);

                // Check if shapes are selected
                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select at least one shape to paste the position.", "Paste Position");
                    return;
                }

                // Apply copied position to all selected shapes
                foreach (PowerPoint.Shape shp in sel.ShapeRange)
                {
                    try
                    {
                        shp.Left = copiedLeft;
                        shp.Top = copiedTop;
                    }
                    catch
                    {
                        // Skip shapes that can’t be moved
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "PastePositionMacro Error");
            }
        }

        public void SelectSameOutline(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a single shape.", "Select Same Outline");
                    return;
                }

                // Get the selected shape
                PowerPoint.Shape selectedShape = sel.ShapeRange[1];
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Get reference outline properties
                int outlineColor = selectedShape.Line.ForeColor.RGB;
                float outlineWeight = selectedShape.Line.Weight;
                float outlineTransparency = selectedShape.Line.Transparency;

                // Create a list to collect all matching shapes
                List<PowerPoint.Shape> matchingShapes = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shp in slide.Shapes)
                {
                    try
                    {
                        if (shp.Line.Visible == Office.MsoTriState.msoTrue &&
                            shp.Line.ForeColor.RGB == outlineColor &&
                            Math.Abs(shp.Line.Weight - outlineWeight) < 0.001 &&
                            Math.Abs(shp.Line.Transparency - outlineTransparency) < 0.001)
                        {
                            matchingShapes.Add(shp);
                        }
                    }
                    catch
                    {
                        // Some shapes may not support Line properties
                        continue;
                    }
                }

                if (matchingShapes.Count > 0)
                {
                    // Convert list to ShapeRange and select all at once
                    PowerPoint.ShapeRange range = slide.Shapes.Range(matchingShapes.Select(s => s.Name).ToArray());
                    range.Select();
                }
                else
                {
                    MessageBox.Show("No shapes found with the same outline.", "Select Same Outline");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SelectSameOutlineMacro Error");
            }
        }

        public void AlignLeft(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Left = 0;
                    }
                    else
                    {
                        float refLeft = shapeRange[1].Left;
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Left = refLeft;
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Left = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignLeft Error");
            }
        }

        public void AlignRight(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                float slideWidth = currentSlide.Master.Width;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Left = slideWidth - shapeRange[1].Width;
                    }
                    else
                    {
                        float refRight = shapeRange[1].Left + shapeRange[1].Width;
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Left = refRight - shp.Width;
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Left = slideWidth - shp.Width;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignRight Error");
            }
        }

        public void AlignTop(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Top = 0;
                    }
                    else
                    {
                        float refTop = shapeRange[1].Top;
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Top = refTop;
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Top = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignTop Error");
            }
        }

        public void AlignBottom(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                float slideHeight = currentSlide.Master.Height;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Top = slideHeight - shapeRange[1].Height;
                    }
                    else
                    {
                        float refBottom = shapeRange[1].Top + shapeRange[1].Height;
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Top = refBottom - shp.Height;
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Top = slideHeight - shp.Height;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignBottom Error");
            }
        }

        public void AlignCentre(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                float slideWidth = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Left = (slideWidth - shapeRange[1].Width) / 2;
                    }
                    else
                    {
                        float refCentre = shapeRange[1].Left + (shapeRange[1].Width / 2);
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Left = refCentre - (shp.Width / 2);
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Left = (slideWidth - shp.Width) / 2;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignCentre Error");
            }
        }

        public void AlignMiddle(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                float slideHeight = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    if (shapeRange.Count == 1)
                    {
                        shapeRange[1].Top = (slideHeight - shapeRange[1].Height) / 2;
                    }
                    else
                    {
                        float refMiddle = shapeRange[1].Top + (shapeRange[1].Height / 2);
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            shp.Top = refMiddle - (shp.Height / 2);
                        }
                    }
                }
                else
                {
                    foreach (PowerPoint.Shape shp in currentSlide.Shapes)
                    {
                        shp.Top = (slideHeight - shp.Height) / 2;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "AlignMiddle Error");
            }
        }

        public void DistributeHorizontal(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    selection.ShapeRange.Distribute(PowerPoint.MsoDistributeCmd.msoDistributeHorizontally, Microsoft.Office.Core.MsoTriState.msoFalse);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "DistributeHorizontal Error");
            }
        }

        public void DistributeVertical(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    selection.ShapeRange.Distribute(PowerPoint.MsoDistributeCmd.msoDistributeVertically, Microsoft.Office.Core.MsoTriState.msoFalse);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "DistributeVertical Error");
            }
        }

        public void SameWidth(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    float refWidth = shapeRange[1].Width;
                    foreach (PowerPoint.Shape shp in shapeRange)
                    {
                        shp.Width = refWidth;
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least one shape.", "SameWidth Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SameWidth Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SameHeight(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    float refHeight = shapeRange[1].Height;
                    foreach (PowerPoint.Shape shp in shapeRange)
                    {
                        shp.Height = refHeight;
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least one shape.", "SameHeight Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SameHeight Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SameWidthHeight(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange shapeRange = selection.ShapeRange;
                    float refWidth = shapeRange[1].Width;
                    float refHeight = shapeRange[1].Height;
                    foreach (PowerPoint.Shape shp in shapeRange)
                    {
                        shp.Width = refWidth;
                        shp.Height = refHeight;
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least one shape.", "SameWidthHeight Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SameWidthHeight Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Unrotate(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select one or more shapes before running the macro.", "Unrotate Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                foreach (PowerPoint.Shape shp in selection.ShapeRange)
                {
                    if (shp.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape || shp.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        if (shp.Rotation != 0)
                        {
                            // Store original rotation and size
                            float originalRotation = shp.Rotation;
                            float originalWidth = shp.Width;
                            float originalHeight = shp.Height;

                            // Get rotated corner positions
                            float[,] corners = GetRotatedCorners(shp);

                            // Find the corner closest to top-left of slide (0,0)
                            double minDist = double.MaxValue;
                            double anchorX = 0, anchorY = 0;
                            for (int i = 0; i < 4; i++)
                            {
                                double dist = corners[i, 0] * corners[i, 0] + corners[i, 1] * corners[i, 1];
                                if (dist < minDist)
                                {
                                    minDist = dist;
                                    anchorX = corners[i, 0];
                                    anchorY = corners[i, 1];
                                }
                            }

                            // Unrotate and swap dimensions
                            shp.Rotation = 0;
                            shp.Width = originalHeight;
                            shp.Height = originalWidth;

                            // Move shape so its new top-left corner matches the original anchor corner
                            shp.Left = (float)anchorX;
                            shp.Top = (float)anchorY;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Unrotate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private float[,] GetRotatedCorners(PowerPoint.Shape shp)
        {
            float[,] corners = new float[4, 2];
            float centerX = shp.Left + shp.Width / 2;
            float centerY = shp.Top + shp.Height / 2;
            float halfWidth = shp.Width / 2;
            float halfHeight = shp.Height / 2;
            float rotation = shp.Rotation * (float)(Math.PI / 180); 

            // Define unrotated corners relative to center
            float[,] unrotatedCorners = new float[,]
            {
                { -halfWidth, -halfHeight }, 
                { halfWidth, -halfHeight },  
                { halfWidth, halfHeight },  
                { -halfWidth, halfHeight }  
            };

            // Apply rotation matrix to each corner
            for (int i = 0; i < 4; i++)
            {
                float x = unrotatedCorners[i, 0];
                float y = unrotatedCorners[i, 1];
                corners[i, 0] = centerX + (float)(x * Math.Cos(rotation) - y * Math.Sin(rotation));
                corners[i, 1] = centerY + (float)(x * Math.Sin(rotation) + y * Math.Cos(rotation));
            }

            return corners;
        }

        public void TouchTop(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;

                PowerPoint.Slide sld = app.ActiveWindow.Selection.SlideRange[1];
                float bottomPos = app.ActiveWindow.Selection.ShapeRange[1].Top + app.ActiveWindow.Selection.ShapeRange[1].Height;

                for (int i = 2; i <= app.ActiveWindow.Selection.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape shp = app.ActiveWindow.Selection.ShapeRange[i];
                    shp.Top = bottomPos;
                    bottomPos = shp.Top + shp.Height;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "TouchTop Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void TouchBottom(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;

                PowerPoint.Slide sld = app.ActiveWindow.Selection.SlideRange[1];
                float topPos = app.ActiveWindow.Selection.ShapeRange[1].Top;

                for (int i = 2; i <= app.ActiveWindow.Selection.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape shp = app.ActiveWindow.Selection.ShapeRange[i];
                    shp.Top = topPos - shp.Height;
                    topPos = shp.Top;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "TouchBottom Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void TouchLeft(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;

                PowerPoint.Slide sld = app.ActiveWindow.Selection.SlideRange[1];
                float leftPos = app.ActiveWindow.Selection.ShapeRange[1].Left;

                for (int i = 2; i <= app.ActiveWindow.Selection.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape shp = app.ActiveWindow.Selection.ShapeRange[i];
                    shp.Left = leftPos - shp.Width;
                    leftPos = shp.Left;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "TouchLeft Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void TouchRight(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;

                PowerPoint.Slide sld = app.ActiveWindow.Selection.SlideRange[1];
                float rightPos = app.ActiveWindow.Selection.ShapeRange[1].Left + app.ActiveWindow.Selection.ShapeRange[1].Width;

                for (int i = 2; i <= app.ActiveWindow.Selection.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape shp = app.ActiveWindow.Selection.ShapeRange[i];
                    shp.Left = rightPos;
                    rightPos = shp.Left + shp.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "TouchRight Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RightSpaceAdjustment(Office.IRibbonControl control)
        {
            try
            {
                using (SpacingAdjustment spacingForm = new SpacingAdjustment())
                {
                    // Show spacing dialog
                    DialogResult result = spacingForm.ShowDialog();
                    if (result == DialogResult.Cancel)
                        return;

                    float cmValue = spacingForm.SpacingCmValue;
                    float marginValue = cmValue * 28.34645669f; 

                    PowerPoint.Application app = Globals.ThisAddIn.Application;
                    var win = app.ActiveWindow;
                    if (win == null) return;

                    PowerPoint.Selection sel = win.Selection;
                    if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        MessageBox.Show("Please select shapes to distribute horizontally.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    // Collect shapes in the order the user selected them (not sorted)
                    List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
                    for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        shapes.Add(sel.ShapeRange[i]);

                    if (shapes.Count < 2)
                        return; 

                    // Keep the first selected shape fixed
                    PowerPoint.Shape firstShape = shapes[0];
                    double currentLeft = firstShape.Left + firstShape.Width + marginValue;

                    // Move each subsequent shape to the right of the previous one
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        PowerPoint.Shape shp = shapes[i];
                        shp.Left = (float)currentLeft;
                        currentLeft += shp.Width + marginValue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "RightSpaceAdjustment Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void BottomSpaceAdjustment(Office.IRibbonControl control)
        {
            try
            {
                using (VerticalSpacingAdjustment spacingForm = new VerticalSpacingAdjustment())
                {
                    DialogResult result = spacingForm.ShowDialog();

                    if (result == DialogResult.Cancel)
                        return;

                    float cmValue = spacingForm.SpacingCmValue;
                    float spacingPoints = cmValue * 28.34645669f; 

                    PowerPoint.Application app = Globals.ThisAddIn.Application;
                    var win = app.ActiveWindow;
                    if (win == null) return;

                    PowerPoint.Selection sel = win.Selection;
                    if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        MessageBox.Show("Please select shapes to distribute.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    double topPos = sel.ShapeRange[1].Top;
                    int count = sel.ShapeRange.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shp = sel.ShapeRange[i];
                        shp.Top = (float)topPos;
                        topPos += shp.Height + spacingPoints;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "BottomSpaceAdjustment Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatTable(Office.IRibbonControl control)
        {
            try
            {
                using (var frm = new FormatTableForm())
                {
                    // The form will now actually close and return OK
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        ApplyTableFormatting(frm.HeadingRows, frm.TotalRows, frm.BandedOption, frm.FirstColumnBold);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "FormatTable Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyTableFormatting(int headingRows, int totalRows, string bandedOption, bool firstColBold)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;
                if (sel == null) return;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        if (shp.HasTable != Office.MsoTriState.msoFalse)
                        {
                            var tbl = shp.Table;
                            int rowCount = tbl.Rows.Count;
                            int colCount = tbl.Columns.Count;

                            // Reset table margins, borders, font
                            for (int r = 1; r <= rowCount; r++)
                            {
                                for (int c = 1; c <= colCount; c++)
                                {
                                    var cell = tbl.Cell(r, c);
                                    var tf = cell.Shape.TextFrame;
                                    tf.MarginTop = tf.MarginBottom = tf.MarginLeft = tf.MarginRight = 4.252f;
                                    tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;

                                    foreach (PowerPoint.PpBorderType border in new[] {
                                        PowerPoint.PpBorderType.ppBorderLeft,
                                        PowerPoint.PpBorderType.ppBorderRight,
                                        PowerPoint.PpBorderType.ppBorderTop,
                                        PowerPoint.PpBorderType.ppBorderBottom })
                                    {
                                        var b = cell.Borders[border];
                                        b.Weight = 1;
                                        b.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                        b.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                                    }

                                    var font = cell.Shape.TextFrame.TextRange.Font;
                                    font.Name = "Helvetica Now Text";
                                    font.Size = 10;
                                    font.Bold = Office.MsoTriState.msoFalse;
                                    font.Italic = Office.MsoTriState.msoFalse;
                                    font.Underline = Office.MsoTriState.msoFalse;
                                    font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                }
                            }

                            // Apply banding
                            if (bandedOption == "Rows")
                            {
                                for (int r = headingRows + 1; r <= rowCount - totalRows; r++)
                                {
                                    for (int c = 1; c <= colCount; c++)
                                    {
                                        var fill = tbl.Cell(r, c).Shape.Fill;
                                        fill.ForeColor.RGB = ((r - headingRows) % 2 == 1)
                                            ? System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(237, 241, 243))
                                            : System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(246, 248, 249));
                                        fill.Transparency = 0;
                                    }
                                }
                            }
                            else 
                            {
                                for (int c = 1; c <= colCount; c++)
                                {
                                    for (int r = 1; r <= rowCount; r++)
                                    {
                                        var fill = tbl.Cell(r, c).Shape.Fill;
                                        fill.ForeColor.RGB = (c % 2 == 1)
                                            ? System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(237, 241, 243))
                                            : System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(246, 248, 249));
                                        fill.Transparency = 0;
                                    }
                                }
                            }

                            // First column style
                            if (firstColBold)
                            {
                                for (int r = 1; r <= rowCount; r++)
                                {
                                    var cell = tbl.Cell(r, 1);
                                    cell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(205, 219, 222));
                                    cell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                }
                            }

                            // Heading rows
                            if (headingRows > 0)
                            {
                                for (int r = 1; r <= headingRows; r++)
                                {
                                    for (int c = 1; c <= colCount; c++)
                                    {
                                        var cell = tbl.Cell(r, c);
                                        cell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(205, 219, 222));
                                        cell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                        cell.Shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;

                                        if (r == headingRows)
                                        {
                                            var btm = cell.Borders[PowerPoint.PpBorderType.ppBorderBottom];
                                            btm.Weight = 3;
                                            btm.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                            btm.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                                        }
                                    }
                                }
                            }

                            // Total rows
                            if (totalRows > 0)
                            {
                                for (int r = rowCount - totalRows + 1; r <= rowCount; r++)
                                {
                                    for (int c = 1; c <= colCount; c++)
                                    {
                                        var cell = tbl.Cell(r, c);
                                        cell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(205, 219, 222));
                                        cell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Formatting Error: {ex.Message}", "ApplyTableFormatting Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } 

        public void TableTitle(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Table Title", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tableShape = shape;
                        float topPosition = tableShape.Top;
                        float leftPosition = tableShape.Left;
                        float width = tableShape.Width;

                        // Default title text
                        string tableTitle = "[Please enter table title]";

                        // Create text box above the table
                        PowerPoint.Shape textBox = tableShape.Parent.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            leftPosition,
                            topPosition - 20, 
                            width,
                            20
                        );

                        // Format textbox
                        var textFrame = textBox.TextFrame;
                        textFrame.MarginTop = 0;
                        textFrame.MarginBottom = 2;
                        textFrame.MarginLeft = 0;
                        textFrame.MarginRight = 0;
                        textFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

                        var textRange = textFrame.TextRange;
                        textRange.Text = tableTitle;
                        textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;

                        var font = textRange.Font;
                        font.Name = "Helvetica Now Text";
                        font.Size = 14;
                        font.Bold = Office.MsoTriState.msoTrue;
                        font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                        // Bring the text box above the table visually
                        textBox.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding table title:\n{ex.Message}", "Table Title Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void TableText(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Open the FontSize dialog
                using (var form = new FontSize())
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return; 

                    int fontSize = form.SelectedFontSize;

                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        if (shape.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var table = shape.Table;

                            // Loop through each cell and set the font size
                            for (int r = 1; r <= table.Rows.Count; r++)
                            {
                                for (int c = 1; c <= table.Columns.Count; c++)
                                {
                                    var cell = table.Cell(r, c);
                                    var textRange = cell.Shape.TextFrame.TextRange;
                                    textRange.Font.Size = fontSize;
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying font size:\n{ex.Message}", "Font Size Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void BandedRows(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int rowColour1 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(246, 248, 249));
                int rowColour2 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(237, 241, 243));
                int ignoreColour1 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(205, 219, 222));
                int ignoreColour2 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        if (shp.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var tbl = shp.Table;
                            for (int rowIndex = 1; rowIndex <= tbl.Rows.Count; rowIndex++)
                            {
                                for (int colIndex = 1; colIndex <= tbl.Columns.Count; colIndex++)
                                {
                                    var fill = tbl.Cell(rowIndex, colIndex).Shape.Fill;

                                    if ((fill.ForeColor.RGB != ignoreColour1 && fill.ForeColor.RGB != ignoreColour2) ||
                                        fill.Type != Office.MsoFillType.msoFillSolid)
                                    {
                                        if (rowIndex % 2 == 1)
                                        {
                                            fill.ForeColor.RGB = rowColour1;
                                            fill.Transparency = 0;
                                        }
                                        else
                                        {
                                            fill.ForeColor.RGB = rowColour2;
                                            fill.Transparency = 0;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    var shp = sel.ShapeRange[1];
                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = shp.Table;
                        for (int rowIndex = 1; rowIndex <= tbl.Rows.Count; rowIndex++)
                        {
                            for (int colIndex = 1; colIndex <= tbl.Columns.Count; colIndex++)
                            {
                                var fill = tbl.Cell(rowIndex, colIndex).Shape.Fill;

                                if ((fill.ForeColor.RGB != ignoreColour1 && fill.ForeColor.RGB != ignoreColour2) ||
                                    fill.Type != Office.MsoFillType.msoFillSolid)
                                {
                                    if (rowIndex % 2 == 1)
                                    {
                                        fill.ForeColor.RGB = rowColour1;
                                        fill.Transparency = 0;
                                    }
                                    else
                                    {
                                        fill.ForeColor.RGB = rowColour2;
                                        fill.Transparency = 0;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying banded rows:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void BandedColumns(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int colColour1 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(237, 241, 243));
                int colColour2 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(246, 248, 249));
                int ignoreColour1 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(205, 219, 222));
                int ignoreColour2 = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shp in sel.ShapeRange)
                    {
                        if (shp.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var tbl = shp.Table;
                            for (int colIndex = 1; colIndex <= tbl.Columns.Count; colIndex++)
                            {
                                for (int rowIndex = 1; rowIndex <= tbl.Rows.Count; rowIndex++)
                                {
                                    var fill = tbl.Cell(rowIndex, colIndex).Shape.Fill;

                                    if ((fill.ForeColor.RGB != ignoreColour1 && fill.ForeColor.RGB != ignoreColour2) ||
                                        fill.Type != Office.MsoFillType.msoFillSolid)
                                    {
                                        if (colIndex % 2 == 1)
                                        {
                                            fill.ForeColor.RGB = colColour1;
                                            fill.Transparency = 0;
                                        }
                                        else
                                        {
                                            fill.ForeColor.RGB = colColour2;
                                            fill.Transparency = 0;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    var shp = sel.ShapeRange[1];
                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = shp.Table;
                        for (int colIndex = 1; colIndex <= tbl.Columns.Count; colIndex++)
                        {
                            for (int rowIndex = 1; rowIndex <= tbl.Rows.Count; rowIndex++)
                            {
                                var fill = tbl.Cell(rowIndex, colIndex).Shape.Fill;

                                if ((fill.ForeColor.RGB != ignoreColour1 && fill.ForeColor.RGB != ignoreColour2) ||
                                    fill.Type != Office.MsoFillType.msoFillSolid)
                                {
                                    if (colIndex % 2 == 1)
                                    {
                                        fill.ForeColor.RGB = colColour1;
                                        fill.Transparency = 0;
                                    }
                                    else
                                    {
                                        fill.ForeColor.RGB = colColour2;
                                        fill.Transparency = 0;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying banded columns:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Highlight(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (sel.ShapeRange.Count > 0)
                {
                    PowerPoint.Shape shp = sel.ShapeRange[1];

                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = shp.Table;

                        int orangeRGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(209, 73, 0));

                        // Find the range of selected cells
                        int firstRow = tbl.Rows.Count;
                        int lastRow = 1;
                        int firstCol = tbl.Columns.Count;
                        int lastCol = 1;

                        for (int row = 1; row <= tbl.Rows.Count; row++)
                        {
                            for (int col = 1; col <= tbl.Columns.Count; col++)
                            {
                                PowerPoint.Cell cell = tbl.Cell(row, col);
                                if (cell.Selected)
                                {
                                    firstRow = Math.Min(firstRow, row);
                                    lastRow = Math.Max(lastRow, row);
                                    firstCol = Math.Min(firstCol, col);
                                    lastCol = Math.Max(lastCol, col);
                                }
                            }
                        }

                        // Check if any cells are selected
                        if (firstRow <= lastRow && firstCol <= lastCol)
                        {
                            // Apply border to the entire outer edge of the selected range
                            for (int row = firstRow; row <= lastRow; row++)
                            {
                                PowerPoint.Cell cell = tbl.Cell(row, firstCol);
                                cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = orangeRGB;
                                cell.Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = 2;

                                cell = tbl.Cell(row, lastCol);
                                cell.Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = orangeRGB;
                                cell.Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = 2;
                            }

                            for (int col = firstCol; col <= lastCol; col++)
                            {
                                PowerPoint.Cell cell = tbl.Cell(firstRow, col);
                                cell.Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = orangeRGB;
                                cell.Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 2;

                                cell = tbl.Cell(lastRow, col);
                                cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = orangeRGB;
                                cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 2;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Selected shape does not contain a table.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying highlight:\n{ex.Message}", "Highlight Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Horizontal(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (sel.ShapeRange.Count > 0)
                {
                    PowerPoint.Shape shp = sel.ShapeRange[1];
                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = shp.Table;
                        int borderColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(172, 192, 195));

                        for (int i = 1; i <= tbl.Rows.Count; i++)
                        {
                            for (int j = 1; j <= tbl.Columns.Count; j++)
                            {
                                var cell = tbl.Cell(i, j);
                                if (cell.Selected)
                                {
                                    var border = cell.Borders[PowerPoint.PpBorderType.ppBorderBottom];
                                    border.ForeColor.RGB = borderColor;
                                    border.Weight = 1f;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying horizontal border:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Vertical(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (sel.ShapeRange.Count > 0)
                {
                    PowerPoint.Shape shp = sel.ShapeRange[1];
                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var tbl = shp.Table;
                        int borderColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(172, 192, 195));

                        for (int i = 1; i <= tbl.Rows.Count; i++)
                        {
                            for (int j = 1; j <= tbl.Columns.Count; j++)
                            {
                                var cell = tbl.Cell(i, j);
                                if (cell.Selected)
                                {
                                    var border = cell.Borders[PowerPoint.PpBorderType.ppBorderRight];
                                    border.ForeColor.RGB = borderColor;
                                    border.Weight = 1f;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Selected shape does not contain a table.", "Font Size", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying vertical border:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Total(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWindow = app.ActiveWindow;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if there is an active window and slide
                if (activeWindow == null || activeWindow.View.Slide == null)
                    return;

                PowerPoint.Slide slide = activeWindow.View.Slide;
                bool tableFound = false;

                // Loop through each shape in the active slide
                foreach (PowerPoint.Shape pptShape in slide.Shapes)
                {
                    // Check if the shape is a table
                    if (pptShape.HasTable == MsoTriState.msoTrue)
                    {
                        PowerPoint.Table pptTable = pptShape.Table;

                        // Check if any cell in the table is selected
                        for (int i = 1; i <= pptTable.Rows.Count; i++)
                        {
                            for (int j = 1; j <= pptTable.Columns.Count; j++)
                            {
                                if (pptTable.Cell(i, j).Selected)
                                {
                                    tableFound = true;
                                    break;
                                }
                            }
                            if (tableFound) break;
                        }

                        // If a selected cell is found, format the table
                        if (tableFound)
                        {
                            List<int> selectedRows = new List<int>();

                            // Collect selected row indices
                            for (int i = 1; i <= pptTable.Rows.Count; i++)
                            {
                                for (int j = 1; j <= pptTable.Columns.Count; j++)
                                {
                                    if (pptTable.Cell(i, j).Selected)
                                    {
                                        if (!selectedRows.Contains(i))
                                            selectedRows.Add(i);
                                        break;
                                    }
                                }
                            }

                            // Format each selected row
                            for (int i = 0; i < selectedRows.Count; i++)
                            {
                                int rowIndex = selectedRows[i];
                                // Validate rowIndex is within bounds
                                if (rowIndex < 1 || rowIndex > pptTable.Rows.Count)
                                    continue;

                                for (int k = 1; k <= pptTable.Rows[rowIndex].Cells.Count; k++)
                                {
                                    PowerPoint.Cell cell = pptTable.Rows[rowIndex].Cells[k];

                                    // Set fill color and transparency
                                    cell.Shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(205, 219, 222));
                                    cell.Shape.Fill.Transparency = 0;

                                    // Set text to bold
                                    cell.Shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;

                                    // Set top border for the first selected row
                                    if (i == 0)
                                    {
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = 3;
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderTop].DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                                    }

                                    // Set bottom border
                                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 1;
                                    cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineSolid;

                                    // Set borders between rows
                                    if (rowIndex > 1 && rowIndex < pptTable.Rows.Count)
                                    {
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = 1;
                                        cell.Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                                    }
                                }
                            }

                            // Exit the loop after formatting the selected table
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error applying subtotal format:\n{ex.Message}", "Subtotal Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void CCWidth(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var win = app.ActiveWindow;

                if (win == null || win.Selection == null || win.Selection.ShapeRange == null || win.Selection.ShapeRange.Count != 1)
                {
                    MessageBox.Show("Please select exactly one table.", "Copy Column Widths",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.Shape shape = win.Selection.ShapeRange[1];
                if (shape.HasTable != MsoTriState.msoTrue)
                {
                    MessageBox.Show("The selected shape is not a table.", "Copy Column Widths",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.Table tbl = shape.Table;
                var presentation = app.ActivePresentation;

                // Clear previous stored widths
                for (int i = 0; i < tbl.Columns.Count; i++)
                {
                    string tagName = $"ColumnWidth_{i}";
                    if (!string.IsNullOrEmpty(presentation.Tags[tagName]))
                        presentation.Tags.Delete(tagName);
                }

                // Store column widths
                for (int i = 1; i <= tbl.Columns.Count; i++)
                {
                    string tagName = $"ColumnWidth_{i - 1}";
                    presentation.Tags.Add(tagName, tbl.Columns[i].Width.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying column widths:\n{ex.Message}", "Copy Column Widths Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ACWidth(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var win = app.ActiveWindow;

                if (win == null || win.Selection == null || win.Selection.ShapeRange == null || win.Selection.ShapeRange.Count == 0)
                {
                    MessageBox.Show("Please select at least one table.", "Apply Column Widths",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var presentation = app.ActivePresentation;
                bool appliedAny = false;
                int processedTables = 0;

                // Loop through all selected shapes
                foreach (PowerPoint.Shape shape in win.Selection.ShapeRange)
                {
                    if (shape.HasTable != MsoTriState.msoTrue)
                        continue; 

                    PowerPoint.Table tbl = shape.Table;

                    for (int i = 1; i <= tbl.Columns.Count; i++)
                    {
                        string tagName = $"ColumnWidth_{i - 1}";
                        string value = presentation.Tags[tagName];

                        if (!string.IsNullOrEmpty(value) && float.TryParse(value, out float width))
                        {
                            tbl.Columns[i].Width = width;
                            appliedAny = true;
                        }
                    }

                    processedTables++;
                }

                if (processedTables == 0)
                {
                    MessageBox.Show("No table found in the current selection.", "Apply Column Widths",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (!appliedAny)
                {
                    MessageBox.Show("No stored column widths found. Please copy first.", "Apply Column Widths",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying column widths:\n{ex.Message}", "Apply Column Widths Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void TableMargin15(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;
                var activeWindow = app.ActiveWindow;
                var slide = activeWindow.View.Slide;

                if (slide == null)
                {
                    MessageBox.Show("No active slide found.", "Table Margin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var shapes = selection.ShapeRange;
                int tablesFound = 0;

                // Convert 0.15 cm to points
                float marginPoints = 0.15f * 28.3464567f;

                foreach (PowerPoint.Shape shp in shapes)
                {
                    if (shp.HasTable == MsoTriState.msoTrue)
                    {
                        tablesFound++;
                        PowerPoint.Table tbl = shp.Table;

                        // Loop through each cell and set margins
                        for (int i = 1; i <= tbl.Rows.Count; i++)
                        {
                            for (int j = 1; j <= tbl.Columns.Count; j++)
                            {
                                var textFrame = tbl.Cell(i, j).Shape.TextFrame;
                                var paragraphFormat = textFrame.TextRange.ParagraphFormat;

                                paragraphFormat.SpaceBefore = 0;
                                paragraphFormat.SpaceAfter = 6;
                                paragraphFormat.SpaceWithin = 1;

                                textFrame.MarginLeft = marginPoints;
                                textFrame.MarginRight = marginPoints;
                                textFrame.MarginTop = marginPoints;
                                textFrame.MarginBottom = marginPoints;
                            }
                        }
                    }
                }

                if (tablesFound == 0)
                {
                    MessageBox.Show("The selected shape is not a table.", "Please select a table.",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying table margins:\n{ex.Message}",
                    "Table Margin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CustomMarginTable(Office.IRibbonControl control)
        {
            try
            {
                using (var dialog = new CustomMarginMacro())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // Convert cm to points (1 cm = 28.3464567 points)
                        float marginPoints = dialog.MarginValueCm * 28.3464567f;

                        bool applyLeft = dialog.ApplyLeft;
                        bool applyRight = dialog.ApplyRight;
                        bool applyTop = dialog.ApplyTop;
                        bool applyBottom = dialog.ApplyBottom;

                        var app = Globals.ThisAddIn.Application;
                        var selection = app.ActiveWindow.Selection;
                        var activeWindow = app.ActiveWindow;
                        var slide = activeWindow.View.Slide;

                        if (slide == null)
                        {
                            MessageBox.Show("No active slide found.", "Table Margin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        int tablesFound = 0;
                        int cellsModified = 0;

                        foreach (PowerPoint.Shape shape in selection.ShapeRange)
                        {
                            if (shape.HasTable == MsoTriState.msoTrue)
                            {
                                tablesFound++;
                                PowerPoint.Table tbl = shape.Table;

                                for (int i = 1; i <= tbl.Rows.Count; i++)
                                {
                                    for (int j = 1; j <= tbl.Columns.Count; j++)
                                    {
                                        var cell = tbl.Cell(i, j);
                                        var textFrame = cell.Shape.TextFrame;

                                        if (textFrame != null)
                                        {
                                            if (applyLeft) textFrame.MarginLeft = marginPoints;
                                            if (applyRight) textFrame.MarginRight = marginPoints;
                                            if (applyTop) textFrame.MarginTop = marginPoints;
                                            if (applyBottom) textFrame.MarginBottom = marginPoints;
                                            cellsModified++;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("The selected shape is not a table.", "Please select a table.",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "CustomMarginTable", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CopyMargin(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var win = app.ActiveWindow;

                if (win?.Selection == null || win.Selection.ShapeRange == null || win.Selection.ShapeRange.Count != 1)
                {
                    MessageBox.Show("Please select exactly one table.", "Copy Margin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.Shape shape = win.Selection.ShapeRange[1];
                if (shape.HasTable != MsoTriState.msoTrue)
                {
                    MessageBox.Show("The selected shape is not a table.", "Copy Margin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.Table tbl = shape.Table;
                float marginTop = 0f, marginBottom = 0f, marginLeft = 0f, marginRight = 0f;

                // Collect margins (use first valid cell as reference)
                for (int r = 1; r <= tbl.Rows.Count; r++)
                {
                    for (int c = 1; c <= tbl.Columns.Count; c++)
                    {
                        try
                        {
                            var tf = tbl.Cell(r, c).Shape.TextFrame2;
                            if (tf != null)
                            {
                                marginTop = (float)tf.MarginTop;
                                marginBottom = (float)tf.MarginBottom;
                                marginLeft = (float)tf.MarginLeft;
                                marginRight = (float)tf.MarginRight;
                                goto Done; 
                            }
                        }
                        catch { }
                    }
                }

            Done:
                var presentation = app.ActivePresentation;
                presentation.Tags.Add("MarginTop", marginTop.ToString());
                presentation.Tags.Add("MarginBottom", marginBottom.ToString());
                presentation.Tags.Add("MarginLeft", marginLeft.ToString());
                presentation.Tags.Add("MarginRight", marginRight.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying margins:\n{ex.Message}", "Copy Margin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void PasteMargin(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var win = app.ActiveWindow;

                if (win?.Selection == null || win.Selection.ShapeRange == null || win.Selection.ShapeRange.Count == 0)
                {
                    MessageBox.Show("Please select at least one table.", "Paste Margin", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var presentation = app.ActivePresentation;

                // Check if margins were previously copied
                if (string.IsNullOrEmpty(presentation.Tags["MarginTop"]))
                {
                    MessageBox.Show("No copied margins found. Please run CopyMargin first.", "Paste Margin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Retrieve stored margins
                float marginTop = float.Parse(presentation.Tags["MarginTop"]);
                float marginBottom = float.Parse(presentation.Tags["MarginBottom"]);
                float marginLeft = float.Parse(presentation.Tags["MarginLeft"]);
                float marginRight = float.Parse(presentation.Tags["MarginRight"]);

                int processedTables = 0;
                int totalCells = 0;

                // Loop through all selected shapes
                foreach (PowerPoint.Shape shape in win.Selection.ShapeRange)
                {
                    if (shape.HasTable != MsoTriState.msoTrue)
                        continue; 

                    PowerPoint.Table tbl = shape.Table;

                    for (int r = 1; r <= tbl.Rows.Count; r++)
                    {
                        for (int c = 1; c <= tbl.Columns.Count; c++)
                        {
                            try
                            {
                                var tf = tbl.Cell(r, c).Shape.TextFrame2;
                                if (tf != null)
                                {
                                    tf.MarginTop = marginTop;
                                    tf.MarginBottom = marginBottom;
                                    tf.MarginLeft = marginLeft;
                                    tf.MarginRight = marginRight;
                                    totalCells++;
                                }
                            }
                            catch{ }
                        }
                    }

                    processedTables++;
                }

                if (processedTables == 0)
                {
                    MessageBox.Show("No table found in the current selection.", "Paste Margin",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error pasting margins:\n{ex.Message}", "Paste Margin Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void FormatGraphFunction(Office.IRibbonControl control)
        {
            try
            {
                var application = Globals.ThisAddIn.Application as PowerPoint.Application;
                if (application == null) return;

                var activeWindow = application.ActiveWindow;
                if (activeWindow == null) return;

                var selection = activeWindow.Selection;
                if (selection == null) return;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shapeRange = selection.ShapeRange;
                    for (int s = 1; s <= shapeRange.Count; s++)
                    {
                        var shp = shapeRange[s] as PowerPoint.Shape;
                        if (shp == null) continue;

                        if (IsTrue(shp.HasChart))
                        {
                            var graph = shp.Chart;

                            // Remove backgrounds
                            graph.ChartArea.Format.Fill.Visible = Office.MsoTriState.msoFalse;
                            graph.PlotArea.Format.Fill.Visible = Office.MsoTriState.msoFalse;

                            // ChartArea text formatting
                            var chartAreaFont = graph.ChartArea.Format.TextFrame2.TextRange.Font;
                            ApplyFont(chartAreaFont, "Helvetica Now Text", 10.0f, Color.FromArgb(93, 109, 120));

                            // Chart title
                            if (IsTrue(graph.HasTitle))
                            {
                                var titleFont = graph.ChartTitle.Format.TextFrame2.TextRange.Font;
                                ApplyFont(titleFont, "Helvetica Now Text", 10.0f, Color.FromArgb(93, 109, 120));
                            }

                            // Legend
                            if (IsTrue(graph.HasLegend))
                            {
                                var legend = graph.Legend;
                                legend.Position = PowerPoint.XlLegendPosition.xlLegendPositionBottom;
                                var legendFont = legend.Format.TextFrame2.TextRange.Font;
                                ApplyFont(legendFont, "Helvetica Now Text", 10.0f, Color.FromArgb(93, 109, 120));
                            }

                            // SeriesCollection (call the method and iterate by 1-based index)
                            dynamic seriesCollection = null;
                            try { seriesCollection = graph.SeriesCollection(); } catch { seriesCollection = null; }

                            int seriesCount = 0;
                            if (seriesCollection != null)
                            {
                                try { seriesCount = (int)seriesCollection.Count; } catch { seriesCount = 0; }

                                for (int i = 1; i <= seriesCount; i++)
                                {
                                    dynamic series = null;
                                    try { series = seriesCollection.Item(i); } catch { continue; }
                                    if (series == null) continue;

                                    if (IsTrue(GetMember(series, "HasDataLabels")))
                                    {
                                        dynamic dataLabels = null;
                                        try { dataLabels = series.DataLabels(); } catch { dataLabels = null; }
                                        if (dataLabels != null)
                                        {
                                            try
                                            {
                                                var dlFont = dataLabels.Format.TextFrame2.TextRange.Font;
                                                ApplyFont(dlFont, "Helvetica Now Text", 10.0f, Color.FromArgb(93, 109, 120));
                                            }
                                            catch { }
                                        }
                                    }

                                    if (IsTrue(GetMember(series, "HasErrorBars")))
                                    {
                                        try { series.ErrorBars.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(93, 109, 120)); }
                                        catch { }
                                    }
                                }
                            }

                            // Data table
                            try
                            {
                                if (IsTrue(graph.HasDataTable))
                                {
                                    var dataTableFont = graph.DataTable.Format.TextFrame2.TextRange.Font;
                                    ApplyFont(dataTableFont, "Helvetica Now Text", 10.0f, Color.FromArgb(93, 109, 120));
                                }
                            }
                            catch { }

                            // Secondary axis
                            try
                            {
                                var secondaryValueAxis = graph.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary);
                                if (secondaryValueAxis != null)
                                {
                                    secondaryValueAxis.Format.Line.Visible = Office.MsoTriState.msoFalse;
                                    secondaryValueAxis.MajorTickMark = PowerPoint.XlTickMark.xlTickMarkNone;
                                    secondaryValueAxis.MinorTickMark = PowerPoint.XlTickMark.xlTickMarkNone;
                                    secondaryValueAxis.HasMinorGridlines = false;
                                    secondaryValueAxis.HasMajorGridlines = false;
                                }
                            }
                            catch { }

                            // Primary value axis + major gridlines
                            try
                            {
                                var primaryValueAxis = graph.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary);
                                if (primaryValueAxis != null)
                                {
                                    primaryValueAxis.Format.Line.Visible = Office.MsoTriState.msoFalse;
                                    primaryValueAxis.MajorTickMark = PowerPoint.XlTickMark.xlTickMarkNone;
                                    primaryValueAxis.MinorTickMark = PowerPoint.XlTickMark.xlTickMarkNone;
                                    primaryValueAxis.HasMinorGridlines = false;
                                    primaryValueAxis.HasMajorGridlines = true;

                                    var oMajorGridlines = primaryValueAxis.MajorGridlines;
                                    oMajorGridlines.Format.Line.Weight = 0.5f;
                                    oMajorGridlines.Format.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid; // fully qualified
                                    oMajorGridlines.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(229, 239, 240));
                                }
                            }
                            catch { }

                            // Primary category axis
                            try
                            {
                                var primaryCategoryAxis = graph.Axes(PowerPoint.XlAxisType.xlCategory, PowerPoint.XlAxisGroup.xlPrimary);
                                if (primaryCategoryAxis != null)
                                {
                                    try { primaryCategoryAxis.MajorGridlines.Delete(); } catch { }
                                    try { primaryCategoryAxis.MinorGridlines.Delete(); } catch { }

                                    var categoryLine = primaryCategoryAxis.Format.Line;
                                    categoryLine.Visible = Office.MsoTriState.msoTrue;
                                    categoryLine.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(172, 192, 196));
                                    categoryLine.Weight = 0.5f;
                                    categoryLine.DashStyle = Office.MsoLineDashStyle.msoLineSolid; // fully qualified

                                    primaryCategoryAxis.MajorTickMark = PowerPoint.XlTickMark.xlTickMarkOutside;
                                    primaryCategoryAxis.TickLabelPosition = PowerPoint.XlTickLabelPosition.xlTickLabelPositionLow;
                                    try { primaryCategoryAxis.TickLabels.Offset = 0; } catch { }
                                }
                            }
                            catch { }
                        } 
                    }
                } 
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }

        // Safe truth evaluation for bool / MsoTriState / numeric COM returns
        private bool IsTrue(object value)
        {
            if (value == null) return false;
            if (value is bool b) return b;
            try
            {
                // MsoTriState true = -1, false = 0 (and mixed = -2) — nonzero indicates truthiness
                int iv = Convert.ToInt32(value);
                return iv != 0;
            }
            catch
            {
                try { return Convert.ToBoolean(value); } catch { return false; }
            }
        }

        // Helper to get a property value safely from a dynamic COM object (avoids exceptions)
        private object GetMember(dynamic obj, string memberName)
        {
            try
            {
                var t = obj.GetType();
                var prop = t.GetProperty(memberName);
                if (prop != null) return prop.GetValue(obj, null);
                // fallback: invoke as dynamic (some interop members throw on direct reflection)
                try { return obj.GetType().InvokeMember(memberName, System.Reflection.BindingFlags.GetProperty, null, obj, null); } catch { }
                return null;
            }
            catch { return null; }
        }

        // Helper to apply font settings (keeps COM exceptions local)
        private void ApplyFont(dynamic font, string name, float size, Color color)
        {
            if (font == null) return;
            try { font.Name = name; } catch { }
            try { font.Size = size; } catch { }
            try { font.Fill.ForeColor.RGB = ColorTranslator.ToOle(color); } catch { }
            try { font.Bold = Office.MsoTriState.msoFalse; } catch { }
        }

        private void FormatSeriesElements(Excel.Chart chart)
        {
            foreach (Excel.Series series in chart.SeriesCollection())
            {
                try
                {
                    bool hasDataLabels = false;
                    try { hasDataLabels = series.HasDataLabels; } catch { }

                    if (hasDataLabels)
                    {
                        var dataLabels = series.DataLabels(Type.Missing);
                        var font = dataLabels.Format.TextFrame2.TextRange.Font;
                        font.Name = "Helvetica Now Text";
                        font.Size = 10;
                        font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(93, 109, 120));
                        font.Bold = Office.MsoTriState.msoFalse;
                    }

                    bool hasErrorBars = false;
                    try { hasErrorBars = series.HasErrorBars; } catch { }

                    if (hasErrorBars && series.ErrorBars != null)
                    {
                        var errorBars = series.ErrorBars;
                        var lineColor = errorBars.Format.Line.ForeColor;
                        lineColor.RGB = ColorTranslator.ToOle(Color.FromArgb(93, 109, 120));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error formatting series: {ex.Message}");
                }
            }
        }

        public void ChartTitle(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;

                if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart)
                        {
                            PowerPoint.Shape chartShape = shape;
                            PowerPoint.Chart chart = chartShape.Chart;

                            float topPosition = chartShape.Top;
                            float leftPosition = chartShape.Left;
                            float width = chartShape.Width;

                            string chartTitle = "";
                            try
                            {
                                if (chart.HasTitle)
                                {
                                    chartTitle = chart.ChartTitle.Text;
                                }
                            }
                            catch
                            {
                                chartTitle = "";
                            }

                            // Add textbox above the chart
                            PowerPoint.Shape textBox = chartShape.Parent.Shapes.AddTextbox(
                                Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                leftPosition,
                                topPosition - 20,
                                width,
                                20
                            );

                            textBox.TextFrame.TextRange.Text = string.IsNullOrEmpty(chartTitle)
                                ? "[Please enter chart title]"
                                : chartTitle;

                            // Set textbox formatting
                            textBox.TextFrame.MarginTop = 0;
                            textBox.TextFrame.MarginBottom = 0;
                            textBox.TextFrame.MarginLeft = 0;
                            textBox.TextFrame.MarginRight = 0;

                            var font = textBox.TextFrame.TextRange.Font;
                            font.Name = "Helvetica Now Text";
                            font.Size = 14;
                            font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                            font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                            // Remove chart's built-in title if it exists
                            if (!string.IsNullOrEmpty(chartTitle))
                            {
                                chart.HasTitle = false;
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Please select at least one chart.", "Chart Title Macro",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating chart title:\n{ex.Message}", "Chart Title Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void GraphText(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select at least one chart.", "Graph Text",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var form = new FontSize())
                {
                    if (form.ShowDialog() != DialogResult.OK)
                        return;

                    int fontSize = form.SelectedFontSize;

                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            var chart = shape.Chart;
                            if (chart == null)
                                continue;

                            // Chart area
                            var chartFont = chart.ChartArea.Format.TextFrame2.TextRange.Font;
                            chartFont.Name = "Helvetica Now Text";
                            chartFont.Size = fontSize;
                            chartFont.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(93, 109, 120));
                            chartFont.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;

                            // Chart title
                            if (chart.HasTitle)
                            {
                                var titleFont = chart.ChartTitle.Format.TextFrame2.TextRange.Font;
                                titleFont.Name = "Helvetica Now Text";
                                titleFont.Size = fontSize;
                                titleFont.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(93, 109, 120));
                                titleFont.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;
                            }

                            // Legend
                            if (chart.HasLegend)
                            {
                                var legendFont = chart.Legend.Format.TextFrame2.TextRange.Font;
                                legendFont.Name = "Helvetica Now Text";
                                legendFont.Size = fontSize;
                                legendFont.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(93, 109, 120));
                                legendFont.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;
                            }

                            // Data labels for each series (safe access)
                            dynamic seriesCollection = chart.SeriesCollection();
                            int seriesCount = seriesCollection.Count;

                            for (int i = 1; i <= seriesCount; i++)
                            {
                                dynamic series = seriesCollection.Item(i);
                                try
                                {
                                    var dataLabels = series.DataLabels();
                                    if (dataLabels != null)
                                    {
                                        var labelFont = dataLabels.Format.TextFrame2.TextRange.Font;
                                        labelFont.Name = "Helvetica Now Text";
                                        labelFont.Size = fontSize;
                                        labelFont.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(93, 109, 120));
                                        labelFont.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;
                                    }
                                }
                                catch { }
                            }

                            // Data table
                            if (chart.HasDataTable)
                            {
                                var dataTableFont = chart.DataTable.Format.TextFrame2.TextRange.Font;
                                dataTableFont.Name = "Helvetica Now Text";
                                dataTableFont.Size = fontSize;
                                dataTableFont.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(93, 109, 120));
                                dataTableFont.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying graph text formatting:\n{ex.Message}", "Graph Text Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ColourSeries(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select at least one chart.", "Colour Series",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Define color hex list
                string[] colors = { 
                    "73E2D8", "0084BB", "101E7F", "A70070", "EA2238", "FFA600", "12A88A", "29B0C3", "0055A8",
                    "6E027F", "000000", "8ABD45", "D14900", "82939A", "73E2D8", "ACC0C4", "5D6D78", "ACC0C4" 
                };

                // Excel chart type constants
                const int xlLine = 4;
                const int xlLineMarkers = 65;
                const int xlXYScatterLines = 74;
                const int xlXYScatterLinesNoMarkers = 75;
                const int xlBarClustered = 57;
                const int xlColumnClustered = 51;
                const int xlWaterfall = 119;
                const int xlBarStacked100 = 59;
                const int xlColumnStacked = 52;
                const int xlColumnStacked100 = 53;
                const int xlPie = 5;
                const int xlDoughnut = 80;
                const int xlArea = 1;
                const int xlBubble = 15;
                const int xlStock = 73;
                const int xlSurface = 83;

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        dynamic chart = shape.Chart;
                        dynamic seriesCollection = chart.SeriesCollection();

                        int seriesCount = seriesCollection.Count;

                        if (seriesCount > 1)
                        {
                            // Multiple series: color by series
                            for (int i = 1; i <= seriesCount; i++)
                            {
                                dynamic series = seriesCollection.Item(i);
                                int colorIndex = (i - 1) % colors.Length;
                                int rgb = System.Drawing.ColorTranslator.ToOle(HexToColor(colors[colorIndex]));
                                int chartType = (int)series.ChartType;

                                if (chartType == xlLine || chartType == xlLineMarkers ||
                                    chartType == xlXYScatterLines || chartType == xlXYScatterLinesNoMarkers)
                                {
                                    // Line charts
                                    series.Format.Line.ForeColor.RGB = rgb;
                                    series.Format.Line.Weight = 1.5f;
                                    series.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid;
                                    series.Format.Line.Transparency = 0f;

                                    try
                                    {
                                        if ((int)series.MarkerStyle != -4142) 
                                        {
                                            series.MarkerForegroundColor = rgb;
                                            series.MarkerBackgroundColor = rgb;
                                        }
                                    }
                                    catch { }
                                }
                                else if (chartType == xlBarClustered || chartType == xlColumnClustered ||
                                         chartType == xlWaterfall || chartType == xlBarStacked100 ||
                                         chartType == xlColumnStacked || chartType == xlColumnStacked100)
                                {
                                    // Bar/column charts
                                    series.Format.Fill.ForeColor.RGB = rgb;
                                    series.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                                }
                                else if (chartType == xlPie || chartType == xlDoughnut)
                                {
                                    // Pie or doughnut charts — color each slice
                                    int pointCount = series.Points().Count;
                                    for (int j = 1; j <= pointCount; j++)
                                    {
                                        dynamic point = series.Points(j);
                                        int pointColorIndex = (j - 1) % colors.Length;
                                        int pointRgb = System.Drawing.ColorTranslator.ToOle(HexToColor(colors[pointColorIndex]));
                                        point.Format.Fill.ForeColor.RGB = pointRgb;
                                        point.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                                    }
                                }
                                else if (chartType == xlArea || chartType == xlBubble ||
                                         chartType == xlStock || chartType == xlSurface)
                                {
                                    // Area, bubble, stock, surface charts
                                    series.Format.Fill.ForeColor.RGB = rgb;
                                }
                                else
                                {
                                    // Default for any other type
                                    series.Format.Fill.ForeColor.RGB = rgb;
                                }
                            }
                        }
                        else if (seriesCount == 1)
                        {
                            // Single series: color by points
                            dynamic series = seriesCollection.Item(1);
                            int pointCount = series.Points().Count;
                            for (int j = 1; j <= pointCount; j++)
                            {
                                dynamic point = series.Points(j);
                                int pointColorIndex = (j - 1) % colors.Length;
                                int pointRgb = System.Drawing.ColorTranslator.ToOle(HexToColor(colors[pointColorIndex]));
                                point.Format.Fill.ForeColor.RGB = pointRgb;
                                point.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying series colors:\n{ex.Message}", "Colour Series Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper: Convert hex string (e.g., "73E2D8") to Color
        private System.Drawing.Color HexToColor(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return System.Drawing.Color.Black;
            if (!hex.StartsWith("#")) hex = "#" + hex;
            return System.Drawing.ColorTranslator.FromHtml(hex);
        }

        public void FindEmbededGraph(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var presentation = app.ActivePresentation;

                if (presentation == null)
                {
                    MessageBox.Show("No active presentation found.", "Embedded Graphs", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool foundGraphs = false;
                StringBuilder message = new StringBuilder("Embedded Graphs Found:\n\n");

                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    int slideNumber = slide.SlideIndex;

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        try
                        {
                            if (shape.HasChart == Office.MsoTriState.msoTrue)
                            {
                                var chart = shape.Chart;

                                if (!chart.ChartData.IsLinked)
                                {
                                    foundGraphs = true;

                                    // Use Try-Finally to ensure Excel instance closes properly
                                    Excel.Workbook workbook = null;
                                    try
                                    {
                                        chart.ChartData.Activate();
                                        workbook = chart.ChartData.Workbook;
                                        string workbookName = workbook.Name;
                                        message.AppendLine($"Slide {slideNumber}: {workbookName}");
                                    }
                                    finally
                                    {
                                        // Ensure workbook and Excel instance are closed
                                        if (workbook != null)
                                        {
                                            workbook.Close(false);
                                            Marshal.ReleaseComObject(workbook);
                                            workbook = null;
                                        }

                                        // Clean up Excel process to allow next chart
                                        GC.Collect();
                                        GC.WaitForPendingFinalizers();
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            continue; 
                        }
                    }
                }

                if (foundGraphs)
                    MessageBox.Show(message.ToString(), "Embedded Graphs", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show("No embedded graphs found.", "Embedded Graphs", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while searching for graphs:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UnembedGraph(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null || sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a single embedded graph.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                PowerPoint.Shape shape = sel.ShapeRange[1];

                if (shape.HasChart != Office.MsoTriState.msoTrue)
                {
                    MessageBox.Show("The selected shape is not a chart.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                var chart = shape.Chart;

                float chartWidth = shape.Width;
                float chartHeight = shape.Height;
                float chartLeft = shape.Left;
                float chartTop = shape.Top;

                shape.Copy();

                chart.ChartData.Activate();
                dynamic workbook = chart.ChartData.Workbook;
                dynamic worksheet = workbook.Sheets(1);

                worksheet.Paste(worksheet.Range["A1"]);
                dynamic embeddedChart = worksheet.ChartObjects(worksheet.ChartObjects().Count);

                embeddedChart.Chart.ChartArea.Copy();

                PowerPoint.Shape newShape = app.ActiveWindow.View.Slide.Shapes.Paste()[1];
                newShape.Left = chartLeft;
                newShape.Top = chartTop;
                newShape.Width = chartWidth;
                newShape.Height = chartHeight;

                workbook.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                shape.Delete();

                MessageBox.Show("Embedded graph unembedded successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while unembedding graph:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CopyLink(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.ActiveWindow.Selection;

                if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a linked object to copy its link.", "Copy Link",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                foreach (PowerPoint.Shape shp in selection.ShapeRange)
                {
                    if (shp.Type == Office.MsoShapeType.msoLinkedOLEObject || shp.Type == Office.MsoShapeType.msoChart)
                    {
                        if (shp.LinkFormat == null)
                        {
                            MessageBox.Show("The selected object is not linked.", "Copy Link",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }

                        string linkPath = shp.LinkFormat.SourceFullName;

                        if (!string.IsNullOrEmpty(linkPath))
                        {
                            // Pass link to dialog
                            CopyLink form = new CopyLink(linkPath);
                            form.ShowDialog();
                            return;
                        }
                        else
                        {
                            MessageBox.Show("The selected object has no valid link.", "Copy Link",
                                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("The selected object is not linked.", "Copy Link",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("The selected object has no valid link.", "Copy Link",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //public void CopyLink(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        var app = Globals.ThisAddIn.Application;
        //        var selection = app.ActiveWindow.Selection;

        //        // Ensure a shape is selected
        //        if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
        //        {
        //            MessageBox.Show("Please select a linked object to copy its link.", "Copy Link",
        //                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //            return;
        //        }

        //        foreach (PowerPoint.Shape shape in selection.ShapeRange)
        //        {
        //            try
        //            {
        //                // Only attempt to access LinkFormat for linked objects
        //                if (shape.Type == Office.MsoShapeType.msoLinkedOLEObject)
        //                {
        //                    var linkFormat = shape.LinkFormat;

        //                    if (linkFormat == null)
        //                    {
        //                        MessageBox.Show("The selected object has no valid link.", "Copy Link",
        //                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                        return;
        //                    }

        //                    string linkPath = linkFormat.SourceFullName;

        //                    if (!string.IsNullOrEmpty(linkPath))
        //                    {
        //                        Clipboard.SetText(linkPath);
        //                        MessageBox.Show($"Linked file path copied to clipboard:\n\n{linkPath}", "Copy Link",
        //                            MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("The selected object has no valid link.", "Copy Link",
        //                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                    }

        //                    return;
        //                }
        //                else if (shape.Type == Office.MsoShapeType.msoChart)
        //                {
        //                    // Charts in PowerPoint are usually embedded, not linked
        //                    MessageBox.Show("The selected object has no valid link.", "Copy Link",
        //                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                    return;
        //                }
        //            }
        //            catch (System.Runtime.InteropServices.COMException)
        //            {
        //                // Handle PowerPoint COM "Invalid request" error safely
        //                MessageBox.Show("The selected object has no valid link.", "Copy Link",
        //                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //                return;
        //            }
        //        }

        //        MessageBox.Show("Please select a linked chart or OLE object.", "Copy Link",
        //            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }
        //    catch (Exception)
        //    {
        //        // Catch any unexpected errors gracefully
        //        MessageBox.Show("The selected object has no valid link.", "Copy Link",
        //            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }
        //}

        public void WidescreenGuide(Office.IRibbonControl control)
        {
            try
            {
                string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/Graph-guides_Widescreen.pptx";
                string tempPath = Path.GetTempPath();
                string baseFileName = Path.Combine(tempPath, "Graph-guides_Widescreen");
                string filePath = "";
                int i = 1;

                // Generate a unique filename
                do
                {
                    filePath = $"{baseFileName}_{i}.pptx";
                    i++;
                } while (File.Exists(filePath));

                // Download the file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(url, filePath);
                }

                // Open the presentation in PowerPoint
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.Presentations.Open(filePath,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (WebException)
            {
                MessageBox.Show("Unable to download the file. Please check your internet connection or the URL.",
                    "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening widescreen guide:\n{ex.Message}",
                    "Widescreen Guide Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void A4Guide(Office.IRibbonControl control)
        {
            try
            {
                string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/Graph-guides_A4.pptx";
                string tempPath = Path.GetTempPath();
                string baseFileName = Path.Combine(tempPath, "Graph-guides_A4");
                string filePath = "";
                int i = 1;

                // Generate a unique filename
                do
                {
                    filePath = $"{baseFileName}_{i}.pptx";
                    i++;
                } while (File.Exists(filePath));

                // Download the file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(url, filePath);
                }

                // Open the presentation in PowerPoint
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.Presentations.Open(filePath,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (WebException)
            {
                MessageBox.Show("Unable to download the file. Please check your internet connection or the URL.",
                    "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening A4 guide:\n{ex.Message}",
                    "A4 Guide Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void USGuide(Office.IRibbonControl control)
        {
            try
            {
                string url = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/Graph-guides_US-Letter.pptx";
                string tempPath = Path.GetTempPath();
                string baseFileName = Path.Combine(tempPath, "Graph-guides_US-Letter");
                string filePath = "";
                int i = 1;

                // Generate a unique filename
                do
                {
                    filePath = $"{baseFileName}_{i}.pptx";
                    i++;
                } while (File.Exists(filePath));

                // Download the file
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(url, filePath);
                }

                // Open the presentation in PowerPoint
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.Presentations.Open(filePath,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (WebException)
            {
                MessageBox.Show("Unable to download the file. Please check your internet connection or the URL.",
                    "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening US Letter guide:\n{ex.Message}",
                    "US Guide Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void InsertNote(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Dimensions 
                float width = 136.06f; 
                float height = 68.04f;  
                float leftPos = presentation.PageSetup.SlideWidth - width;  
                float topPos = 0f;  

                // Add text box
                PowerPoint.Shape noteShape = slide.Shapes.AddTextbox(
                    Orientation: Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    Left: leftPos,
                    Top: topPos,
                    Width: width,
                    Height: height
                );

                // Set text and format
                var textRange = noteShape.TextFrame.TextRange;
                textRange.Text = "[Insert Note]";
                textRange.Font.Name = "Helvetica Now Text";
                textRange.Font.Size = 12;
                textRange.Font.Bold = Office.MsoTriState.msoTrue;
                textRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                // Text frame alignment and margins
                var textFrame = noteShape.TextFrame;
                textFrame.HorizontalAnchor = Office.MsoHorizontalAnchor.msoAnchorCenter;
                textFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                textFrame.MarginBottom = 5.67f;
                textFrame.MarginTop = 5.67f;
                textFrame.MarginLeft = 5.67f;
                textFrame.MarginRight = 5.67f;
                textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                // Box appearance
                noteShape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(235, 0, 23)); // Red
                noteShape.Line.Visible = Office.MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting note:\n{ex.Message}",
                    "Insert Note Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SetUKLanguage(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;
                pres.DefaultLanguageID = Office.MsoLanguageID.msoLanguageIDEnglishUK;

                foreach (PowerPoint.Slide sld in pres.Slides)
                {
                    foreach (PowerPoint.Shape shp in sld.Shapes)
                    {
                        ApplyLanguageToShape(shp, Office.MsoLanguageID.msoLanguageIDEnglishUK);
                    }
                }

                MessageBox.Show("All text set to English (UK).", "Language Applied",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying UK language:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SetUSLanguage(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;
                pres.DefaultLanguageID = Office.MsoLanguageID.msoLanguageIDEnglishUS;

                foreach (PowerPoint.Slide sld in pres.Slides)
                {
                    foreach (PowerPoint.Shape shp in sld.Shapes)
                    {
                        ApplyLanguageToShape(shp, Office.MsoLanguageID.msoLanguageIDEnglishUS);
                    }
                }

                MessageBox.Show("All text set to English (US).", "Language Applied",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error applying US language:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyLanguageToShape(PowerPoint.Shape shp, Office.MsoLanguageID langID)
        {
            try
            {
                if (shp.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (PowerPoint.Shape subShp in shp.GroupItems)
                    {
                        ApplyLanguageToShape(subShp, langID);
                    }
                }
                else if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    if (shp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        shp.TextFrame.TextRange.LanguageID = langID;
                }
                else if (shp.HasTable == Office.MsoTriState.msoTrue)
                {
                    PowerPoint.Table tbl = shp.Table;
                    for (int i = 1; i <= tbl.Rows.Count; i++)
                    {
                        for (int j = 1; j <= tbl.Columns.Count; j++)
                        {
                            tbl.Cell(i, j).Shape.TextFrame.TextRange.LanguageID = langID;
                        }
                    }
                }
                else if (shp.HasChart == Office.MsoTriState.msoTrue)
                {
                    var chart = shp.Chart;
                    try
                    {
                        if (chart.HasTitle)
                            chart.ChartTitle.Format.TextFrame2.TextRange.LanguageID = langID;

                        if (chart.Axes(Excel.XlAxisType.xlCategory).HasTitle)
                            chart.Axes(Excel.XlAxisType.xlCategory).AxisTitle.Format.TextFrame2.TextRange.LanguageID = langID;

                        if (chart.Axes(Excel.XlAxisType.xlValue).HasTitle)
                            chart.Axes(Excel.XlAxisType.xlValue).AxisTitle.Format.TextFrame2.TextRange.LanguageID = langID;
                    }
                    catch { }
                }
            }
            catch { }
        }

        public void SetDocumentTitle(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;

                if (pres == null)
                {
                    MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fileName = pres.Name;
                int dotIndex = fileName.LastIndexOf(".");
                if (dotIndex > 0)
                    fileName = fileName.Substring(0, dotIndex);

                pres.BuiltInDocumentProperties["Title"].Value = fileName;
                pres.BuiltInDocumentProperties["Author"].Value = "";

                MessageBox.Show($"Document title set to: {fileName}", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting document title:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void DeleteComments(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;

                if (pres == null)
                {
                    MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int deletedCount = 0;

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    try
                    {
                        while (slide.Comments.Count > 0)
                        {
                            slide.Comments[1].Delete();
                            deletedCount++;
                        }
                    }
                    catch { }
                }

                MessageBox.Show(deletedCount > 0
                    ? $"{deletedCount} comments deleted successfully."
                    : "No comments found in the presentation.",
                    "Delete Comments", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting comments:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveSlideNotes(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation pres = app.ActivePresentation;

                if (pres == null)
                {
                    MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int clearedCount = 0;

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    try
                    {
                        PowerPoint.Shape notesShape = slide.NotesPage.Shapes.Placeholders[2];
                        if (notesShape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            notesShape.TextFrame.TextRange.Text = "";
                            clearedCount++;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error removing notes:\n{ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper cleanup method
        private void TryCleanup(string tempZipPath, string tempExtractFolder)
        {
            try
            {
                if (File.Exists(tempZipPath)) File.Delete(tempZipPath);
            }
            catch { }

            try
            {
                if (Directory.Exists(tempExtractFolder)) Directory.Delete(tempExtractFolder, true);
            }
            catch { }
        }

        public void FindImages(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;

                string originalFilePath = presentation.FullName;
                string tempFolderPath = Environment.GetEnvironmentVariable("TEMP");

                string tempPptxPath = Path.Combine(tempFolderPath, "temp_presentation.pptx");
                string tempZipPath = Path.Combine(tempFolderPath, "temp_presentation.zip");
                string tempExtractFolder = Path.Combine(tempFolderPath, "temp_presentation_extracted");
                string mediaFolderPath = Path.Combine(tempExtractFolder, @"ppt\media");

                // Clean old temporary files
                try
                {
                    if (File.Exists(tempPptxPath)) File.Delete(tempPptxPath);
                    if (File.Exists(tempZipPath)) File.Delete(tempZipPath);
                    if (Directory.Exists(tempExtractFolder))
                        Directory.Delete(tempExtractFolder, true);
                }
                catch { }

                // Save a copy of the presentation
                presentation.SaveCopyAs(tempPptxPath);

                // Rename .pptx to .zip
                File.Move(tempPptxPath, tempZipPath);

                // Create extract folder
                Directory.CreateDirectory(tempExtractFolder);

                // Extract using PowerShell 
                string psExtract = $"Expand-Archive -Path '{tempZipPath}' -DestinationPath '{tempExtractFolder}' -Force";
                Process.Start(new ProcessStartInfo("powershell.exe", $"-Command \"{psExtract}\"")
                {
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });

                // Wait up to 5 seconds for extraction
                int timeout = 0;
                while (!Directory.Exists(mediaFolderPath) && timeout < 5000)
                {
                    Thread.Sleep(200);
                    timeout += 200;
                }

                // Empty media folder (if found)
                if (Directory.Exists(mediaFolderPath))
                {
                    foreach (var file in Directory.GetFiles(mediaFolderPath))
                    {
                        try { File.Delete(file); } catch { }
                    }

                    // Open File Explorer
                    Process.Start("explorer.exe", mediaFolderPath);

                    // Apply details view and sort by size
                    string psSort =
                        $"$folderPath = '{mediaFolderPath.Replace("\\", "\\\\")}'; " +
                        "$shell = New-Object -ComObject Shell.Application; " +
                        "$folder = $shell.Namespace($folderPath); " +
                        "$items = $folder.Items(); " +
                        "$items.Sort(2, -1);";

                    Process.Start(new ProcessStartInfo("powershell.exe", "-Command \"" + psSort + "\"")
                    {
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden
                    });
                }
                else
                {
                    MessageBox.Show("No media folder found in the presentation. Please check the file.",
                        "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Find Images",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ShowExcel(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                PowerPoint.Presentation presentation = app.ActivePresentation;

                string originalFilePath = presentation.FullName;
                string tempFolderPath = Environment.GetEnvironmentVariable("TEMP");

                string tempPptxPath = Path.Combine(tempFolderPath, "temp_presentation.pptx");
                string tempZipPath = Path.Combine(tempFolderPath, "temp_presentation.zip");
                string tempExtractFolder = Path.Combine(tempFolderPath, "temp_presentation_extracted");
                string embeddingsFolderPath = Path.Combine(tempExtractFolder, @"ppt\embeddings");

                // Remove old temporary files/folders
                try
                {
                    if (File.Exists(tempPptxPath)) File.Delete(tempPptxPath);
                    if (File.Exists(tempZipPath)) File.Delete(tempZipPath);
                    if (Directory.Exists(tempExtractFolder))
                        Directory.Delete(tempExtractFolder, true);
                }
                catch { }

                // Save a copy of the file as .pptx
                presentation.SaveCopyAs(tempPptxPath);

                // Rename .pptx to .zip
                File.Move(tempPptxPath, tempZipPath);

                // Create extraction folder
                Directory.CreateDirectory(tempExtractFolder);

                // Extract .zip contents using PowerShell
                string psExtract = $"Expand-Archive -Path '{tempZipPath}' -DestinationPath '{tempExtractFolder}' -Force";
                Process.Start(new ProcessStartInfo("powershell.exe", $"-Command \"{psExtract}\"")
                {
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });

                // Wait up to 5 seconds for extraction
                int elapsed = 0;
                while (!Directory.Exists(embeddingsFolderPath) && elapsed < 5000)
                {
                    Thread.Sleep(200);
                    elapsed += 200;
                }

                // Delete existing files inside embeddings folder
                if (Directory.Exists(embeddingsFolderPath))
                {
                    foreach (var file in Directory.GetFiles(embeddingsFolderPath))
                    {
                        try { File.Delete(file); } catch { }
                    }
                }

                // Open folder in File Explorer
                if (Directory.Exists(embeddingsFolderPath))
                {
                    Process.Start("explorer.exe", embeddingsFolderPath);
                }
                else
                {
                    MessageBox.Show("No embedded Excel files found in the presentation.",
                        "Show Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Show Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Colours(Office.IRibbonControl control)
        {
            try
            {
                // Force TLS 1.2/1.3 for modern HTTPS
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                string strURL = "https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/Colours-%e2%80%93-Hex-s.pptx";
                string tempFolder = Environment.GetEnvironmentVariable("TEMP");
                string basePath = Path.Combine(tempFolder, "Graph-guides_Widescreen");
                string tempFile;

                int i = 1;
                do
                {
                    tempFile = $"{basePath}_{i}.pptx";
                    i++;
                } while (File.Exists(tempFile));

                // Download the PPTX file
                using (var client = new System.Net.WebClient())
                {
                    client.DownloadFile(strURL, tempFile);
                }

                // Open the downloaded file in PowerPoint
                var app = Globals.ThisAddIn.Application;
                app.Presentations.Open(tempFile);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Failed to open Colours presentation:\n" + ex.Message,
                    "Download Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void BreakTable(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                // Validate selection
                if (sel == null ||
                    (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     sel.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show("Please select a table first.", "Highlight", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                PowerPoint.ShapeRange selectedShapes = app.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape shape in selectedShapes)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        PowerPoint.Table table = shape.Table;

                        for (int i = 1; i <= table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= table.Columns.Count; j++)
                            {
                                PowerPoint.Cell cell = table.Cell(i, j);

                                var textBox = shape.Parent.Shapes.AddTextbox(
                                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                    cell.Shape.Left,
                                    cell.Shape.Top,
                                    cell.Shape.Width,
                                    cell.Shape.Height);

                                textBox.TextFrame.TextRange.Text = cell.Shape.TextFrame.TextRange.Text;
                            }
                        }

                        shape.Delete(); 
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while breaking the table:\n" + ex.Message,
                    "Break Table Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void Bracket(Office.IRibbonControl control)
        {
            InsertShapeFromSlide("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", 4, "Bracket");
        }

        public void Chevron(Office.IRibbonControl control)
        {
            InsertShapeFromSlide("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", 5, "Chevron");
        }

        public void BigStat(Office.IRibbonControl control)
        {
            InsertShapeFromSlide("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", 6, "BigStat");
        }

        public void LitStat(Office.IRibbonControl control)
        {
            InsertShapeFromSlide("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", 7, "LitStat");
        }

        public void Quote(Office.IRibbonControl control)
        {
            InsertShapeFromSlide("https://workingtogether.aon.com/WorkingTogether/media/worktog/Reinsurance/PPT-elements.pptx", 8, "Quote");
        }

        private void InsertShapeFromSlide(string fileUrl, int slideNumber, string macroName)
        {
            // Force TLS 1.2/1.3 for modern HTTPS
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation destPresentation = null;
            PowerPoint.Presentation sourcePresentation = null;
            PowerPoint.Slide destSlide = null;
            PowerPoint.Slide sourceSlide = null;

            string tempFolder = Environment.GetEnvironmentVariable("TEMP");
            string tempFilePath = Path.Combine(tempFolder, "PPT-elements.pptx");

            try
            {
                destPresentation = app.ActivePresentation;
                destSlide = app.ActiveWindow.View.Slide;

                // Download PPTX file
                using (var client = new System.Net.WebClient())
                {
                    client.DownloadFile(fileUrl, tempFilePath);
                }

                // Open source presentation (hidden)
                sourcePresentation = app.Presentations.Open(tempFilePath,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);

                // Get the specified slide
                sourceSlide = sourcePresentation.Slides[slideNumber];

                bool shapeFound = false;
                PowerPoint.Shape pastedShape = null;

                // Copy first shape
                foreach (PowerPoint.Shape shape in sourceSlide.Shapes)
                {
                    shape.Copy();
                    shapeFound = true;
                    break;
                }

                if (!shapeFound)
                {
                    MessageBox.Show($"No shapes found on slide {slideNumber}.",
                        macroName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    sourcePresentation.Close();
                    return;
                }

                // Paste into destination slide
                pastedShape = destSlide.Shapes.Paste()[1];

                // Center the shape
                pastedShape.Left = (float)((destSlide.Parent.PageSetup.SlideWidth - pastedShape.Width) / 2);
                pastedShape.Top = (float)((destSlide.Parent.PageSetup.SlideHeight - pastedShape.Height) / 2);

                // Close source presentation
                sourcePresentation.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred in {macroName}:\n{ex.Message}",
                    macroName + " Macro Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void Red(Office.IRibbonControl control)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide sld = app.ActivePresentation.Slides[app.ActiveWindow.View.Slide.SlideIndex];

            // Convert 1.6 cm to points
            float widthPts = (float)(1.6 * 28.35);
            float heightPts = (float)(1.6 * 28.35);

            // Add a circle (oval with equal width and height)
            PowerPoint.Shape shp = sld.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                0, 0, widthPts, heightPts
            );

            // Set fill color to RGB(235, 0, 23)
            shp.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(235, 0, 23));
            shp.Fill.Solid();

            // Remove outline
            shp.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        public void Amber(Office.IRibbonControl control)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide sld = app.ActivePresentation.Slides[app.ActiveWindow.View.Slide.SlideIndex];

            // Convert 1.6 cm to points
            float widthPts = (float)(1.6 * 28.35);
            float heightPts = (float)(1.6 * 28.35);

            // Add a circle (oval with equal width and height)
            PowerPoint.Shape shp = sld.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                0, 0, widthPts, heightPts
            );

            // Set fill color to RGB(255, 166, 0)
            shp.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 166, 0));
            shp.Fill.Solid();

            // Remove outline
            shp.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        public void Green(Office.IRibbonControl control)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide sld = app.ActivePresentation.Slides[app.ActiveWindow.View.Slide.SlideIndex];

            // Convert 1.6 cm to points
            float widthPts = (float)(1.6 * 28.35);
            float heightPts = (float)(1.6 * 28.35);

            // Add a circle (oval with equal width and height)
            PowerPoint.Shape shp = sld.Shapes.AddShape(
                Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                0, 0, widthPts, heightPts
            );

            // Set fill color to RGB(138, 189, 69)
            shp.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(138, 189, 69));
            shp.Fill.Solid();

            // Remove outline
            shp.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        // Convert System.Drawing.Image -> stdole.IPictureDisp
        private class PictureConverter : System.Windows.Forms.AxHost
        {
            private PictureConverter() : base("dummy") { }

            public static stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }
    }
}