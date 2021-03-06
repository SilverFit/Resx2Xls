namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Windows.Forms;
    using AppSettings = Resx2Xls.Properties.Settings;
    using Excel = Microsoft.Office.Interop.Excel;
    using ResX = Resx2Xls.Properties.Resources;

    public partial class Resx2XlsForm : Form
    {
        public Resx2XlsForm()
        {
            CultureInfo ci = CultureInfo.GetCultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            System.Threading.Thread.CurrentThread.CurrentUICulture = ci;

            InitializeComponent();

            this.textBoxFolder.Text = AppSettings.Default.FolderPath;
            this.textBoxScreenshots.Text = AppSettings.Default.ScreenshotPath;
            this.textBox_ExcludeKey.Text = StringHelper.FromCollection(AppSettings.Default.ExcludeKeys);
            this.textBox_ExcludeComment.Text = StringHelper.FromCollection(AppSettings.Default.ExcludeComments);
            this.textBox_ExcludeFilename.Text = StringHelper.FromCollection(AppSettings.Default.ExcludeFilenames);
            this.hideCommentColumnCheckbox.Checked = AppSettings.Default.HideComments;
            this.hideKeyColumnCheckbox.Checked = AppSettings.Default.HideKeys;
            this.purgeTranslation_CheckBox.Checked = AppSettings.Default.PurgeNonexistant;
            this.checkBoxSubFolders.Checked = AppSettings.Default.ScanSubfolders;

            this.FillCultures();
        }

        private enum Conversion
        {
            ResXToXls,
            XlsToResx,
        };

        private enum WizardStep
        {
            Start = 0,
            ResX1 = 1,
            ResX2 = 2,
            ResX3 = 3,
            Xls1 = 4,
            Finish = 5,
        }

        private Conversion ConversionType
        {
            get
            {
                if (this.radioButtonCreateXls.Checked)
                {
                    return Conversion.ResXToXls;
                }
                else if (this.radioButtonGenerateResx.Checked)
                {
                    return Conversion.XlsToResx;
                }
                else
                {
                    throw new InvalidOperationException("Unknown conversion type");
                }
            }
        }

        public void ResxToXls(string path, string screenshotPath, bool deepSearch, bool purge, string outputPath,
                              List<CultureInfo> cultures, List<string> excludeKeyList, List<string> excludeCommentList)
        {
            if (!Directory.Exists(path))
                return;

            Cursor = Cursors.WaitCursor;

            this.AddSummaryLine();
            this.AddSummaryLine(ResX.parsing_resx);
            var resxdata = ResxData.FromResx(path, deepSearch, purge, cultures, excludeKeyList, excludeCommentList);
            resxdata.ToXls(outputPath, screenshotPath, this.AddSummaryLine);
            this.ShowXls(outputPath);
            
            Cursor = Cursors.Default;
        }

        private void AddSummaryLine(string text = "")
        {
            this.textBoxSummary.Text += text + Environment.NewLine;
            this.textBoxSummary.SelectionStart = this.textBoxSummary.Text.Length;
            this.textBoxSummary.ScrollToCaret();

            Application.DoEvents();
        }

        private void XlsToResx(string xlsFile)
        {
            if (!File.Exists(xlsFile))
                return;

            Cursor = Cursors.WaitCursor;

            string path = new FileInfo(xlsFile).DirectoryName;

            this.AddSummaryLine();
            this.AddSummaryLine(ResX.parsing_excel);
            var rd = ResxData.FromXls(xlsFile);
            rd.ToResx(path, this.AddSummaryLine);

            Cursor = Cursors.Default;
        }

        private void FillCultures()
        {
            var cultures = CultureInfo.GetCultures(CultureTypes.SpecificCultures)
                                      .OrderBy(c => c.EnglishName)
                                      .ToArray();
            
            this.listBoxCultures.Items.AddRange(cultures);

            var selectedCultures = AppSettings.Default.CultureList.Split(';')
                                                                  .Select(n => CultureInfo.GetCultureInfo(n))
                                                                  .ToArray();

            this.listBoxCulturesSelected.Items.AddRange(selectedCultures);
        }

        private void SelectCultures()
        {
            var cultures = this.listBoxCultures.SelectedItems.Cast<CultureInfo>().ToArray();
            foreach (var culture in cultures)
            {
                if (this.listBoxCulturesSelected.Items.IndexOf(culture) == -1)
                {
                    this.listBoxCulturesSelected.Items.Add(culture);
                }
            }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFolder.Text = this.folderBrowserDialog.SelectedPath;
            }
        }

        private void browseButtonScreenshots_Click(object sender, EventArgs e)
        {
            if (this.screenshotFolderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBoxScreenshots.Text = this.screenshotFolderBrowserDialog.SelectedPath;
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            SelectCultures();
        }

        private void listBoxCultures_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            SelectCultures();
        }

        private void buttonBrowseXls_Click(object sender, EventArgs e)
        {
            if (this.openFileDialogXls.ShowDialog() == DialogResult.OK)
            {
                this.textBoxXls.Text = this.openFileDialogXls.FileName;
            }
        }

        private void listBoxSelected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.listBoxCulturesSelected.SelectedItems.Count > 0)
            {
                this.listBoxCulturesSelected.Items.Remove(this.listBoxCulturesSelected.SelectedItems[0]);
            }
        }

        public void ShowXls(string path)
        {
            if (!File.Exists(path))
                return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(
                Filename: path,
                Origin: Excel.XlPlatform.xlWindows, 
                Editable: true,
                AddToMru: true);

            app.Visible = true;
        }

        private void FinishWizard()
        {
            List<string> excludeKeyList = StringHelper.ListFromCollection(AppSettings.Default.ExcludeKeys);
            List<string> excludeCommentList = StringHelper.ListFromCollection(AppSettings.Default.ExcludeComments);

            List<CultureInfo> cultures = AppSettings.Default.CultureList.Split(';')
                                                                        .Select(n => CultureInfo.GetCultureInfo(n))
                                                                        .ToList();

            switch (ConversionType)
            {
                case Conversion.ResXToXls:

                    if (String.IsNullOrEmpty(this.textBoxFolder.Text))
                    {
                        MessageBox.Show(
                            this,
                            "You must select a the .Net Project root wich contains your updated resx files",
                            "Update",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        this.wizardControl1.CurrentStepIndex = this.intermediateStepProject.StepIndex;
                        this.wizardControl1.Enabled = true;
                        return;
                    }

                    if (this.saveFileDialogXls.ShowDialog() == DialogResult.OK)
                    {
                        Application.DoEvents();
                        string outputPath = this.saveFileDialogXls.FileName;
                        ResxToXls(this.textBoxFolder.Text, this.textBoxScreenshots.Text, AppSettings.Default.ScanSubfolders,
                                  AppSettings.Default.PurgeNonexistant, outputPath, cultures, excludeKeyList, excludeCommentList);
                        MessageBox.Show(
                            this,
                            "Excel Document created.",
                            "Create",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        this.SaveAndClose();
                    }
                    else
                    {
                        this.wizardControl1.Enabled = true;
                    }
                    break;

                case Conversion.XlsToResx:
                    XlsToResx(this.textBoxXls.Text);
                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.SaveAndClose();
                    break;

                default:
                    throw new InvalidOperationException();
            }
        }

        private void SaveAndClose()
        {
            AppSettings.Default.Save();
            this.Close();
        }

        private void wizardControl1_NextButtonClick(WizardBase.WizardControl sender, WizardBase.WizardNextButtonClickEventArgs args)
        {
            switch ((WizardStep)this.wizardControl1.CurrentStepIndex)
            {
                case WizardStep.Start:
                    switch (this.ConversionType)
                    {
                        case Conversion.ResXToXls:
                            args.NextStepIndex = (int)WizardStep.ResX1;
                            break;
                        case Conversion.XlsToResx:
                            args.NextStepIndex = (int)WizardStep.Xls1;
                            break;
                        default:
                            throw new InvalidOperationException("Unknown conversion");
                    }
                    break;

                case WizardStep.ResX1:
                    if (!Directory.Exists(this.textBoxFolder.Text))
                    {
                        MessageBox.Show(
                            this,
                            "Select an existing project directory to continue",
                            "Project directory",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                        args.Cancel = true;
                    }
                    else
                    {
                        AppSettings.Default.FolderPath = this.textBoxFolder.Text;
                        AppSettings.Default.ScreenshotPath = this.textBoxScreenshots.Text;
                        AppSettings.Default.ScanSubfolders = this.checkBoxSubFolders.Checked;
                    }
                    break;

                case WizardStep.ResX2:
                    if (this.listBoxCulturesSelected.Items.Count == 0)
                    {
                        MessageBox.Show(
                            this,
                            "Please select at least one target culture",
                            "Target culture",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);
                        args.Cancel = true;
                    }
                    else
                    {
                        var cultures = this.listBoxCulturesSelected.Items.Cast<CultureInfo>()
                                                                         .Select(c => c.Name)
                                                                         .ToArray();
                        AppSettings.Default.CultureList = string.Join(";", cultures);
                    }
                    break;
                    
                case WizardStep.ResX3:
                    AppSettings.Default.ExcludeKeys = StringHelper.ToCollection(this.textBox_ExcludeKey.Text);
                    AppSettings.Default.ExcludeComments = StringHelper.ToCollection(this.textBox_ExcludeComment.Text);
                    AppSettings.Default.ExcludeFilenames = StringHelper.ToCollection(this.textBox_ExcludeFilename.Text);

                    AppSettings.Default.HideKeys = this.hideKeyColumnCheckbox.Checked;
                    AppSettings.Default.HideComments = this.hideCommentColumnCheckbox.Checked;
                    AppSettings.Default.PurgeNonexistant = this.purgeTranslation_CheckBox.Checked;

                    args.NextStepIndex = (int)WizardStep.Finish;
                    this.textBoxSummary.Text = ResX.summary_create_excel + Environment.NewLine;
                    break;

                case WizardStep.Xls1:
                    if (File.Exists(this.textBoxXls.Text))
                    {
                        this.textBoxSummary.Text = ResX.summary_create_resx + Environment.NewLine;
                    }
                    else
                    {
                        MessageBox.Show(
                            this,
                            "Select an existing Excel document to continue",
                            "Excel path",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation);

                        args.Cancel = true;
                    }
                    break;
            }
        }

        private void wizardControl1_BackButtonClick(WizardBase.WizardControl sender, WizardBase.WizardClickEventArgs args)
        {
            switch ((WizardStep)this.wizardControl1.CurrentStepIndex)
            {
                case WizardStep.Finish:

                    switch (ConversionType)
                    {
                        case Conversion.ResXToXls:
                            this.wizardControl1.CurrentStepIndex = (int)WizardStep.ResX3;
                            args.Cancel = true;
                            break;
                        default:
                            break;
                    }
                    break;
                case WizardStep.Xls1:

                    switch (ConversionType)
                    {
                        case Conversion.XlsToResx:
                            this.wizardControl1.CurrentStepIndex = (int)WizardStep.Start;
                            args.Cancel = true;
                            break;
                        default:
                            break;

                    }
                    break;
            }
        }

        private void wizardControl1_FinishButtonClick(object sender, EventArgs e)
        {
            this.wizardControl1.Enabled = false;
            FinishWizard();
        }

        /// <summary>
        /// Cancel button pressed, quit form
        /// </summary>
        /// <param name="sender">cancel button</param>
        /// <param name="e">bogus event args</param>
        private void wizardControl1_CancelButtonClick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}