namespace Resx2Xls
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel;
    using AppSettings = Resx2Xls.Properties.Settings;
    using ResX = Resx2Xls.Properties.Resources;
    using System.Diagnostics;

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
            this.textBoxExclude.Text = AppSettings.Default.ExcludeList;
            this.checkBoxFolderNaming.Checked = AppSettings.Default.FolderNamespaceNaming;
            this.hideCommentColumnCheckbox.Checked = AppSettings.Default.HideComments;
            this.hideKeyColumnCheckbox.Checked = AppSettings.Default.HideKeys;

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

        public void ResxToXls(
            string path,
            string screenshotPath,
            bool deepSearch,
            bool purge,
            string outputPath,
            List<CultureInfo> cultures,
            List<string> excludeFilter,
            bool useFolderNamespacePrefix)
        {
            if (!Directory.Exists(path))
                return;

            Cursor = Cursors.WaitCursor;

            this.AddSummaryLine();
            this.AddSummaryLine(ResX.parsing_resx);
            var resxdata = ResxData.FromResx(path, deepSearch, purge, cultures, excludeFilter, useFolderNamespacePrefix);
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
            
            string cList = AppSettings.Default.CultureList;

            string[] cultureList = cList.Split(';');

            foreach (string cult in cultureList)
            {
                CultureInfo info = new CultureInfo(cult);

                this.listBoxSelected.Items.Add(info);
            }
        }

        private void AddCultures()
        {
            for (int i = 0; i < this.listBoxCultures.SelectedItems.Count; i++)
            {
                CultureInfo ci = (CultureInfo)this.listBoxCultures.SelectedItems[i];

                if (this.listBoxSelected.Items.IndexOf(ci) == -1)
                    this.listBoxSelected.Items.Add(ci);
            }
        }

        private void SaveCultures()
        {
            string cultures = String.Empty;
            for (int i = 0; i < this.listBoxSelected.Items.Count; i++)
            {
                CultureInfo info = (CultureInfo)this.listBoxSelected.Items[i];

                if (cultures != String.Empty)
                    cultures = cultures + ";";

                cultures = cultures + info.Name;
            }

            AppSettings.Default.CultureList = cultures;
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
            AddCultures();
        }

        private void listBoxCultures_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AddCultures();
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
            if (this.listBoxSelected.SelectedItems.Count > 0)
            {
                this.listBoxSelected.Items.Remove(this.listBoxSelected.SelectedItems[0]);
            }
        }

        private void textBoxExclude_TextChanged(object sender, EventArgs e)
        {
            AppSettings.Default.ExcludeList = this.textBoxExclude.Text;

        }

        private void Resx2XlsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveCultures();
            AppSettings.Default.FolderNamespaceNaming = this.checkBoxFolderNaming.Checked;
            AppSettings.Default.Save();
        }

        private void textBoxFolder_TextChanged(object sender, EventArgs e)
        {
            AppSettings.Default.FolderPath = this.textBoxFolder.Text;
        }

        private void textBoxScreenshots_TextChanged(object sender, EventArgs e)
        {
            AppSettings.Default.ScreenshotPath = this.textBoxScreenshots.Text;
        }

        public void ShowXls(string path)
        {
            if (!File.Exists(path))
                return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(
                Filename: path,
                Format: 5,
                Origin: Excel.XlPlatform.xlWindows, 
                Editable: true,
                AddToMru: true);

            app.Visible = true;
        }

        private void FinishWizard()
        {
            // Set settings here, no need to pass along
            AppSettings.Default.HideKeys = hideKeyColumnCheckbox.Checked;
            AppSettings.Default.HideComments = hideCommentColumnCheckbox.Checked;

            var excludeFilter = new List<string>(this.textBoxExclude.Text.Split(';'));

            List<CultureInfo> cultures = this.listBoxSelected.Items.Cast<CultureInfo>().ToList();

            switch (ConversionType)
            {
                case Conversion.ResXToXls:

                    if (String.IsNullOrEmpty(this.textBoxFolder.Text))
                    {
                        MessageBox.Show(
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
                        ResxToXls(
                            this.textBoxFolder.Text,
                            this.textBoxScreenshots.Text,
                            this.checkBoxSubFolders.Checked,
                            this.purgeTranslation_CheckBox.Checked,
                            outputPath,
                            cultures,
                            excludeFilter,
                            this.checkBoxFolderNaming.Checked);
                        MessageBox.Show(
                            this,
                            "Excel Document created.",
                            "Create",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        this.Close();
                    }
                    else
                    {
                        this.wizardControl1.Enabled = true;
                    }
                    break;

                case Conversion.XlsToResx:
                    XlsToResx(this.textBoxXls.Text);
                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.Close();
                    break;

                default:
                    throw new InvalidOperationException();
            }
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

                case WizardStep.ResX3:
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
                            "Select an Excel document to continue",
                            "Update",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

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