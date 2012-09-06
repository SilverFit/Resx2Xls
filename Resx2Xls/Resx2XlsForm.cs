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

    public partial class Resx2XlsForm : Form
    {
        enum ResxToXlsOperation
        {
            Create,
            Build,
        };

        private ResxToXlsOperation _operation;

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

            FillCultures();

            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
        }

        void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.radioButtonCreateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);

            if (this.radioButtonCreateXls.Checked)
            {
                _operation = ResxToXlsOperation.Create;
            }
            if (this.radioButtonBuildXls.Checked)
            {
                _operation = ResxToXlsOperation.Build;
            }
            if (((RadioButton)sender).Checked)
            {
                if (((RadioButton)sender) == this.radioButtonCreateXls)
                {
                    this.radioButtonBuildXls.Checked = false;
                }

                if (((RadioButton)sender) == this.radioButtonBuildXls)
                {
                    this.radioButtonCreateXls.Checked = false;
                }
            }
            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
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

            var cultures = new List<CultureInfo>();
            for (int i = 0; i < this.listBoxSelected.Items.Count; i++)
            {
                cultures.Add((CultureInfo)this.listBoxSelected.Items[i]);
            }

            switch (_operation)
            {
                case ResxToXlsOperation.Create:

                    if (String.IsNullOrEmpty(this.textBoxFolder.Text))
                    {
                        MessageBox.Show(
                            "You must select a the .Net Project root wich contains your updated resx files",
                            "Update",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        this.wizardControl1.CurrentStepIndex = this.intermediateStepProject.StepIndex;
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
                            "Excel Document created.",
                            "Create",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    break;
                case ResxToXlsOperation.Build:
                    if (String.IsNullOrEmpty(this.textBoxXls.Text))
                    {
                        MessageBox.Show(
                            "You must select the Excel document to update",
                            "Update",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                        this.wizardControl1.CurrentStepIndex = this.intermediateStepXlsSelect.StepIndex;
                        return;
                    }

                    XlsToResx(this.textBoxXls.Text);
                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;

                default:
                    throw new InvalidOperationException();
            }

            this.Close();
        }

        private void wizardControl1_NextButtonClick(WizardBase.WizardControl sender, WizardBase.WizardNextButtonClickEventArgs args)
        {
            int index = this.wizardControl1.CurrentStepIndex;

            int offset = 1; // è un bug? se non faccio così

            switch (index)
            {
                case 0:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 1 - offset;
                            break;
                        case ResxToXlsOperation.Build:
                            this.wizardControl1.CurrentStepIndex = 4 - offset;
                            break;
                        default:
                            break;
                    }
                    break;

                case 1:

                    switch (_operation)
                    {
                        default:
                            break;
                    }
                    break;


                case 3:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 5 - offset;
                            this.textBoxSummary.Text = ResX.summary_create_excel;
                            break;
                        default:
                            break;
                    }
                    break;

                case 4:
                    this.textBoxSummary.Text = ResX.summary_create_resx;
                    break;
            }
        }

        private void wizardControl1_BackButtonClick(WizardBase.WizardControl sender, WizardBase.WizardClickEventArgs args)
        {
            int index = this.wizardControl1.CurrentStepIndex;

            int offset = 1; // è un bug? se non faccio così

            switch (index)
            {
                case 5:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Create:
                            this.wizardControl1.CurrentStepIndex = 3 + offset;
                            break;
                        default:
                            break;
                    }
                    break;
                case 4:

                    switch (_operation)
                    {
                        case ResxToXlsOperation.Build:
                            this.wizardControl1.CurrentStepIndex = 0 + offset;
                            break;
                        default:
                            break;

                    }
                    break;
            }
        }

        private void wizardControl1_FinishButtonClick(object sender, EventArgs e)
        {
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