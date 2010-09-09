
namespace Resx2Xls
{
    using System;
    using System.Globalization;
    using System.IO;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class Resx2XlsForm : Form
    {
        enum ResxToXlsOperation
        {
            Create,
            Build,
        };

        private ResxToXlsOperation _operation;

        string _summary1;
        string _summary2;

        public Resx2XlsForm()
        {
            CultureInfo ci = new CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            System.Threading.Thread.CurrentThread.CurrentUICulture = ci;

            InitializeComponent();

            this.textBoxFolder.Text = Properties.Settings.Default.FolderPath;
            this.textBoxExclude.Text = Properties.Settings.Default.ExcludeList;
            this.checkBoxFolderNaming.Checked = Properties.Settings.Default.FolderNamespaceNaming;
            
            FillCultures();

            this.radioButtonCreateXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged += new EventHandler(radioButton_CheckedChanged);

            _summary1 = "Operation:\r\nCreate a new Excel document ready for localization.";
            _summary2 = "Operation:\r\nBuild your localized Resource files from a Filled Excel Document.";

            this.textBoxSummary.Text = _summary1;
        }

        void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            this.radioButtonCreateXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);
            this.radioButtonBuildXls.CheckedChanged -= new EventHandler(radioButton_CheckedChanged);

            if (this.radioButtonCreateXls.Checked)
            {
                _operation = ResxToXlsOperation.Create;
                this.textBoxSummary.Text = _summary1;
            }
            if (this.radioButtonBuildXls.Checked)
            {
                _operation = ResxToXlsOperation.Build;
                this.textBoxSummary.Text = _summary2;
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
            bool deepSearch,
            bool purge,
            string xslFile,
            string[] cultures,
            string[] excludeList,
            bool useFolderNamespacePrefix)
        {
            if (!System.IO.Directory.Exists(path))
                return;

            ResxData rd = ResxData.FromResx(path, deepSearch, purge, cultures, excludeList, useFolderNamespacePrefix);

            rd.ToXls(xslFile);

            ShowXls(xslFile);
        }

        private void XlsToResx(string xlsFile)
        {
            if (!File.Exists(xlsFile))
                return;

            string path = new FileInfo(xlsFile).DirectoryName;

            var rd = ResxData.FromXls(xlsFile);
            rd.ToResx(path);
        }

        private void FillCultures()
        {
            CultureInfo[] array = CultureInfo.GetCultures(CultureTypes.AllCultures);
            Array.Sort(array, new CultureInfoComparer());
            foreach (CultureInfo info in array)
            {
                if (info.Equals(CultureInfo.InvariantCulture))
                {
                    //this.listBoxCultures.Items.Add(info, "Default (Invariant Language)");
                }
                else
                {
                    this.listBoxCultures.Items.Add(info);
                }

            }

            string cList = Properties.Settings.Default.CultureList;

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

            Properties.Settings.Default.CultureList = cultures;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (this.folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxFolder.Text = this.folderBrowserDialog.SelectedPath;
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
            Properties.Settings.Default.ExcludeList = this.textBoxExclude.Text;

        }

        private void Resx2XlsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveCultures();
            Properties.Settings.Default.FolderNamespaceNaming = this.checkBoxFolderNaming.Checked;
            Properties.Settings.Default.Save();
        }

        private void textBoxFolder_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.FolderPath = this.textBoxFolder.Text;
        }

        public void ShowXls(string xslFilePath)
        {
            if (!System.IO.File.Exists(xslFilePath))
                return;

            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xslFilePath,
                                                   0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                                                   true, false, 0, true, false, false);

            app.Visible = true;
        }

        private void FinishWizard()
        {
            Cursor = Cursors.WaitCursor;

            string[] excludeList = this.textBoxExclude.Text.Split(';');

            string[] cultures = null;

            cultures = new string[this.listBoxSelected.Items.Count];
            for (int i = 0; i < this.listBoxSelected.Items.Count; i++)
            {
                cultures[i] = ((CultureInfo)this.listBoxSelected.Items[i]).Name;
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
                        string path = this.saveFileDialogXls.FileName;
                        ResxToXls(
                            this.textBoxFolder.Text,
                            this.checkBoxSubFolders.Checked,
                            this.purgeTranslation_CheckBox.Checked,
                            path,
                            cultures,
                            excludeList,
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

            Cursor = Cursors.Default;

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
                            break;
                        default:
                            break;
                    }
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