namespace Resx2Xls
{
    partial class Resx2XlsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialogXls = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialogXls = new System.Windows.Forms.SaveFileDialog();
            this.wizardControl1 = new WizardBase.WizardControl();
            this.startStep1 = new WizardBase.StartStep();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButtonGenerateResx = new System.Windows.Forms.RadioButton();
            this.radioButtonCreateXls = new System.Windows.Forms.RadioButton();
            this.intermediateStepProject = new WizardBase.IntermediateStep();
            this.label7 = new System.Windows.Forms.Label();
            this.textBoxScreenshots = new System.Windows.Forms.TextBox();
            this.browseButtonScreenshots = new System.Windows.Forms.Button();
            this.labelFolder = new System.Windows.Forms.Label();
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.checkBoxSubFolders = new System.Windows.Forms.CheckBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.intermediateStepCultures = new WizardBase.IntermediateStep();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.listBoxCultures = new System.Windows.Forms.ListBox();
            this.listBoxCulturesSelected = new System.Windows.Forms.ListBox();
            this.intermediateStepOptions = new WizardBase.IntermediateStep();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.textBox_ExcludeFilename = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textBox_ExcludeKey = new System.Windows.Forms.TextBox();
            this.textBox_ExcludeComment = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.purgeTranslation_CheckBox = new System.Windows.Forms.CheckBox();
            this.hideCommentColumnCheckbox = new System.Windows.Forms.CheckBox();
            this.hideKeyColumnCheckbox = new System.Windows.Forms.CheckBox();
            this.intermediateStepXlsSelect = new WizardBase.IntermediateStep();
            this.labelXlsFile = new System.Windows.Forms.Label();
            this.textBoxXls = new System.Windows.Forms.TextBox();
            this.buttonBrowseXls = new System.Windows.Forms.Button();
            this.finishStep1 = new WizardBase.FinishStep();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxSummary = new System.Windows.Forms.TextBox();
            this.screenshotFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.startStep1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.intermediateStepProject.SuspendLayout();
            this.intermediateStepCultures.SuspendLayout();
            this.intermediateStepOptions.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.intermediateStepXlsSelect.SuspendLayout();
            this.finishStep1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialogXls
            // 
            this.openFileDialogXls.DefaultExt = "xlsx";
            this.openFileDialogXls.Filter = "Excel Workbook (*.xlsx,*.xls)|*.xlsx;*.xls";
            // 
            // saveFileDialogXls
            // 
            this.saveFileDialogXls.DefaultExt = "xls";
            this.saveFileDialogXls.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls";
            // 
            // wizardControl1
            // 
            this.wizardControl1.BackButtonEnabled = false;
            this.wizardControl1.BackButtonVisible = true;
            this.wizardControl1.CancelButtonEnabled = true;
            this.wizardControl1.CancelButtonVisible = true;
            this.wizardControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wizardControl1.HelpButtonEnabled = true;
            this.wizardControl1.HelpButtonVisible = false;
            this.wizardControl1.Location = new System.Drawing.Point(0, 0);
            this.wizardControl1.Name = "wizardControl1";
            this.wizardControl1.NextButtonEnabled = true;
            this.wizardControl1.NextButtonVisible = true;
            this.wizardControl1.Size = new System.Drawing.Size(704, 466);
            this.wizardControl1.WizardSteps.Add(this.startStep1);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepProject);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepCultures);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepOptions);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepXlsSelect);
            this.wizardControl1.WizardSteps.Add(this.finishStep1);
            this.wizardControl1.BackButtonClick += new WizardBase.WizardClickEventHandler(this.wizardControl1_BackButtonClick);
            this.wizardControl1.CancelButtonClick += new System.EventHandler(this.wizardControl1_CancelButtonClick);
            this.wizardControl1.FinishButtonClick += new System.EventHandler(this.wizardControl1_FinishButtonClick);
            this.wizardControl1.NextButtonClick += new WizardBase.WizardNextButtonClickEventHandler(this.wizardControl1_NextButtonClick);
            // 
            // startStep1
            // 
            this.startStep1.BindingImage = global::Resx2Xls.Properties.Resources.leftbar;
            this.startStep1.Controls.Add(this.groupBox1);
            this.startStep1.Icon = global::Resx2Xls.Properties.Resources.icon;
            this.startStep1.Name = "startStep1";
            this.startStep1.Subtitle = "This wizard helps you to localize your .Net Project";
            this.startStep1.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.startStep1.Title = "Welcome to the Resx to Xls Wizard.";
            this.startStep1.TitleFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonGenerateResx);
            this.groupBox1.Controls.Add(this.radioButtonCreateXls);
            this.groupBox1.Location = new System.Drawing.Point(198, 93);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(373, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // radioButtonGenerateResx
            // 
            this.radioButtonGenerateResx.AutoSize = true;
            this.radioButtonGenerateResx.Location = new System.Drawing.Point(45, 52);
            this.radioButtonGenerateResx.Name = "radioButtonGenerateResx";
            this.radioButtonGenerateResx.Size = new System.Drawing.Size(258, 17);
            this.radioButtonGenerateResx.TabIndex = 1;
            this.radioButtonGenerateResx.Text = "Generate resx files from localized Excel document";
            this.radioButtonGenerateResx.UseVisualStyleBackColor = true;
            // 
            // radioButtonCreateXls
            // 
            this.radioButtonCreateXls.AutoSize = true;
            this.radioButtonCreateXls.Checked = true;
            this.radioButtonCreateXls.Location = new System.Drawing.Point(45, 29);
            this.radioButtonCreateXls.Name = "radioButtonCreateXls";
            this.radioButtonCreateXls.Size = new System.Drawing.Size(267, 17);
            this.radioButtonCreateXls.TabIndex = 0;
            this.radioButtonCreateXls.TabStop = true;
            this.radioButtonCreateXls.Text = "Create a new Excel document ready to be localized";
            this.radioButtonCreateXls.UseVisualStyleBackColor = true;
            // 
            // intermediateStepProject
            // 
            this.intermediateStepProject.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepProject.Controls.Add(this.label7);
            this.intermediateStepProject.Controls.Add(this.textBoxScreenshots);
            this.intermediateStepProject.Controls.Add(this.browseButtonScreenshots);
            this.intermediateStepProject.Controls.Add(this.labelFolder);
            this.intermediateStepProject.Controls.Add(this.textBoxFolder);
            this.intermediateStepProject.Controls.Add(this.checkBoxSubFolders);
            this.intermediateStepProject.Controls.Add(this.buttonBrowse);
            this.intermediateStepProject.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepProject.Name = "intermediateStepProject";
            this.intermediateStepProject.Subtitle = "Browse the root folder of your .Net Project..";
            this.intermediateStepProject.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepProject.Title = "Select your .Net Project.";
            this.intermediateStepProject.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label7.Location = new System.Drawing.Point(20, 230);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(104, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Screenshot directory";
            // 
            // textBoxScreenshots
            // 
            this.textBoxScreenshots.Location = new System.Drawing.Point(23, 246);
            this.textBoxScreenshots.Name = "textBoxScreenshots";
            this.textBoxScreenshots.Size = new System.Drawing.Size(438, 20);
            this.textBoxScreenshots.TabIndex = 13;
            // 
            // browseButtonScreenshots
            // 
            this.browseButtonScreenshots.ForeColor = System.Drawing.SystemColors.ControlText;
            this.browseButtonScreenshots.Location = new System.Drawing.Point(467, 245);
            this.browseButtonScreenshots.Name = "browseButtonScreenshots";
            this.browseButtonScreenshots.Size = new System.Drawing.Size(75, 23);
            this.browseButtonScreenshots.TabIndex = 15;
            this.browseButtonScreenshots.Text = "Browse";
            this.browseButtonScreenshots.UseVisualStyleBackColor = true;
            this.browseButtonScreenshots.Click += new System.EventHandler(this.browseButtonScreenshots_Click);
            // 
            // labelFolder
            // 
            this.labelFolder.AutoSize = true;
            this.labelFolder.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelFolder.Location = new System.Drawing.Point(20, 120);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(217, 13);
            this.labelFolder.TabIndex = 10;
            this.labelFolder.Text = "Project Root (that contains neutral resx files):";
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Location = new System.Drawing.Point(23, 136);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.Size = new System.Drawing.Size(438, 20);
            this.textBoxFolder.TabIndex = 9;
            // 
            // checkBoxSubFolders
            // 
            this.checkBoxSubFolders.AutoSize = true;
            this.checkBoxSubFolders.Checked = true;
            this.checkBoxSubFolders.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSubFolders.ForeColor = System.Drawing.SystemColors.ControlText;
            this.checkBoxSubFolders.Location = new System.Drawing.Point(23, 162);
            this.checkBoxSubFolders.Name = "checkBoxSubFolders";
            this.checkBoxSubFolders.Size = new System.Drawing.Size(110, 17);
            this.checkBoxSubFolders.TabIndex = 12;
            this.checkBoxSubFolders.Text = "Scan Sub Folders";
            this.checkBoxSubFolders.UseVisualStyleBackColor = true;
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonBrowse.Location = new System.Drawing.Point(467, 135);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowse.TabIndex = 11;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // intermediateStepCultures
            // 
            this.intermediateStepCultures.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepCultures.Controls.Add(this.label5);
            this.intermediateStepCultures.Controls.Add(this.label4);
            this.intermediateStepCultures.Controls.Add(this.label3);
            this.intermediateStepCultures.Controls.Add(this.label1);
            this.intermediateStepCultures.Controls.Add(this.buttonAdd);
            this.intermediateStepCultures.Controls.Add(this.listBoxCultures);
            this.intermediateStepCultures.Controls.Add(this.listBoxCulturesSelected);
            this.intermediateStepCultures.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepCultures.Name = "intermediateStepCultures";
            this.intermediateStepCultures.Subtitle = "This step creates a new xls file that contains all your resource keys.";
            this.intermediateStepCultures.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepCultures.Title = "Select the Cultures that you want include in the project.";
            this.intermediateStepCultures.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label5.Location = new System.Drawing.Point(317, 380);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(160, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Double click to remove a culture";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label4.Location = new System.Drawing.Point(64, 380);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(143, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Double click to add a culture";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label3.Location = new System.Drawing.Point(317, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Selected Cultures:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(64, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Available Cultures:";
            // 
            // buttonAdd
            // 
            this.buttonAdd.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonAdd.Location = new System.Drawing.Point(249, 220);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(52, 30);
            this.buttonAdd.TabIndex = 7;
            this.buttonAdd.Text = ">>";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // listBoxCultures
            // 
            this.listBoxCultures.DisplayMember = "EnglishName";
            this.listBoxCultures.FormattingEnabled = true;
            this.listBoxCultures.Location = new System.Drawing.Point(67, 100);
            this.listBoxCultures.Name = "listBoxCultures";
            this.listBoxCultures.Size = new System.Drawing.Size(164, 277);
            this.listBoxCultures.TabIndex = 4;
            this.listBoxCultures.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxCultures_MouseDoubleClick);
            // 
            // listBoxCulturesSelected
            // 
            this.listBoxCulturesSelected.DisplayMember = "EnglishName";
            this.listBoxCulturesSelected.FormattingEnabled = true;
            this.listBoxCulturesSelected.Location = new System.Drawing.Point(319, 100);
            this.listBoxCulturesSelected.Name = "listBoxCulturesSelected";
            this.listBoxCulturesSelected.Size = new System.Drawing.Size(164, 277);
            this.listBoxCulturesSelected.TabIndex = 6;
            this.listBoxCulturesSelected.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxSelected_MouseDoubleClick);
            // 
            // intermediateStepOptions
            // 
            this.intermediateStepOptions.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepOptions.Controls.Add(this.groupBox3);
            this.intermediateStepOptions.Controls.Add(this.groupBox2);
            this.intermediateStepOptions.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepOptions.Name = "intermediateStepOptions";
            this.intermediateStepOptions.Subtitle = "Advanced configuration.";
            this.intermediateStepOptions.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepOptions.Title = "Options.";
            this.intermediateStepOptions.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox_ExcludeFilename);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.textBox_ExcludeKey);
            this.groupBox3.Controls.Add(this.textBox_ExcludeComment);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Location = new System.Drawing.Point(53, 101);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(594, 177);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Ignore translation entry (regular expressions)";
            // 
            // textBox_ExcludeFilename
            // 
            this.textBox_ExcludeFilename.Location = new System.Drawing.Point(393, 41);
            this.textBox_ExcludeFilename.Multiline = true;
            this.textBox_ExcludeFilename.Name = "textBox_ExcludeFilename";
            this.textBox_ExcludeFilename.Size = new System.Drawing.Size(179, 121);
            this.textBox_ExcludeFilename.TabIndex = 20;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label9.Location = new System.Drawing.Point(390, 25);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(95, 13);
            this.label9.TabIndex = 21;
            this.label9.Text = "Resx filename filter";
            // 
            // textBox_ExcludeKey
            // 
            this.textBox_ExcludeKey.Location = new System.Drawing.Point(23, 41);
            this.textBox_ExcludeKey.Multiline = true;
            this.textBox_ExcludeKey.Name = "textBox_ExcludeKey";
            this.textBox_ExcludeKey.Size = new System.Drawing.Size(179, 121);
            this.textBox_ExcludeKey.TabIndex = 13;
            // 
            // textBox_ExcludeComment
            // 
            this.textBox_ExcludeComment.Location = new System.Drawing.Point(208, 41);
            this.textBox_ExcludeComment.Multiline = true;
            this.textBox_ExcludeComment.Name = "textBox_ExcludeComment";
            this.textBox_ExcludeComment.Size = new System.Drawing.Size(179, 121);
            this.textBox_ExcludeComment.TabIndex = 18;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label2.Location = new System.Drawing.Point(20, 25);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 14;
            this.label2.Text = "Key filter";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label8.Location = new System.Drawing.Point(205, 25);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(73, 13);
            this.label8.TabIndex = 19;
            this.label8.Text = "Comment filter";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.purgeTranslation_CheckBox);
            this.groupBox2.Controls.Add(this.hideCommentColumnCheckbox);
            this.groupBox2.Controls.Add(this.hideKeyColumnCheckbox);
            this.groupBox2.Location = new System.Drawing.Point(53, 284);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(594, 87);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Advanced options";
            // 
            // purgeTranslation_CheckBox
            // 
            this.purgeTranslation_CheckBox.AutoSize = true;
            this.purgeTranslation_CheckBox.ForeColor = System.Drawing.SystemColors.ControlText;
            this.purgeTranslation_CheckBox.Location = new System.Drawing.Point(9, 19);
            this.purgeTranslation_CheckBox.Name = "purgeTranslation_CheckBox";
            this.purgeTranslation_CheckBox.Size = new System.Drawing.Size(198, 17);
            this.purgeTranslation_CheckBox.TabIndex = 15;
            this.purgeTranslation_CheckBox.Text = "Purge nonexistant keys in translation";
            this.purgeTranslation_CheckBox.UseVisualStyleBackColor = true;
            // 
            // hideCommentColumnCheckbox
            // 
            this.hideCommentColumnCheckbox.AutoSize = true;
            this.hideCommentColumnCheckbox.ForeColor = System.Drawing.SystemColors.ControlText;
            this.hideCommentColumnCheckbox.Location = new System.Drawing.Point(9, 65);
            this.hideCommentColumnCheckbox.Name = "hideCommentColumnCheckbox";
            this.hideCommentColumnCheckbox.Size = new System.Drawing.Size(131, 17);
            this.hideCommentColumnCheckbox.TabIndex = 16;
            this.hideCommentColumnCheckbox.Text = "Hide comment column";
            this.hideCommentColumnCheckbox.UseVisualStyleBackColor = true;
            // 
            // hideKeyColumnCheckbox
            // 
            this.hideKeyColumnCheckbox.AutoSize = true;
            this.hideKeyColumnCheckbox.ForeColor = System.Drawing.SystemColors.ControlText;
            this.hideKeyColumnCheckbox.Location = new System.Drawing.Point(9, 42);
            this.hideKeyColumnCheckbox.Name = "hideKeyColumnCheckbox";
            this.hideKeyColumnCheckbox.Size = new System.Drawing.Size(105, 17);
            this.hideKeyColumnCheckbox.TabIndex = 17;
            this.hideKeyColumnCheckbox.Text = "Hide key column";
            this.hideKeyColumnCheckbox.UseVisualStyleBackColor = true;
            // 
            // intermediateStepXlsSelect
            // 
            this.intermediateStepXlsSelect.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepXlsSelect.Controls.Add(this.labelXlsFile);
            this.intermediateStepXlsSelect.Controls.Add(this.textBoxXls);
            this.intermediateStepXlsSelect.Controls.Add(this.buttonBrowseXls);
            this.intermediateStepXlsSelect.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepXlsSelect.Name = "intermediateStepXlsSelect";
            this.intermediateStepXlsSelect.Subtitle = "Give a valid xls document that contains localization info.";
            this.intermediateStepXlsSelect.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepXlsSelect.Title = "Select your Excel document.";
            this.intermediateStepXlsSelect.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // labelXlsFile
            // 
            this.labelXlsFile.AutoSize = true;
            this.labelXlsFile.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelXlsFile.Location = new System.Drawing.Point(32, 94);
            this.labelXlsFile.Name = "labelXlsFile";
            this.labelXlsFile.Size = new System.Drawing.Size(96, 13);
            this.labelXlsFile.TabIndex = 2;
            this.labelXlsFile.Text = "Excel resource file:";
            // 
            // textBoxXls
            // 
            this.textBoxXls.Location = new System.Drawing.Point(35, 110);
            this.textBoxXls.Name = "textBoxXls";
            this.textBoxXls.Size = new System.Drawing.Size(385, 20);
            this.textBoxXls.TabIndex = 0;
            // 
            // buttonBrowseXls
            // 
            this.buttonBrowseXls.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonBrowseXls.Location = new System.Drawing.Point(426, 108);
            this.buttonBrowseXls.Name = "buttonBrowseXls";
            this.buttonBrowseXls.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseXls.TabIndex = 1;
            this.buttonBrowseXls.Text = "Browse";
            this.buttonBrowseXls.UseVisualStyleBackColor = true;
            this.buttonBrowseXls.Click += new System.EventHandler(this.buttonBrowseXls_Click);
            // 
            // finishStep1
            // 
            this.finishStep1.BackgroundImage = global::Resx2Xls.Properties.Resources.finishbar;
            this.finishStep1.Controls.Add(this.label6);
            this.finishStep1.Controls.Add(this.textBoxSummary);
            this.finishStep1.Name = "finishStep1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 88);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Summary:";
            // 
            // textBoxSummary
            // 
            this.textBoxSummary.Location = new System.Drawing.Point(27, 103);
            this.textBoxSummary.Multiline = true;
            this.textBoxSummary.Name = "textBoxSummary";
            this.textBoxSummary.ReadOnly = true;
            this.textBoxSummary.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxSummary.Size = new System.Drawing.Size(646, 255);
            this.textBoxSummary.TabIndex = 0;
            // 
            // Resx2XlsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 466);
            this.Controls.Add(this.wizardControl1);
            this.Name = "Resx2XlsForm";
            this.Text = "Resx To Xls";
            this.startStep1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.intermediateStepProject.ResumeLayout(false);
            this.intermediateStepProject.PerformLayout();
            this.intermediateStepCultures.ResumeLayout(false);
            this.intermediateStepCultures.PerformLayout();
            this.intermediateStepOptions.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.intermediateStepXlsSelect.ResumeLayout(false);
            this.intermediateStepXlsSelect.PerformLayout();
            this.finishStep1.ResumeLayout(false);
            this.finishStep1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBoxCultures;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.ListBox listBoxCulturesSelected;
        private System.Windows.Forms.OpenFileDialog openFileDialogXls;
        private System.Windows.Forms.Label labelXlsFile;
        private System.Windows.Forms.Button buttonBrowseXls;
        private System.Windows.Forms.TextBox textBoxXls;
        private System.Windows.Forms.CheckBox checkBoxSubFolders;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.Label labelFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_ExcludeKey;
        private System.Windows.Forms.SaveFileDialog saveFileDialogXls;
        private WizardBase.WizardControl wizardControl1;
        private WizardBase.StartStep startStep1;
        private WizardBase.IntermediateStep intermediateStepProject;
        private WizardBase.IntermediateStep intermediateStepCultures;
        private WizardBase.FinishStep finishStep1;
        private WizardBase.IntermediateStep intermediateStepXlsSelect;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButtonGenerateResx;
        private System.Windows.Forms.RadioButton radioButtonCreateXls;
        private WizardBase.IntermediateStep intermediateStepOptions;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxSummary;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.CheckBox purgeTranslation_CheckBox;
        private System.Windows.Forms.CheckBox hideKeyColumnCheckbox;
        private System.Windows.Forms.CheckBox hideCommentColumnCheckbox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBoxScreenshots;
        private System.Windows.Forms.Button browseButtonScreenshots;
        private System.Windows.Forms.FolderBrowserDialog screenshotFolderBrowserDialog;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox_ExcludeComment;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBox_ExcludeFilename;
        private System.Windows.Forms.Label label9;
    }
}

