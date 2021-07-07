/*
 * Created by SharpDevelop.
 * User: zsianti
 * Date: 14.08.2012
 * Time: 07:54
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace OutlookGoogleSync
{
    partial class MainForm
    {
        /// <summary>
        /// Designer variable used to keep track of non-visual components.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Disposes resources used by the form.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// This method is required for Windows Forms designer support.
        /// Do not change the method contents inside the source code editor. The Forms designer might
        /// not be able to load this method if it was changed manually.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageSync = new System.Windows.Forms.TabPage();
            this.buttonDeleteAll = new System.Windows.Forms.Button();
            this.textBoxLogs = new System.Windows.Forms.TextBox();
            this.buttonSyncNow = new System.Windows.Forms.Button();
            this.tabPageSettings = new System.Windows.Forms.TabPage();
            this.groupBoxWhenCreating = new System.Windows.Forms.GroupBox();
            this.checkBoxAddReminders = new System.Windows.Forms.CheckBox();
            this.checkBoxAddAttendees = new System.Windows.Forms.CheckBox();
            this.checkBoxAddDescription = new System.Windows.Forms.CheckBox();
            this.groupBoxOptions = new System.Windows.Forms.GroupBox();
            this.checkBoxMinimizeToTray = new System.Windows.Forms.CheckBox();
            this.checkBoxStartInTray = new System.Windows.Forms.CheckBox();
            this.checkBoxCreateFiles = new System.Windows.Forms.CheckBox();
            this.groupBoxSyncRegularly = new System.Windows.Forms.GroupBox();
            this.checkBoxShowBubbleTooltips = new System.Windows.Forms.CheckBox();
            this.checkBoxSyncEveryHour = new System.Windows.Forms.CheckBox();
            this.textBoxMinuteOffsets = new System.Windows.Forms.TextBox();
            this.groupBoxGoogleCalendar = new System.Windows.Forms.GroupBox();
            this.labelUseGoogleCalendar = new System.Windows.Forms.Label();
            this.buttonGetMyCalendars = new System.Windows.Forms.Button();
            this.comboBoxCalendars = new System.Windows.Forms.ComboBox();
            this.buttonSave = new System.Windows.Forms.Button();
            this.groupBoxSyncDataRange = new System.Windows.Forms.GroupBox();
            this.numericUpDownDaysInTheFuture = new System.Windows.Forms.NumericUpDown();
            this.numericUpDownDaysInThePast = new System.Windows.Forms.NumericUpDown();
            this.labelDaysInTheFuture = new System.Windows.Forms.Label();
            this.labelDaysInThePast = new System.Windows.Forms.Label();
            this.tabPageAbout = new System.Windows.Forms.TabPage();
            this.linkLabelUrl = new System.Windows.Forms.LinkLabel();
            this.labelAbout = new System.Windows.Forms.Label();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.tabControl1.SuspendLayout();
            this.tabPageSync.SuspendLayout();
            this.tabPageSettings.SuspendLayout();
            this.groupBoxWhenCreating.SuspendLayout();
            this.groupBoxOptions.SuspendLayout();
            this.groupBoxSyncRegularly.SuspendLayout();
            this.groupBoxGoogleCalendar.SuspendLayout();
            this.groupBoxSyncDataRange.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDaysInTheFuture)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDaysInThePast)).BeginInit();
            this.tabPageAbout.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageSync);
            this.tabControl1.Controls.Add(this.tabPageSettings);
            this.tabControl1.Controls.Add(this.tabPageAbout);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(519, 529);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPageSync
            // 
            this.tabPageSync.Controls.Add(this.buttonDeleteAll);
            this.tabPageSync.Controls.Add(this.textBoxLogs);
            this.tabPageSync.Controls.Add(this.buttonSyncNow);
            this.tabPageSync.Location = new System.Drawing.Point(4, 22);
            this.tabPageSync.Name = "tabPageSync";
            this.tabPageSync.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.tabPageSync.Size = new System.Drawing.Size(511, 503);
            this.tabPageSync.TabIndex = 0;
            this.tabPageSync.Text = "Sync";
            this.tabPageSync.UseVisualStyleBackColor = true;
            // 
            // buttonDeleteAll
            // 
            this.buttonDeleteAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonDeleteAll.Location = new System.Drawing.Point(350, 467);
            this.buttonDeleteAll.Name = "buttonDeleteAll";
            this.buttonDeleteAll.Size = new System.Drawing.Size(158, 31);
            this.buttonDeleteAll.TabIndex = 2;
            this.buttonDeleteAll.Text = "Delete All Sync Entries";
            this.buttonDeleteAll.UseVisualStyleBackColor = true;
            this.buttonDeleteAll.Click += new System.EventHandler(this.buttonDeleteAll_Click);
            // 
            // textBoxLogs
            // 
            this.textBoxLogs.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLogs.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxLogs.Location = new System.Drawing.Point(7, 6);
            this.textBoxLogs.Multiline = true;
            this.textBoxLogs.Name = "textBoxLogs";
            this.textBoxLogs.Size = new System.Drawing.Size(501, 455);
            this.textBoxLogs.TabIndex = 3;
            // 
            // buttonSyncNow
            // 
            this.buttonSyncNow.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonSyncNow.Location = new System.Drawing.Point(5, 467);
            this.buttonSyncNow.Name = "buttonSyncNow";
            this.buttonSyncNow.Size = new System.Drawing.Size(98, 31);
            this.buttonSyncNow.TabIndex = 0;
            this.buttonSyncNow.Text = "Sync now";
            this.buttonSyncNow.UseVisualStyleBackColor = true;
            this.buttonSyncNow.Click += new System.EventHandler(this.buttonSyncNow_Click);
            // 
            // tabPageSettings
            // 
            this.tabPageSettings.Controls.Add(this.groupBoxWhenCreating);
            this.tabPageSettings.Controls.Add(this.groupBoxOptions);
            this.tabPageSettings.Controls.Add(this.groupBoxSyncRegularly);
            this.tabPageSettings.Controls.Add(this.groupBoxGoogleCalendar);
            this.tabPageSettings.Controls.Add(this.buttonSave);
            this.tabPageSettings.Controls.Add(this.groupBoxSyncDataRange);
            this.tabPageSettings.Location = new System.Drawing.Point(4, 22);
            this.tabPageSettings.Name = "tabPageSettings";
            this.tabPageSettings.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.tabPageSettings.Size = new System.Drawing.Size(511, 503);
            this.tabPageSettings.TabIndex = 1;
            this.tabPageSettings.Text = "Settings";
            this.tabPageSettings.UseVisualStyleBackColor = true;
            // 
            // groupBoxWhenCreating
            // 
            this.groupBoxWhenCreating.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxWhenCreating.Controls.Add(this.checkBoxAddReminders);
            this.groupBoxWhenCreating.Controls.Add(this.checkBoxAddAttendees);
            this.groupBoxWhenCreating.Controls.Add(this.checkBoxAddDescription);
            this.groupBoxWhenCreating.Location = new System.Drawing.Point(6, 171);
            this.groupBoxWhenCreating.Name = "groupBoxWhenCreating";
            this.groupBoxWhenCreating.Size = new System.Drawing.Size(499, 112);
            this.groupBoxWhenCreating.TabIndex = 15;
            this.groupBoxWhenCreating.TabStop = false;
            this.groupBoxWhenCreating.Text = "When creating Google Calendar Entries...   ";
            // 
            // checkBoxAddReminders
            // 
            this.checkBoxAddReminders.Location = new System.Drawing.Point(12, 79);
            this.checkBoxAddReminders.Name = "checkBoxAddReminders";
            this.checkBoxAddReminders.Size = new System.Drawing.Size(139, 24);
            this.checkBoxAddReminders.TabIndex = 8;
            this.checkBoxAddReminders.Text = "Add Reminders";
            this.checkBoxAddReminders.UseVisualStyleBackColor = true;
            this.checkBoxAddReminders.CheckedChanged += new System.EventHandler(this.checkBoxAddReminders_CheckedChanged);
            // 
            // checkBoxAddAttendees
            // 
            this.checkBoxAddAttendees.Location = new System.Drawing.Point(12, 49);
            this.checkBoxAddAttendees.Name = "checkBoxAddAttendees";
            this.checkBoxAddAttendees.Size = new System.Drawing.Size(235, 24);
            this.checkBoxAddAttendees.TabIndex = 6;
            this.checkBoxAddAttendees.Text = "Add Attendees at the end of the Description";
            this.checkBoxAddAttendees.UseVisualStyleBackColor = true;
            this.checkBoxAddAttendees.CheckedChanged += new System.EventHandler(this.cbAddAttendees_CheckedChanged);
            // 
            // checkBoxAddDescription
            // 
            this.checkBoxAddDescription.Location = new System.Drawing.Point(12, 19);
            this.checkBoxAddDescription.Name = "checkBoxAddDescription";
            this.checkBoxAddDescription.Size = new System.Drawing.Size(209, 24);
            this.checkBoxAddDescription.TabIndex = 1;
            this.checkBoxAddDescription.Text = "Add Description";
            this.checkBoxAddDescription.UseVisualStyleBackColor = true;
            this.checkBoxAddDescription.CheckedChanged += new System.EventHandler(this.checkBoxAddDescription_CheckedChanged);
            // 
            // groupBoxOptions
            // 
            this.groupBoxOptions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxOptions.Controls.Add(this.checkBoxMinimizeToTray);
            this.groupBoxOptions.Controls.Add(this.checkBoxStartInTray);
            this.groupBoxOptions.Controls.Add(this.checkBoxCreateFiles);
            this.groupBoxOptions.Location = new System.Drawing.Point(6, 289);
            this.groupBoxOptions.Name = "groupBoxOptions";
            this.groupBoxOptions.Size = new System.Drawing.Size(499, 115);
            this.groupBoxOptions.TabIndex = 20;
            this.groupBoxOptions.TabStop = false;
            this.groupBoxOptions.Text = "Options";
            // 
            // checkBoxMinimizeToTray
            // 
            this.checkBoxMinimizeToTray.Location = new System.Drawing.Point(12, 49);
            this.checkBoxMinimizeToTray.Name = "checkBoxMinimizeToTray";
            this.checkBoxMinimizeToTray.Size = new System.Drawing.Size(104, 24);
            this.checkBoxMinimizeToTray.TabIndex = 5;
            this.checkBoxMinimizeToTray.Text = "Minimize to Tray";
            this.checkBoxMinimizeToTray.UseVisualStyleBackColor = true;
            this.checkBoxMinimizeToTray.CheckedChanged += new System.EventHandler(this.checkBoxMinimizeToTray_CheckedChanged);
            // 
            // checkBoxStartInTray
            // 
            this.checkBoxStartInTray.Location = new System.Drawing.Point(12, 19);
            this.checkBoxStartInTray.Name = "checkBoxStartInTray";
            this.checkBoxStartInTray.Size = new System.Drawing.Size(104, 24);
            this.checkBoxStartInTray.TabIndex = 1;
            this.checkBoxStartInTray.Text = "Start in Tray";
            this.checkBoxStartInTray.UseVisualStyleBackColor = true;
            this.checkBoxStartInTray.CheckedChanged += new System.EventHandler(this.checkBoxStartInTray_CheckedChanged);
            // 
            // checkBoxCreateFiles
            // 
            this.checkBoxCreateFiles.Location = new System.Drawing.Point(12, 79);
            this.checkBoxCreateFiles.Name = "checkBoxCreateFiles";
            this.checkBoxCreateFiles.Size = new System.Drawing.Size(235, 24);
            this.checkBoxCreateFiles.TabIndex = 7;
            this.checkBoxCreateFiles.Text = "Create text files with found/identified entries";
            this.checkBoxCreateFiles.UseVisualStyleBackColor = true;
            this.checkBoxCreateFiles.CheckedChanged += new System.EventHandler(this.cbCreateFiles_CheckedChanged);
            // 
            // groupBoxSyncRegularly
            // 
            this.groupBoxSyncRegularly.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxSyncRegularly.Controls.Add(this.checkBoxShowBubbleTooltips);
            this.groupBoxSyncRegularly.Controls.Add(this.checkBoxSyncEveryHour);
            this.groupBoxSyncRegularly.Controls.Add(this.textBoxMinuteOffsets);
            this.groupBoxSyncRegularly.Location = new System.Drawing.Point(177, 80);
            this.groupBoxSyncRegularly.Name = "groupBoxSyncRegularly";
            this.groupBoxSyncRegularly.Size = new System.Drawing.Size(328, 85);
            this.groupBoxSyncRegularly.TabIndex = 10;
            this.groupBoxSyncRegularly.TabStop = false;
            this.groupBoxSyncRegularly.Text = "Sync Regularly";
            // 
            // checkBoxShowBubbleTooltips
            // 
            this.checkBoxShowBubbleTooltips.Location = new System.Drawing.Point(6, 49);
            this.checkBoxShowBubbleTooltips.Name = "checkBoxShowBubbleTooltips";
            this.checkBoxShowBubbleTooltips.Size = new System.Drawing.Size(259, 24);
            this.checkBoxShowBubbleTooltips.TabIndex = 13;
            this.checkBoxShowBubbleTooltips.Text = "Show Bubble Tooltip in Taskbar when Syncing";
            this.checkBoxShowBubbleTooltips.UseVisualStyleBackColor = true;
            this.checkBoxShowBubbleTooltips.CheckedChanged += new System.EventHandler(this.checkBoxShowBubbleTooltips_CheckedChanged);
            // 
            // checkBoxSyncEveryHour
            // 
            this.checkBoxSyncEveryHour.Location = new System.Drawing.Point(6, 19);
            this.checkBoxSyncEveryHour.Name = "checkBoxSyncEveryHour";
            this.checkBoxSyncEveryHour.Size = new System.Drawing.Size(221, 24);
            this.checkBoxSyncEveryHour.TabIndex = 6;
            this.checkBoxSyncEveryHour.Text = "Sync every hour at these Minute Offset(s)";
            this.checkBoxSyncEveryHour.UseVisualStyleBackColor = true;
            this.checkBoxSyncEveryHour.CheckedChanged += new System.EventHandler(this.checkBoxSyncEveryHour_CheckedChanged);
            // 
            // textBoxMinuteOffsets
            // 
            this.textBoxMinuteOffsets.Location = new System.Drawing.Point(231, 21);
            this.textBoxMinuteOffsets.Name = "textBoxMinuteOffsets";
            this.textBoxMinuteOffsets.Size = new System.Drawing.Size(67, 20);
            this.textBoxMinuteOffsets.TabIndex = 9;
            this.textBoxMinuteOffsets.TextChanged += new System.EventHandler(this.textBoxMinuteOffsets_TextChanged);
            // 
            // groupBoxGoogleCalendar
            // 
            this.groupBoxGoogleCalendar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxGoogleCalendar.Controls.Add(this.labelUseGoogleCalendar);
            this.groupBoxGoogleCalendar.Controls.Add(this.buttonGetMyCalendars);
            this.groupBoxGoogleCalendar.Controls.Add(this.comboBoxCalendars);
            this.groupBoxGoogleCalendar.Location = new System.Drawing.Point(6, 6);
            this.groupBoxGoogleCalendar.Name = "groupBoxGoogleCalendar";
            this.groupBoxGoogleCalendar.Size = new System.Drawing.Size(499, 68);
            this.groupBoxGoogleCalendar.TabIndex = 0;
            this.groupBoxGoogleCalendar.TabStop = false;
            this.groupBoxGoogleCalendar.Text = "Google Calendar";
            // 
            // labelUseGoogleCalendar
            // 
            this.labelUseGoogleCalendar.Location = new System.Drawing.Point(6, 33);
            this.labelUseGoogleCalendar.Name = "labelUseGoogleCalendar";
            this.labelUseGoogleCalendar.Size = new System.Drawing.Size(112, 23);
            this.labelUseGoogleCalendar.TabIndex = 3;
            this.labelUseGoogleCalendar.Text = "Use Google Calendar:";
            // 
            // buttonGetMyCalendars
            // 
            this.buttonGetMyCalendars.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonGetMyCalendars.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonGetMyCalendars.Location = new System.Drawing.Point(379, 19);
            this.buttonGetMyCalendars.Name = "buttonGetMyCalendars";
            this.buttonGetMyCalendars.Size = new System.Drawing.Size(114, 40);
            this.buttonGetMyCalendars.TabIndex = 5;
            this.buttonGetMyCalendars.Text = "Get My\r\nGoogle Calendars";
            this.buttonGetMyCalendars.UseVisualStyleBackColor = true;
            this.buttonGetMyCalendars.Click += new System.EventHandler(this.buttonGetMyGoogleCalendars_Click);
            // 
            // comboBoxCalendars
            // 
            this.comboBoxCalendars.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxCalendars.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCalendars.FormattingEnabled = true;
            this.comboBoxCalendars.Location = new System.Drawing.Point(123, 31);
            this.comboBoxCalendars.Name = "comboBoxCalendars";
            this.comboBoxCalendars.Size = new System.Drawing.Size(249, 21);
            this.comboBoxCalendars.TabIndex = 1;
            this.comboBoxCalendars.SelectedIndexChanged += new System.EventHandler(this.comboBoxCalendars_SelectedIndexChanged);
            // 
            // buttonSave
            // 
            this.buttonSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonSave.Location = new System.Drawing.Point(5, 467);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(75, 31);
            this.buttonSave.TabIndex = 30;
            this.buttonSave.Text = "Save";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // groupBoxSyncDataRange
            // 
            this.groupBoxSyncDataRange.Controls.Add(this.numericUpDownDaysInTheFuture);
            this.groupBoxSyncDataRange.Controls.Add(this.numericUpDownDaysInThePast);
            this.groupBoxSyncDataRange.Controls.Add(this.labelDaysInTheFuture);
            this.groupBoxSyncDataRange.Controls.Add(this.labelDaysInThePast);
            this.groupBoxSyncDataRange.Location = new System.Drawing.Point(6, 80);
            this.groupBoxSyncDataRange.Name = "groupBoxSyncDataRange";
            this.groupBoxSyncDataRange.Size = new System.Drawing.Size(165, 85);
            this.groupBoxSyncDataRange.TabIndex = 5;
            this.groupBoxSyncDataRange.TabStop = false;
            this.groupBoxSyncDataRange.Text = "Sync Date Range";
            // 
            // numericUpDownDaysInTheFuture
            // 
            this.numericUpDownDaysInTheFuture.Location = new System.Drawing.Point(105, 50);
            this.numericUpDownDaysInTheFuture.Name = "numericUpDownDaysInTheFuture";
            this.numericUpDownDaysInTheFuture.Size = new System.Drawing.Size(46, 20);
            this.numericUpDownDaysInTheFuture.TabIndex = 4;
            this.numericUpDownDaysInTheFuture.ValueChanged += new System.EventHandler(this.numericUpDownDaysInTheFuture_ValueChanged);
            // 
            // numericUpDownDaysInThePast
            // 
            this.numericUpDownDaysInThePast.Location = new System.Drawing.Point(105, 21);
            this.numericUpDownDaysInThePast.Name = "numericUpDownDaysInThePast";
            this.numericUpDownDaysInThePast.Size = new System.Drawing.Size(46, 20);
            this.numericUpDownDaysInThePast.TabIndex = 3;
            this.numericUpDownDaysInThePast.ValueChanged += new System.EventHandler(this.numericUpDownDaysInThePast_ValueChanged);
            // 
            // labelDaysInTheFuture
            // 
            this.labelDaysInTheFuture.Location = new System.Drawing.Point(6, 54);
            this.labelDaysInTheFuture.Name = "labelDaysInTheFuture";
            this.labelDaysInTheFuture.Size = new System.Drawing.Size(100, 23);
            this.labelDaysInTheFuture.TabIndex = 0;
            this.labelDaysInTheFuture.Text = "Days in the Future";
            // 
            // labelDaysInThePast
            // 
            this.labelDaysInThePast.Location = new System.Drawing.Point(6, 24);
            this.labelDaysInThePast.Name = "labelDaysInThePast";
            this.labelDaysInThePast.Size = new System.Drawing.Size(100, 23);
            this.labelDaysInThePast.TabIndex = 0;
            this.labelDaysInThePast.Text = "Days in the Past";
            // 
            // tabPageAbout
            // 
            this.tabPageAbout.Controls.Add(this.linkLabelUrl);
            this.tabPageAbout.Controls.Add(this.labelAbout);
            this.tabPageAbout.Location = new System.Drawing.Point(4, 22);
            this.tabPageAbout.Name = "tabPageAbout";
            this.tabPageAbout.Padding = new System.Windows.Forms.Padding(3, 3, 3, 3);
            this.tabPageAbout.Size = new System.Drawing.Size(511, 503);
            this.tabPageAbout.TabIndex = 2;
            this.tabPageAbout.Text = "About";
            this.tabPageAbout.UseVisualStyleBackColor = true;
            // 
            // linkLabelUrl
            // 
            this.linkLabelUrl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabelUrl.Location = new System.Drawing.Point(5, 157);
            this.linkLabelUrl.Name = "linkLabelUrl";
            this.linkLabelUrl.Size = new System.Drawing.Size(505, 23);
            this.linkLabelUrl.TabIndex = 2;
            this.linkLabelUrl.TabStop = true;
            this.linkLabelUrl.Text = "https://github.com/David-Engel/OutlookGoogleCalendarSync";
            this.linkLabelUrl.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.linkLabelUrl.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelUrl_LinkClicked);
            // 
            // labelAbout
            // 
            this.labelAbout.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelAbout.Location = new System.Drawing.Point(3, 32);
            this.labelAbout.Name = "labelAbout";
            this.labelAbout.Size = new System.Drawing.Size(508, 118);
            this.labelAbout.TabIndex = 1;
            this.labelAbout.Text = "David\'s Outlook to Google Calendar Sync\r\n\r\nVersion {version}\r\n\r\nby David Engel\r\n\r" +
    "\nOriginal Credit:\r\nOutlookGoogleSync by Zissis Siantidis\r\n";
            this.labelAbout.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
            this.notifyIcon1.Text = "OutlookGoogleSync";
            this.notifyIcon1.Click += new System.EventHandler(this.notifyIcon1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(519, 529);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "David\'s Outlook to Google Calendar Sync";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            this.tabControl1.ResumeLayout(false);
            this.tabPageSync.ResumeLayout(false);
            this.tabPageSync.PerformLayout();
            this.tabPageSettings.ResumeLayout(false);
            this.groupBoxWhenCreating.ResumeLayout(false);
            this.groupBoxOptions.ResumeLayout(false);
            this.groupBoxSyncRegularly.ResumeLayout(false);
            this.groupBoxSyncRegularly.PerformLayout();
            this.groupBoxGoogleCalendar.ResumeLayout(false);
            this.groupBoxSyncDataRange.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDaysInTheFuture)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownDaysInThePast)).EndInit();
            this.tabPageAbout.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        private System.Windows.Forms.CheckBox checkBoxAddReminders;
        private System.Windows.Forms.CheckBox checkBoxAddDescription;
        private System.Windows.Forms.CheckBox checkBoxShowBubbleTooltips;
        private System.Windows.Forms.CheckBox checkBoxSyncEveryHour;
        private System.Windows.Forms.CheckBox checkBoxMinimizeToTray;
        private System.Windows.Forms.CheckBox checkBoxStartInTray;
        private System.Windows.Forms.GroupBox groupBoxOptions;
        private System.Windows.Forms.GroupBox groupBoxWhenCreating;
        private System.Windows.Forms.LinkLabel linkLabelUrl;
        private System.Windows.Forms.TabPage tabPageAbout;
        private System.Windows.Forms.TextBox textBoxMinuteOffsets;
        private System.Windows.Forms.GroupBox groupBoxSyncRegularly;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.Label labelAbout;
        private System.Windows.Forms.CheckBox checkBoxAddAttendees;
        private System.Windows.Forms.CheckBox checkBoxCreateFiles;
        private System.Windows.Forms.TextBox textBoxLogs;
        private System.Windows.Forms.GroupBox groupBoxGoogleCalendar;
        private System.Windows.Forms.Label labelUseGoogleCalendar;
        public System.Windows.Forms.ComboBox comboBoxCalendars;
        private System.Windows.Forms.Label labelDaysInThePast;
        private System.Windows.Forms.Label labelDaysInTheFuture;
        private System.Windows.Forms.GroupBox groupBoxSyncDataRange;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Button buttonGetMyCalendars;
        private System.Windows.Forms.TabPage tabPageSettings;
        private System.Windows.Forms.Button buttonSyncNow;
        private System.Windows.Forms.TabPage tabPageSync;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.NumericUpDown numericUpDownDaysInTheFuture;
        private System.Windows.Forms.NumericUpDown numericUpDownDaysInThePast;
        private System.Windows.Forms.Button buttonDeleteAll;









    }
}
