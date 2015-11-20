using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.ComponentModel;
using System.Drawing;
using CsvHelper;

namespace ExportMSWordAutoCorrect
{
    public partial class Form1 : Form
    {
        Word.Application application;
        static List<AutoCorrects> items;
        CheckBox checkboxHeader;
        static int TotalClicked = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtLocation.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MSWord_Autocorrects.csv";
        }

        #region DataGridView Events
        void checkboxHeader_Click(object sender, EventArgs e)
        {
            bool value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;
           
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                Row.Cells[0].Value = value;
            }
            TotalClicked = value ? items.Count : 0;
            dataGridView1.RefreshEdit();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex > -1)
            {
                
                    //Escalate Editmode
                    this.dataGridView1.EndEdit();
                    string re_value = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue.ToString();
                    if (re_value.ToLower() == "true")
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "false";
                        checkboxHeader.Checked = false;
                        TotalClicked--;
                        
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "true";
                        TotalClicked++;
                        if (TotalClicked == items.Count)
                        {
                            checkboxHeader.Checked = true;
                        }
                    }

                
            }
        }
        #endregion

        #region Button Events
        private void btnFetch_Click(object sender, EventArgs e)
        {
            btnFetch.Enabled = false;
            btnClear.Enabled = false;
            btnSave.Enabled = false;
            btnSelectFolder.Enabled = false;
            try
            {
                progressBar1.Maximum = 100;
                progressBar1.Step = 1;
                progressBar1.Value = 0;
                // Clean old DataGridView 
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Refresh();
                lblACCount.Text = string.Empty;
                backgroundWorker.RunWorkerAsync();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtLocation.Text = folderBrowserDialog1.SelectedPath + "\\MSWord_Autocorrects.csv";
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            List<AutoCorrects> selected = new List<AutoCorrects>();
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                if (Row.Cells[0].FormattedValue.ToString().ToLower() == "true")
                {
                    selected.Add(new AutoCorrects
                    {
                        Key = Row.Cells[1].FormattedValue.ToString(),
                        Value = Row.Cells[2].FormattedValue.ToString()
                    });
                }
            }
            if (selected.Count > 0)
            {
                if(File.Exists(txtLocation.Text))
                    File.Delete(txtLocation.Text);
                using (TextWriter writer = File.CreateText(txtLocation.Text))
                {
                    var csv = new CsvWriter(writer);
                    csv.Configuration.HasHeaderRecord = false;
                    csv.Configuration.QuoteAllFields = true;
                    csv.WriteRecords(selected);
                }
                MessageBox.Show("File saved successfully !");
            }
            else
            {
                MessageBox.Show("No AutoCorrects Selected");
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            lblACCount.Text = string.Empty;
        }
        #endregion

        #region Get MS Word AutoCorrects
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            application = new Word.Application();
            var backgroundWorker = sender as BackgroundWorker;
            GetAutoCorrectEntries(application);
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            object missing = Type.Missing;
            // prevents Word from trying to save the Normal.dotm template
            application.NormalTemplate.Saved = true;
            ((Microsoft.Office.Interop.Word._Application)application).Quit(missing, missing, missing);

            btnFetch.Enabled = true;
            btnClear.Enabled = true;
            btnSave.Enabled = true;
            btnSelectFolder.Enabled = true;
            progressBar1.Value = 0;
            if (items.Count > 0)
            {
                var bindingList = new BindingList<AutoCorrects>(items);
                var source = new BindingSource(bindingList, null);
                dataGridView1.DataSource = source;
                GenerateDGV();
                lblACCount.Text = this.dataGridView1.Rows.Count + " AutoCorrects found";
            }
        }

        private void GenerateDGV()
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            // Adding column to select/deselect all auto-corrects
            DataGridViewCheckBoxColumn checkboxColumn = new DataGridViewCheckBoxColumn();
            checkboxColumn.Width = 30;
            checkboxColumn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns.Insert(0, checkboxColumn);

            // add checkbox header
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            // set checkbox header to center of header cell. +1 pixel to position correctly.
            rect.X = rect.Location.X + (rect.Width / 4);

            checkboxHeader = new CheckBox();
            checkboxHeader.Name = "checkboxHeader";
            checkboxHeader.Size = new Size(18, 18);
            checkboxHeader.Location = rect.Location;
            checkboxHeader.Checked = true;
            checkboxHeader.Click += checkboxHeader_Click;

            dataGridView1.Controls.Add(checkboxHeader);


            this.dataGridView1.Columns[0].Width = 30;
            this.dataGridView1.Columns[1].Width = 250;
            this.dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[0].Value = true;
            }
            TotalClicked = dataGridView1.Rows.Count;
        }

        private int GetAutoCorrectEntries(Word.Application application)
        {
            int TotalACEntries = -1;
            try
            {
                // If MS Word is not installed properly or some registry issue
                // http://answers.microsoft.com/en-us/office/forum/office_2010-word/unable-to-cast-com-object-of-type/34efcd57-e819-4a83-9cdd-501ab15b0971
                TotalACEntries = application.AutoCorrect.Entries.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show("In the Windows Control Panel, go to Programs & Features. Select the Office 2010 icon and click Change. On the first page of the wizard, click Repair and then Continue. Let it finish and run app again.");
                return -1;
            }
            items = new List<AutoCorrects>();
            for (int x = 1; x <= TotalACEntries; x++)
            {
                object index = (object)x;
                items.Add(new AutoCorrects
                {
                    Key = application.AutoCorrect.Entries.get_Item(ref index).Name,
                    Value = application.AutoCorrect.Entries.get_Item(ref index).Value
                });
                backgroundWorker.ReportProgress((x * 100) / TotalACEntries);
            }
            return TotalACEntries;
        }
        #endregion

    }

    public class AutoCorrects
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

}
