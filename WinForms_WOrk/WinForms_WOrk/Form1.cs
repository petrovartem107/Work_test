using ExcelDataReader;
using System.Data;
using System.IO;

namespace WinForms_WOrk

{
    public partial class Form1 : Form
    {
        private string fileNmane = string.Empty;

        private DataTableCollection tableCollection = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void îòêðûòüToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = openFileDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    fileNmane = openFileDialog1.FileName;

                    Text = fileNmane;
                }

                else
                {
                    throw new Exception("Ôàéë íå âûáðàí!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Îøèáêà!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void OpenExcelFile (string path)
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            tableCollection = db.Tables;

            toolStripComboBox1.Items.Clear();

            foreach(DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }
        }
    }
}