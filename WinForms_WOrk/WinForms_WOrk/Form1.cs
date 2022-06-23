using ExcelDataReader;
using System.Data;

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

        private void �������ToolStripMenuItem_Click(object sender, EventArgs e)
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
                    throw new Exception("���� �� ������!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "������!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}