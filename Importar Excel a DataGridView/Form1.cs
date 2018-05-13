using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Importar_Excel_a_DataGridView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            if(op.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                txt_Path.Text = op.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string PatCpnn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txt_Path.Text + "; Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(PatCpnn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter("Select * from ["+txt_Sheet.Text+"$]", conn);
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            dataGridView1.DataSource = dt;
        }
    }
}
