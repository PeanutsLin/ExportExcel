using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportExcel
{
    public partial class WPSExportForm : Form
    {
        private WPSExcel wpsExcel = new WPSExcel();
        public WPSExportForm()
        {
            InitializeComponent();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {       
            OpenFileDialog oFD = new OpenFileDialog();

            oFD.Title = "打开Excel";
            oFD.Filter = "WPS Excel文件|*.xls";

            if (oFD.ShowDialog() == DialogResult.OK)
            {
                if (oFD.FileName.Length == 0)
                {
                    MessageBox.Show("请输入打开的文件名");
                    return;
                }

                try
                {
                    for (int i = 0; i < 10; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            dataGridViewExcel[j, i].Value = "";
                        }
                    }

                    wpsExcel.OpenExcel(oFD.FileName, 1);

                    for (int i = 0; i < 10; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            dataGridViewExcel[j, i].Value = wpsExcel.Read(i + 1, j + 1);                         
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }   
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            wpsExcel.CloseExcel();

            this.Close();
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            SaveFileDialog sFD = new SaveFileDialog();

            sFD.Title = "创建Excel";
            sFD.Filter = "WPS Excel文件|*.xls";

            if (sFD.ShowDialog() == DialogResult.OK)
            {
                if (sFD.FileName.Length == 0)
                {
                    MessageBox.Show("请输入创建的文件名");
                    return;
                }
                wpsExcel.CreateExcel(sFD.FileName);
            }
        }

        private void WPSExportForm_Load(object sender, EventArgs e)
        {
            dataGridViewExcel.RowHeadersVisible = false;
            dataGridViewExcel.ColumnHeadersVisible = false;

            for (int i = 0; i < 10; i++)
            {
                dataGridViewExcel.Rows.Add("");
            }

            dataGridViewExcel.AllowUserToResizeRows = false;
            dataGridViewExcel.AllowUserToResizeColumns = false;
        }
      
        private void buttonWrite_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    if (dataGridViewExcel[j, i].Value != null && dataGridViewExcel[j, i].Value.ToString() != string.Empty)
                    {
                        wpsExcel.Write(i + 1, j + 1, dataGridViewExcel[j, i].Value.ToString());
                    }
                }
            }
        }
    }
}
