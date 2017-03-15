using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadXlsFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void LoadXLS(String fileName, String SheetName)
        {
            //var fileName = string.Format("{0}\\data.xls", Directory.GetCurrentDirectory());
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            
            var adapter = new OleDbDataAdapter("SELECT * FROM [" + SheetName + "$]", connectionString);
            var ds = new DataSet();

            adapter.Fill(ds, SheetName);

            DataTable data = ds.Tables[SheetName];  //<<資料都在這裡 DataTable data


            List<String> colName = new List<String>(); //取得所有 Colnum Name
            DataColumnCollection cols = data.Columns;
            for (int i = 0; i < cols.Count; i++)
            {
                Console.WriteLine(cols[i].ToString());
                colName.Add(cols[i].ToString());
            }


            //一筆一筆row取出來
            DataView dvEmp = new DataView(data);
            foreach (DataRowView rowView in dvEmp)
            {
                DataRow row = rowView.Row;
                //取出這筆row裡面的 column欄位
                for (int i = 0; i < colName.Count; i++)
                {
                    Console.WriteLine(String.Format("{0}:{1}", colName[i], row[i].ToString()));
                }
                Console.WriteLine("---------------------------------");
            }



            //dvEmp.RowFilter = "Type like 'C%'";
            dataGridView1.DataSource = dvEmp;

        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel File|*.xls";
            openFileDialog1.Title = "Open an Excel File";
            openFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.
            if (openFileDialog1.FileName != "")
            {
                String SheetName = "Driver";
                LoadXLS(openFileDialog1.FileName, SheetName);
            }
        }
    }
}
