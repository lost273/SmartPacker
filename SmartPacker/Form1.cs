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
using System.Xml.Serialization;

namespace SmartPacker {
    public partial class Form1 : Form {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1};IMEX=1'";
        DataTable dataFromFile = new DataTable("table");
        public Form1 () {
            InitializeComponent();
        }

        private void OpenButton_Click (object sender, EventArgs e) {
            openFileDialog.ShowDialog();
        }

        private void openFileDialog_FileOk (object sender, CancelEventArgs e) {
            string filePath = openFileDialog.FileName;
            string extension = Path.GetExtension(filePath);
            string header = "NO";
            string conStr;
            List<string> sheetNames = new List<string>();

            conStr = string.Empty;
            switch (extension) {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            // get the names of the sheets
            using (OleDbConnection con = new OleDbConnection(conStr)) {
                using (OleDbCommand cmd = new OleDbCommand()) {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    // fill the list
                    for (int i = 0; i < dtExcelSchema.Rows.Count; i++) {
                        sheetNames.Add(dtExcelSchema.Rows[i]["TABLE_NAME"].ToString());
                    }
                    con.Close();
                }
            }

            //Read Data from the Sheets
            using (OleDbConnection con = new OleDbConnection(conStr)) {
                using (OleDbCommand cmd = new OleDbCommand()) {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter()) {
                        con.Open();
                        foreach (string Name in sheetNames) {
                            cmd.CommandText = "SELECT * From [" + Name + "]";
                            cmd.Connection = con;
                            oda.SelectCommand = cmd;
                            oda.Fill(dataFromFile);
                        }
                        con.Close();
                    }
                }
            }
        }

        private void SaveButton_Click(object sender, EventArgs e) {
            saveFileDialog.ShowDialog();
        }

        private void SaveFileDialog_FileOk(object sender, CancelEventArgs e) {
            dataFromFile.WriteXml(saveFileDialog.FileName);

            DataTable data = new DataTable();
            data.ReadXml(saveFileDialog.FileName);

            //XmlSerializer ser = new XmlSerializer(typeof(DataSet));

            // Creates a DataSet; adds a table, column, and ten rows.
            //DataSet ds = new DataSet("DataSet");
            //DataTable t = new DataTable("table1");
            //DataColumn c = new DataColumn("thing");
            //t.Columns.Add(c);
            //ds.Tables.Add(t);
            //DataRow r;
            //for (int i = 0; i < 10; i++) {
            //    r = t.NewRow();
            //    r[0] = "Thing " + i;
            //    t.Rows.Add(r);
            //}
            //TextWriter writer = new StreamWriter(saveFileDialog.FileName);
            //ser.Serialize(writer, ds);
            //writer.Close();
        }
    }
}
