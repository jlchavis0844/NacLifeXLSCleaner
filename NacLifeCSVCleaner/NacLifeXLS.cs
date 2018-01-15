using System.Linq;
using ExcelDataReader;
using System.Data;
using System.IO;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace NacLifeXLSCleaner {
    internal class NacLifeXLS {

        public Microsoft.Office.Interop.Excel.Application oXL;
        public Microsoft.Office.Interop.Excel._Workbook oWB;
        public Microsoft.Office.Interop.Excel._Worksheet oSheet;
        public Microsoft.Office.Interop.Excel.Range oRng;
        public string fileName = "";

        public NacLifeXLS(string filePath) {
            fileName = Path.GetFileName(filePath);
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            DataSet result = excelReader.AsDataSet();

            DataTable table = result.Tables[0];
            Console.WriteLine(table.Rows.Count);
            List<DataRow> delRows = new List<DataRow>();
            List<DataRow> commRows = new List<DataRow>();
            List<DataRow> advRows = new List<DataRow>();


            foreach (DataRow row in table.Rows) {
                if (row[12].ToString() == "") {
                    delRows.Add(row);
                }
                else if (row[14].ToString() == "") {
                    commRows.Add(row);
                }
                else {
                    advRows.Add(row);
                }
            }
            Console.WriteLine(table.Rows.Count);

            DataTable newTable = new DataTable();
            DataColumn polCol = new DataColumn();
            newTable.Columns.Add("Policy", typeof(string));
            newTable.Columns.Add("Fullname", typeof(string));
            newTable.Columns.Add("Sfx Prod", typeof(string));
            newTable.Columns.Add("Premium", typeof(double));
            newTable.Columns.Add("mmyy", typeof(string));
            newTable.Columns.Add("Rate %", typeof(double));
            newTable.Columns.Add("Rate", typeof(double));
            newTable.Columns.Add("Rate2", typeof(string));
            newTable.Columns.Add("Code", typeof(string));
            newTable.Columns.Add("Commission", typeof(double));
            newTable.Columns.Add("Renewal", typeof(double));
            newTable.Columns.Add("Issue Date", typeof(string));
            newTable.Columns.Add("processed", typeof(bool));
            DataRow newRow;



            for (int i = 1; i < commRows.Count; i++) {
                newRow = newTable.NewRow();
                newRow["Policy"] = commRows[i][0];
                newRow["Fullname"] = commRows[i][1];
                newRow["Sfx Prod"] = commRows[i][4];
                newRow["Premium"] = commRows[i][7];
                DateTime tDates;
                DateTime.TryParse(commRows[i][6].ToString(), out tDates);
                newRow["mmyy"] = tDates.ToShortDateString();
                newRow["Rate %"] = (double)commRows[i][11] * 100;
                newRow["Rate"] = commRows[i][12];
                newRow["Rate2"] = "";
                string code = commRows[i][8].ToString().Trim();
                newRow["Code"] = code;

                if (code == "Renewal") {
                    newRow["Commission"] = 0;
                    newRow["Renewal"] = commRows[i][13];
                }
                else {
                    newRow["Renewal"] = 0;
                    newRow["Commission"] = commRows[i][13];
                }
                DateTime.TryParse(commRows[i][5].ToString(), out tDates);
                newRow["Issue Date"] = tDates.ToShortDateString();
                newRow["processed"] = false;
                newTable.Rows.Add(newRow);
            }
            //Console.ReadLine();

            var cultureInfo = new System.Globalization.CultureInfo("en-US");

            //sort policy asc, premium desc
            DataView temp = new DataView(newTable);
            temp.Sort = "Policy ASC, Premium Desc";
            newTable = temp.ToTable();

            for (int i = 1; i < advRows.Count; i++) {
                string policy = advRows[i][0].ToString(); ;
                Double commTotal = (double)advRows[i][11];
                commTotal = (-1) * Math.Abs(commTotal);
                commTotal += (double)advRows[i][12];
                commTotal += (double)advRows[i][10];

                var cMatches = from myRow in newTable.AsEnumerable()
                               where myRow.Field<string>("Policy") == policy
                               && myRow.Field<bool>("processed") == false
                               select myRow;

                if (cMatches.Count() == 0) {
                    MessageBox.Show("Could not find advance match for " + policy, "MATCH NOT FOUND", MessageBoxButtons.OK);
                }
                else {
                    if ((string)cMatches.First()["Code"] == "Renewal") {
                        double cRen = (double)cMatches.First()["Renewal"];
                        cMatches.First()["Renewal"] = (cRen + commTotal);
                        cMatches.First()["Processed"] = true;
                    }
                    else {
                        double cTemp = (double)cMatches.First()["Commission"];
                        cMatches.First()["Commission"] = (cTemp + commTotal);
                        cMatches.First()["Processed"] = true;
                    }
                }

            }

            newTable.Columns.Remove("Processed");
            temp = new DataView(newTable);
            temp.RowFilter = "Commission <> 0 OR Renewal <> 0";
            newTable = temp.ToTable();

            DataTable uniques = newTable.Clone();

            while(newTable.Rows.Count > 0) {
                IEnumerable<DataRow> totals = from row in newTable.AsEnumerable()
                             where row.Field<string>("Policy") == newTable.Rows[0].Field<string>("Policy") &&
                             row.Field<double>("Premium") == newTable.Rows[0].Field<double>("Premium")
                             select row;
                if(totals.Count() == 1) {
                    uniques.Rows.Add(newTable.Rows[0].ItemArray);
                    newTable.Rows.Remove(newTable.Rows[0]);
                } else {
                    DataRow  tRow = uniques.NewRow();
                    tRow["Policy"] = newTable.Rows[0]["Policy"];
                    tRow["Fullname"] = newTable.Rows[0]["Fullname"];
                    tRow["Sfx Prod"] = newTable.Rows[0]["Sfx Prod"];
                    tRow["Premium"] = newTable.Rows[0]["Premium"];
                    tRow["mmyy"] = newTable.Rows[0]["mmyy"];
                    //tRow["Rate %"] = newTable.Rows[0]["Rate %"];
                    tRow["Rate %"] = totals.Sum(g => g.Field<double>("Rate %"));
                    tRow["Rate"] = newTable.Rows[0]["Rate"];
                    tRow["Rate2"] = newTable.Rows[0]["Rate2"];
                    tRow["Code"] = newTable.Rows[0]["Code"];
                    tRow["Commission"] = totals.Sum(g => g.Field<double>("Commission"));
                    tRow["Renewal"] = totals.Sum(g => g.Field<double>("Renewal"));
                    tRow["Issue Date"] = newTable.Rows[0]["Issue Date"];
                    uniques.Rows.Add(tRow);

                    var myList = totals.ToList();
                    foreach (var thisRow in myList) {
                        newTable.Rows.Remove(thisRow);
                    }                    
                }
            }
            temp = new DataView(uniques);
            temp.RowFilter = "Commission <> 0 OR Renewal <> 0";
            uniques = temp.ToTable();
            stream.Close();

            writeToExcel(uniques);
        }
        public void writeToExcel(DataTable dt) {
            string outFile = "";
            try {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;
                oXL.UserControl = false;
                oXL.DisplayAlerts = false;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Policy"; //A
                oSheet.Cells[1, 2] = "Fullname"; //B
                oSheet.Cells[1, 3] = "Sfx Prod"; //C
                oSheet.Cells[1, 4] = "Premium"; //D
                oSheet.Cells[1, 5] = "mmyy"; //E
                oSheet.Cells[1, 6] = "Rate %"; //F
                oSheet.Cells[1, 7] = "Rate"; //G
                oSheet.Cells[1, 8] = "Rate2"; //H
                oSheet.Cells[1, 9] = "Code"; //I
                oSheet.Cells[1, 10] = "Commission"; //J
                oSheet.Cells[1, 11] = "Renewal"; //K
                oSheet.Cells[1, 12] = "Issue Date"; //L

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "L1").Font.Bold = true;
                oSheet.get_Range("A1", "L1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                for (int i = 0; i < dt.Rows.Count; i++) {
                    object[] outPut = dt.Rows[i].ItemArray;
                    oSheet.get_Range("A" + (i + 2), "L" + (i + 2)).Value2 = outPut;
                }
                oRng = oSheet.get_Range("A1", "L1");
                oRng.EntireColumn.AutoFit();
                oXL.Visible = false;
                oXL.UserControl = false;

                outFile = GetSavePath();

                oWB.SaveAs(outFile,
                    56, //Seems to work better than default excel 16
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);

                //System.Diagnostics.Process.Start(outFile);
            }
            catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message, "Error");
            }
            finally {
                if (oWB != null)
                    oWB.Close();
                if (File.Exists(outFile))
                    System.Diagnostics.Process.Start(outFile);
            }
        }

        public string GetSavePath() {

            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.InitialDirectory = "H:\\Desktop\\";
            saveFileDialog1.Filter = "xlsx|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = fileName.Replace(".xls", "_out.xls");

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) {
                return saveFileDialog1.FileName;
            }
            //else System.Windows.Application.Exit();
            return "";
        }
    }
}