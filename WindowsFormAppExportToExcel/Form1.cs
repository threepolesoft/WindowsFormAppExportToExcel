using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormAppExportToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<stdList> stdList = new List<stdList>();

        private void Form1_Load(object sender, EventArgs e)
        {
            stdList.Add(new stdList
            {
                id = "181967001",
                name = "Name 1",
                roll = "001",
                mob = "01758545548"
            });

            stdList.Add(new stdList
            {
                id = "181967002",
                name = "Name 2",
                roll = "002",
                mob = "01758545548"
            });

            stdList.Add(new stdList
            {
                id = "181967003",
                name = "Name 3",
                roll = "003",
                mob = "01758545548"
            });

            stdList.Add(new stdList
            {
                id = "181967004",
                name = "Name 4",
                roll = "004",
                mob = "01758545548"
            });

            stdList.Add(new stdList
            {
                id = "181967005",
                name = "Name 5",
                roll = "005",
                mob = "01758545548"
            });

            stdList.Add(new stdList
            {
                id = "181967006",
                name = "Name 6",
                roll = "006",
                mob = "01758545548"
            });

            dataGridView1.DataSource = stdList.ToList();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                List<stdList> list = ((DataPara)e.Argument).stdList;
                string fileName = ((DataPara)e.Argument).fileName;


                Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
                Worksheet ws = (Worksheet)excel.ActiveSheet;
                excel.Visible = false;

                int index = 1;
                int procee = list.Count;


                ws.Cells[1, 1] = "Id";
                ws.Cells[1, 2] = "Name";
                ws.Cells[1, 3] = "Roll";
                ws.Cells[1, 4] = "Mob.";

                foreach (stdList item in list)
                {
                    if (!backgroundWorker1.CancellationPending)
                    {
                        backgroundWorker1.ReportProgress(index++ * 100 / procee);

                        ws.Cells[index, 1] = item.id;
                        ws.Cells[index, 2] = item.name;
                        ws.Cells[index, 3] = item.roll;
                        ws.Cells[index, 4] = item.mob;


                    }
                }

                //ws.SaveAs(fileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange,
                //            XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);



                ws.SaveAs(fileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlExclusive,
                    XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                excel.Quit();

                MessageBox.Show("Done");
            }
            catch (Exception)
            {

                excel.Quit();
                MessageBox.Show("Fail");
            }

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            progressBar1.Update();
        }


        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error==null)
            {
                Thread.Sleep(100);
            }
        }

        struct DataPara
        {
           public  List<stdList> stdList;
            public string fileName { get; set; }
        }

        DataPara _inputDataPara;

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy)
            {
                return;
            }
            else
            {
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xlsx" })
                {
                    if (sfd.ShowDialog()==DialogResult.OK)
                    {
                        _inputDataPara.fileName = sfd.FileName;
                        _inputDataPara.stdList = stdList.ToList();

                        progressBar1.Minimum = 0;
                        progressBar1.Value = 0;
                        backgroundWorker1.RunWorkerAsync(_inputDataPara);
                    }
                }
            }
        }
    }
}
