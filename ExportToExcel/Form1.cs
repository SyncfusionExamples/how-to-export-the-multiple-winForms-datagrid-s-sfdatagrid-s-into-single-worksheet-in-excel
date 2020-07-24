using SfDataGrid_Demo;
using Syncfusion.WinForms.Controls;
using Syncfusion.WinForms.DataGrid;
using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;
using Syncfusion.WinForms.DataGrid.Enums;
using System.Drawing;
using Syncfusion.WinForms.DataGridConverter;
using Syncfusion.XlsIO;
using System.IO;

namespace SfDataGrid_Demo
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Constructor

        /// <summary>
        /// Initializes the new instance for the Form.
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            sfDataGrid1.DataSource = new OrderInfoCollection().OrdersListDetails;

            sfDataGrid2.DataSource = new OrderInfoCollection().OrdersListDetails1;
        }
        #endregion

        private void OnExportButton_Click(object sender, EventArgs e)
        {
            var options = new ExcelExportingOptions();
            
            var excelEngine = sfDataGrid1.ExportToExcel(sfDataGrid1.View, options);
            var workBook1 = excelEngine.Excel.Workbooks[0];
            var worksheet1 = workBook1.Worksheets[0];

            excelEngine = sfDataGrid2.ExportToExcel(sfDataGrid2.View, options);
            var workBook2 = excelEngine.Excel.Workbooks[0];
            var worksheet2 = workBook2.Worksheets[0];

            var columnCount = sfDataGrid2.Columns.Count;
            //Merge the One SfDataGrid WorkSheet into the other SfDataGrid WorkSheet
            worksheet2.UsedRange.CopyTo(worksheet1[1, columnCount + 1]);
            workBook1.SaveAs("sample.xlsx");

            SaveFileDialog saveFilterDialog = new SaveFileDialog
            {
                FilterIndex = 2,
                Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx|Excel 2013 File(*.xlsx)|*.xlsx"
            };

            if (saveFilterDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (Stream stream = saveFilterDialog.OpenFile())
                {
                    if (saveFilterDialog.FilterIndex == 1)
                        workBook1.Version = ExcelVersion.Excel97to2003;
                    else if (saveFilterDialog.FilterIndex == 2)
                        workBook1.Version = ExcelVersion.Excel2010;
                    else
                        workBook1.Version = ExcelVersion.Excel2013;
                    workBook1.SaveAs(stream);
                }

                //Message box confirmation to view the created workbook.
                if (MessageBox.Show(this, "Do you want to view the workbook?", "Workbook has been created",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {

                    //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                    System.Diagnostics.Process.Start(saveFilterDialog.FileName);
                }
            }
        }
    }
}
