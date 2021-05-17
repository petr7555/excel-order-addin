using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

            foreach (Excel.Worksheet displayWorksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                RibbonDropDownItem item =  Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = displayWorksheet.Name; 
                table1Combo.Items.Add(item);
            }

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("hello");
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void comboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
