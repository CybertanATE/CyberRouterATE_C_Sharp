///---------------------------------------------------------------------------------------
///  This code was created by CyberTan Jin Wang.
///  File           : RouterTestResultTable.cs
///  Update         : 2016-09-21
///  Description    : This is router test result table
///  Modified       : 2016-09-21 Initial version  
///                   
///---------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;

namespace CyberRouterATE
{
    public partial class RouterTestMain : Form
    {
        /* Declare global variable */        

        /* Declare variable to indicate Delete button exist in the datagridview1 column 0*/
        
        private void InitRouterTestResultTable(int columncount, string[] headerText, string condition)
        {
            labRvRResultInfoTitle.Text = condition;
            SetupRvRTestResultDataGridView(columncount, headerText);            
        }

        public void SetupRvRTestResultDataGridView(int columncount, string[] headerText)
        {
            dgvRvRTestResultTable.ColumnCount = columncount;
            dgvRvRTestResultTable.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dgvRvRTestResultTable.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dgvRvRTestResultTable.Name = "Test Result Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgvRvRTestResultTable.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 0; i < columncount; i++)
            {
                dgvRvRTestResultTable.Columns[i].Name = headerText[i];
            }

            //dgvRouterTestResultTable.Columns[0].Name = "Band";
            //dgvRouterTestResultTable.Columns[1].Name = "Mode";
            //dgvRouterTestResultTable.Columns[2].Name = "Channel";
            //dgvRouterTestResultTable.Columns[3].Name = "SSID";
            //dgvRouterTestResultTable.Columns[4].Name = "Security";    
            //dgvRouterTestResultTable.Columns[5].Name = "Key";                     
            //dgvRouterTestResultTable.Columns[6].Name = "Start";
            //dgvRouterTestResultTable.Columns[7].Name = "Stop";
            //dgvRouterTestResultTable.Columns[8].Name = "Step";
            //dgvRouterTestResultTable.Columns[9].Name = "TX tst File";
            //dgvRouterTestResultTable.Columns[10].Name = "RX tst File";
            //dgvRouterTestResultTable.Columns[11].Name = "BI-Dir tst File";

            dgvRvRTestResultTable.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvRvRTestResultTable.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvRvRTestResultTable.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False; //標題列換行, true -換行, false-不換行
            dgvRvRTestResultTable.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//標題列置中

            //dgvRouterTestResultTable.Columns[0].Width = 80;
            //dgvRouterTestResultTable.Columns[1].Width = 80;
            //dgvRouterTestResultTable.Columns[2].Width = 80;
            //dgvRouterTestResultTable.Columns[3].Width = 120;
            //dgvRouterTestResultTable.Columns[4].Width = 80;
            //dgvRouterTestResultTable.Columns[5].Width = 80;
            //dgvRouterTestResultTable.Columns[6].Width = 80;
            //dgvRouterTestResultTable.Columns[7].Width = 80;
            //dgvRouterTestResultTable.Columns[8].Width = 80;
            //dgvRouterTestResultTable.Columns[9].Width = 120;
            //dgvRouterTestResultTable.Columns[10].Width = 120;
            //dgvRouterTestResultTable.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvRvRTestResultTable.Columns[columncount-1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;


            dgvRvRTestResultTable.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            //dgvRouterTestResultTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; //Use the column width setting
            dgvRvRTestResultTable.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            /*
            string[] row = { (dataGridView1.NewRowIndex + 1).ToString(), numericUpDown1.Value.ToString(), numericUpDown2.Value.ToString(), numericUpDown3.Value.ToString(), textBox1.Text, textBox2.Text };
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoResizeColumns();

            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;
            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Rows.Add(row);
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
            */
        }      
    }
}
