using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RvRTest
{
    public partial class RvRMain : Form
    {
        /* Declare global variable */
        int Channel_2_4G_Bound = 14 ;
        int[] Channel_5G_20M = { 36, 40, 44, 48, 149, 153, 157, 161, 165};
        int[] Channel_5G_40M = { 36, 40, 44, 48, 149, 153, 157, 161};
        int[] Channel_5G_80M = { 36, 40, 44, 48, 149, 153, 157, 161};
        string[] Mode_2_4G = {"20M", "40M"};
        string[] Mode_5G = {"20M", "40M", "80M"};
        int[] Attenuation_2_4G; 
        int[] Attenuatin_5G;
        ListBox gpib_IP;
        int gpib_no;

        /* Declare variable to indicate Delete button exist in the datagridview1 column 0*/
        bool hasDeleteButton = false;


        private void InitTestCondition()
        {
            rbtn_TestCondition_24G.Checked = true;
            txt_TestCondition_SSID.Text = "RvRTest24";

            SetupDataGridView();
            //SetupAttenuation();

            /* Add 2.4G mode 20M, 40M */
            /*
            dud_TestCondition_Mode.Items.Add("20M");
            dud_TestCondition_Mode.Items.Add("40M");
            dud_TestCondition_Mode.SelectedIndex = 0;
            cbox_TestCondition_Mode.Items.Add("20M");
            cbox_TestCondition_Mode.Items.Add("40M");
            cbox_TestCondition_Mode.SelectedIndex = 0;
            */

            /* Add 2.4G channel 1-13 */
            /*
            for (int i = 1; i < Channel_2_4G_Bound+1; i++)
            {
                dud_TestCondition_Channel.Items.Add(i.ToString());
                cbox_TestCondition_Channel.Items.Add(i.ToString());
            }
            dud_TestCondition_Channel.SelectedIndex = 0;
            cbox_TestCondition_Channel.SelectedIndex = 0;
            */
        }

        private void btn_TestCondition_AddSetting_Click(object sender, EventArgs e)
        {
            /* Check all data has value */
            if (txt_TestCondition_SSID.Text == "") txt_TestCondition_SSID.Text = "RvRTest";

            /* Check at least one tst file has selected */
            if (txt_TestCondition_TXTst.Text == "" && txt_TestCondition_RXTst.Text == "" && txt_TestCondition_BITst.Text == "")
            {
                MessageBox.Show("Selece at least one TST File!!", "Warning");
                return;
            }

            /* Check if attenuation stop value < start value */
            if(nud_TestCondition_AtteuationMax.Value <= nud_TestCondition_AtteuationMin.Value)
            {
                MessageBox.Show("Attenuation Stop Value must greater than Start Value.", "Warning");
                    return ;
            }

            /* Check if Delete button exist in the datagridview */
            if (hasDeleteButton)
            {
                hasDeleteButton = false;
                dataGridView1.Columns.RemoveAt(0);
            }

            /* Add data to datagridview */
            string a1 = rbtn_TestCondition_24G.Checked? "2.4G":"5G"; //Band
            string a2 = cbox_TestCondition_Mode.SelectedItem.ToString(); //Mode
            string a3 = cbox_TestCondition_Channel.SelectedItem.ToString(); //Channel
            string a4 = txt_TestCondition_SSID.Text;//SSID
            string a5 = nud_TestCondition_AtteuationMin.Value.ToString(); //Attenuation min
            string a6 = nud_TestCondition_AtteuationMax.Value.ToString() ; //Attenuation max
            string a7 = nud_TestCondition_AtteuationStep.Value.ToString() ; //Attenuation step
            string a8 = txt_TestCondition_TXTst.Text; //TX tst file
            string a9 = txt_TestCondition_RXTst.Text; //RX tst File
            string a10 = txt_TestCondition_BITst.Text; //Bi-Direction tst File*/

            string[] row = {a1, a2, a3, a4, a5, a6, a7, a8, a9, a10 };            
            dataGridView1.Rows.Add(row);
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
        
        }

        private void btn_TestCondition_Edit_Click(object sender, EventArgs e)
        {
            hasDeleteButton = true;
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            btn.Width = 66;
            btn.HeaderText = "Delete";
            btn.Text = "Delete";
            btn.Name = "btnDelete";
            btn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            btn.UseColumnTextForButtonValue = true;

            dataGridView1.Columns.Insert(0, btn);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "btnDelete" && hasDeleteButton == true)
            {
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }

        private void txt_TestCondition_TXTst_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "Tx.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txt_TestCondition_TXTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open TX TST file: " + ex.Message);
                }
            }    
        }

        private void txt_TestCondition_RXTst_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "Rx.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txt_TestCondition_RXTst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open RX TST file: " + ex.Message);
                }
            }    
        }

        private void txt_TestCondition_BITst_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "tst";
            openFileDialog1.Title = "Choose Test tst file ";
            openFileDialog1.Filter = "tst files(*.txt)|*.tst|All files(*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "TX-RX.tst";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    txt_TestCondition_BITst.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not open Bi-Direction TST file: " + ex.Message);
                }
            }     
        }


        private void rbtn_TestCondition_24G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtn_TestCondition_24G.Checked == true)
            {        
                cbox_TestCondition_Channel.Items.Clear();
                cbox_TestCondition_Mode.Items.Clear();

               int i;
                for(i=0 ;i< Mode_2_4G.Length ; i++)
                {
                    cbox_TestCondition_Mode.Items.Add(Mode_2_4G[i]);
                }
                
                cbox_TestCondition_Mode.SelectedIndex = 0;


                for (i = 1; i < Channel_2_4G_Bound+1; i++)
                {                   
                    cbox_TestCondition_Channel.Items.Add(i.ToString());
                }
                
                cbox_TestCondition_Channel.SelectedIndex = 0;
                
            }
        }

        private void rbtn_TestCondition_5G_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtn_TestCondition_5G.Checked == true)
            {               
                cbox_TestCondition_Channel.Items.Clear();
                cbox_TestCondition_Mode.Items.Clear();

                int i;
                for (i = 0; i < Mode_5G.Length; i++)
                {
                    cbox_TestCondition_Mode.Items.Add(Mode_5G[i]);
                }
                cbox_TestCondition_Mode.SelectedIndex = 0;

                for (i = 0; i < Channel_5G_20M.Length; i++)
                {
                    cbox_TestCondition_Channel.Items.Add(Channel_5G_20M[i]);
                }               
                cbox_TestCondition_Channel.SelectedIndex = 0;
            }
        }

        private void cbox_TestCondition_Mode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbtn_TestCondition_5G.Checked == true)
            {
                cbox_TestCondition_Channel.Items.Clear();
                int i;
                
                if (cbox_TestCondition_Mode.SelectedItem.ToString() == "40M")
                {
                    for (i = 0; i < Channel_5G_40M.Length; i++)
                    {
                        cbox_TestCondition_Channel.Items.Add(Channel_5G_40M[i]);
                    }
                    cbox_TestCondition_Channel.SelectedIndex = 0;

                }
                if (cbox_TestCondition_Mode.SelectedItem.ToString() == "80M")
                {
                    for (i = 0; i < Channel_5G_80M.Length; i++)
                    {
                        cbox_TestCondition_Channel.Items.Add(Channel_5G_80M[i]);
                    }
                    cbox_TestCondition_Channel.SelectedIndex = 0;
                }                
            }
        }

        public void SetupDataGridView()
        {
            dataGridView1.ColumnCount = 10;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

            dataGridView1.Name = "Test Condition Setting";
            //dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        
            dataGridView1.Columns[0].Name = "Band";
            dataGridView1.Columns[1].Name = "Mode";
            dataGridView1.Columns[2].Name = "Channel";
            dataGridView1.Columns[3].Name = "SSID";
            dataGridView1.Columns[4].Name = "Start";
            dataGridView1.Columns[5].Name = "Stop";
            dataGridView1.Columns[6].Name = "Step";
            dataGridView1.Columns[7].Name = "TX tst File";
            dataGridView1.Columns[8].Name = "RX tst File";
            dataGridView1.Columns[9].Name = "BI-Dir tst File";
            
            /*
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns[3].Width = 80;
            dataGridView1.Columns[4].Width = 50;
            dataGridView1.Columns[5].Width = 50; 
            dataGridView1.Columns[6].Width = 50;
            dataGridView1.Columns[7].Width = 120;
            dataGridView1.Columns[8].Width = 120;
            dataGridView1.Columns[9].Width = 120;

             */
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


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


        private void DefaultAttenuation()
        {
            Attenuation_2_4G =new int[] { 1, 2, 4, 4, 10, 20, 30, 30 };
            Attenuatin_5G =new int[] { 1, 2, 4, 4, 10, 20, 40, 40 };
            
        
        }







    }
}
