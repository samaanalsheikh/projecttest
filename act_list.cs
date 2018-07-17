using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using xls = Microsoft.Office.Interop.Excel;
using System.Threading;     // For setting the Localization of the thread to fit
using ProgressBarExample;
using GridPrintPreviewLib;
using System.Data.OleDb;
using System.Drawing.Printing;

using System.Data.SqlClient;


namespace product_plan
{
    public partial class act_list : Form
    {
        raw_material_card raw_mat;

        public act_list(raw_material_card raw_mat)
        {
            InitializeComponent();
            this.raw_mat = raw_mat;
        }

        public act_list()
        {
            InitializeComponent();
        }

        private void act_list_Load(object sender, EventArgs e)
        {
            ///////////////// Dg Mat Draw /////////////

            dg_mat.Columns.Add("Code No.", "Code No.");
            dg_mat.Columns.Add("Material Name", "Material Name");
            dg_mat.Columns.Add("Warehouse Stock", "Warehouse Stock");
            dg_mat.Columns.Add("Stock", "Stock");
            dg_mat.Columns.Add("Unit", " Unit");
            dg_mat.Columns.Add("id", "id");
            dg_mat.Columns.Add("Total In", "Total In");
            dg_mat.Columns.Add("Total out", "Total out");
            dg_mat.Columns[0].Width = 60;
            dg_mat.Columns[1].Width = 230;
            dg_mat.Columns[2].Width = 80;
            dg_mat.Columns[4].Width = 30;

            dg_mat.Columns[5].Visible = false;
            dg_mat.Columns[2].Visible = false;


            ///////////////// Dg prod Draw /////////////////

            dg_prod.Columns.Add("Id", "Id");
            dg_prod.Columns.Add("Code No.", "Code No.");
            dg_prod.Columns.Add("Product Name", "Product Name");
            dg_prod.Columns.Add("Requirement", "Requirement");
            dg_prod.Columns.Add("Num of batches", "Num of batches");
            dg_prod.Columns[0].Width = 35;
            dg_prod.Columns[1].Width = 80;
            dg_prod.Columns[2].Width = 180;
            dg_prod.Columns[3].Width = 80;

            dg_prod.Columns[4].Visible = false;

            ///////////  connection  ////////////// 

            conn c = new conn();
            c.connect();

            ///////////////  connection UT  //////////////// 

            c.connect_UT();

            txt_user_name.Text = System.IO.File.ReadAllText(@"user.fb");
            string form_name = this.Name.ToString();
            string form_id = c.get_form_id_by_name(form_name).ToString();
            txt_user_id.Text = c.get_user_id_by_name(txt_user_name.Text);

            try { if (c.get_component_name_from_permissions_by_component_id("1152", txt_user_id.Text, form_id) == "1") { toolStrip_Supplier_act.Enabled = true; } else { toolStrip_Supplier_act.Enabled = false; } }
            catch (Exception) { toolStrip_Supplier_act.Enabled = false; }
            try { if (c.get_component_name_from_permissions_by_component_id("5401", txt_user_id.Text, form_id) == "1") { act_list_btn_quantity.Enabled = true; } else { act_list_btn_quantity.Enabled = false; } }
            catch (Exception) { act_list_btn_quantity.Enabled = false; }
            try { if (c.get_component_name_from_permissions_by_component_id("5402", txt_user_id.Text, form_id) == "1") { txt_move_to_card_pointer.Text = "1"; } else { txt_move_to_card_pointer.Text = "0"; } }
            catch (Exception) { txt_move_to_card_pointer.Text = "0"; }


            ///////////  Fill ACT mat  ////////////// 

            DataSet ds_ma_act = new DataSet();
            c.Fill_mat_act(ds_ma_act);

            for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
            {
                dg_mat.Rows.Add();
                dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                dg_mat.Rows[x].Cells[3].Value = ".";
                dg_mat.Rows[x].Cells[6].Value = ".";
                dg_mat.Rows[x].Cells[7].Value = ".";
                dg_mat.Rows[x].Cells[3].Style.BackColor = Color.LightGray;
            }
            dg_mat.Rows.OfType<DataGridViewRow>().First().Selected = false;
            dg_prod.Rows.OfType<DataGridViewRow>().First().Selected = false;
        }

        private void dg_mat_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            raw_material_card raw_mat = new raw_material_card();
            try
            {
                if (txt_move_to_card_pointer.Text == "1")
                {
                    raw_mat.txt_mat_typ.Text = "1";
                    raw_mat.txt_mat_id.Text = dg_mat.Rows[e.RowIndex].Cells[5].Value.ToString();
                    raw_mat.Show();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error 1757 ");
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            txt_search_by_code.Text = "";
            dg_mat.Rows.Clear();

            ///////////  connection  ////////////// 

            conn c = new conn();
            c.connect();

            ///////////  Fill ACT mat  ////////////// 

            DataSet ds_ma_act = new DataSet();
            c.Fill_mat_act_name_LIKE(ds_ma_act, txt_search_by_name.Text);

            for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
            {
                dg_mat.Rows.Add();
                dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                dg_mat.Rows[x].Cells[6].Value = ".";
                dg_mat.Rows[x].Cells[7].Value = ".";
                dg_mat.Rows[x].Cells[3].Value = ".";

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {


        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txt_mat_id.Text == "")
            {
                MessageBox.Show("  !....الرجاء اختيار مادة أولا ");
            }
            else
            {
                receiving_supplier r_supp = new receiving_supplier();
                r_supp.txt_mat_type.Text = "1";

                r_supp.txt_mat_id.Text = txt_mat_id.Text.ToString();
                r_supp.Show();

            }
        }

        private void txt_search_by_code_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txt_search_by_name.Text = "";
                dg_mat.Rows.Clear();

                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill ACT mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act_by_id(ds_ma_act, int.Parse(txt_search_by_code.Text));

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                    dg_mat.Rows[x].Cells[6].Value = ".";
                    dg_mat.Rows[x].Cells[7].Value = ".";
                    dg_mat.Rows[x].Cells[3].Value = ".";

                }
            }
            catch (Exception)
            {
                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill ACT mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act(ds_ma_act);

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                }
            }
        }

        private void btn_show_quantity_Click(object sender, EventArgs e)
        {

        }

        private void txt_search_by_code_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                dg_mat.Select();
            }
        }

        private void txt_search_by_name_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                dg_mat.Select();
            }
        }

        private void btn_show_quantity_Click_1(object sender, EventArgs e)
        {


        }

        private void btn_serch_Click(object sender, EventArgs e)
        {
            dg_mat.Rows.Clear();

            ///////////  connection  ////////////// 

            conn c = new conn();
            c.connect();

            ///////////  Fill ACT mat  ////////////// 

            DataSet ds_ma_act = new DataSet();
            c.Fill_mat_act(ds_ma_act);

            for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
            {
                dg_mat.Rows.Add();
                dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();

                float total_in = float.Parse(c.sum_quantity_receiving_report_by_mat_type_mat_id_and_date_less_than("1", dg_mat.Rows[x].Cells[5].Value.ToString(), dateTimePicker1.Value.ToString("dd/MM/yyyy"), "r").ToString());
                float total_out = float.Parse(c.sum_quan_by_mat_id_and_mat_type_and_dep_id_and_date_less_than("1", dg_mat.Rows[x].Cells[5].Value.ToString(), "r", dateTimePicker1.Value.ToString("dd/MM/yyyy")));
                dg_mat.Rows[x].Cells[3].Value = (total_in - total_out).ToString();
                dg_mat.Rows[x].Cells[6].Value = total_in.ToString();
                dg_mat.Rows[x].Cells[7].Value = total_out.ToString();
                decimal dd = decimal.Parse(dg_mat.Rows[x].Cells[2].Value.ToString());
                dg_mat.Rows[x].Cells[2].Value = string.Format("{0:n3}", dd);
                decimal ddd = decimal.Parse(dg_mat.Rows[x].Cells[3].Value.ToString());
                dg_mat.Rows[x].Cells[3].Value = string.Format("{0:n3}", ddd);

            }
        }

        private void toolStrip_Supplier_act_Click(object sender, EventArgs e)
        {

        }

        private void txt_search_by_name_TextChanged(object sender, EventArgs e)
        {
            try
            {
                txt_search_by_code.Text = "";
                dg_mat.Rows.Clear();

                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill ACT mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act_name_LIKE(ds_ma_act, txt_search_by_name.Text);

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                    dg_mat.Rows[x].Cells[6].Value = ".";
                    dg_mat.Rows[x].Cells[7].Value = ".";
                    dg_mat.Rows[x].Cells[3].Value = ".";

                }
            }
            catch (Exception)
            {
                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill AS mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act(ds_ma_act);

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                }
            }
        }

        private void txt_search_by_code_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                txt_search_by_name.Text = "";
                dg_mat.Rows.Clear();

                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill ACT mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act_by_id(ds_ma_act, int.Parse(txt_search_by_code.Text));

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                    dg_mat.Rows[x].Cells[6].Value = ".";
                    dg_mat.Rows[x].Cells[7].Value = ".";
                    dg_mat.Rows[x].Cells[3].Value = ".";

                }
            }
            catch (Exception)
            {
                ///////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                ///////////  Fill AS mat  ////////////// 

                DataSet ds_ma_act = new DataSet();
                c.Fill_mat_act(ds_ma_act);

                for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
                {
                    dg_mat.Rows.Add();
                    dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                    dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                    dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                    dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();
                }
            }
        }
        private bool isProcessRunning = false;
        private void button4_Click(object sender, EventArgs e)
        {
            ProgressDialog progressDialog = new ProgressDialog();
            progressDialog.progressBar1.Maximum = dg_mat.Rows.Count;
            progressDialog.label1.Text = "الرجاء الإنتظار , جاري حساب الكميات";

            conn c = new conn();
            c.connect();

            if (isProcessRunning)
            {
                MessageBox.Show("A process is already running.");
                return;
            }


            Thread backgroundThread = new Thread(
                new ThreadStart(() =>
                {

                    isProcessRunning = true;

                    int n = 0;

                    ///////////  Fill Quantity  ////////////// 

                    for (int x = 0; x < dg_mat.Rows.Count - 1; x++)
                    {

                        double total_in = double.Parse(c.sum_quantity_receiving_report_by_mat_type_mat_id("1", dg_mat.Rows[x].Cells[5].Value.ToString(), "r").ToString());
                        double total_out = double.Parse(c.sum_quan_by_mat_id_and_mat_type_and_dep_id("1", dg_mat.Rows[x].Cells[5].Value.ToString(), "r"));
                        dg_mat.Rows[x].Cells[3].Value = string.Format("{0:n5}", (total_in - total_out));
                        dg_mat.Rows[x].Cells[6].Value = string.Format("{0:n5}", total_in);
                        dg_mat.Rows[x].Cells[7].Value = string.Format("{0:n5}", total_out);
                        progressDialog.progressBar1.BeginInvoke(new Action(() => progressDialog.progressBar1.Value = n));
                        n++;
                    }
                    if (progressDialog.InvokeRequired)
                        progressDialog.BeginInvoke(new Action(() => progressDialog.Close()));
                    isProcessRunning = false;
                }
            ));

            backgroundThread.Start();
            progressDialog.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (txt_mat_id.Text == "")
            {
                MessageBox.Show("  !....الرجاء اختيار مادة أولا ");
            }
            else
            {
                receiving_supplier r_supp = new receiving_supplier();
                r_supp.txt_mat_type.Text = "1";

                r_supp.txt_mat_id.Text = txt_mat_id.Text.ToString();
                r_supp.Show();

            }
        }

        private void act_list_btn_xlxs_Click(object sender, EventArgs e)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.ActiveWindow.DisplayGridlines = false;

            try
            {
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet2"];
                worksheet.Delete();
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet3"];
                worksheet.Delete();
            }
            catch (Exception) { }

            try
            {
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "ACT List";
                worksheet.DisplayRightToLeft = false;

                worksheet.get_Range("A1", "Z1000").Font.Name = "Arial";
                worksheet.get_Range("A1", "Z1000").Font.Size = 10;
                worksheet.get_Range("A1", "Z1000").Font.Bold = true;
                worksheet.get_Range("A1", "G1").Merge(Type.Missing);
                worksheet.get_Range("A1", "G1").RowHeight = 17.5;
                worksheet.get_Range("A2", "G2").RowHeight = 15;
                worksheet.get_Range("A3", "G3").RowHeight = 21;
                worksheet.get_Range("A4", "G4").RowHeight = 5;
                worksheet.Cells[1, 1] = "Raw Materials List";
                worksheet.get_Range("A1", "G1").Font.Bold = true;
                worksheet.get_Range("A1", "G1").Font.Italic = true;
                worksheet.get_Range("A1", "G1").Font.Size = 10;
                worksheet.get_Range("A1", "G1").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                worksheet.get_Range("A1", "G1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                worksheet.get_Range("A6", "Z1000").Font.Name = "Arial";
                worksheet.get_Range("A6", "Z1000").Font.Size = 10;
                worksheet.get_Range("A6", "Z1000").Font.Bold = false;

                worksheet.get_Range("A5", "G1000").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("D5", "D1000").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("E5", "E1000").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("C5", "C1000").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                worksheet.Cells[2, 1] = "Stock Balance at Date: " + dateTimePicker1.Text;

                worksheet.get_Range("A2", "G2").Font.Bold = false;
                worksheet.get_Range("A2", "G2").Font.Size = 10;
                worksheet.get_Range("A5", "G5").RowHeight = 20;
                worksheet.get_Range("A5", "G5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("A5", "G5").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                worksheet.get_Range("G2", "G2").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;

                worksheet.get_Range("A3", "G3").Merge(Type.Missing);
                worksheet.get_Range("A3", "G3").Font.Bold = true;
                worksheet.get_Range("A3", "G3").Font.Size = 11;
                worksheet.get_Range("A3", "G3").Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                worksheet.get_Range("A3", "G3").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("A3", "G3").VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.Cells[3, 1] = "ACT List";

                worksheet.get_Range("A1", "A250").ColumnWidth = 3.7;
                worksheet.get_Range("B1", "B250").ColumnWidth = 9;
                worksheet.get_Range("C1", "C250").ColumnWidth = 35;
                worksheet.get_Range("D1", "D250").ColumnWidth = 8;
                worksheet.get_Range("E1", "E250").ColumnWidth = 4;
                worksheet.get_Range("F1", "F250").ColumnWidth = 8;
                worksheet.get_Range("G1", "G250").ColumnWidth = 8;

                worksheet.Cells[5, 1] = "ID";
                worksheet.Cells[5, 2] = "Code No.";
                worksheet.Cells[5, 3] = "Material Name.";
                worksheet.Cells[5, 4] = "Stock.";
                worksheet.Cells[5, 5] = "Unit";
                worksheet.Cells[5, 6] = "Total In";
                worksheet.Cells[5, 7] = "Total Out";

                int row_count = 1;
                for (int i = 0; i < dg_mat.Rows.Count; i++)
                {
                    if (dg_mat.Rows[i].Cells[1].Value != null)
                    {
                        worksheet.Cells[row_count + 5, 1] = row_count;
                        worksheet.Cells[row_count + 5, 2] = dg_mat.Rows[i].Cells[0].Value;
                        worksheet.Cells[row_count + 5, 3] = dg_mat.Rows[i].Cells[1].Value;
                        worksheet.Cells[row_count + 5, 4] = dg_mat.Rows[i].Cells[3].Value;
                        worksheet.Cells[row_count + 5, 5] = dg_mat.Rows[i].Cells[4].Value;
                        worksheet.Cells[row_count + 5, 6] = dg_mat.Rows[i].Cells[6].Value;
                        worksheet.Cells[row_count + 5, 7] = dg_mat.Rows[i].Cells[7].Value;

                        row_count++;
                    }

                }

                worksheet.get_Range("A5", "G" + (row_count + 4)).Borders.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
                worksheet.get_Range("A" + (row_count + 5), "G" + (row_count + 5)).RowHeight = 8;
                try
                {
                    worksheet.PageSetup.PrintTitleRows = "$5:$5";
                    worksheet.PageSetup.LeftFooter = "&\"Arial Unicode MS\"&8 Page &P of &N ";
                    worksheet.PageSetup.RightFooter = "&\"Arial Unicode MS\"&8 Converted from ERP At:" + DateTime.Now + " By: " + txt_user_name.Text;
                }
                catch (Exception) { }

                string fileName = String.Empty;
                saveFileDialog1.FileName = "Act List Balance-" + dateTimePicker1.Value.ToString("dd-MM-yyyy");
                saveFileDialog1.Filter = "Excel files |*.xls|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    workbook.Close(null, null, null);
                    app.Quit();
                }

                else
                {
                    workbook.Close(false, Type.Missing, Type.Missing);

                }

            }
            catch (Exception) { MessageBox.Show("Error 1874 :Either the file is already open or you have cancel the operation"); }

        }

        private void dg_mat_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = dg_mat.HitTest(e.X, e.Y).RowIndex;
                DataObject d = dg_mat.GetClipboardContent();
                Clipboard.SetDataObject(d);
                contextMenuStrip1.Show(new Point(Cursor.Position.X, Cursor.Position.Y));
            }


        }

        private void dg_prod_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int currentMouseOverRow = dg_prod.HitTest(e.X, e.Y).RowIndex;
                DataObject d = dg_prod.GetClipboardContent();
                Clipboard.SetDataObject(d);
                contextMenuStrip1.Show(new Point(Cursor.Position.X, Cursor.Position.Y));
            }

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                dg_prod.Rows.Clear();
                if (textBox1.Text != "")
                {
                    ///////////  connection  ////////////// 

                    conn c = new conn();
                    c.connect();

                    //////////////////////////////////////

                    DataSet ds_out = new DataSet();


                    //////////////////////////////////////
                    c.Fill_rec_rep_by_dep_id_lotno(ds_out, "r", textBox1.Text);
                    for (int x = 0; x < ds_out.Tables[0].Rows.Count; x++)
                    {

                        dg_prod.Rows.Add();
                        dg_prod.Rows[x].Cells[0].Value = (x + 1).ToString();
                        try
                        {
                            DateTime date = DateTime.Parse(ds_out.Tables[0].Rows[x][14].ToString());
                            dg_prod.Rows[x].Cells[1].Value = date.ToString("dd/MM/yyyy");
                        }
                        catch (Exception)
                        {
                        }
                        dg_prod.Rows[x].Cells[2].Value = ds_out.Tables[0].Rows[x][4].ToString();
                        dg_prod.Rows[x].Cells[1].Value = ds_out.Tables[0].Rows[x][2].ToString();


                    }
                }
            }
            catch (Exception) { dg_prod.Rows.Clear(); }
        }

        private void dg_mat_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                txt_mat_id.Text = dg_mat.Rows[e.RowIndex].Cells[5].Value.ToString();
                dg_prod.Rows.Clear();

                /////////  connection  ////////////// 

                conn c = new conn();
                c.connect();

                /////////  Fill Prod  ////////////// 

                int row_id = 1;
                DataSet ds_prod = new DataSet();
                c.Fill_prod_by_mat_req(ds_prod, 1, int.Parse(dg_mat.Rows[e.RowIndex].Cells[5].Value.ToString()));

                for (int x = 0; x < ds_prod.Tables[0].Rows.Count; x++)
                {
                    dg_prod.Rows.Add();
                    dg_prod.Rows[x].Cells[0].Value = row_id++.ToString();
                    dg_prod.Rows[x].Cells[2].Value = ds_prod.Tables[0].Rows[x][1].ToString();
                    dg_prod.Rows[x].Cells[3].Value = ds_prod.Tables[0].Rows[x][5].ToString();
                }

                for (int z = 0; z < dg_prod.Rows.Count - 1; z++)
                {
                    dg_prod.Rows[z].Cells[1].Value = c.get_prod_code_no_by_id(int.Parse(dg_prod.Rows[z].Cells[2].Value.ToString()));
                    dg_prod.Rows[z].Cells[2].Value = c.get_prod_name_by_id(int.Parse(dg_prod.Rows[z].Cells[2].Value.ToString()));


                }
            }
            catch (Exception) { dg_prod.Rows.Clear(); }
        }

        private void dg_mat_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            txtrow_index.Text = e.RowIndex.ToString();
        }

        private void dg_mat_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {

                raw_material_card raw_mat = new raw_material_card();
                try
                {
                    raw_mat.txt_mat_typ.Text = "1";
                    raw_mat.txt_mat_id.Text = dg_mat.Rows[int.Parse(txtrow_index.Text)].Cells[5].Value.ToString();
                    raw_mat.Show();
                }
                catch (Exception)
                {
                    MessageBox.Show("Error 1757 ");
                }




            }
        }

       
        private void act_list_btn_print_Click(object sender, EventArgs e)
        {
           
                 

    }
        private void printDocument_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Graphics graphic = ev.Graphics;
            foreach (DataRow row in dg_mat.Rows)
            {
                string text = row.ToString(); //or whatever you want from the current row
            graphic.DrawString(text, new Font("Times New Roman", 14, FontStyle.Bold), Brushes.Black, 20, 225); 
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            dg_mat.Rows.Clear();

            ///////////  connection  ////////////// 

            conn c = new conn();
            c.connect();

            ///////////  Fill ACT mat  ////////////// 

            DataSet ds_ma_act = new DataSet();
            c.Fill_mat_act(ds_ma_act);

            for (int x = 0; x < ds_ma_act.Tables[0].Rows.Count; x++)
            {
                dg_mat.Rows.Add();
                dg_mat.Rows[x].Cells[0].Value = ds_ma_act.Tables[0].Rows[x][7].ToString();
                dg_mat.Rows[x].Cells[1].Value = ds_ma_act.Tables[0].Rows[x][1].ToString();
                dg_mat.Rows[x].Cells[2].Value = ds_ma_act.Tables[0].Rows[x][4].ToString();
                dg_mat.Rows[x].Cells[4].Value = ds_ma_act.Tables[0].Rows[x][6].ToString();
                dg_mat.Rows[x].Cells[5].Value = ds_ma_act.Tables[0].Rows[x][0].ToString();

                float total_in = float.Parse(c.sum_quantity_receiving_report_by_mat_type_mat_id_and_date_between("1", dg_mat.Rows[x].Cells[5].Value.ToString(), dt_from.Value.ToString("dd/MM/yyyy"), dt_to.Value.ToString("dd/MM/yyyy")).ToString());
                float total_out = float.Parse(c.sum_quan_by_mat_id_and_mat_type_and_dep_id_and_lot_no_and_date_between("1", dg_mat.Rows[x].Cells[5].Value.ToString(), "r", dt_from.Value.ToString("dd/MM/yyyy"), dt_to.Value.ToString("dd/MM/yyyy")));
                dg_mat.Rows[x].Cells[3].Value = (total_in - total_out).ToString();
                dg_mat.Rows[x].Cells[6].Value = total_in.ToString();
                dg_mat.Rows[x].Cells[7].Value = total_out.ToString();
                decimal dd = decimal.Parse(dg_mat.Rows[x].Cells[2].Value.ToString());
                dg_mat.Rows[x].Cells[2].Value = string.Format("{0:n3}", dd);
                decimal ddd = decimal.Parse(dg_mat.Rows[x].Cells[3].Value.ToString());
                dg_mat.Rows[x].Cells[3].Value = string.Format("{0:n3}", ddd);

            }
        }
    }
}
    