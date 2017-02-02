using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BeetleBase
{
    public partial class Form4 : Form
    {
        public void form4editenable(bool a)
        {
            this.comboBox1.Enabled = a;
            this.comboBox2.Enabled = a;
            this.comboBox3.Enabled = a;
            this.comboBox4.Enabled = a;
            this.comboBox5.Enabled = a;
            this.comboBox6.Enabled = a;
            this.textBox4.Enabled = a;
            this.textBox5.Enabled = a;
            this.textBox8.Enabled = a;
            this.textBox9.Enabled = a;
            this.comboBox7.Enabled = a;
            this.comboBox8.Enabled = a;
            this.comboBox9.Enabled = a;
            this.textBox13.Enabled = a;
            this.comboBox10.Enabled = a;
            this.comboBox11.Enabled = a;
            this.comboBox12.Enabled = a;
            this.button1.Enabled = !a;
            this.button2.Enabled = a;
            this.button3.Enabled = a;
            this.dataGridView1.Enabled = !a;
            this.textBox1.Enabled = !a;
            this.button7.Enabled = !a;
        }

        public Form4(mutual mutual, DB thefile, Form2 aa)
        {
            InitializeComponent();
            form4editenable(false);
            this.mutual = mutual;
            this.thefile = thefile;
            this.aa = aa;
            this.button5.Enabled = false;
            this.dataGridView1.DataSource = this.thefile.main2.Tables[0];
            DataSet dropdowns = new DataSet();
            OleDbCommand first = new OleDbCommand("SELECT * FROM [COLLECTIONS- Drop Down Capture Type]", this.thefile.dbo);
            OleDbDataAdapter dropdown = new OleDbDataAdapter(first);
            dropdown.Fill(dropdowns, "capturetype");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Collector/Museum]";
            dropdown.Fill(dropdowns, "collectormuseum");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Country]";
            dropdown.Fill(dropdowns, "country");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Experiment]";
            dropdown.Fill(dropdowns, "experiment");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Fungus]";
            dropdown.Fill(dropdowns, "fungus");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Province]";
            dropdown.Fill(dropdowns, "province");
            int capturetypecount = dropdowns.Tables["capturetype"].Rows.Count;
            int collectormuseumcount = dropdowns.Tables["collectormuseum"].Rows.Count;
            int countrycount = dropdowns.Tables["country"].Rows.Count;
            int experimentcount = dropdowns.Tables["experiment"].Rows.Count;
            int funguscount = dropdowns.Tables["fungus"].Rows.Count;
            int provincecount = dropdowns.Tables["province"].Rows.Count;
            for (int i = 0; i < capturetypecount; i++)
            {
                comboBox2.Items.Add(dropdowns.Tables["capturetype"].Rows[i][0]);
            }
            for (int i = 0; i < countrycount; i++)
            {
                comboBox5.Items.Add(dropdowns.Tables["country"].Rows[i][0]);
            }
            for (int i = 0; i < collectormuseumcount; i++)
            {
                comboBox6.Items.Add(dropdowns.Tables["collectormuseum"].Rows[i][0]);
            }
            for (int i = 0; i < experimentcount; i++)
            {
                comboBox1.Items.Add(dropdowns.Tables["experiment"].Rows[i][0]);
            }
            for (int i = 0; i < funguscount; i++)
            {
                comboBox3.Items.Add(dropdowns.Tables["fungus"].Rows[i][0]);
            }
            for (int i = 0; i < provincecount; i++)
            {
                comboBox4.Items.Add(dropdowns.Tables["province"].Rows[i][0]);
            }
            dropdown.Dispose();
        }

        public mutual mutual;
        public DB thefile;
        public Form2 aa;
        public void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            string cmd = "SELECT * FROM [COLLECTIONS] WHERE vial = " + textBox1.Text;
            OleDbCommand vialsearch = new OleDbCommand(cmd, this.thefile.dbo);
            OleDbDataAdapter vialadapter = new OleDbDataAdapter(vialsearch);
            DataSet vials = new DataSet();
            vialadapter.Fill(vials, "INIT");
            vials.Tables["INIT"].Columns["date"].DataType = System.Type.GetType("System.DateTime");
            this.dataGridView1.DataSource = vials.Tables["INIT"];
            vialadapter.Dispose();
        }

        public void textBox1_KeyUp2(object sender, KeyEventArgs e)
        {
            string cmd = "SELECT * FROM [COLLECTIONS] WHERE vial = " + textBox1.Text + "";
            OleDbCommand vialsearch = new OleDbCommand(cmd, this.thefile.dbo);
            OleDbDataAdapter vialadapter = new OleDbDataAdapter(vialsearch);
            DataSet vials = new DataSet();
            vialadapter.Fill(vials);
            this.dataGridView1.DataSource = vials.Tables[0];
            vialadapter.Dispose();
        }

        public void button1_Click(object sender, EventArgs e)
        {
            this.editting = true;
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }
            form4editenable(true);
            col = editted[0].Cells;

            char[] delimiterChars = { '/' };
            string[] words1 = col[10].Value.ToString().Split(delimiterChars);
            if (words1.Length > 1)
            {
                if (Int32.Parse(words1[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words1[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            string[] words2 = col[13].Value.ToString().Split(delimiterChars);
            if (words2.Length > 1)
            {
                if (Int32.Parse(words2[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words2[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            groupBox1.Text = "Edit Vial " + col[0].Value.ToString();
            this.currentvial = col[0].Value.ToString();
            comboBox1.Text = col[1].Value.ToString();
            textBox4.Text = col[2].Value.ToString();
            textBox5.Text = col[3].Value.ToString();
            comboBox2.Text = col[4].Value.ToString();
            comboBox3.Text = col[5].Value.ToString();
            textBox8.Text = col[6].Value.ToString();
            textBox9.Text = col[7].Value.ToString();
            comboBox4.Text = col[8].Value.ToString();
            comboBox5.Text = col[9].Value.ToString();
            //            textBox12.Text = col[10].Value.ToString();
            if (words1.Length > 1)
            {
                comboBox7.Text = words1[2];
                comboBox8.Text = words1[0];
                comboBox9.Text = words1[1];
            }
            textBox13.Text = col[11].Value.ToString();
            //            textBox15.Text = col[12].Value.ToString();
            if (words2.Length > 1)
            {
                comboBox10.Text = words2[2];
                comboBox11.Text = words2[0];
                comboBox12.Text = words2[1];
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public string currentvial;
        public bool editting = false;

        private void button2_Click(object sender, EventArgs e)
        {
            this.editting = false;
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string updatemaster = "UPDATE [COLLECTIONS] SET ";
            updatemaster += " [experiment] = " + comboBox1.Text + "'";
            if (textBox4.Text == "")
            {
                //                updatemaster += ", [Count] = null";
            }
            else
            {
                updatemaster += ", [field_vial] = " + textBox4.Text;
            }
            updatemaster += ", [host_or_trap] = " + textBox5.Text + "'";
            updatemaster += ", [capture->storage] = '" + comboBox2.Text + "'";
            updatemaster += ", [fungus] = '" + comboBox3.Text + "'";
            updatemaster += ", [locality] = '" + textBox8.Text + "'";
            updatemaster += ", [county] = '" + textBox9.Text + "'";
            updatemaster += ", [province] = '" + comboBox4.Text + "'";
            updatemaster += ", [country] = '" + comboBox5.Text + "'";
            if
                (
                comboBox7.Text != ""
                && comboBox7.Text != " "
                && comboBox7.Text != "  "
                && comboBox8.Text != ""
                && comboBox8.Text != " "
                && comboBox8.Text != "  "
                && comboBox9.Text != ""
                && comboBox9.Text != " "
                && comboBox9.Text != "  "
                )
            {
                updatemaster += ", [date] = '" + comboBox8.Text + "/" + comboBox9.Text + "/" + comboBox7.Text + "'";
            }
            updatemaster += ", [VIAL_note] = '" + textBox13.Text + "'";
            updatemaster += ", [collector/museum] = '" + comboBox6 + "'";
            if 
                (
                comboBox11.Text != "" 
                && comboBox11.Text != " " 
                && comboBox11.Text != "  "
                && comboBox12.Text != ""
                && comboBox12.Text != " "
                && comboBox12.Text != "  "
                && comboBox10.Text != ""
                && comboBox10.Text != " "
                && comboBox10.Text != "  "
                )
            {
                updatemaster += ", [date collected] = '" + comboBox11.Text + "/" + comboBox12.Text + "/" + comboBox10.Text + "'";
            }
            updatemaster += " WHERE vial = " + this.currentvial;
            try
            {
                OleDbCommand up = new OleDbCommand(updatemaster, this.thefile.dbo);
                OleDbDataAdapter upd = new OleDbDataAdapter();
                upd.UpdateCommand = up;
                upd.UpdateCommand.ExecuteNonQuery();
            }
            catch (OleDbException)
            {
                MessageBox.Show("Unable to write. Check to make sure information is valid!");
            }
            textBox1_KeyUp(null, null);
            button3_Click(null, null);
            form4editenable(false);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.editting = false;
            this.comboBox1.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";
            this.comboBox2.Text = "";
            this.comboBox3.Text = "";
            this.comboBox4.Text = "";
            this.comboBox5.Text = "";
            this.comboBox6.Text = "";
            this.textBox8.Text = "";
            this.textBox9.Text = "";
            this.comboBox7.Text = "";
            this.comboBox8.Text = "";
            this.comboBox9.Text = "";
            this.textBox13.Text = "";
            this.comboBox10.Text = "";
            this.comboBox11.Text = "";
            this.comboBox12.Text = "";
            this.button5.Enabled = false;
            this.button4.Enabled = true;
            form4editenable(false);
            groupBox1.Text = "Vial Info";
        }

        public void button4_Click(object sender, EventArgs e)
        {
            this.comboBox1.Text = "";
            this.textBox4.Text = "";
            this.textBox5.Text = "";
            this.comboBox2.Text = "";
            this.comboBox3.Text = "no";
            this.comboBox4.Text = "";
            this.comboBox5.Text = "";
            this.comboBox6.Text = "";
            this.textBox8.Text = "";
            this.textBox9.Text = "";
            this.comboBox7.Text = "";
            this.comboBox8.Text = "";
            this.comboBox9.Text = "";
            this.textBox13.Text = "";
            this.comboBox10.Text = "";
            this.comboBox11.Text = "";
            this.comboBox12.Text = "";
            form4editenable(true);
            button2.Enabled = false;
            button3.Enabled = true;
            dataGridView1.ClearSelection();
            button4.Enabled = false;
            button5.Enabled = true;
            string year = DateTime.Today.Year.ToString();
            string month = DateTime.Today.Month.ToString("D2");
            string day = DateTime.Today.Day.ToString("D2");
            comboBox7.Text = year;
            comboBox8.Text = month;
            comboBox9.Text = day;
        }

        public void button5_Click(object sender, EventArgs e)
        {
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string insertmaster = "INSERT INTO [COLLECTIONS] ([experiment],";
            if (textBox4.Text != "" && textBox4.Text != " " && textBox4.Text != "  ")
            {
                insertmaster += " [field_vial],";
            }
            insertmaster += "[host_or_trap], [capture->storage], [fungus], [locality], [county], [province], [country], [date], [VIAL_note], [collector/museum], [date collected]) VALUES (";
            insertmaster += "'" + comboBox1.Text + "', ";
            if (textBox4.Text != "" && textBox4.Text != " " && textBox4.Text != "  ")
            {
                insertmaster += textBox4.Text + ", ";
            }
            insertmaster += "'" + textBox5.Text + "', ";
            insertmaster += "'" + comboBox2.Text + "', ";
            insertmaster += "'" + comboBox3.Text + "', ";
            insertmaster += "'" + textBox8.Text + "', ";
            insertmaster += "'" + textBox9.Text + "', ";
            insertmaster += "'" + comboBox4.Text + "', ";
            insertmaster += "'" + comboBox5.Text + "', ";
            //            insertmaster += "'" + textBox12.Text + "', ";
            insertmaster += "'" + comboBox8.Text + "/" + comboBox9.Text + "/" + comboBox7.Text + "', ";
            insertmaster += "'" + textBox13.Text + "', ";
            insertmaster += "'" + comboBox6.Text + "', ";
            //            insertmaster += "'" + textBox15.Text + "'); ";
            insertmaster += "'" + comboBox11.Text + "/" + comboBox12.Text + "/" + comboBox10.Text + "'";
            insertmaster += ")";
            try
            {
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                OleDbCommand thebiginsert = new OleDbCommand(insertmaster, this.thefile.dbo);
                OleDbDataAdapter inserter = new OleDbDataAdapter();
                inserter.InsertCommand = thebiginsert;
                inserter.InsertCommand.ExecuteNonQuery();
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                string newest = "SELECT TOP 1 vial FROM [COLLECTIONS] ORDER BY vial DESC";
                OleDbCommand selectnew = new OleDbCommand(newest, this.thefile.dbo);
                OleDbDataAdapter sn = new OleDbDataAdapter(selectnew);
                DataSet latest = new DataSet();
                sn.Fill(latest);
                textBox1.Text = latest.Tables[0].Rows[0][0].ToString();
                button3_Click(null, null);
                form4editenable(false);
                button4.Enabled = true;
                button5.Enabled = false;
                textBox1_KeyUp(null, null);
            }
            catch (OleDbException err)
            {
                MessageBox.Show(err.ToString());
//                MessageBox.Show("Error inserting Data! Did you check to make sure everything is filled out, valid, and in the right format? E.g. Is the date mm/dd/yyyy? Is the fungus a member of it's respective dropdown menu?");
//                MessageBox.Show(insertmaster);
            }

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form editDropDowns = new Form5(this.thefile, this, aa);
            editDropDowns.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete vial?", "Delete Vial",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            {
                return;
            }
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }
            col = editted[0].Cells;
            if (col.Count > 0)
            {
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                var todelete = col[0].Value.ToString();
                string deletemaster = "DELETE FROM [COLLECTIONS] WHERE VIAL = " + todelete;
                string deletemaster2 = "DELETE FROM [SPECIES_IN_COLLECTIONS] WHERE VIAL = " + todelete;
                OleDbCommand thebigdelete2 = new OleDbCommand(deletemaster2, this.thefile.dbo);
                OleDbDataAdapter deleter2 = new OleDbDataAdapter();
                deleter2.DeleteCommand = thebigdelete2;
                deleter2.DeleteCommand.ExecuteNonQuery();
                OleDbCommand thebigdelete = new OleDbCommand(deletemaster, this.thefile.dbo);
                OleDbDataAdapter deleter = new OleDbDataAdapter();
                deleter.DeleteCommand = thebigdelete;
                deleter.DeleteCommand.ExecuteNonQuery();
                textBox1_KeyUp(null, null);
            }
        }
    }
}
