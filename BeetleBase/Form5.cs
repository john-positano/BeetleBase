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
using System.Runtime.InteropServices;

namespace BeetleBase
{
    public partial class Form5 : Form
    {
        public Form5(DB thefile, Form4 a, Form2 aa)
        {
            InitializeComponent();
            this.treeView1.ExpandAll();
            this.thefile = thefile;
            this.ds = new DataSet();
            this.a = a;
            this.aa = aa;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int AllocConsole();

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                //            MessageBox.Show(e.Node.Name);
                this.ds.Clear();
                this.ds = new DataSet();
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                dataGridView1.AllowUserToDeleteRows = true;
                this.currentname = e.Node.Name;
                if (e.Node.Name == "CaptureType")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Capture Type]";
                    this.currenttable = "[COLLECTIONS- Drop Down Capture Type]";
                    this.cn = "capture type";

                }
                else if (e.Node.Name == "Country")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Country]";
                    this.currenttable = "[COLLECTIONS- Drop Down Country]";
                    this.cn = "country";
                }
                else if (e.Node.Name == "Experiment")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Experiment]";
                    this.currenttable = "[COLLECTIONS- Drop Down Experiment]";
                    this.cn = "experiment";
                }
                else if (e.Node.Name == "Fungus")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Fungus]";
                    this.currenttable = "[COLLECTIONS- Drop Down Fungus]";
                    this.cn = "fungus";
                }
                else if (e.Node.Name == "Province")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Province]";
                    this.currenttable = "[COLLECTIONS- Drop Down Province]";
                    this.cn = "province";
                }
                else if (e.Node.Name == "Identifiers")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Identifiers]";
                    this.currenttable = "[COLLECTIONS- Drop Down Identifiers]";
                    this.cn = "identifier";
                }
                else if (e.Node.Name == "SpeciesTable")
                {
                    this.cmd = "SELECT [key], [SpCode], [Tribe], [Genus], [species] FROM [Species_Table_NEW]";
                    this.currenttable = "[Species_Table_NEW]";
                    this.cn = "speciestable";
                }
                else
                {
                    return;
                }
                OleDbCommand vialsearch = new OleDbCommand(this.cmd, this.thefile.dbo);
                this.vialadapter = new OleDbDataAdapter(vialsearch);
                DataSet vials = new DataSet();
                vialadapter.Fill(ds, "TABLE");
                this.dataGridView1.DataSource = ds.Tables[0];
            }
            catch (OleDbException err)
            {
                MessageBox.Show(err.ToString());
            }
        }

        public DB thefile;
        public Form2 aa;

        private void button1_Click(object sender, EventArgs e)
        {
            //            AllocConsole();
            DataSet s = ds.GetChanges();
            OleDbDataAdapter inserter = new OleDbDataAdapter();
            if (s != null)
            {
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                int srowcount = s.Tables["TABLE"].Rows.Count;
                for (int i = 0; i < srowcount; i++)
                {
//                    MessageBox.Show(s.Tables["TABLE"].Rows[i][0].ToString().Trim());
                    if (this.cn == "speciestable")
                    {
                        if 
                        (
                            s.Tables["TABLE"].Rows[i][1].ToString().Trim() == "" &&
                            s.Tables["TABLE"].Rows[i][4].ToString().Trim() == ""
                        )
                        {
                            return;
                        }
                        string cmd = "INSERT INTO "
                        + this.currenttable
                        + " ([key], [SpCode], [Tribe], [Genus], [Species])"
                        + " VALUES ("
                        + s.Tables["TABLE"].Rows[i][0].ToString() + ", "
                        + s.Tables["TABLE"].Rows[i][1].ToString() + ", '"
                        + s.Tables["TABLE"].Rows[i][2].ToString() + "', '"
                        + s.Tables["TABLE"].Rows[i][3].ToString() + "', '"
                        + s.Tables["TABLE"].Rows[i][4].ToString()
                        + "');";
                        OleDbCommand a = new OleDbCommand(cmd, this.thefile.dbo);
                        inserter.InsertCommand = a;
                        inserter.InsertCommand.ExecuteNonQuery();
                    }
                    else if (s.Tables["TABLE"].Rows[i][0].ToString().Trim() != "")
                    {
                        //                    Console.WriteLine(s.Tables["TABLE"].Rows[i][0].ToString());
                        string cmd = "INSERT INTO "
                        + this.currenttable
                        + " (["
                        + this.cn
                        + "]) VALUES ('"
                        + s.Tables["TABLE"].Rows[i][0].ToString()
                        + "');";
                        OleDbCommand a = new OleDbCommand(cmd, this.thefile.dbo);
                        inserter.InsertCommand = a;
                        inserter.InsertCommand.ExecuteNonQuery();
                    }
                }
            }
        }

        public string cmd;
        public DataSet ds;
        public OleDbDataAdapter vialadapter;
        public string currenttable;
        public string currentname;
        public string cn;
        public Form4 a;

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
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
            try
            {
                dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Identifiers]";
                dropdown.Fill(dropdowns, "identifiers");
            } catch (OleDbException err)
            { MessageBox.Show(err.ToString()); }
            a.comboBox1.Items.Clear();
            a.comboBox2.Items.Clear();
            a.comboBox3.Items.Clear();
            a.comboBox4.Items.Clear();
            a.comboBox5.Items.Clear();
            a.comboBox6.Items.Clear();
            aa.identifiercombo.Items.Clear();
            int capturetypecount = dropdowns.Tables["capturetype"].Rows.Count;
            int collectormuseumcount = dropdowns.Tables["collectormuseum"].Rows.Count;
            int countrycount = dropdowns.Tables["country"].Rows.Count;
            int experimentcount = dropdowns.Tables["experiment"].Rows.Count;
            int funguscount = dropdowns.Tables["fungus"].Rows.Count;
            int provincecount = dropdowns.Tables["province"].Rows.Count;
            int identifierscount = dropdowns.Tables["identifiers"].Rows.Count;
            for (int i = 0; i < capturetypecount; i++)
            {
                a.comboBox2.Items.Add(dropdowns.Tables["capturetype"].Rows[i][0]);
            }
            for (int i = 0; i < countrycount; i++)
            {
                a.comboBox5.Items.Add(dropdowns.Tables["country"].Rows[i][0]);
            }
            for (int i = 0; i < collectormuseumcount; i++)
            {
                a.comboBox6.Items.Add(dropdowns.Tables["collectormuseum"].Rows[i][0]);
            }
            for (int i = 0; i < experimentcount; i++)
            {
                a.comboBox1.Items.Add(dropdowns.Tables["experiment"].Rows[i][0]);
            }
            for (int i = 0; i < funguscount; i++)
            {
                a.comboBox3.Items.Add(dropdowns.Tables["fungus"].Rows[i][0]);
            }
            for (int i = 0; i < provincecount; i++)
            {
                a.comboBox4.Items.Add(dropdowns.Tables["province"].Rows[i][0]);
            }
            for (int i = 0; i < identifierscount; i++)
            {
                aa.identifiercombo.Items.Add(dropdowns.Tables["identifiers"].Rows[i][0]);
            }
            dropdown.Dispose();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > 0 && e.ColumnIndex > 0)
            {
                dataGridView1.Rows[e.RowIndex].ReadOnly = true;
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim() == "")
                {
                    dataGridView1.Rows[e.RowIndex].ReadOnly = false;
                }
            }
        }
    }
}
