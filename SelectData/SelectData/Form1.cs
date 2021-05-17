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
namespace SelectData
{
    public partial class Form1 : Form
    {
        OleDbConnection cn;
        string strcon = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\QData.mdb";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // เปิดการเชื่อมต่อ
            cn = new OleDbConnection(strcon);

            //สร้างคำสั่ง SQL เพื่อดึงข้อมูล
            string txt = "select * from QList";
            //สร้าง Datatable เพื่อรองรับชุดข้อมูล
            DataTable dt = new DataTable();
            //สร้างคำสั่ง Query ข้อมูล
            OleDbDataAdapter da = new OleDbDataAdapter(txt, cn);
            // เมื่อดึงข้อมูลเสร็จแล้วให้ Foreach array เก็บไว้ที่ Datatable
            da.Fill(dt);

            //ให้ข้อมูลแสดง บน Datagrid
            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int id;

            cn = new OleDbConnection(strcon);
            cn.Open();

            string txt = "SELECT MAX(QNo) FROM QList";

            OleDbCommand cmd = new OleDbCommand(txt, cn);
            OleDbDataReader dr = cmd.ExecuteReader();

            if (dr.Read())

            {
                string strid = dr[0].ToString();

                if (strid != "")
                {
                    id = Convert.ToInt32(strid) + 1;
                    textBox1.Text = id.ToString();
                }
                else
                {
                    id = 1;
                    textBox1.Text = id.ToString();
                }
            }

        }
    }
}
