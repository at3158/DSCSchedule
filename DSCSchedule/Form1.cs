using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace DSCSchedule
{
    public partial class Form1 : Form
    {
        int count = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "START! ";

            //timer1.Interval = 1000; // 設定每秒觸發一次
            timer1.Interval = 1000*60; // 設定每分觸發一次
            timer1.Enabled = true; // 啟動 Timer

        }

        #region timer1_Tick 判斷
        private void timer1_Tick(object sender, EventArgs e)
        {
            count++;
            textBox1.Text = count.ToString() + Environment.NewLine;
            // 更新-標準售價 
            if (checkBox1.Checked == true)
            {
                if (radioButton1.Checked==true)
                {
                    if (dateTimePicker1.Value.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString()))
                    {
                        ExeSQL1();
                    }
                }
                else if (radioButton2.Checked == true)
                {
                    if ((dateTimePicker1.Value.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString()))&&(DateTime.Now.Day.ToString().Equals(numericUpDown1.Value.ToString())))
                    {
                        ExeSQL1();
                    }
                }
                else
                {
                    count = 0;
                    textBox1.Text = "START! ";
                }
                
            }
            else
            {
                count = 0;
                textBox1.Text = "START! ";
            }

            // 更新-成本+運費
            if (checkBox2.Checked == true)
            {
                if (radioButton3.Checked == true)
                {
                    if (dateTimePicker2.Value.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString()))
                    {
                        ExeSQL2();
                    }
                }
                else if (radioButton4.Checked == true)
                {
                    if ((dateTimePicker2.Value.ToShortTimeString().Equals(DateTime.Now.ToShortTimeString())) && (DateTime.Now.Day.ToString().Equals(numericUpDown2.Value.ToString())))
                    {
                        ExeSQL2();
                    }
                }
                else
                {
                    count = 0;
                    textBox1.Text = "START! ";
                }

            }
            else
            {
                count = 0;
                textBox1.Text = "START! ";
            }

        }

        #endregion

        #region 執行SQL
        /// <summary>
        /// 更新-標準售價 
        /// </summary>
        public void ExeSQL1()
        {

            SqlConnection sqlConn;
            SqlCommand sqlComm = new SqlCommand();
            //Basic UPDATE method with Parameters           
            //sqlComm.CommandText = @"UPDATE tableName SET paramColumn='@paramName' WHERE conditionColumn='@conditionName'";
            //sqlComm.Parameters.Add("@paramName", SqlDbType.VarChar);
            //sqlComm.Parameters.Add("@conditionName", SqlDbType.VarChar);

            StringBuilder sbSql = new StringBuilder();
            sbSql.Append(" UPDATE INVMB SET MB047=0");
            

            foreach (var item in checkedListBox1.CheckedItems)
            {                
                if (item.ToString().Equals("TAX2016"))
                {
                    try
                    {                       

                        sqlConn = new SqlConnection("sqlConnString");
                        sqlComm = sqlConn.CreateCommand();
                        sqlComm.CommandText = @sbSql.ToString();
                        sqlConn.Open();
                        sqlComm.ExecuteNonQuery();
                        sqlConn.Close();

                        textBox1.Text = textBox1.Text + "更新-標準售價  Success!" + Environment.NewLine;
                    }
                    catch
                    {
                        textBox1.Text = textBox1.Text + "更新-標準售價  fail" + Environment.NewLine;
                    }
                   

                }
                if (item.ToString().Equals("東京著衣2015"))
                {
                    try
                    {

                        sqlConn = new SqlConnection("sqlConnString");
                        sqlComm = sqlConn.CreateCommand();
                        sqlComm.CommandText = @sbSql.ToString();
                        sqlConn.Open();
                        sqlComm.ExecuteNonQuery();
                        sqlConn.Close();

                        textBox1.Text = textBox1.Text + "更新-標準售價  Success!" + Environment.NewLine;
                    }
                    catch
                    {
                        textBox1.Text = textBox1.Text + "更新-標準售價  fail" + Environment.NewLine;
                    }
                    
                }

            }
           
        }
        /// <summary>
        /// 更新- 成本+運費.
        /// </summary>
        public void ExeSQL2()
        {
            SqlConnection sqlConn;new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            StringBuilder sbSql = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            sbSql.Append(" ");
            string NowYearMonth = null;

            sbSql.Append(" UPDATE INVMB SET MB053=ISNULL(ROUND(((LB004 / NULLIF(LB003,0))*1.056),2),0)");


            foreach (var item in checkedListBox2.CheckedItems)
            {
                if (item.ToString().Equals("TAX2016"))
                {
                    try
                    {
                        sqlConn = new SqlConnection("sqlConnString");
                        adapter = new SqlDataAdapter(@"SELECT TOP 1 MA011 FROM CMSMA WITH (NOLOCK)", sqlConn);
                        sqlCmdBuilder = new SqlCommandBuilder(adapter);
                        DataSet ds = new DataSet();
                        sqlConn.Open();
                        adapter.Fill(ds);
                        sqlConn.Close();

                        NowYearMonth = ds.Tables[0].Rows[0]["MA011"].ToString();

                        sbSql.Replace("@NowYearMonth", NowYearMonth);
                        sqlConn = new SqlConnection("sqlConnString");
                        sqlComm = sqlConn.CreateCommand();
                        sqlComm.CommandText = @sbSql.ToString();
                        sqlConn.Open();
                        sqlComm.ExecuteNonQuery();
                        sqlConn.Close();

                        textBox1.Text = textBox1.Text + "更新-成本+運費  Success! " + NowYearMonth.ToString() + Environment.NewLine;
                    }
                    catch
                    {
                        textBox1.Text = textBox1.Text + "更新-成本+運費  fail" + Environment.NewLine;
                    }


                }
                if (item.ToString().Equals("東京著衣2015"))
                {
                    try
                    {

                        sqlConn = new SqlConnection("sqlConnString");
                        adapter = new SqlDataAdapter(@"SELECT TOP 1 MA011 FROM CMSMA WITH (NOLOCK)", sqlConn);
                        sqlCmdBuilder = new SqlCommandBuilder(adapter);
                        DataSet ds = new DataSet();
                        sqlConn.Open();
                        adapter.Fill(ds);
                        sqlConn.Close();

                        NowYearMonth = ds.Tables[0].Rows[0]["MA011"].ToString();

                        sbSql.Replace("@NowYearMonth", NowYearMonth);
                        sqlConn = new SqlConnection("sqlConnString");
                        sqlComm = sqlConn.CreateCommand();
                        sqlComm.CommandText = @sbSql.ToString();
                        sqlConn.Open();
                        sqlComm.ExecuteNonQuery();
                        sqlConn.Close();

                        textBox1.Text = textBox1.Text + "更新-成本+運費  Success! " + NowYearMonth.ToString() + Environment.NewLine;
                    }
                    catch
                    {
                        textBox1.Text = textBox1.Text + "更新-成本+運費  fail" + Environment.NewLine;
                    }

                }

            }
        }

        #endregion

        #region 手動執行
        /// <summary>
        /// 手動-更新-標準售價 
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            ExeSQL1();
        }

        /// <summary>
        /// 手動-更新-成本+運費
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            ExeSQL2();
        }

        #endregion


    }
}
