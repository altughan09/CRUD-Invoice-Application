using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace CrudApp
{
public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        dataGridView1.RowTemplate.Height = 50;
        string path = @"c:\Users\Altug\Desktop\import tracking file.xlsx";
        string connection = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + path + "; Extended Properties = Excel 12.0;";
        string Command = "Select * from [Sheet1$]";
        OleDbConnection con = new OleDbConnection(connection);
        con.Open();
        OleDbCommand cmd = new OleDbCommand(Command, con);
        OleDbDataAdapter db = new OleDbDataAdapter(cmd);
        DataTable dt = new DataTable();
        db.Fill(dt);
        dataGridView1.DataSource = dt;
        con.Close();

    }

    private void button1_Click(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox3.Text))
        {
            MessageBox.Show("Please fill out all the text areas", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else
        {
            string path = @"c:\Users\Altug\Desktop\import tracking file.xlsx";
            string connection = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + path + "; Extended Properties = Excel 12.0;";

            string Command = "INSERT INTO [Sheet1$]([Invoice_Date], [Invoice_No], [Amount], [Currency]) VALUES (@value1,@value2,@value3,@value4)";
            OleDbConnection con = new OleDbConnection(connection);
            try
            {
                con.Open();
                OleDbCommand cmd = new OleDbCommand(Command, con);
                cmd.Parameters.AddWithValue("@value1", textBox1.Text);
                cmd.Parameters.AddWithValue("@value2", textBox2.Text);
                cmd.Parameters.AddWithValue("@value3", textBox3.Text);
                cmd.Parameters.AddWithValue("@value4", textBox4.Text);
                cmd.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.ToString(), "Error!", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            // Get the entered value using  OleDbAdapter  
            System.Data.OleDb.OleDbDataAdapter cmd2;
            cmd2 = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", con);
            // copy the Excel value to the DataSet   
            DataSet ds = new System.Data.DataSet();
            cmd2.Fill(ds);
            // Finally display the entered Item using dataGriedView   
            dataGridView1.DataSource = ds.Tables[0];
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            con.Close();
        }
    }
}
}
