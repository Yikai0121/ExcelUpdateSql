using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Z.Dapper.Plus;

namespace ExcelUpdateSql
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            //dataGridView1.DataSource = dt;
            if (dt != null)
            {
                List<Customer> customers = new List<Customer>();
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    Customer customer = new Customer();
                    customer.CustomerID = dt.Rows[i]["CustomerID"].ToString();
                    customer.CompanyName = dt.Rows[i]["CompanyName"].ToString();
                    customers.Add(customer);

                }
                customerBindingSource.DataSource = customers;
            }
        }
          DataTableCollection tableCollection;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    using(var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using(IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (DataTable table in tableCollection)
                            {
                                cboSheet.Items.Add(table.TableName);
                            }
                        }
                    }
                }
            }


        }

        private void btnimport_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionString = "Server=PC-MISI5\\SQLEXPRESS;Database=Customer;Trusted_Connection=True;MultipleActiveResultSets=True";
                DapperPlusManager.Entity<Customer>().Table("Customers");
                List<Customer> customers = customerBindingSource.DataSource as List<Customer>;
                if(customers != null)
                {
                    using(IDbConnection db = new SqlConnection(connectionString))
                    {
                        db.BulkInsert(customers);
                    }
                }
                MessageBox.Show("上傳成功");
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
