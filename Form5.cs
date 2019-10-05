using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;


namespace LEGAL
{
    public partial class Form5 : Form
    {

        string ConnectionStringForAddCzPodnik = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;

        public Form5()
        {
            InitializeComponent();
            FillComboBoxCountry();
        }

        private void FillComboBoxCountry()
        {
            string[] ArrForComboBoxType = new string[3];
            ArrForComboBoxType[0] = "Czech Republic";
            ArrForComboBoxType[1] = "Slovakia";
            ArrForComboBoxType[2] = "Belgie";

            for (int i = 0; i < ArrForComboBoxType.Length; i++)
            {
                comboBoxCountry.Items.Add(ArrForComboBoxType[i]);
            }
        }

        private int GetCountryId(string country)
        {
            int id;
            if(country == "Czech Republic")
            {
                id = 1;
            }else if(country == "Slovakia")
            {
                id = 2;
            }else
            {
                id = 3;
            }
            return id;
        }


        private void GetArrayFromComboBoxes() 
        {

            SqlConnection ConnectionForm4 = new SqlConnection(ConnectionStringForAddCzPodnik);

            string CompanyName, CompanyAddres, CompanyRepresentant, CompanyIc;
            int CompanyID = 666; //сюда запишу то что отдаст запрос на макс ID

            CompanyName = textBoxForCZCompanzName.Text;
            CompanyAddres = textBoxAddresCzCompany.Text;
            CompanyRepresentant = textBoxRepresentant.Text;
            CompanyIc = textBoxIC.Text;



            string SqlForFindLastId = "SELECT MAX(CompanyId) From CompanyCustomers";

            ConnectionForm4.Open();
            SqlCommand CommandForCreateId = new SqlCommand(SqlForFindLastId, ConnectionForm4);
            SqlDataReader reader = CommandForCreateId.ExecuteReader();

            while (reader.Read())
            {
                CompanyID = Convert.ToInt32(reader.GetValue(0)) + 1;
            }
            reader.Close();
            
            string SqlExpressionAddTooDb = "insert into CompanyCustomers (CompanyID, CompanyName, CompanyAddress, CompanyIC, CompanyRepresentant, Country) VALUES(N'" + CompanyID + "', N'" + CompanyName + "', N' " + CompanyAddres + "', N'" + CompanyIc + "', N'" + CompanyRepresentant + "', N'"+ GetCountryId(comboBoxCountry.Text) + "');";


            using (ConnectionForm4)
            {

                SqlCommand CommandForCreateNewCompany = new SqlCommand(SqlExpressionAddTooDb, ConnectionForm4);
                int number = CommandForCreateNewCompany.ExecuteNonQuery();
                MessageBox.Show(CompanyName + " была добавлена.");
                ConnectionForm4.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetArrayFromComboBoxes();
        }

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            Control.DeleteFromSeznam(this.Name);
        }

    }
}
