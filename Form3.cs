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
    public partial class Form3 : Form
    {
        string connectionStringForm3 = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;
        string sqlExpressionForFindingAllPlCompany = "SELECT CompanyName FROM CompaniesPL";
        string sqlExpressionForFindingALLCZCompany = "SELECT CompanyName FROM CompanyCustomers ORDER BY CompanyName";
        Fernando WhiteBelt = new Fernando();

        public Form3()
        {
            InitializeComponent();
            FillcomboBoxForPlCompanyForm3();
            FillComboBoxForCzCompanyForm3();
            FillComboBoxForType();
            FillComboBoxCurrency();
            // FillcomboBoxForPayMethodCompanyForm3();
           


        }

        private void FillComboBoxForType()
        {
            string[] ArrForComboBoxType = new string[3];
            ArrForComboBoxType[0] = "usł.";
            ArrForComboBoxType[1] = "m2";
            ArrForComboBoxType[2] = "m3";

            for(int i = 0; i < ArrForComboBoxType.Length;i++)
            {
                comboBoxForType.Items.Add(ArrForComboBoxType[i]);
            }
            
        }

        private void FillComboBoxCurrency()
        {
            string[] ArrForComboBoxType = new string[3];
            ArrForComboBoxType[0] = "PLN";
            ArrForComboBoxType[1] = "EUR";
            ArrForComboBoxType[2] = "CZK";

            for (int i = 0; i < ArrForComboBoxType.Length; i++)
            {
                comboBoxCurrency.Items.Add(ArrForComboBoxType[i]);
            }
        }

        private void FillComboBoxForCzCompanyForm3()
        {
            SqlConnection connectionForFillComboBoxCzComp = new SqlConnection(connectionStringForm3);
            connectionForFillComboBoxCzComp.Open();
            SqlCommand command = new SqlCommand(sqlExpressionForFindingALLCZCompany, connectionForFillComboBoxCzComp);
            SqlDataReader reader = command.ExecuteReader();
            if(reader.HasRows)
            {
                while(reader.Read())
                {
                    comboBoxForCZCompanies.Items.Add(reader.GetValue(0));
                }
            }
            reader.Close();
            connectionForFillComboBoxCzComp.Close();
        }

        private void FillcomboBoxForPlCompanyForm3()
        {
            SqlConnection connection = new SqlConnection(connectionStringForm3);
            connection.Open();
            SqlCommand command = new SqlCommand(sqlExpressionForFindingAllPlCompany, connection);
            SqlDataReader reader = command.ExecuteReader();

            if(reader.HasRows)
            {
                while(reader.Read())
                {
                    comboBoxForPlCompanyForm3.Items.Add(reader.GetValue(0));
                }
            }
            reader.Close();
            connection.Close();

        }

        private string NumberOfFaktura()
        {
            string text = File.ReadAllText("C:\\Legalizace\\Templates\\fakturaNumb.txt", Encoding.Default);
            int newNumber = Convert.ToInt32(text) + 1;
            string newText = Convert.ToString(newNumber);
            File.WriteAllText("C:\\Legalizace\\Templates\\fakturaNumb.txt", newText);
            string date = DateTime.Now.ToString("Myyyy");
            string id = "0" + date + newText;

            return id;
        }

        private string[] AllDataFromFormToArray() //общий массив с данными на отдачу
        {
            string[] arrPolCompData = new string[3]; //заполнили польскую сторону
            string[] arrCzCompData = new string[3]; //заполнили cz сторону
            string[] Arr;
            arrPolCompData = TakeDataFormDbTable();
            arrCzCompData = TakeDataFromDBTableFromTableCZ();

            if (checkBoxPrzelew.Checked) { 

                Arr = new string[22];
                decimal priceForOne = Convert.ToDecimal(textBoxPrice.Text);
                decimal priceForAll = priceForOne * Convert.ToDecimal(textBoxForQant.Text);

            Arr[0] = textBoxForNameGoods.Text;
            Arr[1] = comboBoxForType.Text;
            Arr[2] = textBoxForQant.Text;
            Arr[3] = textBoxPrice.Text;
            Arr[4] = comboBoxCurrency.Text;
            Arr[5] = dateTimePicker1.Text; //monthCalendarForNewCompany.SelectionStart.Day.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Month.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Year.ToString();
            Arr[6] = dateTimePicker2.Text;
            Arr[7] = arrPolCompData[0]; // название польской компании
            Arr[8] = arrPolCompData[1]; // адресс польской компании
            Arr[9] = arrPolCompData[2]; // ич польской компании
            Arr[10] = arrCzCompData[0];// название cz компании
            Arr[11] = arrCzCompData[1];// адресс cz компании
            Arr[12] = arrCzCompData[2];// ич cz компании
            Arr[13] = NumberOfFaktura(); // --- номер фактуры
            Arr[14] = textBoxPayTime.Text; // до когда нужно заплатить
            Arr[15] = textBoxKonto.Text;
            Arr[16] = textBoxIBAN.Text;
            Arr[17] = textBoxKontoNumber.Text;
            Arr[18] = textBoxSWIFT.Text;
            Arr[19] = "przelew";
            Arr[20] = Convert.ToString(priceForAll);
            Arr[21] = WhiteBelt.NumbersToWords(Convert.ToInt32(priceForAll));



            }
            else
            {
                decimal priceForOne = Convert.ToDecimal(textBoxPrice.Text);
                decimal priceForAll = priceForOne * Convert.ToDecimal(textBoxForQant.Text);

                Arr = new string[17];
                Arr[0] = textBoxForNameGoods.Text;
                Arr[1] = comboBoxForType.Text;
                Arr[2] = textBoxForQant.Text;
                Arr[3] = textBoxPrice.Text;
                Arr[4] = comboBoxCurrency.Text;
                Arr[5] = dateTimePicker1.Text; //monthCalendarForNewCompany.SelectionStart.Day.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Month.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Year.ToString();
                Arr[6] = dateTimePicker2.Text;
                Arr[7] = arrPolCompData[0]; // название польской компании
                Arr[8] = arrPolCompData[1]; // адресс польской компании
                Arr[9] = arrPolCompData[2]; // ич польской компании
                Arr[10] = arrCzCompData[0];// название cz компании
                Arr[11] = arrCzCompData[1];// адресс cz компании
                Arr[12] = arrCzCompData[2];// ич cz компании
                Arr[13] = NumberOfFaktura(); // --- номер фактуры
                Arr[14] = "gotówka";
                Arr[15] = Convert.ToString(priceForAll);
                Arr[16] = WhiteBelt.NumbersToWords(Convert.ToInt32(priceForAll));
            }

            return Arr; 

    }

        private string[] TakeDataFormDbTable()
        {
            string[] Arra = new string[3];
            
            string CompanyNameFromForm3 = comboBoxForPlCompanyForm3.Text;
            string TextForQuereName = "SELECT CompanyName FROM CompaniesPL WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для Name
            string TextForQuereAdress = "SELECT CompanyAdress FROM CompaniesPL WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для Adress
           //string TextForQuereRegion = "SELECT CompanyRegion FROM CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для Region
           // string TextForQuereKRS = "SELECT CompanyName KRS CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для KRS
            string TextForQuereNIP = "SELECT CompanyNIP FROM CompaniesPL WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для NIP
           //string TextForQuereREPRESENTANT = "SELECT CompanyRepresentant FROM CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для REPRESENTANT

            SqlConnection ConnectionForTAkeDataFromForm3 = new SqlConnection(connectionStringForm3);
            ConnectionForTAkeDataFromForm3.Open();

            SqlCommand commandName = new SqlCommand(TextForQuereName, ConnectionForTAkeDataFromForm3);
            SqlDataReader reader = commandName.ExecuteReader();

            while(reader.Read())
            {
                Arra[0] = Convert.ToString(reader.GetValue(0));
            }

            reader.Close();

            SqlCommand commandAdress = new SqlCommand(TextForQuereAdress, ConnectionForTAkeDataFromForm3);
            SqlDataReader readerTwo = commandAdress.ExecuteReader();

            while (readerTwo.Read())
            {
                Arra[1] = Convert.ToString(readerTwo.GetValue(0));
            }

            readerTwo.Close();

            SqlCommand commandNIP = new SqlCommand(TextForQuereNIP, ConnectionForTAkeDataFromForm3);
            SqlDataReader readerThree = commandNIP.ExecuteReader();

            while (readerThree.Read())
            {
                Arra[2] = Convert.ToString(readerThree.GetValue(0));
            }

            readerThree.Close();

            return Arra;

        }

        private string[] TakeDataFromDBTableFromTableCZ()
        {
            string[] Arra = new string[3];

            string CompanyNameFromForm3 = comboBoxForCZCompanies.Text;
            string TextForQuereName = "SELECT CompanyName FROM CompanyCustomers WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для Name
            string TextForQuereAdress = "SELECT CompanyAddress FROM CompanyCustomers WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для Adress
            //string TextForQuereRegion = "SELECT CompanyRegion FROM CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для Region
            // string TextForQuereKRS = "SELECT CompanyName KRS CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для KRS
            string TextForQuereIC = "SELECT CompanyIC FROM CompanyCustomers WHERE CompanyName=N'" + CompanyNameFromForm3 + "'"; // запрос для NIP
            //string TextForQuereREPRESENTANT = "SELECT CompanyRepresentant FROM CompaniesPL WHERE CompanyName='" + CompanyNameFromForm3 + "'"; // запрос для REPRESENTANT

            SqlConnection ConnectionForTAkeDataFromForm3 = new SqlConnection(connectionStringForm3);
            ConnectionForTAkeDataFromForm3.Open();

            SqlCommand commandName = new SqlCommand(TextForQuereName, ConnectionForTAkeDataFromForm3);
            SqlDataReader reader = commandName.ExecuteReader();

            while (reader.Read())
            {
                Arra[0] = Convert.ToString(reader.GetValue(0));
            }

            reader.Close();

            SqlCommand commandAdress = new SqlCommand(TextForQuereAdress, ConnectionForTAkeDataFromForm3);
            SqlDataReader readerTwo = commandAdress.ExecuteReader();

            while (readerTwo.Read())
            {
                Arra[1] = Convert.ToString(readerTwo.GetValue(0));
            }

            readerTwo.Close();

            SqlCommand commandIC = new SqlCommand(TextForQuereIC, ConnectionForTAkeDataFromForm3);
            SqlDataReader readerThree = commandIC.ExecuteReader();

            while (readerThree.Read())
            {
                Arra[2] = Convert.ToString(readerThree.GetValue(0));
            }

            readerThree.Close();

            return Arra;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            WhiteBelt.ChangeTextForFaktura(AllDataFromFormToArray());
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            Control.DeleteFromSeznam(this.Name);
        }
        
    }
}
