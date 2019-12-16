using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Configuration;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;

namespace LEGAL
{
    public partial class Form2 : Form

    {
        //"Data Source=ED-PC\\SQL2014EXPRESS;Initial Catalog=WorkDBFirstTry;Integrated Security=True";
        Fernando blackBelt = new Fernando();
        string connectionString = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;
        string sqlExpression = "SELECT CZ FROM Professions";
        //string sqlExpressionForFindingPl = "SELECT ProfessionPL FROM Professions WHERE ProfessionCZ = '"+"'";


        

        public struct PlCompany
        {
            public string name;
            public string adress;
            public long region;
            public int krs;
            public long nip;
            public string representant;
        }

        public struct PovolaniDict
        {
            public string PlPovolani;
            public string CzPovolani;
        }
      


       PlCompany[] plCompanysArr = new PlCompany[6];
       PovolaniDict[] povolaniDictArr = new PovolaniDict[28];
        
       //public 

       public PlCompany FullStruct()
        {
            
            plCompanysArr[0].name = "AGRAT SP.Z O.O";
            plCompanysArr[0].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[0].region = 361501834;
            plCompanysArr[0].krs = 0000557289;
            plCompanysArr[0].nip = 8982211670;
            plCompanysArr[0].representant = "Ganna Ivanchenko";

            plCompanysArr[1].name = "POLFIZ SP.Z O.O";
            plCompanysArr[1].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[1].region = 36162602700000;
            plCompanysArr[1].krs = 0000560386;
            plCompanysArr[1].nip = 8992768049;
            plCompanysArr[1].representant = "Ganna Ivanchenko";

            plCompanysArr[2].name = "PORTOFINO SP.Z O.O";
            plCompanysArr[2].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[2].region = 362256288;
            plCompanysArr[2].krs = 0000571220;
            plCompanysArr[2].nip = 8971812213;
            plCompanysArr[2].representant = "Larysa Kudinova";

            plCompanysArr[3].name = "VENEZIA SP.Z O.O";
            plCompanysArr[3].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[3].region = 362260261;
            plCompanysArr[3].krs = 0000571242;
            plCompanysArr[3].nip = 8971812207;
            plCompanysArr[3].representant = "Larysa Kudinova";

            plCompanysArr[4].name = "PANAMERA SP.Z O.O";
            plCompanysArr[4].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[4].region = 368701915;
            plCompanysArr[4].krs = 0000702537;
            plCompanysArr[4].nip = 8971847803;
            plCompanysArr[4].representant = "Yulia Polizhak";

            plCompanysArr[5].name = "WOX SP.Z O.O";
            plCompanysArr[5].adress = "Wrocławiu 50-148 przy ul. Wita Stwosza 16";
            plCompanysArr[5].region = 382915864;
            plCompanysArr[5].krs = 0000778713;
            plCompanysArr[5].nip = 8971865149;
            plCompanysArr[5].representant = "Yevgeniy Nam";

            PlCompany PlCompanyTuChtoOtdam;
            PlCompanyTuChtoOtdam.name = "Pdd";
            PlCompanyTuChtoOtdam.adress = "sxs";
            PlCompanyTuChtoOtdam.region = 666;
            PlCompanyTuChtoOtdam.krs = 777;
            PlCompanyTuChtoOtdam.nip = 464;
            PlCompanyTuChtoOtdam.representant = "KRATOS";

            if (comboBoxForOldCompany.Text == "AGRAT SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[0].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[0].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[0].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[0].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[0].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[0].representant;
            }

            if (comboBoxForOldCompany.Text == "POLFIZ SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[1].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[1].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[1].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[1].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[1].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[1].representant;
            }

            if (comboBoxForOldCompany.Text == "PORTOFINO SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[2].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[2].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[2].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[2].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[2].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[2].representant;
            }

            if (comboBoxForOldCompany.Text == "VENEZIA SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[3].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[3].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[3].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[3].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[3].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[3].representant;
            }

            if (comboBoxForOldCompany.Text == "PANAMERA SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[4].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[4].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[4].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[4].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[4].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[4].representant;
            }

            if (comboBoxForOldCompany.Text == "WOX SP.Z O.O")
            {
                PlCompanyTuChtoOtdam.name = plCompanysArr[5].name;
                PlCompanyTuChtoOtdam.adress = plCompanysArr[5].adress;
                PlCompanyTuChtoOtdam.region = plCompanysArr[5].region;
                PlCompanyTuChtoOtdam.krs = plCompanysArr[5].krs;
                PlCompanyTuChtoOtdam.nip = plCompanysArr[5].nip;
                PlCompanyTuChtoOtdam.representant = plCompanysArr[5].representant;
            }

            return PlCompanyTuChtoOtdam;

        }
       
       

        
        


        public Form2()
        {
            InitializeComponent();
            FillComboBoxFromList();
            FillSecondComboBox();
        }


        public void FillComboBoxFromList()
        {   
            List<string> plCompanysName = new List<string>();

            plCompanysName.Add("AGRAT SP.Z O.O");
            plCompanysName.Add("VENEZIA SP.Z O.O");
            plCompanysName.Add("PORTOFINO SP.Z O.O");
            plCompanysName.Add("POLFIZ SP.Z O.O");
            plCompanysName.Add("PANAMERA SP.Z O.O");
            plCompanysName.Add("WOX SP.Z O.O");

            for (int i = 0; i < plCompanysName.Count; i++)
            {
                comboBoxForOldCompany.Items.Add(plCompanysName[i]);
            }
        }


        public void FillSecondComboBox()
        {
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = new SqlCommand(sqlExpression, connection);
            SqlDataReader reader = command.ExecuteReader();

            if(reader.HasRows)
            {
                while(reader.Read())
                {
                    comboBoxForNewCompanyTypeOfWork.Items.Add(reader.GetValue(0));
                }
                
            }

            reader.Close();
            connection.Close();
        }


        private void Form2_Load(object sender, EventArgs e)
        {
          

        }

        private void buttonForCreateContract_Click(object sender, EventArgs e)
        {
            object templateForNewCompanyContractPL = "C:\\Legalizace\\Templates\\SmlouvaNewCompanyPL.docx";
            object templateForNewCompanyContractCZ = "C:\\Legalizace\\Templates\\SmlouvaNewCompanyCZ.docx";
            string[] data = new string[6];
            data = readAllTextBoxToArr();
            blackBelt.changeTextForContracts(templateForNewCompanyContractCZ, data);
            blackBelt.changeTextForContracts(templateForNewCompanyContractPL, data);
            //файл для шаблона контракта

        }


        private string[] takeDataFromDb()
        {
            string[] tmp = new string[6];
            string choosenCompanyName;
            choosenCompanyName = comboBoxForOldCompany.Text;


            return tmp;
        }

        private string[] readAllTextBoxToArr()
        {
            string[] data = new string[12];

            
            
            string tutenhamon;
            tutenhamon = comboBoxForNewCompanyTypeOfWork.Text;

            string sqlExpressionFindingCz = "SELECT PLURAL.CZ FROM Plural INNER JOIN Professions ON Plural.ID = Professions.ID_Plural WHERE Professions.CZ = (N'" + tutenhamon + "')";  //ЗАПРОС ДЛЯ НАХОЖДЕНИЯ ПРОФЕССИИ ВО МН.Ч
            string sqlExpressionForFindingPl = "SELECT PLURAL.PL FROM Plural INNER JOIN Professions ON Plural.ID = Professions.ID_Plural WHERE Professions.CZ = (N'" + tutenhamon + "')"; //ЗАПРОС ДЛЯ НАХОЖДЕНИЯ ПРОФЕССИИ ВО МН.Ч

            //"SELECT PLURAL.PL FROM Plural INNER JOIN Profession ON Plural.ID = Profession.ID_Plural WHERE Profession.CZ = '" + tutenhamon + "'"


            SqlConnection secondConnection = new SqlConnection();
            secondConnection.ConnectionString = connectionString;
            
            secondConnection.Open();
            SqlCommand secondSqlCommand = new SqlCommand(sqlExpressionForFindingPl, secondConnection);// для польского
            SqlCommand thirdSqlCommand = new SqlCommand(sqlExpressionFindingCz, secondConnection);

  


            SqlDataReader readerTwo = secondSqlCommand.ExecuteReader();
            
            
            PlCompany ourPlCompany = FullStruct();

            data[0] = ourPlCompany.name;
            data[1] = ourPlCompany.nip.ToString();
            data[2] = ourPlCompany.region.ToString();
            data[3] = ourPlCompany.representant;
            data[4] = textBoxForNewCompanyName.Text;
            data[5] = textBoxForNewCompanyAdress.Text;
            data[6] = textBoxForNewCompanyIC.Text;
            data[7] = textBoxForNewCompanyRepresent.Text;
            data[8] = comboBoxForNewCompanyTypeOfWork.Text;
            data[9] = monthCalendarForNewCompany.SelectionStart.Day.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Month.ToString() + "." + monthCalendarForNewCompany.SelectionStart.Year.ToString();
            data[10] = ourPlCompany.krs.ToString();
            while (readerTwo.Read()) 
            {
                data[11] = Convert.ToString(readerTwo.GetValue(0));
            }
            readerTwo.Close();

            SqlDataReader readerThree = thirdSqlCommand.ExecuteReader();

            while (readerThree.Read())
            {
                data[8] = Convert.ToString(readerThree.GetValue(0));
            }
            readerThree.Close();


            //for (int i = 0; i < 5; i++)
            //{
            //    data[i] = textBoxForNewCompanyName.Text;
            //}

            return data;
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Control.DeleteFromSeznam(this.Name);
        }
    }
}


