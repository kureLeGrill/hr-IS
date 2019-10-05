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
using System.Configuration;


namespace LEGAL
{
    public partial class Form6 : Form
    {
        
        string ConnectionStringForAddProfession = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;

        public Form6()
        {
            InitializeComponent();
        }
        
        string ProfNameCZ, ProfNamePL, ProfNameCZPlural, ProfNamePLPlural, WageZL, Hours, ProfNameNl, ProfNameNlPlural;

        private void Form6_FormClosing(object sender, FormClosingEventArgs e)
        {
            Control.DeleteFromSeznam(this.Name);
        }

        string[] professionDataForPlural = new string[2];
        string[] professionDataForUsal = new string[3];

        private void FillArrayProfessionDataF()
        {
            SqlConnection ConnectionForForm6 = new SqlConnection(ConnectionStringForAddProfession);
            ProfNameCZ = textBoxForCZCompanzName.Text;
            ProfNamePL = textBox1.Text;
            ProfNameCZPlural = textBoxAForCzProfessionPlural.Text;
            ProfNamePLPlural = textBox2.Text;
            WageZL = textBoxWage.Text;
            Hours = textBoxHours.Text;
            ProfNameNl = textBoxForNL.Text;
            ProfNameNlPlural = textBoxForNlPlural.Text;

            int CompanyIdPL = 100;
            string RequestForfindingLastAvailableId = "select MAX(ID) from Plural"; // находим макс айди из существцющих у профессий с профессий в мн числе             
            ConnectionForForm6.Open();
            SqlCommand CommandForGetIdPlural = new SqlCommand(RequestForfindingLastAvailableId, ConnectionForForm6);
         
            SqlDataReader reader = CommandForGetIdPlural.ExecuteReader();
           // SqlDataReader readerTwo = CommandForGetIDSingle.ExecuteReader();

            while(reader.Read())
            {
                CompanyIdPL = Convert.ToInt32(reader.GetValue(0)) + 1;
            }
            reader.Close();

            string RequestForCreatingNewPluralProfession = "insert into Plural(Id, Cz, Pl, Nl ) VALUES(" + CompanyIdPL + ", N'" + ProfNameCZPlural + "', N'" + ProfNamePLPlural + "', N'" + ProfNameNlPlural + "')"; // добавляем профессию во мн числе

            using (ConnectionForForm6)
            {
                SqlCommand CommandToAddNewProf = new SqlCommand(RequestForCreatingNewPluralProfession, ConnectionForForm6);
                int number = CommandToAddNewProf.ExecuteNonQuery();
                MessageBox.Show(ProfNameCZPlural + " была добавлена.");
                ConnectionForForm6.Close();
            }

            FillArrayTwo(CompanyIdPL);

        }

        public void FillArrayTwo(int CompanyIdPL)
        {
            int CompanyIdSingle = 99;
            string RequestForFidingLastAvailableIdForProf = "select MAX(ID) from Professions"; // ищем последний доступный айди
            SqlConnection ConnectionForForm6 = new SqlConnection(ConnectionStringForAddProfession);
            ConnectionForForm6.Open();

            SqlCommand CommandForGetIDSingle = new SqlCommand(RequestForFidingLastAvailableIdForProf, ConnectionForForm6);
            SqlDataReader reader = CommandForGetIDSingle.ExecuteReader();

            while (reader.Read())
            {
                CompanyIdSingle = Convert.ToInt32(reader.GetValue(0)) + 1;
            }
            reader.Close();

            string RequestForInsertingSingleFormProfession = "insert into Professions(Id, Cz, Pl, Nl, Salary, Id_Plural, Hours) Values (" + CompanyIdSingle + ", N'" + ProfNameCZ + "', N'" + ProfNamePL + "',N'" + ProfNameNl + "', N'" + WageZL + "', N'" + CompanyIdPL + "', N'" + Hours + "')"; // вставляем последнюю проф

            using (ConnectionForForm6)
            {
                SqlCommand CommandToAddNewProf = new SqlCommand(RequestForInsertingSingleFormProfession, ConnectionForForm6);
                int number = CommandToAddNewProf.ExecuteNonQuery();

                ConnectionForForm6.Close();
                MessageBox.Show(ProfNameCZ + " была добавлена.");
            }
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            FillArrayProfessionDataF();
        }
    }
}
