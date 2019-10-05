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

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Data.SqlClient;
using System.Collections.Generic;

namespace LEGAL
{
    
    public partial class Form1 : Form
    {

        //public Control Jonh = new Control();hh
        Fernando AllMethods = new Fernando(); // все вспомогательные функции вынесены в отдельный класс Fernando
        Dictionary<int, string> Countries = new Dictionary<int, string>(3); //словарь список стран, заполнен странами в методе  - button 1 click -
        Dictionary<int, string> DocumentType = new Dictionary<int, string>(2); //словарь список типа документа по которому приехал человек, заполнен странами в методе  - button 1 click -
        Dictionary<int, string> CountryOfWorking = new Dictionary<int, string>(2); // словарь из которого вставляеться 
        string filePath = "C:\\Legalizace";
        SqlConn conn = new SqlConn();
        SqlConnection connection;

        public Form1() 
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            string ConnectionStringForDataGridView = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;
            SqlConnection ConnectionForDataGridView = new SqlConnection(ConnectionStringForDataGridView);
            ConnectionForDataGridView.Open();

            string sqlExpressionForFindAllProfessions = "select Professions.Cz, Professions.Salary, Professions.Hours from Professions order by cz";

            SqlCommand CommandToGetAllInfProfession = new SqlCommand(sqlExpressionForFindAllProfessions, ConnectionForDataGridView);
            SqlDataReader ReaderFoDataGridView = CommandToGetAllInfProfession.ExecuteReader();

            List<string[]> data = new List<string[]>(); // динамический массив в котрый сложим профессии

            while(ReaderFoDataGridView.Read()) //заполняем динамический массив в который скалдываем професии 
            {
                data.Add(new string[3]);
                data[data.Count - 1][0] = ReaderFoDataGridView[0].ToString();
                data[data.Count - 1][1] = ReaderFoDataGridView[1].ToString();
                data[data.Count - 1][2] = ReaderFoDataGridView[2].ToString();
            }
            ReaderFoDataGridView.Close();
            ConnectionForDataGridView.Close();

            foreach(string[] s in data) // из динамического массива складываем в ДатаГридВью
            {
                dataGridView1.Rows.Add(s);
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ///////////////////////////////////////////////////////////// WAR ZONE
            ////////////////////////////////////////////////////////////
            CountryOfWorking.Add(4, "Republiki Czeskiej");
            CountryOfWorking.Add(5, "Słowacja");
            CountryOfWorking.Add(6, "Belgii");

            Countries.Add(1, "Czech Republic");
            Countries.Add(2, "Slovakia");
            Countries.Add(3, "Belgie");

            DocumentType.Add(1, "BIOPASSPORT");
            DocumentType.Add(2, "VIZA");
            // тут будут 3 шаблона CZ -SMLOUVA- PL -UMOVA- -ANEX-

            //  C:\\Users\\Dell\\Documents\\!!!Некит\\Form.xlsx

            object templateForCzSmlouva = filePath + "\\Templates\\HellOfADrug\\smlouva.docx";
            object templateForPlUmova = filePath + "\\Templates\\HellOfADrug\\umowa.docx";
            object templateForPlAneks = filePath + "\\Templates\\HellOfADrug\\AneksPl.docx";
            object templateForCestjak = filePath + "\\Templates\\HellOfADrug\\CESTAK.docx";

            ////-----------Бельгийские шаблоны-----------////
            object templateForBelSmlouva = filePath + "\\Templates\\HellOfADrug\\nlUmowa.docx";
            object templateForBelSmlouvaPl = filePath + "\\Templates\\HellOfADrug\\nlUmowaPl.docx";

            string connectionString = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString; //строка подключения

            int maxNumbWorkers = countPeopleInExcelFile() - 1;
            string[] ztmp = new string[15];
            string[] arrForPlCompanyInformation = new string[8];

            for (int i = 0; i<maxNumbWorkers;i++)
            {
                ztmp = readExcel(i);

                if(ztmp[6] == "PORTOFINO SP.Z O.O" || ztmp[6] == "VENEZIA SP.Z O.O" || ztmp[6] == "AGRAT SP.Z O.O" || ztmp[6] == "POLFIZ SP.Z O.O" || AllMethods.checkSpaceInFolderName(ztmp[1]) || ztmp[7]!=null) 
                {
                    SqlConnection connection = new SqlConnection(connectionString);
                    //SqlConnection connect = conn.OpenSqlConn(connectionString);
                    string sqlExpressionForFindingCz = "SELECT CompanyName, CompanyAdress, CompanyRegion, CompanyKRS, CompanyNIP, CompanyRepresentant, CertifikatNumber, CertifikatDate FROM CompaniesPL where CompanyName = (N'" + ztmp[6] + "')";

                    // vot tu oshibka byla connection.ConnectionString = connectionString;
                    connection.Open();
                    SqlCommand CommandForToGetAllInfAboutPlCompany = new SqlCommand(sqlExpressionForFindingCz, connection);

                    SqlDataReader ReaderReadDataFromCommandForGettingAllInformationForPlCompany = CommandForToGetAllInfAboutPlCompany.ExecuteReader();

                    while (ReaderReadDataFromCommandForGettingAllInformationForPlCompany.Read())
                    {
                        arrForPlCompanyInformation[0] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(0));
                        arrForPlCompanyInformation[1] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(1));
                        arrForPlCompanyInformation[2] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(2));
                        arrForPlCompanyInformation[3] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(3));
                        arrForPlCompanyInformation[4] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(4));
                        arrForPlCompanyInformation[5] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(5));
                       // arrForPlCompanyInformation[6] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(6));
                        string ll = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(6));
                        int cont = int.Parse(ll.Replace(" ", string.Empty));
                        arrForPlCompanyInformation[6] = Convert.ToString(cont);
                        DateTime date = new DateTime();
                        date = Convert.ToDateTime(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(7));
                        arrForPlCompanyInformation[7] = date.ToString("dd/MM/yyyy"); 
                       // arrForPlCompanyInformation[7] = Convert.ToString(ReaderReadDataFromCommandForGettingAllInformationForPlCompany.GetValue(7));
                       // string date = DateTime.Now.ToString("Myyyy");
                    }
                    ReaderReadDataFromCommandForGettingAllInformationForPlCompany.Close();
                    connection.Close();

                    if (ztmp[5] == null)
                    {
                        
                       //connection = conn.OpenSqlConn(connectionString);
                        
                        
                        SqlConnection connectionTwo = new SqlConnection();
                        string SqlExpressionForFindingAddressIfItsNull = "Select CompanyAddress from CompanyCustomers where CompanyName = @test"; //  "(N'" + ztmp[4] + "')";
                        connectionTwo.ConnectionString = connectionString;
                        connectionTwo.Open();
                        SqlCommand CommandToGetAdress = new SqlCommand(SqlExpressionForFindingAddressIfItsNull, connectionTwo);
                        CommandToGetAdress.Parameters.AddWithValue("@test",ztmp[4]);
                        SqlDataReader ReaderReadDataFromCommandForGettingPlAdress = CommandToGetAdress.ExecuteReader();

                        while (ReaderReadDataFromCommandForGettingPlAdress.Read())
                        {
                            ztmp[5] = Convert.ToString(ReaderReadDataFromCommandForGettingPlAdress.GetValue(0));

                        }
                        ReaderReadDataFromCommandForGettingPlAdress.Close();
                        connectionTwo.Close();
                        //conn.CloseSqlConn(connection);
                    }

                    if(ztmp[8] == "" || ztmp[8] == null)
                    {
                        SqlConnection ConnectionForFindingPaymentPerHour = new SqlConnection();
                        string SQlExpressionForFindingPaymentPerHour = "select Salary FROM Professions where Professions.CZ = (N'" + ztmp[7] + "')";
                        ConnectionForFindingPaymentPerHour.ConnectionString = connectionString;
                        ConnectionForFindingPaymentPerHour.Open();
                        SqlCommand CommandToGetHours = new SqlCommand(SQlExpressionForFindingPaymentPerHour, ConnectionForFindingPaymentPerHour);

                        SqlDataReader reader = CommandToGetHours.ExecuteReader();

                        while(reader.Read())
                        {
                            ztmp[8] = Convert.ToString(reader.GetValue(0));
                        }
                        reader.Close();
                        ConnectionForFindingPaymentPerHour.Close();

                    }

                    if(ztmp[9] == "" || ztmp[9] == null)
                    {
                        SqlConnection ConnectionForPlProffesion = new SqlConnection();
                        string SqlExpresseonForFindingProffesionInPolish = "select Pl FROM Professions where Professions.CZ = (N'" + ztmp[7] + "')";
                        ConnectionForPlProffesion.ConnectionString = connectionString;
                        ConnectionForPlProffesion.Open();

                        SqlCommand CommandGetPlProffesion = new SqlCommand(SqlExpresseonForFindingProffesionInPolish, ConnectionForPlProffesion);

                        SqlDataReader redaderTwo = CommandGetPlProffesion.ExecuteReader();
                        while(redaderTwo.Read())
                        {
                            ztmp[9] = Convert.ToString(redaderTwo.GetValue(0));
                        }

                        redaderTwo.Close();
                        ConnectionForPlProffesion.Close();

                    }

                    if(ztmp[10] == "" || ztmp[10] == null)
                    {
                        ztmp[10] = DateTime.Now.ToString("dd/M/yyyy");
                    }

                    if(ztmp[11] == "" || ztmp[11] == null)
                    {
                        SqlConnection ConnectionForFindingPlProffesionInCaps = new SqlConnection();
                        string SqlExpressionDorFindingProffesionInPolishCaps = "select Pl FROM Professions where Professions.CZ = (N'" + ztmp[7] + "')";
                        ConnectionForFindingPlProffesionInCaps.ConnectionString = connectionString;
                        ConnectionForFindingPlProffesionInCaps.Open();

                        SqlCommand CommandGetPlProfInCaps = new SqlCommand(SqlExpressionDorFindingProffesionInPolishCaps, ConnectionForFindingPlProffesionInCaps);

                        SqlDataReader redaderTwo = CommandGetPlProfInCaps.ExecuteReader();
                        while (redaderTwo.Read())
                        {
                            string jonh = Convert.ToString(redaderTwo.GetValue(0));

                            ztmp[11] = jonh.ToUpper();
                        }

                        redaderTwo.Close();
                        ConnectionForFindingPlProffesionInCaps.Close();
                    }

                    ztmp[14] = "zagluskaEsliPusto";

                    if(checkCountry(ztmp[4]) == Countries[3])
                    {
                        string NlProffession;
                        SqlConnection ConnectionForFindingNlProffession = new SqlConnection();
                        string SqlExpressionDorFindingProffesionInPolishCaps = "select Nl FROM Professions where Professions.Cz = (N'" + ztmp[7] + "')";
                        ConnectionForFindingNlProffession.ConnectionString = connectionString;
                        ConnectionForFindingNlProffession.Open();
                        SqlCommand CommandToGetNlProffession = new SqlCommand(SqlExpressionDorFindingProffesionInPolishCaps, ConnectionForFindingNlProffession);
                        object ObjectForGettingNlProffession = CommandToGetNlProffession.ExecuteScalar();
                        NlProffession = ObjectForGettingNlProffession.ToString();

                        ztmp[14] = NlProffession;
                        

                    }



                    ///////начинаем вставлять, шаблоны их три штуки
                    ///
                    if (!checkBox1.Checked)
                    {
                        if (checkCountry(ztmp[4])== Countries[3])
                        {
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForBelSmlouva);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForBelSmlouvaPl);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForPlAneks);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForCestjak);
                        }else
                        {
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForCzSmlouva);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForPlUmova);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForPlAneks);
                            TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForCestjak);
                        }

                            
                    }
                    else
                    {
                        TakeDataFromExcelPutItToWord(ztmp, arrForPlCompanyInformation, templateForCestjak);
                    }

                }else
                {
                    MessageBox.Show("Проверь PL название фирмы и только один пробел должен быть между именем и фамилией.");
                    break;
                }   
                
            }

            void TakeDataFromExcelPutItToWord(string[] CompanyHeadData, string[] WorkerData, object template ) // функция получает массив с данными и вставляет в три шаблона 
            {

                Microsoft.Office.Interop.Word.Application appMy = new Microsoft.Office.Interop.Word.Application();

                
                Microsoft.Office.Interop.Word.Document docMy = null;

                object missingMy = Type.Missing;

                docMy = appMy.Documents.Open(template, missingMy, missingMy);
                appMy.Selection.Find.ClearFormatting();
                appMy.Selection.Find.Replacement.ClearFormatting();

                int xxx = CompanyHeadData.Length + WorkerData.Length; //длина общего массива в котором храним два других массива, первый массив с данными из эксель второй данные по фирме их дб, и третий БУДЕТ с данными по Цестяку
                string[] tmpForData = new string[xxx];
                
                for(int z = 0; z< CompanyHeadData.Length;z++)
                {
                    tmpForData[z] = CompanyHeadData[z];
                }

                for(int k = 0; k< WorkerData.Length;k++)
                {
                    tmpForData[k + CompanyHeadData.Length] = WorkerData[k];
                }
                string date = DateTime.Now.ToString("Myyyy");

                string readNumber()
                {
                    string text = File.ReadAllText(filePath + "\\Templates\\docNumb.txt", Encoding.Default);
                    int newNumber = Convert.ToInt32(text) + 1;
                    string newText = Convert.ToString(newNumber);
                    File.WriteAllText(filePath + "\\Templates\\docNumb.txt", newText);

                    return newText;
                }

              


                appMy.Selection.Find.Execute("<zzz>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[0], 2);
                appMy.Selection.Find.Execute("<surname>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[1], 2);
                appMy.Selection.Find.Execute("<dateOfBirth>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[2], 2);
                appMy.Selection.Find.Execute("<passNumber>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[3], 2);
                appMy.Selection.Find.Execute("<companyCz>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[4], 2);
                appMy.Selection.Find.Execute("<companyCzAddress>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[5], 2);
                appMy.Selection.Find.Execute("<PlCompanyNameIfFull>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[6], 2);
                appMy.Selection.Find.Execute("<employmentCz>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[7], 2);
                //appMy.Selection.Find.Execute("<hoursPerMounth>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[8], 2);
                appMy.Selection.Find.Execute("<employmentPL>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[9], 2);
                //appMy.Selection.Find.Execute("<AddressOfLivingDontUSe>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[10], 2);
                appMy.Selection.Find.Execute("<dateOfStartWorkig>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[10], 2); //лишняя надо убрать
                appMy.Selection.Find.Execute("<workPlCestjak>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[11], 2);
                appMy.Selection.Find.Execute("<PlCompanyName>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[15], 2);
                appMy.Selection.Find.Execute("<PlCompanyAddress>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[16], 2);
                appMy.Selection.Find.Execute("<REGON>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[17], 2);
                appMy.Selection.Find.Execute("<KRS>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[18], 2);
                appMy.Selection.Find.Execute("<NIP>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[19], 2);
                appMy.Selection.Find.Execute("<Representant>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[20], 2);// + отсюда дописываю сегодняшнее инфо
                appMy.Selection.Find.Execute("<CertifikateNumber>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[21], 2);// + отсюда дописываю сегодняшнее инфо
                appMy.Selection.Find.Execute("<CertifikateDate>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[22], 2);// + отсюда дописываю сегодняшнее инфо
                appMy.Selection.Find.Execute("<employmentNl>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[14], 2);// + отсюда дописываю сегодняшнее инфо

                
                if (checkCountry(tmpForData[4]) == Countries[1])
                {
                    appMy.Selection.Find.Execute("<CountryOfWorking>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, Convert.ToString(CountryOfWorking[4]), 2);
                }
                else if (checkCountry(tmpForData[4]) == Countries[2])
                {
                    appMy.Selection.Find.Execute("<CountryOfWorking>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, Convert.ToString(CountryOfWorking[5]), 2);
                }
                else if (checkCountry(tmpForData[4]) == Countries[3])
                {
                    appMy.Selection.Find.Execute("<CountryOfWorking>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, Convert.ToString(CountryOfWorking[6]), 2);
                }


                if (checkCountry(tmpForData[4]) == Countries[1] || checkCountry(tmpForData[4]) == Countries[2])
                {                  
                    appMy.Selection.Find.Execute("<hoursPerMounth>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, Convert.ToDouble(tmpForData[8]), 2);
                }
                else if (checkCountry(tmpForData[4]) == Countries[3])
                {
                    appMy.Selection.Find.Execute("<hoursPerMounth>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, "59", 2);
                }
                
                if (ztmp[13] == "" || ztmp[13] == null)
                {
                    appMy.Selection.Find.Execute("<dateOfStartWorkigPlusTwoY>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, AllMethods.bhBjj(tmpForData[10]), 2);
                }else
                {
                    appMy.Selection.Find.Execute("<dateOfStartWorkigPlusTwoY>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, tmpForData[13], 2);
                }

                
                appMy.Selection.Find.Execute("<IdDate>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, readNumber(), 2);
                appMy.Selection.Find.Execute("<cestNumber>", missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, missingMy, date, 2);

                Random jj = new Random();
                
                string zzz = Convert.ToString(jj.Next(1,1000));

                if (!checkBox1.Checked)
                {
                    if (!Directory.Exists(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3])) // проверяем если есть папка, чисто по логике не должно быть, по-этому сразу можно сохранять туда smlovuCZ
                    {
                        //создаем папку и сразу сохраняем

                        string PathForFolder = filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3];
                        string nlNameForSmlouva = "\\SmlouvaNL";
                        string plNameForSmlouva = "\\SmlouvaCZ";
                        DirectoryInfo DirInfo = new DirectoryInfo(PathForFolder);
                        DirInfo.Create();


                        object FilePathForFakturaNL = (object)PathForFolder + nlNameForSmlouva + tmpForData[15] + ".docx";
                        object FilePathForFakturaPL = (object)PathForFolder + plNameForSmlouva + tmpForData[15] + ".docx";


                        if (checkCountry(ztmp[4]) == Countries[3])
                        {
                            

                            docMy.SaveAs2(FilePathForFakturaNL, missingMy, missingMy, missingMy);

                            //MessageBox.Show("Files Are Created!");
                            docMy.Close(false, missingMy, missingMy);
                            appMy.Quit(false, false, false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                        }else
                        {
                            docMy.SaveAs2(FilePathForFakturaPL, missingMy, missingMy, missingMy);

                            //MessageBox.Show("Files Are Created!");
                            docMy.Close(false, missingMy, missingMy);
                            appMy.Quit(false, false, false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                        }
                    }
                    else
                    {
                        //папка есть значит просто сохранеем в уже созданную папку

                        //проверяю уже существует файл контракта если файла контракта нет то
                        if (!File.Exists(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\UmovaPL" + tmpForData[15] + ".docx"))
                        {
                            object FilePathForFaktura = (object)filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\UmovaPL" + tmpForData[15] + ".docx";

                            docMy.SaveAs2(FilePathForFaktura, missingMy, missingMy, missingMy);

                            //MessageBox.Show("Files Are Created!");
                            docMy.Close(false, missingMy, missingMy);
                            appMy.Quit(false, false, false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                        }
                        else
                        {
                            if(!File.Exists(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\AnexPl" + tmpForData[15] + ".docx"))
                            {
                                object FilePathForFaktura = (object)filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\AnexPl" + tmpForData[15] + ".docx";

                                docMy.SaveAs2(FilePathForFaktura, missingMy, missingMy, missingMy);

                                //MessageBox.Show("Files Are Created!");
                                docMy.Close(false, missingMy, missingMy);
                                appMy.Quit(false, false, false);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                            }
                            else
                            {
                                object FilePathForFaktura = (object)filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\CESTAK" + tmpForData[15] + "_" + tmpForData[3] + ".docx";

                                docMy.SaveAs2(FilePathForFaktura, missingMy, missingMy, missingMy);
                                
                                //MessageBox.Show("Files Are Created!");
                                docMy.Close(false, missingMy, missingMy);
                                appMy.Quit(false, false, false);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                            }

                            
                        }if(!File.Exists(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\CESTAK" + tmpForData[15] + "_" + tmpForData[3] + ".docx") || !File.Exists(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\CESTAK" + tmpForData[15] + "_" + tmpForData[3] + ".docx"))
                        {
                            if(checkCountry(ztmp[4]) == Countries[1])
                            {
                                using (StreamWriter sw = File.CreateText(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\" + tmpForData[1] + "_"+ tmpForData[9] + "_" + tmpForData[8]+ "zl_" + checkHours(ztmp[9]) + "_" + "hours_" + "_" + tmpForData[12] + "_" + tmpForData[13] + ".txt"))
                                {
                                    sw.WriteLine(tmpForData[1] + "_" + tmpForData[3] + "\\" + tmpForData[1] + "_" + tmpForData[9] + "_" + tmpForData[8] + "zl" + "_" + tmpForData[12] + tmpForData[13]);
                                }
                            }else if(checkCountry(ztmp[4]) == Countries[3])
                                    {
                                        using (StreamWriter sw = File.CreateText(filePath + "\\Fresh L\\" + tmpForData[1] + "_" + tmpForData[3] + "\\" + tmpForData[1] + "_" + tmpForData[9] + "_" + "59zl" + "_" + "80" + "hours_" + tmpForData[12] + "_" + tmpForData[13] + ".txt"))
                                             {
                                                 sw.WriteLine(tmpForData[1] + "_" + tmpForData[3] + "\\" + tmpForData[1] + "_" + tmpForData[9] + "_" + "59zl" + "_" + tmpForData[12] + tmpForData[13]);
                                             }
                                    }

                        }



                    }

                }else
                {
                    string pathToFolder = filePath + "\\Fresh L\\" + tmpForData[4]; // путь к папке с названием фирмы PL
                    object FilePathForFaktura = (object)filePath + "\\Fresh L\\" + tmpForData[4] + "\\" + tmpForData[1] + "_" + tmpForData[3] + "_CESTAK.docx";

                    if (!File.Exists(pathToFolder))
                    {
                        DirectoryInfo CreateDirectoryObject = new DirectoryInfo(pathToFolder);
                        CreateDirectoryObject.Create();
                        docMy.SaveAs2(FilePathForFaktura, missingMy, missingMy, missingMy);

                        //MessageBox.Show("Files Are Created!");
                        docMy.Close(false, missingMy, missingMy);
                        appMy.Quit(false, false, false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                    }
                    else
                    {
                        docMy.SaveAs2(FilePathForFaktura, missingMy, missingMy, missingMy);
                        //MessageBox.Show("Files Are Created!");
                        docMy.Close(false, missingMy, missingMy);
                        appMy.Quit(false, false, false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appMy);
                    }

                }

            }

            MessageBox.Show("Files Are Created!");
                

            }
        

        private string checkHours(string profName)
        {
            string hours;
            string connectionString = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString;
            SqlConnection connectionForCheckCountry = new SqlConnection(connectionString);
            string queryForFindingCompanyCountry = "select hours from Professions where Professions.Pl = N'" + profName + "'";
            connectionForCheckCountry.Open();
            SqlCommand CommandToGetAllInfProfession = new SqlCommand(queryForFindingCompanyCountry, connectionForCheckCountry);

            object countryg = CommandToGetAllInfProfession.ExecuteScalar();
            hours = countryg.ToString();

            // SqlDataReader ReaderForGetCountry = CommandToGetAllInfProfession.ExecuteReader();

            // country = ReaderForGetCountry.GetString(1);
            connectionForCheckCountry.Close();
            return hours;
        }

        private string checkCountry(string companyName) // проверяет к какой стране принадлежит фирма и отдает название этой страны
        {
            string country;

            string connectionString = ConfigurationManager.ConnectionStrings["LEGAL.Properties.Settings.WorkDBFirstTryConnectionString"].ConnectionString; //строка подключения
            SqlConnection connectionForCheckCountry = new SqlConnection(connectionString);
            string queryForFindingCompanyCountry = "select name from Countries inner join CompanyCustomers on Countries.ID = CompanyCustomers.Country where CompanyCustomers.CompanyName = N'"+ companyName +"'";
            connectionForCheckCountry.Open();
            SqlCommand CommandToGetAllInfProfession = new SqlCommand(queryForFindingCompanyCountry, connectionForCheckCountry);

            object countryg =  CommandToGetAllInfProfession.ExecuteScalar();
            country = countryg.ToString();
            
            // SqlDataReader ReaderForGetCountry = CommandToGetAllInfProfession.ExecuteReader();

           // country = ReaderForGetCountry.GetString(1);
            connectionForCheckCountry.Close();
            return country;
        }
        

        private string[] readExcel(int index)
        {
            string res = filePath + "\\Form.xlsx";
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            index += 2;
            string[] data = new string[15]; //array[] = {ID, Surname, Name, DateOfBith, Passport number, CompanyCZ, CompanyPL, EmploymentCZ, EmploymentPL, ResidenceInPoland, date of start working}
            data[0] = xlWorkSheet.get_Range("A" + index.ToString()).Text;// get ID
            data[1] = xlWorkSheet.get_Range("B" + index.ToString()).Value; // get Surname and Name
            //data[2] = xlWorkSheet.get_Range("C" + index.ToString()).Value; // get Name
            data[2] = xlWorkSheet.get_Range("C" + index.ToString()).Text;// get DateOfBith
            data[3] = xlWorkSheet.get_Range("D" + index.ToString()).Value; // get Passport number
            data[4] = xlWorkSheet.get_Range("E" + index.ToString()).Value; // get CompanyCZ
            data[5] = xlWorkSheet.get_Range("F" + index.ToString()).Value; // get CompanyCZAdress
            data[6] = xlWorkSheet.get_Range("G" + index.ToString()).Value; // get CompanyPL
            data[7] = xlWorkSheet.get_Range("H" + index.ToString()).Value; // get EmploymentCZ
            data[8] = xlWorkSheet.get_Range("I" + index.ToString()).Text; //  Payment per hour
            //data[9] = xlWorkSheet.get_Range("J" + index.ToString()).Text; // get hours per month
            data[9] = xlWorkSheet.get_Range("J" + index.ToString()).Value;// get EmploymentPL
            //data[10] = xlWorkSheet.get_Range("K" + index.ToString()).Value; // get ResidenceInPoland
            data[10] = xlWorkSheet.get_Range("K" + index.ToString()).Text; // date of start working
            data[11] = xlWorkSheet.get_Range("L" + index.ToString()).Text; // cestjak position pl
            data[12] = xlWorkSheet.get_Range("M" + index.ToString()).Text; // document type
            data[13] = xlWorkSheet.get_Range("N" + index.ToString()).Text; // viza do
            data[14] = xlWorkSheet.get_Range("O" + index.ToString()).Text; // viza do


            xlWorkBook.Close();
            xlApp.Quit();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //return data
            return data;
        }

        int countPeopleInExcelFile() //функция считающая количество человек
        {
            string res = filePath + "\\Form.xlsx"; 
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(res, 0, true, 5, "", "", true);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            var data = xlApp.WorksheetFunction.CountA(xlWorkSheet.Columns[1]);

            xlWorkBook.Close();
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            int ndata = Convert.ToInt32(data);
            //return data
            return ndata;
        }

        private void сделатьКонтрактToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            if(Control.ControlListIfAlreadyExist(f2.Name) == false)
            {
                Control.AddFormToList(f2.Name);
                f2.Show();
            }
            
        }

        private void сделатьФактуруToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            if(Control.ControlListIfAlreadyExist(f3.Name) == false)
            {
                Control.AddFormToList(f3.Name);
                f3.Show();
            }
            
        }

        private void фирмаCZToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form5 f5 = new Form5();
            if(Control.ControlListIfAlreadyExist(f5.Name)== false)
            {
                Control.AddFormToList(f5.Name);
                f5.Show();
            }
           
        }
        
        private void профToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form6 f6 = new Form6();
            if (Control.ControlListIfAlreadyExist(f6.Name) == false)
            {
                Control.AddFormToList(f6.Name);
                f6.Show();
            }    
        }
    }
}