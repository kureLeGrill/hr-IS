using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace LEGAL
{
    class Fernando //эксперементальный класс
    {
        public void AddProfession() // метод который добавляет профессию
        {

        }

        public bool checkSpaceInFolderName(string newFolderName)  // проверяет количество пробелов в названии папки и считает их колл
        {
            int spaceCounter = 0;
            char[] arrForStringFolderName;

            arrForStringFolderName = newFolderName.ToCharArray();

            for (int i = 0; i < newFolderName.Length; i++)
            {
                if (arrForStringFolderName[i] == ' ')
                {
                    spaceCounter++;
                }

                if (spaceCounter == 2)
                {
                    return false;
                }

            }
            return true;
        }

        public string jungleBjj(string firstDate) //function for data change by day изменяем дату на день <dateOfStartWorkigMinusDay>
        {
            DateTime result;
            result = DateTime.ParseExact(firstDate, "d", null);
            result = result.AddDays(-1);

            
            return result.ToString("dd.MM.yyyy"); 
        }


        public string bhBjj(string firstDate) //function for data change by year <dateOfStartWorkigMinusDayPlusTwoY>
        {
            DateTime result;
            //result = DateTime.ParseExact(firstDate, "d", null);
            result = Convert.ToDateTime(firstDate);
            result = result.AddDays(+730);
           
            return result.ToString("dd.MM.yyyy");
        }


        public void changeTextForContracts(object templateForNewCompanyContract, string[] data)
        {
            //now create word file into template and fill data
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //load document
            Microsoft.Office.Interop.Word.Document doc = null;

            object missing = Type.Missing;

            doc = app.Documents.Open(templateForNewCompanyContract, missing, missing);
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();

            string[] tmpForData = new string[12];
            tmpForData = data;

            app.Selection.Find.Execute("<PlCompanyName>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[0], 2);
            app.Selection.Find.Execute("<PlCompanyNIP>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[1], 2);
            app.Selection.Find.Execute("<PlCompanyDIC>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[2], 2);
            app.Selection.Find.Execute("<PlCompanyRepresentant>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[3], 2);


            app.Selection.Find.Execute("<newCompanyName>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[4], 2);
            app.Selection.Find.Execute("<newCompanyAdress>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[5], 2);
            app.Selection.Find.Execute("<newCompanyIC>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[6], 2);
            app.Selection.Find.Execute("<newCompanyRepresentant>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[7], 2);
            app.Selection.Find.Execute("<newCompanyTypeOfWork>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[8], 2);
            app.Selection.Find.Execute("<newCompanyDate>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[9], 2);

            app.Selection.Find.Execute("<PlCompanyKRS>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[10], 2);

            app.Selection.Find.Execute("<newCompanyTypeOfWorkPl>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[11], 2);

            Random rand = new Random();
            int ddd;
            ddd = rand.Next();


            object FilePathForContractCz = (object)"Y:\\Legalizace\\Fresh L\\!!!Contracts\\" + tmpForData[1] + "_" + ddd + "Smlouva.docx";
            doc.SaveAs2(FilePathForContractCz, missing, missing, missing);

            MessageBox.Show("Files Are Created!");
            doc.Close(false, missing, missing);
            app.Quit(false, false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }

        public string NumbersToWords(int number) // класс для перевода чисел в слова
        {
            if(number == 0)
            {
                return "zero";
            }

            if(number < 0)
            {
                return "minus" + NumbersToWords(Math.Abs(number));
            }

            string words = "";

            if((number / 1000000) > 0)
            {
                words += NumbersToWords(number / 1000000) + " milion ";
                number %= 1000000;
            }

            if((number / 1000) > 0)
            {
                words += NumbersToWords(number / 1000) + " tysiąc ";
                number %= 1000;
            }

            if((number/100) == 1 )
            {
                words += NumbersToWords(number / 100) + " sto ";
                number %= 100;
            }

            if ((number / 100) > 0)
            {
                words += NumbersToWords(number / 100) + " set ";
                number %= 100;
            }

            if (number> 0)
            {
                if (words != "") words += " i ";
                var unitsMap = new[] { "zero", "jeden", "dwa", "trzy", "cztery", "pięć", "sześć", "siedem", "osiem", "dziewięć", "dziesięć", "jedenaście", "dwanaście", "trzynaście", "czternaście", "piętnaście", "szesnaście", "siedemnaście", "osiemnaście", "dziewiętnaście" };
                var tensMap = new[] { "zero", "dziesięć", "dwadzieścia", "trzydzieści", "czterdzieści", "pięćdziesiąt", "sześćdziesiąt", "siedemdziesiąt", "osiemdziesiąt", "dziewięćdziesiąt" };

                if (number < 20) words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0) words += "-" + unitsMap[number % 10];
                }
            }

            return words;
        }

        public void ChangeTextForFaktura(string[] data)
        {
            if (data.Length == 17)
            {
                object templateForFaktura = "Y:\\Legalizace\\Templates\\faktura.docx";

                //now create word file into template and fill data
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

                //load document
                Microsoft.Office.Interop.Word.Document doc = null;

                object missing = Type.Missing;

                doc = app.Documents.Open(templateForFaktura, missing, missing);
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();

                string[] tmpForData = new string[17];
                tmpForData = data;

                app.Selection.Find.Execute("<GoodsName>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[0], 2);
                app.Selection.Find.Execute("<Type>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[1], 2);
                app.Selection.Find.Execute("<Qant>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[2], 2);
                app.Selection.Find.Execute("<Price>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[3], 2);
                app.Selection.Find.Execute("<Currency>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[4], 2);
                app.Selection.Find.Execute("<DateWystawenia>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[5], 2);
                app.Selection.Find.Execute("<DateSprzedazy>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[6], 2);
                app.Selection.Find.Execute("<NameOfPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[7], 2);
                app.Selection.Find.Execute("<AddresPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[8], 2);
                app.Selection.Find.Execute("<IcPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[9], 2);
                app.Selection.Find.Execute("<NameOfCzCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[10], 2);
                app.Selection.Find.Execute("<AddresOfCZCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[11], 2);
                app.Selection.Find.Execute("<IcOfCZCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[12], 2);
                app.Selection.Find.Execute("<NumberOfFactura>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[13], 2);
                app.Selection.Find.Execute("<gotowka>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[14], 2);
                app.Selection.Find.Execute("<AllCountPrice>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[15], 2);
                app.Selection.Find.Execute("<PriceToWords>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[16], 2);

                object FilePathForFaktura = (object)"Y:\\Legalizace\\Faktury\\" + tmpForData[7] + "_" + tmpForData[10] + "_" + tmpForData[13] +"_"+ tmpForData[14] + "_Faktura.docx";
                doc.SaveAs2(FilePathForFaktura, missing, missing, missing);

                MessageBox.Show("Files Are Created!");
                doc.Close(false, missing, missing);
                app.Quit(false, false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
            else {
                object templateForFakturaTrans = "Y:\\Legalizace\\Templates\\fakturaTrans.docx";

                //now create word file into template and fill data
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

                //load document
                Microsoft.Office.Interop.Word.Document doc = null;

                object missing = Type.Missing;

                doc = app.Documents.Open(templateForFakturaTrans, missing, missing);
                app.Selection.Find.ClearFormatting();
                app.Selection.Find.Replacement.ClearFormatting();

                string[] tmpForData = new string[22];
                tmpForData = data;

                app.Selection.Find.Execute("<GoodsName>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[0], 2);
                app.Selection.Find.Execute("<Type>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[1], 2);
                app.Selection.Find.Execute("<Qant>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[2], 2);
                app.Selection.Find.Execute("<Price>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[3], 2);
                app.Selection.Find.Execute("<Currency>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[4], 2);
                app.Selection.Find.Execute("<DateWystawenia>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[5], 2);
                app.Selection.Find.Execute("<DateSprzedazy>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[6], 2);
                app.Selection.Find.Execute("<NameOfPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[7], 2);
                app.Selection.Find.Execute("<AddresPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[8], 2);
                app.Selection.Find.Execute("<IcPlCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[9], 2);
                app.Selection.Find.Execute("<NameOfCzCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[10], 2);
                app.Selection.Find.Execute("<AddresOfCZCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[11], 2);
                app.Selection.Find.Execute("<IcOfCZCompany>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[12], 2);
                app.Selection.Find.Execute("<NumberOfFactura>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[13], 2);
                app.Selection.Find.Execute("<PayTime>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[14], 2);
                app.Selection.Find.Execute("<Konto>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[15], 2);
                app.Selection.Find.Execute("<IBAN>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[16], 2);
                app.Selection.Find.Execute("<KontoNumber>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[17], 2);
                app.Selection.Find.Execute("<SWIFT>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[18], 2);
                app.Selection.Find.Execute("<przelew>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[19], 2);
                app.Selection.Find.Execute("<AllCountPrice>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[20], 2);
                app.Selection.Find.Execute("<PriceToWords>", missing, missing, missing, missing, missing, missing, missing, missing, tmpForData[21], 2);

                object FilePathForFaktura = (object)"Y:\\Legalizace\\Faktury\\" + tmpForData[7] + "_" + tmpForData[10] +"_"+ tmpForData[20] +"_"+ tmpForData[19] + "_Faktura.docx";
                doc.SaveAs2(FilePathForFaktura, missing, missing, missing);

                MessageBox.Show("Files Are Created!");
                doc.Close(false, missing, missing);
                app.Quit(false, false, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
                

        }

    }
}
