using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LEGAL
{
   public static class Control
    {

       public static List<string> SeznamForm = new List<string>();

        public static void AddFormToList(string formName)
        {
            if (ControlListIfAlreadyExist(formName) == false)
            {
                SeznamForm.Add(formName);
            }
            

        }

        public static Boolean ControlListIfAlreadyExist(string formName)
        {
            
            for(int i=0; i<SeznamForm.Count;i++)
            {
                if (SeznamForm[i] == formName)
                {
                    return true;
                }
                
               
            }

            return false;
        }

        public static void DeleteFromSeznam(string formName)
        {
            for(int i=0;i<SeznamForm.Count;i++)
            {
                if(SeznamForm[i] == formName)
                {
                    SeznamForm.Remove(formName);
                }
             }
        }

    }
}
