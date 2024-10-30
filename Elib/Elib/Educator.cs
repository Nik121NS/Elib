using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Elib
{
    public class Educator
    {
        public string Id;
        public string FIO;
        public String LinkOnPublish;
        public List<Publish> Published;

        public Educator()
        {
            Published = new List<Publish>();
        }
    }

    public class Publish
    {
        public string Id;
        public String Name;
        public String Link;
        public String Anotation;
        public String Citat;
        public String Journal;
        public String Autors;
        public String Status;
        public String RINC;
        public String Core_RINC;
        public String Recenz;
        public String Citir;
        public int CountPage;
    }
}
