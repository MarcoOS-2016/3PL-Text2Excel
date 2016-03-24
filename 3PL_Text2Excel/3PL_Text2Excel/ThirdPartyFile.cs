using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _3PL_Text2Excel
{
    public class ThirdPartyFile
    {
        private string filename;
        private string fullfilename;
        private string fileextension;

        public string FileName
        {
            get { return filename; }
            set { filename = value; }
        }

        public string FullFillName
        {
            get { return fullfilename; }
            set { fullfilename = value; }
        }

        public string FileExtension
        {
            get { return fileextension; }
            set { fileextension = value; }
        }
    }
}
