using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace _3PL_Text2Excel
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        //static MainForm mainForm = null;
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                //mainForm = new MainForm();
                //Application.Run(mainForm);
                Application.Run(new MainForm());
            }
            else
            {
                FileHandler handler = new FileHandler(args[0].ToUpper());
                handler.Process();
            }
        }
    }
}
