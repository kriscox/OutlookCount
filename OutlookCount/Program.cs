

using System;
using Application = System.Windows.Forms.Application;

namespace OutlookCount
{
    class Program
    {

        [STAThreadAttribute]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Main());

            return 0;
       }

 
    }
}


