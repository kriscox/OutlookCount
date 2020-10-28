using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCount
{
    public partial class Main : Form
    {
        static string mailBox = null;
        static Outlook._Application oApp = null;
        static int totalSentMails = 0;
        static int totalReceivedMails = 0;
        static string restrictDate = null;
        static readonly List<CountAgent> CountAgents = new List<CountAgent> { };

        public Main()
        {
            // Start outlook
            try
            {
                // Create the Outlook application.
                // in-line initialization
                oApp = new Outlook.Application();
            }
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught while opening Outlook: ", e);
            }

            InitializeComponent();
        }



        public void Process()
        {
            //Count mail
            CountMailbox();

            // send results
            ShowResults();
        }

        /// <summary>
        ///     DateRange 
        ///     
        /// DateRange defines the range in time to search
        /// 
        /// </summary>
        private static void DateRange(DateTime? givenDate = null)
        {
            // get the begin and end of yesterday

            DateTime dtRestrict = givenDate == null ? DateTime.Now.AddDays(-1) : (DateTime)givenDate;
            DateTime dtEnd = new DateTime(dtRestrict.Year, dtRestrict.Month, dtRestrict.Day, 23, 59, 59, 999);
            DateTime dtStart = new DateTime(dtRestrict.Year, dtRestrict.Month, dtRestrict.Day, 0, 0, 0, 0);
            restrictDate = " [received] >= \"" + dtStart.ToString("g") + "\" AND [received] <= \"" + dtEnd.ToString("g") + "\"";
        }

       /// <summary>
        /// Read the codes form the file given by the file parameter into codeAgents
        /// </summary>
        /// <param name="v"></param>
        private static int ReadCodes(string filename)
        {
            // CHeck if filename is actual a file
            if (string.IsNullOrEmpty(filename) || !File.Exists(filename))
            {
                Console.WriteLine("File " + filename + " does not exitst");
                return 7;
            }

            using (TextFieldParser csvReader = new TextFieldParser(filename))
            {
                csvReader.CommentTokens = new string[] { "#", "//" };
                csvReader.SetDelimiters(new string[] { ",", ";" });
                csvReader.HasFieldsEnclosedInQuotes = false;

                // Skip the row with the column names
                csvReader.ReadLine();

                while (!csvReader.EndOfData)
                {
                    // Read current line fields, pointer moves to the next line.
                    string[] fields = csvReader.ReadFields();
                    CountAgents.Add(new CountAgent(fields[0]));
                }

            }

            return 0;
        }

        /// <summary>
        /// Show result on console
        /// </summary>
        private static void ShowResults()
        {
            string result;
            Console.WriteLine("");
            result = string.Format("Total sent mail :     {0:### ##0}", totalSentMails);
            Console.WriteLine(result);
            result = string.Format("Total received mail : {0:### ##0}", totalReceivedMails);
            Console.WriteLine(result);
            foreach (CountAgent agent in CountAgents)
            {
                result = string.Format("{1,10} has treathed {0:### ##0} mails", agent.Amount, agent.AgentCode);
                Console.WriteLine(result);
            }
        }

        /// <summary>
        /// Count the number of mails per agent in the defined range in the mailbox specified.
        /// </summary>
        private static void CountMailbox()
        {
            // Get the MAPI namespace.
            Outlook.NameSpace NS = oApp.GetNamespace("Mapi");

            // Run over all mailboxes
            foreach (Folder folder in NS.Folders)
                if (folder.Name == mailBox)
                {
                    try
                    {
                        // Sent mails
                        sentMailsCount(restrictDate, folder);

                        // Received mails
                        ReceivedMailsCount(restrictDate, folder);

                    }
                    //catch faulty stores
                    catch (COMException)
                    {
                        continue;
                    };

                }

            return;

            void sentMailsCount(string restrictYesterday, Folder folder)
            {
                Items zItems = null;
                MAPIFolder sentFolder = folder.Folders["Éléments envoyés"];

                Items yItems = sentFolder.Items.Restrict(restrictYesterday);
                totalSentMails = yItems.Count;

                foreach (Object yItem in yItems)
                    if (yItem is MailItem item)
                    {
                        GetAgentFromMail(item.Body);
                    }
                    else continue;

                foreach (Folder subfolder in sentFolder.Folders)
                {
                    zItems = subfolder.Items.Restrict(restrictYesterday);
                    totalSentMails += zItems.Count;

                    foreach (Object yItem in yItems)
                        if (yItem is MailItem item)
                        {
                            GetAgentFromMail(item.Body);
                        }
                        else continue;
                }

            }

            void ReceivedMailsCount(string restrictYesterday, Folder folder)
            {
                Items zItems = null;
                MAPIFolder receiveFolder = folder.Folders["Boîte de réception"];

                foreach (Folder subfolder in receiveFolder.Folders)
                {
                    zItems = subfolder.Items.Restrict(restrictYesterday);
                    totalReceivedMails += zItems.Count;
                }

                zItems = receiveFolder.Items.Restrict(restrictYesterday);
                totalReceivedMails += zItems.Count;
            }
        }

        /// <summary>
        /// Search in the mail for the first code of an agent found
        /// </summary>
        /// <param name="mailBody">body of the mail to search in</param>
        private static void GetAgentFromMail(string mailBody)
        {
            String Name = null;
            int Position = -1;

            // Check for each code if it's included and if its found if it is before the previous one found
            // Position -1 is equal to not found.
            foreach (CountAgent agent in CountAgents)
            {
                int Loc = mailBody.IndexOf(agent.AgentCode);
                if (Loc > -1 && (Position == -1 || Loc < Position))
                {
                    Position = Loc;
                    Name = agent.AgentCode;
                }

            }

            // update the found agent
            if (Position >= 0)
            {
                CountAgent result = CountAgents.Find(
                    delegate (CountAgent agent)
                    {
                        return agent.AgentCode == Name;
                    });
                result.Amount++;
            }
        }

        private void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void Main_Load(object sender, EventArgs e)
        {
            var DataSource = new List<String>();

            // Get the MAPI namespace.
            Outlook.NameSpace NS = oApp.GetNamespace("Mapi");

            // Run over all mailboxes
            foreach (Folder folder in NS.Folders)
            {
                DataSource.Add(folder.Name);
            }

            this.cMailboxes.DataSource = DataSource;
        }
    }

    class CountAgent
    {
        public string AgentCode { get; set; }
        public int Amount { get; set; } = 0;

        public CountAgent(string _agentCode)
        {
            AgentCode = _agentCode;
        }
    }
}
