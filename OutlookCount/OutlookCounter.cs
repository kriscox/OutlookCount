using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Runtime.InteropServices;


namespace OutlookCount
{
    class OutlookCounter
    {
        public string mailBox { get; set; }

        static Outlook._Application oApp = null;
        public int totalSentMails = 0;
        public int totalReceivedMails = 0;
        string restrictDate = null;
        public List<CountAgent> CountAgents = new List<CountAgent> { };

        public OutlookCounter()
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
        }


        public void Process()
        {
            //Count mail
            CountMailbox();

            // send results
            //ShowResults();
        }

        /// <summary>
        ///     DateRange 
        ///     
        /// DateRange defines the range in time to search
        /// 
        /// </summary>
        public void SetDateRange(DateTime dtStart, DateTime dtEnd)
        {
            // get the begin and end of yesterday
            restrictDate = " [received] >= \"" + dtStart.Date.ToString("g") + "\" AND [received] <= \"" + (dtEnd.Date + new TimeSpan(23, 59, 59)).ToString("g") + "\"";
        }

        /// <summary>
        /// Read the codes form the file given by the file parameter into codeAgents
        /// </summary>
        /// <param name="v"></param>
        public int ReadCodes(string filename)
        {
            // Clear current codes.
            CountAgents = new List<CountAgent> { };

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
        private void ShowResults()
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
        private void CountMailbox()
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
        }

        private void ReceivedMailsCount(string restrictYesterday, Folder folder)
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

        private void sentMailsCount(string restrictYesterday, Folder folder)
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

        /// <summary>
        /// Search in the mail for the first code of an agent found
        /// </summary>
        /// <param name="mailBody">body of the mail to search in</param>
        private void GetAgentFromMail(string mailBody)
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

        public List<String> NameSpaces()
        {
            List<String> ReturnValue = new List<String>();

            // Get the MAPI namespace.
            Outlook.NameSpace NS = oApp.GetNamespace("Mapi");

            // Run over all mailboxes
            foreach (Folder folder in NS.Folders)
            {
                ReturnValue.Add(folder.Name);
            }

            return ReturnValue;
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