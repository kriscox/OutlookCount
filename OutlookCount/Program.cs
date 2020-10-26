using Microsoft.Office.Interop.Outlook;
using Microsoft.SqlServer.Server;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookCount
{
    class CountAgent
    {
        public string AgentCode { get; set; }
        public int Amount { get; set; } = 0;

        public CountAgent(string _agentCode)
        {
            AgentCode = _agentCode;
        }
    }

    class Program
    {
        static string mailBox = null;
        static Outlook._Application oApp = null;
        static int totalSentMails = 0;
        static int totalReceivedMails = 0;
        static string restrictDate = null;
        static readonly List<CountAgent> CountAgents = new List<CountAgent> { };

        static int Main(string[] args)
        {
            // Test if input arguments were supplied.
            if ((args.Length != 4) || (args.Length != 6))
            {
                ErrorParameters();
                return 1;
            }

            // Read first two arguments
            int returnValue = ReadTwoArguments(args.Take(2).ToArray());
            if (returnValue != 0)
            {
                return returnValue;
            }

            // Read the next two arguments
            returnValue = ReadTwoArguments(args.Skip(2).Take(2).ToArray());
            if (returnValue != 0)
            {
                return returnValue;
            }

            if (args.Length == 6)
            {
                returnValue = ReadTwoArguments(args.Skip(4).Take(2).ToArray());
                if (returnValue != 0)
                {
                    return returnValue;
                }
            }

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
                return 7;
            }

            // Set daterange to count if not given by parameters
            if (restrictDate == null) DateRange();

            //Count mail
            CountMailbox();

            // send results
            ShowResults();

            return 0;
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
        ///     ErrorParameters
        ///     
        /// Write errormessge for parameters
        /// 
        /// </summary>
        static void ErrorParameters()
        {
            Console.WriteLine("Please enter all obligatory arguments.");
            Console.WriteLine("Usage: outlookcount -e <mailbox name> -c <csv-file containing agent codes> [-d <date>]");
        }

        /// <summary>
        /// Read two arguments from a list and check their validity, if valid, put them in the necessairy objects
        /// </summary>
        /// <param name="args">list of 2 parameters</param>
        /// <returns>
        /// <list>
        ///     <item>0:  Succesfull </item>
        ///     <item>1:  Wrong code in parameter</item>
        ///     <item>5:  Wrong date format</item>
        ///     <item>6:  Error reading codes from file</item>
        ///     <item>7:  Wrong filename</item>
        /// </list>
        /// </returns>
        private static int ReadTwoArguments(string[] args)
        {

            // Read first argument
            if (args[0] == "-e")
            {
                // read mailbox name 
                mailBox = args[1];
            } // first argument Agent Codes
            else if (args[0] == "-c")
            {
                // read the agent codes from file
                int returnvalue = ReadCodes(args[1]);
                if (0 != returnvalue)
                {
                    return returnvalue;
                }
                if (CountAgents.Count() <= 0)
                {
                    Console.WriteLine("Agent codes are not read. Error in file format?");
                    return 6;
                }
            } // wrong Arguments, write error message and quit
            else if (args[0] == "-d")
            {

                DateTime givenDate;
                // Read the date
                try
                {
                    givenDate = DateTime.Parse(args[1]);
                }
                catch (FormatException e)
                {
                    Console.WriteLine("Wrong date format");
                    Console.WriteLine(e.Message);
                    return 5;
                }

                DateRange(givenDate);
            }
            else
            {
                ErrorParameters();
                return 1;
            }

            return 0;
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
    }
}


