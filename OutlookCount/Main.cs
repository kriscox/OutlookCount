using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;

namespace OutlookCount
{
    public partial class Main : Form
    {
        private OutlookCounter outlookCounter;
        //private ListViewColumnSorter lvwColumnSorter;

        public Main()
        {
            outlookCounter = new OutlookCounter();
            InitializeComponent();
        }

        private void OpenFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (outlookCounter == null) return;
            this.cMailboxes.DataSource = outlookCounter.NameSpaces();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = FileNameBox.Text;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileNameBox.Text = openFileDialog1.FileName;
            }

        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            
            if (0 != outlookCounter.ReadCodes(FileNameBox.Text))
            {
                MessageBox.Show("Erreur dans fichier de codes", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (BeginDatePicker.Value > EndDatePicker.Value)
            {
                DateTime temp = BeginDatePicker.Value;
                BeginDatePicker.Value = EndDatePicker.Value;
                EndDatePicker.Value = temp;
            }
            outlookCounter.SetDateRange(BeginDatePicker.Value, EndDatePicker.Value);

            outlookCounter.mailBox = cMailboxes.SelectedItem.ToString();

            Cursor.Current = Cursors.WaitCursor;
            this.Enabled = false;

            outlookCounter.Process();

            ShowResults();

            this.Enabled = true;
            Cursor.Current = Cursors.Default;
        }

        private void ShowResults()
        {
            // empty listview1
            listView1.Items.Clear();
            listView1.ListViewItemSorter = null;

            listView1.Items.Add(new ListViewItem(
                new String[] { "Total mails envoyés", outlookCounter.totalSentMails.ToString() }));
            listView1.Items.Add(new ListViewItem(
                new String[] { "Total mails recus", outlookCounter.totalReceivedMails.ToString() }));
            foreach (CountAgent agent in outlookCounter.CountAgents)
            {
                listView1.Items.Add(new ListViewItem(
                    new String[] { agent.AgentCode, agent.Amount.ToString() }));
            }

            listView1.Enabled = true;
        }

        private void ListView1_KeyUp(object sender, KeyEventArgs e)
        {
            if (sender != listView1) return;

            if (e.Control && e.KeyCode == Keys.C)
                CopySelectedValuesToClipboard();
        }

        private void CopySelectedValuesToClipboard()
        {
            var builder = new System.Text.StringBuilder();
            foreach (ListViewItem item in listView1.Items)
                builder.AppendLine(item.SubItems[1].Text);

            Clipboard.SetText(builder.ToString());
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            agentComparer sorter = listView1.ListViewItemSorter as agentComparer;

            if (sorter == null)
            {
                sorter = new agentComparer(e.Column);
                listView1.ListViewItemSorter = sorter;
            }
            else
            {
                if (e.Column == sorter.Column)
                    sorter.Order = sorter.Order == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
                else
                {
                    sorter.Order = SortOrder.Ascending;
                    sorter.Column = e.Column;
                }
            }

            listView1.Sort();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Main_ResizeEnd(object sender, EventArgs e)
        {
            listView1.Height = Height - 190;
        }
    }
    public class agentComparer : IComparer
    {
        //column used for comparison
        public SortOrder Order = SortOrder.Ascending;
        public int Column { get; set; }

        public agentComparer(int colIndex)
        {
            Column = colIndex;
        }
        public int Compare(object a, object b)
        {

            int result;
            ListViewItem itemA = a as ListViewItem;
            ListViewItem itemB = b as ListViewItem;
            if (itemA == null && itemB == null)
                result = 0;
            else if (itemA == null)
                result = -1;
            else if (itemB == null)
                result = 1;
            if (itemA == itemB)
                result = 0;
            //alphabetic comparison
            if (Column == 0)
                result = String.Compare(itemA.SubItems[0].Text, itemB.SubItems[0].Text);
            else
            {
                int _a = int.Parse(itemA.SubItems[1].Text);
                int _b = int.Parse(itemB.SubItems[1].Text);
                result = _a.CompareTo(_b);
            }

            if (Order == SortOrder.Descending)
                return -result;
            else
                return result;
        }

    }
}
