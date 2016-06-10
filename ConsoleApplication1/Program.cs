using Domino;
using System.Runtime.InteropServices;
using System.Text;
using DWORD = System.UInt32;

using HANDLE = System.UInt32;

using STATUS = System.UInt16;

namespace ConsoleApplication1
{
    internal class Program
    {
        private NotesSession ns;
        private NotesDatabase db;
        private NotesDocument doc;

        [DllImport("nnotes.dll")]
        public static extern STATUS OSPathNetConstruct(string port, string server, string filename, StringBuilder retFullpathName);

        [DllImport("nnotes.dll")]
        public static extern STATUS NSFDbOpen(string path, out HANDLE phDB);

        [DllImport("nnotes.dll")]
        public static extern STATUS NSFDbClose(HANDLE hDb);

        [DllImport("nnotes.dll")]
        public static extern STATUS NSFDbGetUnreadNoteTable(HANDLE hDb, string user, ushort namelen, bool create, out HANDLE hList);

        [DllImport("nnotes.dll")]
        public static extern STATUS NSFDbUpdateUnread(HANDLE hDb, HANDLE hUnreadList);

        //[DllImport("nnotes.dll")]
        //public static extern bool IDIsPresent(HANDLE hTable, DWORD id);

        [DllImport("nnotes.dll")]
        public static extern bool IDScan(HANDLE hTable, bool first, out HANDLE id);

        [DllImport("nnotes.dll")]
        public static extern STATUS OSMemFree(HANDLE h);

        [DllImport("nnotes.dll")]
        public static extern STATUS IDDestroyTable(HANDLE h);

        public int getUnreadMail()
        {
            return 0;
        }

        private static void Main(string[] args)
        {
            Program prog = new Program();
        }

        public Program()
        {
            ns = new NotesSession();
            ns.Initialize("password");

            string mailServer = ns.GetEnvironmentString("MailServer", true);
            string mailFile = ns.GetEnvironmentString("MailFile", true);
            string userName = ns.UserName;

            System.Console.WriteLine($"mailServer: {mailServer}");
            System.Console.WriteLine($"mailFile: {mailFile}");
            System.Console.WriteLine($"userName: {userName}");

            StringBuilder fullpathName = new StringBuilder(512);
            OSPathNetConstruct(null, mailServer, mailFile, fullpathName);
            System.Console.WriteLine($"fullpathName: {fullpathName.ToString()}");

            HANDLE hNotesDB;
            HANDLE hUnreadListTable;

            NSFDbOpen(fullpathName.ToString(), out hNotesDB);
            System.Console.WriteLine($"hNotesDB: {hNotesDB.ToString()}");

            NSFDbGetUnreadNoteTable(hNotesDB, userName, (ushort)userName.Length, true, out hUnreadListTable);
            System.Console.WriteLine($"hUnreadListTable: {hUnreadListTable.ToString()}");
            db = ns.GetDatabase(mailServer, mailFile, false);

            int numUnreadMail = 0;
            bool first = true;
            HANDLE id;

            while (true)
            {
                numUnreadMail = 0; first = true;
                while (IDScan(hUnreadListTable, first, out id))
                {
                    doc = db.GetDocumentByID(id.ToString("X"));
                    string subject = (string)((object[])doc.GetItemValue("Subject"))[0];
                    string sender = (string)((object[])doc.GetItemValue("From"))[0];
                    if (!sender.Equals(""))
                    {
                        System.Console.WriteLine($"   Doc: {subject} / *{sender}*");
                        if (!sender.Equals(userName))
                            numUnreadMail++;
                    }
                    first = false;
                }
                //numUnreadMail -= 3;
                System.Console.WriteLine($"Unread mail: {numUnreadMail.ToString()}");
                System.Threading.Thread.Sleep(3000);
                NSFDbUpdateUnread(hNotesDB, hUnreadListTable);
            }
            IDDestroyTable(hUnreadListTable);
            NSFDbClose(hNotesDB);

            /*
            db = ns.GetDatabase(mailServer, mailFile, false);

            NotesView inbox = db.GetView("($Inbox)");
            doc = inbox.GetFirstDocument();
            System.Console.WriteLine($"Notes database: /{db.ToString()}");

            // NotesViewEntryCollection vc = inbox.GetAllUnreadEntries();

            while (doc != null)
            {
                System.DateTime lastAccessed = doc.LastAccessed;
                System.DateTime lastModified = doc.LastModified;
                System.DateTime created = doc.Created;

                //if ( (lastAccessed.Subtract(lastModified)).TotalSeconds==(double)0.0 && (created.Subtract(lastModified)).TotalSeconds<(double)60.0 )
                if (lastAccessed.CompareTo(lastModified) < 0)

                {
                    string subject = (string)((object[])doc.GetItemValue("Subject"))[0];
                    System.Console.WriteLine($"LastAccessed: {doc.LastAccessed} | LastModified: {doc.LastModified} | Created: {doc.Created} | Subject: {subject}");
                }
                doc = inbox.GetNextDocument(doc);
            }

            db = null;
            ns = null;

    */

            System.Console.WriteLine("Hello world!");
            System.Console.ReadLine();      // as pause
        }
    }
}