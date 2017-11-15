using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using System.Linq;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Data.SqlClient;
using System.Data.Common;
using System.Drawing;

namespace zoonn
{
    public partial class Form1 : Form
    {
        private readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        private readonly Pop3Client pop3Client;
        bool restart = true;
     
        NotifyIcon myIcon = new NotifyIcon();


        public Form1()
        {           
            
            InitializeComponent();
            List<string> asd = new List<string> { };
            FetchAllMessages("pop.gmail.com", 995, true, "-e-mail", "-password", asd);      

        }
     

        public /*static List<Message> */ void FetchAllMessages(string hostname, int port, bool useSsl, string username, string password, List<string> seenUids)
        {
            // The client disconnects from the server when being disposed

            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Fetch all the current uids seen
                List<string> uids = client.GetMessageUids();

                // Create a list we can return with all new messages
                List<Message> newMessages = new List<Message>();

                // All the new messages not seen by the POP3 client
                for (int i = 0; i < uids.Count; i++)
                {
                    string currentUidOnServer = uids[i];
                    if (!seenUids.Contains(currentUidOnServer))
                    {

                        // Download it and add this new uid to seen uids
                        Message unseenMessage = client.GetMessage(i + 1);

                        // Check message header
                        if (unseenMessage.Headers.Subject == "zoniletisim")
                        {
                            // Check also date is today
                            if (unseenMessage.Headers.DateSent.ToShortDateString() == DateTime.Today.ToShortDateString())
                            {
                                // Add the message to the new messages
                                newMessages.Add(unseenMessage);

                                // Add the uid to the seen uids, as it has now been seen
                                seenUids.Add(currentUidOnServer);
                            }

                        }
                        // if subject is not zoniletisim then delete message from gmail
                        //else
                        //{
                        //    client.DeleteAllMessages();

                        //}

                    }
                }

                string body = "";
                string bodyHtml = "";
                var splitedBody = new string[] { };
                List<string> splitList = new List<string>();
                int count = 0;
                foreach (var item in newMessages)
                {
                    // Count for message select
                    count++;

                    // Check message is null
                    if (item != null)
                    {

                        // Get first message
                        OpenPop.Mime.MessagePart messagePart = item.MessagePart.MessageParts[1];

                        // Get mail body
                        body = messagePart.GetBodyAsText();

                        // Change body format 
                        bodyHtml = StripHTML(body);

                        // Get important part with split method
                        splitedBody = bodyHtml.Split(new string[] { "8: Promosyon2" }, StringSplitOptions.RemoveEmptyEntries);
                        splitedBody = splitedBody[1].Split(new string[] { " ", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);


                        var totalString = new string[11];
                        int j = 0;
                        for (int i = 0; i < splitedBody.Length; i++)
                        {                             
                            totalString[j] = splitedBody[i];                           

                            if (j % 10 == 0 && j != 0)
                            {
                                DatabaseSave(totalString);
                                j = -1;
                                Array.Clear(totalString, 0, totalString.Length);
                            }
                                j++;
                        }
                    }   
                }
            }
        }

        // Convert string to html
        public static string StripHTML(string htmlString)
        {

            string pattern = @"<(.|\n)*?>";

            return Regex.Replace(htmlString, pattern, " ");
        }

        public void DatabaseSave(string []data)
        {
            bool save = true;
            int count = 0;

            Back:
            var connection = System.Configuration.ConfigurationManager.ConnectionStrings["Test"].ConnectionString;

            if (connection == string.Empty && connection == "")
            {
                goto Back;
            }

            while (save)
            {
                using (SqlConnection conn = new SqlConnection(connection))
                {
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO Data(Date, Pisirme, Porselen, Path_Sag, Path_Sol, Promosyon1, Aksesuar, Sofra, MutfakAlet, Promosyon2)VALUES(@date, @pisirme, @porselen, @path_sag, @path_sol, @promosyon1, @aksesuar, @sofra, @mutfakalet, @promosyon2)", conn))
                    {
                        conn.Open();
                        string myDate = data[0] + " " + data[1];
                        DateTime date = DateTime.Parse(myDate);                       
                        cmd.Parameters.AddWithValue("@date", date);
                        cmd.Parameters.AddWithValue("@pisirme", data[2]);
                        cmd.Parameters.AddWithValue("@porselen", data[3]);
                        cmd.Parameters.AddWithValue("@path_sag", data[4]);
                        cmd.Parameters.AddWithValue("@path_sol", data[5]);
                        cmd.Parameters.AddWithValue("@promosyon1", data[6]);
                        cmd.Parameters.AddWithValue("@aksesuar", data[7]);
                        cmd.Parameters.AddWithValue("@sofra", data[8]);
                        cmd.Parameters.AddWithValue("@mutfakalet", data[9]);
                        cmd.Parameters.AddWithValue("@promosyon2", data[10]);

                        // Check save is ok
                        int cmd0 = cmd.ExecuteNonQuery();

                        if (cmd0 == 1)
                        {
                            save = false;
                            conn.Close();
                            break;
                        }

                        // Wait 5 second
                        System.Threading.Thread.Sleep(5000);

                        // Add +1 to count. If count is 10, skip next data 
                        count++;
                        if (count == 10)
                        {
                            conn.Close();
                            break;
                        }
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            myIcon.Icon = new Icon("g.ico");
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                Hide();
                myIcon.Text = "Zon iletisim";
                myIcon.Visible = true;
                Visible = false; // Hide form window.
                ShowInTaskbar = false; // Remove from taskbar            
                myIcon.MouseDoubleClick += new MouseEventHandler(MyIcon_MouseDoubleClick);
            }
        }

        void MyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            ShowInTaskbar = true;
            myIcon.Visible = false;
          
        }
    }
}
