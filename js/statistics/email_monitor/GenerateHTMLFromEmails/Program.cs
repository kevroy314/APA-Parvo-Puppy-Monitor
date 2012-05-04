using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using System.Drawing;

namespace GenerateHTMLFromEmails
{
    class Program
    {
        private static readonly Pop3Client pop3Client = new Pop3Client();
        private static readonly Dictionary<int, Message> messages = new Dictionary<int, Message>();
        private static List<DogData> dogs = new List<DogData>();
        private static List<string> uids = new List<string>();
        private static bool monitorRunning = true;
        private static bool monitorQuit = false;
        private static int monitorCount = 0;
        private static string filename = "data.json";
        private static string header = "[\r\n\t{\r\n\t\t" +
		                               "\"id\": \"Puppies\",\r\n\t\t" +
		                               "\"title\": \"Parvo Puppies\",\r\n\t\t" +
		                               "\"focus_date\": \"2012-5-4 12:00\",\r\n\t\t" +
		                               "\"initial_zoom\": \"16\",\r\n\t\t" +
		                               "\"size_importance\": \"true\",\r\n\t\t" +
		                               "\"events\":\r\n\t\t" +
		                               "[";
        private static string footer = "\r\n\t\t]\r\n\t}\r\n]";
        static void Main(string[] args)
        {
            Console.SetWindowSize(40, 8);
            ReadCurrentPageStatus();
            System.Threading.Thread monitorDaemon = new System.Threading.Thread(new System.Threading.ThreadStart(MonitorMailbox));
            monitorDaemon.Start();
            Console.ReadLine();
            Console.WriteLine("Waiting for iteration to finish...");
            monitorRunning = false;
            while (!monitorQuit) System.Threading.Thread.Sleep(100);
            Console.WriteLine("Monitor Stopped, Quitting...");
        }
        private static void MonitorMailbox()
        {
            while (monitorRunning)
            {
                GetMessages();
                ParseEmailToDogData();
                PrintDogDataToHTML();
                monitorCount++;
                System.Threading.Thread.Sleep(1000);
            }
            monitorQuit = true;
        }
        private static void ReadCurrentPageStatus()
        {
            StreamReader reader = new StreamReader(filename);
            string document = reader.ReadToEnd().Replace(" \\", "").Remove(0, header.Length);
            document = document.Remove(document.Length - footer.Length, footer.Length);
            reader.Close();
        }
        private static void GetMessages()
        {
            if (pop3Client.Connected)
                pop3Client.Disconnect();
            pop3Client.Connect("pop.gmail.com", 995, true);
            pop3Client.Authenticate("apaparvolog", "austinpetsalive3141592");
            int count = pop3Client.GetMessageCount();
            Console.Clear();
            Console.WriteLine("Message Count: " + count);
            Console.WriteLine("Total Receieved This Session: " + uids.Count);
            Console.WriteLine("Press Enter To Quit.");

            int success = 0;
            int fail = 0;
            for (int i = count; i >= 1; i -= 1)
            {
                try
                {
                    Message message = pop3Client.GetMessage(i);
                    // Add the message to the dictionary from the messageNumber to the Message
                    messages.Add(i, message);
                    success++;
                }
                catch (Exception e)
                {
                    Console.Write(
                        "TestForm: Message fetching failed: " + e.Message + "\r\n" +
                        "Stack trace:\r\n" +
                        e.StackTrace);
                    fail++;
                }
            }
            List<string> currentUIDS = pop3Client.GetMessageUids();
            uids.AddRange(currentUIDS);
            pop3Client.Disconnect();
        }
        private static void ParseEmailToDogData()
        {
            for (int i = 1; i <= messages.Count; i++)
            {
                List<MessagePart> attachments = messages[i].FindAllAttachments();
                string imageFilename = "";
                if (attachments.Count > 0)
                {
                    attachments[0].Save(new FileInfo("..\\..\\..\\..\\..\\..\\images\\"+attachments[0].FileName));
                    imageFilename = attachments[0].FileName;
                }

                MessagePart message = messages[i].FindAllTextVersions()[0];
                string name = message.GetBodyAsText();
                int body_tag_start = name.IndexOf("<body");
                if (body_tag_start != -1)
                {
                    int body_tag_end = name.IndexOf(">", body_tag_start);
                    int body_end_tag = name.IndexOf("</body>");
                    name = name.Substring(body_tag_end + 1, body_end_tag - body_tag_end - 1).Replace("  ", "").Replace("\t", "").Replace("\n", "").Replace("\r", "");
                    while (name.Contains("<"))
                    {
                        int tag_start = name.IndexOf('<');
                        int tag_end = name.IndexOf('>');
                        name = name.Remove(tag_start, tag_end - tag_start + 1);
                    }
                }

                string description = "None.";

                string from = messages[i].Headers.From.Address;

                DateTime startTime = DateTime.Now;

                DogData d = new DogData(imageFilename, name, startTime, DateTime.Now, from, description);

                dogs.Add(d);
            }
        }
        private static void PrintDogDataToHTML()
        {
            FileStream writer = new FileStream(filename, FileMode.Open);
            for (int i = 0; i < dogs.Count; i++)
            {
                writer.Seek(-footer.Length, SeekOrigin.End);
                byte[] writeBytes = Encoding.ASCII.GetBytes(dogs[i].toHTML());
                writer.Write(writeBytes,0,writeBytes.Length);
            }
            byte[] footerBytes = Encoding.ASCII.GetBytes(footer);
            if(dogs.Count>0)
                writer.Seek(0, SeekOrigin.End);
            else
                writer.Seek(-footer.Length, SeekOrigin.End);
            writer.Write(footerBytes, 0, footerBytes.Length);
            writer.Close();
            messages.Clear();
            dogs.Clear();
        }
    }
    class DogData
    {
        private string iFilename;
        private string n;
        private DateTime sTime;
        private DateTime eTime;
        private string f;
        private string desc;
        public DogData(string imageFilename, string name, DateTime startTime, DateTime endTime, string from, string description)
        {
            iFilename = imageFilename.Trim();
            n = name.Trim(); ;
            sTime = startTime;
            eTime = endTime;
            f = from.Trim();
            desc = description.Trim();
        }
        public string toHTML()
        {
            string startTimeString = sTime.ToString("yyyy-MM-dd HH:mm");
            string endTimeString = "today";//eTime.ToString("yyyy-MM-dd HH:mm");
            string idString = "com.apaparvolog.event-"+startTimeString+n;
            string descString = "<img src=\'./images/" + iFilename + "\' width=50 height=50></img></br>" + desc;
            string importanceString = "30";
            string returnVal = ",\r\n\t\t\t{\r\n\t\t\t\t\"id\": \"" + idString + 
                               "\",\r\n\t\t\t\t\"startdate\": \"" + startTimeString + 
                               "\",\r\n\t\t\t\t\"enddate\": \"" + endTimeString + 
                               "\",\r\n\t\t\t\t\"title\": \"" + n + 
                               "\",\r\n\t\t\t\t\"description\": \"" + descString + 
                               "\",\r\n\t\t\t\t\"importance\": \"" + importanceString + 
                               "\",\r\n\t\t\t\t\"link\": \"" + f + "\"\r\n\t\t\t}";
            returnVal = returnVal.Replace("  ", "").Replace(" <", "<").Replace("> ", ">");
            return returnVal;
        }
    }
}
