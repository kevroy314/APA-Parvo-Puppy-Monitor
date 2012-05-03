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
        private static int outfileHeaderLength = 68;
        private static int outfileFooterLength = 15;
        static void Main(string[] args)
        {
            ReadCurrentPageStatus();
            System.Threading.Thread monitorDaemon = new System.Threading.Thread(new System.Threading.ThreadStart(MonitorMailbox));
            monitorDaemon.Start();
            Console.ReadLine();
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
            StreamReader reader = new StreamReader("data.js");
            string document = reader.ReadToEnd().Replace(" \\", "").Remove(0, outfileHeaderLength);
            document = document.Remove(document.Length - outfileFooterLength, outfileFooterLength);
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
            FileStream writer = new FileStream("data.js", FileMode.Open);
            for (int i = 0; i < dogs.Count; i++)
            {
                writer.Seek(-outfileFooterLength, SeekOrigin.End);
                byte[] writeBytes = Encoding.ASCII.GetBytes(dogs[i].toHTML());
                writer.Write(writeBytes,0,writeBytes.Length);
            }
            string footer = "\"+\"</tbody>\")};";
            byte[] footerBytes = Encoding.ASCII.GetBytes(footer);
            if(dogs.Count>0)
                writer.Seek(0, SeekOrigin.End);
            else
                writer.Seek(-outfileFooterLength, SeekOrigin.End);
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
            iFilename = imageFilename;
            n = name;
            sTime = startTime;
            eTime = endTime;
            f = from;
            desc = description;
        }
        public string toHTML()
        {
            string returnVal = "\"+\r\n\"<tr>\"+\r\n\t\"<td>\"+\r\n\t\"" + sTime.ToString("yyyy-MM-dd HH:mm") + "</td>\"+\r\n\t\"<td>today</td>\"+\r\n\t\"<td>" + n + "</td>\"+\r\n\t\"<td><img src=\'./images/" + iFilename + "\' width=50 height=50></img></br>" + desc + "</td>\"+\r\n\t\"<td></td>\"+\r\n\t\"<td>&nbsp;</td>\"+\r\n\t\"<td>30</td>\"+\r\n\t\"<td>" + f + "</td>\"+\r\n\t\"<td></td>\"+\r\n\"</tr>";
            returnVal = returnVal.Replace("  ", "").Replace(" <", "<").Replace("> ", ">");
            return returnVal;
        }
    }
}
