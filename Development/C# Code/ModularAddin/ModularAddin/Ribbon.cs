using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;

namespace ModularAddin
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public Ribbon()
        {
            
        }
        public void button8_Click(Office.IRibbonControl control)
        {
            //MessageBox.Show("ASDFa");
            Outlook.Application otApp = new Outlook.Application();
            Outlook.MAPIFolder Fol;
            Outlook.NameSpace Ons = otApp.GetNamespace("MAPI");
            Fol = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            string path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            string text = File.ReadAllText(path + "red.txt");
            string[] names = text.Split('\n');
            Outlook.Items item__ = Fol.Items;
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    Outlook.MailItem mi = (Outlook.MailItem)item;
                    
                    mi.Categories = null;
                    mi.Save();
                            //MessageBox.Show(sender+mi.SenderEmailAddress);
                }
            }
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items;
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            mi.Categories = null;
                            mi.Save();
                                    //MessageBox.Show(sender+mi.SenderEmailAddress);
                        }
                    }
                }

            }
            
        }
        public void button7_Click(Office.IRibbonControl control)
        {
            //MessageBox.Show("ASDFa");
            Outlook.Application otApp = new Outlook.Application();
            Outlook.MAPIFolder Fol;
            Outlook.NameSpace Ons = otApp.GetNamespace("MAPI");
            Fol = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            string path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            string text = File.ReadAllText(path + "red.txt");
            string[] names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if(subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        Outlook.MailItem newitem = item as Outlook.MailItem;
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            foreach (string sender in names)
                            {
                                if (sender.Length == 0)
                                    continue;
                                string newsend = sender.Substring(0, sender.Length - 1);
                                int x = sender.Length;
                                int y = mi.SenderEmailAddress.Length;
                                //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                int val = string.Compare(newsend, mi.SenderEmailAddress);
                                if (val == 0)
                                {
                                    float check = runCommand(mi.Body);
                                    if (check < 0.5)
                                    {
                                        mi.Categories = "Red category";
                                    }
                                    else if (check < 0.75)
                                    {
                                        mi.Categories = "Orange category";
                                    }
                                    else
                                    {
                                        mi.Categories = "Yellow category";
                                    }
                                    mi.Save();
                                    //MessageBox.Show(sender+mi.SenderEmailAddress);
                                }
                            }
                        }
                    }
                }
                
            }
            Outlook.Items item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {

                        Outlook.MailItem mi = (Outlook.MailItem)item;

                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body.Replace("\n", "").Replace("\r", ""));
                                if (check < 0.5)
                                {
                                    mi.Categories = "Red category";
                                }
                                else if (check < 0.75)
                                {
                                    mi.Categories = "Orange category";
                                }
                                else
                                {
                                    mi.Categories = "Yellow category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }

                    }
                }
            }

            path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            text = File.ReadAllText(path + "orange.txt");
            names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi.UnRead)
                            {
                                foreach (string sender in names)
                                {
                                    if (sender.Length == 0)
                                        continue;
                                    string newsend = sender.Substring(0, sender.Length - 1);
                                    int x = sender.Length;
                                    int y = mi.SenderEmailAddress.Length;
                                    //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                    int val = string.Compare(newsend, mi.SenderEmailAddress);
                                    if (val == 0)
                                    {
                                        float check = runCommand(mi.Body.Replace("\n", "").Replace("\r", ""));
                                        if (check < 0.5)
                                        {
                                            mi.Categories = "Orange category";
                                        }
                                        else if (check < 0.75)
                                        {
                                            mi.Categories = "Yellow category";
                                        }
                                        else
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        mi.Save();
                                        //MessageBox.Show(sender+mi.SenderEmailAddress);
                                    }
                                }
                            }    
                        }
                    }
                }
            }
            item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {

                        Outlook.MailItem mi = (Outlook.MailItem)item;

                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body.Replace("\n", "").Replace("\r", ""));
                                if (check < 0.5)
                                {
                                    mi.Categories = "Orange category";
                                }
                                else if (check < 0.75)
                                {
                                    mi.Categories = "Yellow category";
                                }
                                else
                                {
                                    mi.Categories = "Green category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }

                    }
                }
            }

            path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            text = File.ReadAllText(path + "yellow.txt");
            names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi.UnRead)
                            {
                                foreach (string sender in names)
                                {
                                    if (sender.Length == 0)
                                        continue;
                                    string newsend = sender.Substring(0, sender.Length - 1);
                                    int x = sender.Length;
                                    int y = mi.SenderEmailAddress.Length;
                                    //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                    int val = string.Compare(newsend, mi.SenderEmailAddress);
                                    if (val == 0)
                                    {
                                        float check = runCommand(mi.Body);
                                        if (check < 0.5)
                                        {
                                            mi.Categories = "Yellow category";
                                        }
                                        else if (check < 0.75)
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        else
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        mi.Save();
                                        //MessageBox.Show(sender+mi.SenderEmailAddress);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {

                        Outlook.MailItem mi = (Outlook.MailItem)item;

                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body);
                                if (check < 0.5)
                                {
                                    mi.Categories = "Yellow category";
                                }
                                else if (check < 0.75)
                                {
                                    mi.Categories = "Green category";
                                }
                                else
                                {
                                    mi.Categories = "Green category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }

                    }
                }
            }
            path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            text = File.ReadAllText(path + "green.txt");
            names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem newitem = item as Outlook.MailItem;
                            if (newitem.UnRead)
                            {
                                Outlook.MailItem mi = (Outlook.MailItem)item;

                                foreach (string sender in names)
                                {
                                    if (sender.Length == 0)
                                        continue;
                                    string newsend = sender.Substring(0, sender.Length - 1);
                                    int x = sender.Length;
                                    int y = mi.SenderEmailAddress.Length;
                                    //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                    int val = string.Compare(newsend, mi.SenderEmailAddress);
                                    if (val == 0)
                                    {
                                        float check = runCommand(mi.Body);
                                        if (check < 0.2)
                                        {
                                            mi.Categories = "Yellow category";
                                        }
                                        else
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        mi.Save();
                                        //MessageBox.Show(sender+mi.SenderEmailAddress);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {

                        Outlook.MailItem mi = (Outlook.MailItem)item;

                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body);
                                if (check < 0.2)
                                {
                                    mi.Categories = "Yellow category";
                                }
                                else
                                {
                                    mi.Categories = "Green category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }

                    }
                }
            }
            path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            text = File.ReadAllText(path + "blue.txt");
            names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi.UnRead)
                            {
                                foreach (string sender in names)
                                {
                                    if (sender.Length == 0)
                                        continue;
                                    string newsend = sender.Substring(0, sender.Length - 1);
                                    int x = sender.Length;
                                    int y = mi.SenderEmailAddress.Length;
                                    //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                    int val = string.Compare(newsend, mi.SenderEmailAddress);
                                    if (val == 0)
                                    {
                                        float check = runCommand(mi.Body);
                                        if (check < 0.2)
                                        {
                                            mi.Categories = "Yellow category";
                                        }
                                        else if (check < 0.3)
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        else
                                        {
                                            mi.Categories = "Blue category";
                                        }
                                        mi.Save();
                                        //MessageBox.Show(sender+mi.SenderEmailAddress);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {

                        Outlook.MailItem mi = (Outlook.MailItem)item;

                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body);
                                if (check < 0.2)
                                {
                                    mi.Categories = "Yellow category";
                                }
                                else if (check < 0.3)
                                {
                                    mi.Categories = "Green category";
                                }
                                else
                                {
                                    mi.Categories = "Blue category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }

                    }
                }
            }
            path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
            text = File.ReadAllText(path + "purple.txt");
            names = text.Split('\n');
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread]=true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi.UnRead)
                            {
                                foreach (string sender in names)
                                {
                                    if (sender.Length == 0)
                                        continue;
                                    string newsend = sender.Substring(0, sender.Length - 1);
                                    int x = sender.Length;
                                    int y = mi.SenderEmailAddress.Length;
                                    //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                                    int val = string.Compare(newsend, mi.SenderEmailAddress);
                                    if (val == 0)
                                    {
                                        float check = runCommand(mi.Body);
                                        if (check < 0.2)
                                        {
                                            mi.Categories = "Yellow category";
                                        }
                                        else if (check < 0.3)
                                        {
                                            mi.Categories = "Green category";
                                        }
                                        else if (check < 0.5)
                                        {
                                            mi.Categories = "Blue category";
                                        }
                                        else
                                        {
                                            mi.Categories = "Purple category";
                                        }
                                        mi.Save();
                                        //MessageBox.Show(sender+mi.SenderEmailAddress);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            item__ = Fol.Items.Restrict("[Unread]=true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    if (newitem.UnRead)
                    {
                        Outlook.MailItem mi = (Outlook.MailItem)item;
                        foreach (string sender in names)
                        {
                            if (sender.Length == 0)
                                continue;
                            string newsend = sender.Substring(0, sender.Length - 1);
                            int x = sender.Length;
                            int y = mi.SenderEmailAddress.Length;
                            //MessageBox.Show(x.ToString()+y.ToString()+newsend+mi.SenderEmailAddress);
                            int val = string.Compare(newsend, mi.SenderEmailAddress);
                            if (val == 0)
                            {
                                float check = runCommand(mi.Body);
                                if (check < 0.2)
                                {
                                    mi.Categories = "Yellow category";
                                }
                                else if (check < 0.3)
                                {
                                    mi.Categories = "Green category";
                                }
                                else if (check < 0.5)
                                {
                                    mi.Categories = "Blue category";
                                }
                                else
                                {
                                    mi.Categories = "Purple category";
                                }
                                mi.Save();
                                //MessageBox.Show(sender+mi.SenderEmailAddress);
                            }
                        }
                    }
                }
            }

            item__ = Fol.Items.Restrict("[Unread] = true");
            foreach (Object item in item__)
            {
                Outlook.MailItem newitem = item as Outlook.MailItem;
                if (item is Outlook.MailItem)
                {
                    Outlook.MailItem mi = (Outlook.MailItem)item;
                    if (mi.Categories==null)
                    {
                        float val = runCommand(mi.Body);
                        if (val < 0.2)
                        {
                            mi.Categories = "New Sender category,Yellow category";
                        }
                        else if (val < 0.5)
                        {
                            mi.Categories = "New Sender category,Green category";
                        }
                        else if (val < 0.75)
                        {
                            mi.Categories = "New Sender category,Blue category";
                        }
                        else
                        {
                            mi.Categories = "New Sender category,Purple category";
                        }
                    }
                    mi.Save();
                    //MessageBox.Show(sender+mi.SenderEmailAddress);
                }
            }
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                if (subfolder is Outlook.Folder)
                {
                    Outlook.Items item_ = subfolder.Items.Restrict("[Unread] = true");
                    foreach (Object item in item_)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi.Categories == null)
                            {
                                float val = runCommand(mi.Body);
                                if (val < 0.2)
                                {
                                    mi.Categories = "New Sender category,Yellow category";
                                }
                                else if (val < 0.5)
                                {
                                    mi.Categories = "New Sender category,Green category";
                                }
                                else if (val < 0.75)
                                {
                                    mi.Categories = "New Sender category,Blue category";
                                }
                                else
                                {
                                    mi.Categories = "New Sender category,Purple category";
                                }
                            }
                            mi.Save();
                            //MessageBox.Show(sender+mi.SenderEmailAddress);
                        }
                    }
                }

            }
        }
        public void button10_Click(Office.IRibbonControl control)
        {
            string url = "http://127.0.0.1:8000/imp.html";
            WebRequest request = HttpWebRequest.Create(url);
            WebResponse response = request.GetResponse();
            url = "http://127.0.0.1:8000/rank.html";
            request = HttpWebRequest.Create(url);
            response = request.GetResponse();
            MessageBox.Show("Your Model training is complete");
        }
        static float runCommand(string s)
        {
            string path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathdjango.txt");
            Process process = new Process();
            process.StartInfo.FileName = path + "helperDjango/dist/execute/execute.exe";
            process.StartInfo.Arguments = "/c "+s;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.ErrorDataReceived += new DataReceivedEventHandler(ErrorOutputHandler);
            process.Start();
            process.BeginErrorReadLine();
            string output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            return float.Parse(output);
        }

        static void ErrorOutputHandler(object sendingProcess, DataReceivedEventArgs outLine)
        {
            //* Do your stuff with the output (write to console/log/StringBuilder)
            Console.WriteLine(outLine.Data);
        }
        public void button9_Click(Office.IRibbonControl control)
        {
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            MessageBox.Show(File.ReadAllText(path + "lastupdatetime.txt"));
        }
        public void editBox1_TextChanged(Office.IRibbonControl control)
        {
            MessageBox.Show("OK");
        }

        public void button6_Click(Office.IRibbonControl control)
        {
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            Process.Start(path + "/open_db.py");
        }
        public void button3_Click(Office.IRibbonControl control)
        {
            //string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            //Process.Start(path + "open_settings.py");
            Process.Start("chrome.exe", "http://127.0.0.1:8000/index.html");
        }
        public void button5_Click(Office.IRibbonControl control)
        {
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            MessageBox.Show("This is an office solution made by IIT Gandhinagar students in an intern at Capgeminie. ");
        }

        public void button4_Click_1(Office.IRibbonControl control)
        {
            MessageBox.Show("Email to dsjoshi1990@gmail.com or jani.dhyey@iitgn.ac.in");
        }
        public void button2_Click(Office.IRibbonControl control)
        {
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            Process.Start(path+"addin.py");
        }

        public void button1_Click(Office.IRibbonControl control)
        {
            Outlook.Application otApp = new Outlook.Application();
            Outlook.MAPIFolder Fol;
            Outlook.NameSpace Ons = otApp.GetNamespace("MAPI");
            Fol = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            int check = 0;
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                Outlook.Folder sub = subfolder;
                int val = string.Compare(sub.Name, "Negative");
                if (val == 0)
                {
                    check = 1;
                }
            }
            if (check == 0)
            {
                Fol.Folders.Add("Negative", Outlook.OlDefaultFolders.olFolderInbox);
            }
            check = 0;
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                Outlook.Folder sub = subfolder;
                int val = string.Compare(sub.Name, "Positive");
                if (val == 0)
                {
                    check = 1;
                }
            }
            if (check == 0)
            {
                Fol.Folders.Add("Positive", Outlook.OlDefaultFolders.olFolderInbox);
            }
            check = 0;
            foreach (Outlook.Folder subfolder in Fol.Folders)
            {
                Outlook.Folder sub = subfolder;
                int val = string.Compare(sub.Name, "Neutral");
                if (val == 0)
                {
                    check = 1;
                }
            }
            if (check == 0)
            {
                Fol.Folders.Add("Neutral", Outlook.OlDefaultFolders.olFolderInbox);
            }
            Outlook.Items item_ = Fol.Items;
            string pathtostop = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathdjango.txt");
            string [] ll = File.ReadAllText(pathtostop+"/comp.txt").Split(',');
            int mystop1 = int.Parse(ll[6]);
            int stop = 0;
            for(int i = item_.Count;i>=1;i--)
            {
                try
                {
                    Outlook.MailItem item = item_[i];
                    if (item is Outlook.MailItem && stop != mystop1)
                    {
                        stop += 1;
                        Outlook.MailItem thismail = (Outlook.MailItem)item;
                        Outlook.MailItem mi = (Outlook.MailItem)item;
                        //MessageBox.Show(mi.PropertyAccessor.BinaryToString((mi.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102")))+stop.ToString());
                        string path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathtest.txt");
                        String Pol = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathpol.txt");
                        File.WriteAllText(path, mi.Body);
                        System.Threading.Thread.Sleep(3000);
                        String value = File.ReadAllText(Pol);
                        int comp = int.Parse(value);
                        String showing = mi.SenderEmailAddress + " " + value;
                        Outlook.MAPIFolder Folderx;
                        if (value == "1")
                        {
                            Folderx = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Positive"];
                        }
                        else if (value == "2")
                        {
                            Folderx = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Neutral"];
                        }
                        else
                        {
                            Folderx = Ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Folders["Negative"];
                        }
                        thismail.Move(Folderx);
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                
            }
            //MessageBox.Show(Ons.)
        }
        public Bitmap bt2_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.images);
        }
        public Bitmap bt1_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources._611MlcxCI2L);
        }
        public Bitmap bt3_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.download__1_);
        }
        public Bitmap bt4_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.download);
        }
        public Bitmap bt5_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.icons8_about_100);
        }
        public Bitmap bt6_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources._100765_200);
        }
        public Bitmap bt7_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.todo_list_1942026_1642356);
        }
        public Bitmap bt8_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.todo_16_1092716);
        }
        public Bitmap bt9_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.download__2_);
        }
        public Bitmap bt10_img(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.deep_learning_1524275_1290822);
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ModularAddin.Ribbon.xml");
        }

        #endregion

        
        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion
        public void Ribbon1_Load()
        {
            
        }

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
        #endregion
    }
}