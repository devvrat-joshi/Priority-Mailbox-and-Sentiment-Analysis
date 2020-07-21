using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Diagnostics;
namespace ModularAddin
{
    public partial class ThisAddIn
    {   
        Outlook.Explorer currentExplorer = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Application.Session.Categories["Red category"]==null)
            {
                Application.Session.Categories.Add("Red category", Outlook.OlCategoryColor.olCategoryColorRed);
            }
            if (Application.Session.Categories["Blue category"] == null)
            {
                Application.Session.Categories.Add("Blue category", Outlook.OlCategoryColor.olCategoryColorBlue);
            }
            if (Application.Session.Categories["Green category"] == null)
            {
                Application.Session.Categories.Add("Green category", Outlook.OlCategoryColor.olCategoryColorGreen);
            }
            if (Application.Session.Categories["Yellow category"] == null)
            {
                Application.Session.Categories.Add("Yellow category", Outlook.OlCategoryColor.olCategoryColorYellow);
            }
            if (Application.Session.Categories["Orange category"] == null)
            {
                Application.Session.Categories.Add("Orange category", Outlook.OlCategoryColor.olCategoryColorOrange);
            }
            if (Application.Session.Categories["Purple category"] == null)
            {
                Application.Session.Categories.Add("Purple category", Outlook.OlCategoryColor.olCategoryColorPurple);
            }
            if (Application.Session.Categories["New Sender category"] == null)
            {
                Application.Session.Categories.Add("New Sender category", Outlook.OlCategoryColor.olCategoryColorSteel);
            }
            string path = File.ReadAllText("c:/Program Files/ModularAddinDjango/pathdjango.txt");
            Process.Start(path + "addin.py");
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }
        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder = this.Application.ActiveExplorer().CurrentFolder;
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =(selObject as Outlook.MailItem);
                        DateTime currentDateTime = DateTime.Now;
                        string path = File.ReadAllText("C:/Program Files/ModularAddinDjango/pathmaildata.txt");
                        path = path + mailItem.PropertyAccessor.BinaryToString(mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102")) + ".txt";
                        if (!File.Exists(path))
                        {
                            File.WriteAllText(path, currentDateTime.ToString());
                        }
                        
                    }
                }
            }
            catch (Exception ex)
            {
                string error = "error";
                Console.WriteLine(error);
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
          return new Ribbon();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            MessageBox.Show("Shutting ModularAddin Processes");
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

/*
 * 
 *Public Sub Initialize_handler()
 Set myOlExp = Application.ActiveExplorer
End Sub

Private Sub myOlExp_SelectionChange()
 Dim currTime As Date
 Dim oMail As Outlook.MailItem
 currTime = time()
 Set myOlExp = Application.ActiveExplorer
 Set myOlSel = myOlExp.Selection
 Dim x As Long
 x = 1
 On Error GoTo done
 If myOlSel.Item(x).Class = OlObjectClass.olMail Then
 Set oMail = myOlSel.Item(x)
 MsgTxt = MsgTxt & "Received Time: " & oMail.ReceivedTime & ", Read Time: " & Date & " " & currTime
 

On Error GoTo oops
Dim TextFile As Integer
Dim FilePath As String
Dim FileContent As String
  FilePath = "C:\Users\Devvrat\Desktop\django_modular\maildata\" & oMail.EntryID & ".count"

  TextFile = FreeFile
  Open FilePath For Input As TextFile
  FileContent = Input(LOF(TextFile), TextFile)
  Close TextFile
retval = CDbl(FileContent)
retval = retval + 1
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("C:\Users\Devvrat\Desktop\django_modular\maildata\" & oMail.EntryID & ".count", True)
a.writeline (retval)
a.Close
GoTo done

oops:
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("C:\Users\Devvrat\Desktop\django_modular\maildata\" & oMail.EntryID & ".txt", True)
a.writeline ("received time: " & oMail.ReceivedTime & ", Open Time: " & Date & " " & currTime & " sender: " & oMail.Sender & " senderaddress: " & oMail.SenderEmailAddress & " body: " & oMail.Body)
a.Close
Set a = fs.CreateTextFile("C:\Users\Devvrat\Desktop\django_modular\maildata\" & oMail.EntryID & ".count", True)
a.writeline ("1")
a.Close
 End If
done:
End Sub

 
 */
