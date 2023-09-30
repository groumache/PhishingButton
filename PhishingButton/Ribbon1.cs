using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  suivez ces étapes pour activer l'élément (XML) Ruban :

// 1. Copiez le bloc de code suivant dans la classe ThisAddin, ThisWorkbook ou ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Créez des méthodes de rappel dans la région "Rappels du ruban" de cette classe pour gérer les actions des utilisateurs
//    comme les clics sur un bouton. Remarque : si vous avez exporté ce ruban à partir du Concepteur de ruban,
//    vous devrez déplacer votre code des gestionnaires d'événements vers les méthodes de rappel et modifiez le code pour qu'il fonctionne avec
//    le modèle de programmation d'extensibilité du ruban (RibbonX).

// 3. Assignez les attributs aux balises de contrôle dans le fichier XML du ruban pour identifier les méthodes de rappel appropriées dans votre code.  

// Pour plus d'informations, consultez la documentation XML du ruban dans l'aide de Visual Studio Tools pour Office.


namespace PhishingButton
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }
        public void PhishingButtonClicked(Office.IRibbonControl control)
        {
            var formPopup = new Form1();
            formPopup.ShowDialog();

            bool attachment_opened = formPopup.AttachmentOpened();
            bool credential_send = formPopup.CredentialsProvided();

            // check if the user submitted the form or cancelled it
            if (formPopup.formSubmitted)
            {
                CreateNewMailToSecurityTeam(control, attachment_opened, credential_send);
            }
        }

        private void CreateNewMailToSecurityTeam(IRibbonControl control, bool attachment_opened, bool credentials_sent)
        {
            Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            // check if the user actually selected an email
            if (selection.Count < 1)
            {
                MessageBox.Show("Please select one email.");
                return;
            }
            else if (selection.Count > 1)
            {
                MessageBox.Show("Please select only one email.");
                return;
            }
            object mailItem = selection[1]; // index starts at 1

            // generate a new email to send to the security team
            MailItem tosend = (MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            tosend.Attachments.Add(mailItem);
            tosend.Subject = "[User Alert] Suspicious mail";
            tosend.To = "phishing.soc@example.com";

            tosend.Body = "Suspicious email sent automatically using the phishing button.\n\n";
            tosend.Body += GetCurrentUserInfos();
            tosend.Body += " - Attachment opened: " + (attachment_opened ? "Yes" : "No");
            tosend.Body += " - Credentials sent: " + (credentials_sent ? "Yes" : "No");

            tosend.Display();
        }

        public String GetCurrentUserInfos()
        {
            String wComputername = System.Environment.MachineName + " (" + System.Environment.OSVersion.ToString() + ")";
            String wUsername = System.Environment.UserDomainName + "\\" + System.Environment.UserName;

            Outlook.ExchangeUser currentUser = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry.GetExchangeUser();

            string str = "Possibly useful information:\n--------------";

            str += "\n - Name: " + currentUser.Name;
            str += "\n - STMP address: " + currentUser.PrimarySmtpAddress;
            str += "\n - Title: " + currentUser.JobTitle;
            str += "\n - Department: " + currentUser.Department;
            str += "\n - Location: " + currentUser.OfficeLocation;
            str += "\n - Business phone: " + currentUser.BusinessTelephoneNumber;
            str += "\n - Mobile phone: " + currentUser.MobileTelephoneNumber;
            str += "\n - Preferred language: " + CultureInfo.CurrentCulture.TwoLetterISOLanguageName;
            str += "\n - Windows username: " + wUsername;
            str += "\n - Computername: " + wComputername;
            str += "\n";

            return str;
        }


        #region Membres IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PhishingButton.Ribbon1.xml");
        }

        #endregion

        #region Rappels du ruban
        //Créez des méthodes de rappel ici. Pour plus d'informations sur l'ajout de méthodes de rappel, consultez https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Programmes d'assistance

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
