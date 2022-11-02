using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security;
//Toegevoegd om LDAP verbinding mogelijk te maken.
using System.DirectoryServices; 
using System.DirectoryServices.AccountManagement;

namespace DS_GUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void progressBar1_Click(object sender, EventArgs e) //foutje 
        {

        }
        
        //Genereer inlognaam
        private void invoerinlognaambutton_Click(object sender, EventArgs e)
        {
            //Kijk of er een nummer in de naam zit
            if (invoernaamtextbox.Text.All(char.IsDigit) && invoernaamtextbox.Text.Length > 0)
            {
                DialogResult result;
                result = MessageBox.Show("Er zit een cijfer in de naam, weet u zeker dat u deze naam wil toevoegen?", "Invoerfout", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (result == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }
            //Check of er wel een naam ingevuld is en of deze te lang is.
            if (invoernaamtextbox.Text.Length == 0)
            {
                MessageBox.Show("Er is geen naam ingevuld, vul deze in om een Inlognaam en wachtwoord te genereren.", "Invoerfout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } else if (invoernaamtextbox.Text.Length <= 5)
            {
                invoerinlognaamtextbox.Text = invoernaamtextbox.Text;
            } else if (invoernaamtextbox.Text.Length > 5)
            {
                string trimmedName = invoernaamtextbox.Text.Replace(" ","");
                string userNameLetters = trimmedName.Substring(0, 5);
                invoerinlognaamtextbox.Text = userNameLetters;
            } 
            //Voeg 5 willekeurige nummers toe aan de gebruikersnaam.
            Random random = new Random();
            for (int i = 0; i < 5; i++) 
            {
                int num = random.Next(1, 10);
                invoerinlognaamtextbox.Text = invoerinlognaamtextbox.Text + num.ToString(); 
            }
            //genereer wachtwoord
            invoerwachtwoordtextbox.Text = ""; //Zorgt ervoor dat het wachtwoordvak leeg is als er opnieuw geklikt wordt.
            string allowedPasswordCharacters = "aAbBcCdDeEfFgGhHiIjJhHkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ01234567890123456789,;:!*$@-_=,;:!*$@-_=";
            int passwordLength = random.Next(8, 15);
            for (int i = 0; i < passwordLength; i++)
            {
                int characterNumber = random.Next(1, allowedPasswordCharacters.Length);
                invoerwachtwoordtextbox.Text = invoerwachtwoordtextbox.Text + allowedPasswordCharacters[characterNumber];
            }
            return;
        }



        //Opslaan user
        private void invoeropslaanbutton_Click(object sender, EventArgs e)
        {
            //Datum checks
            string strDateTime = invoergeboortedatummaskedtextbox.Text;
            string correctedDate = strDateTime.Replace(" ", "0");
            DateTime userDate = DateTime.Now;
            try
            {
                userDate = DateTime.ParseExact(correctedDate, "dd/MM/yyyy", null);

            }
            catch (Exception exception)
            {
                MessageBox.Show(correctedDate + " is geen geldige datum. Foutmelding: " + exception, "Ongeldige datum", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            DateTime tijd = new DateTime(1, 1, 1);
            TimeSpan verschil = DateTime.Now - userDate;
            int jaarVerschil = (tijd + verschil).Year - 1;
            if (jaarVerschil < 4)
            {
                MessageBox.Show("Opgegeven geboortedatum geeft een te jonge leeftijd, datum mag niet in het systeem.", "Datumfout", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            } else if (jaarVerschil > 150)
            {
                MessageBox.Show("Opgegeven geboortedatum geeft een te oude leeftijd, datum mag niet in het systeem.", "Datumfout", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //User toevoegen
            try
            {
                DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://OU=" + invoerstudierichtingcombobox.Text + ",DC=jonard,DC=prive");
                DirectoryEntry childEntry = directoryEntry.Children.Add("CN=" + invoerinlognaamtextbox.Text, "user");
                childEntry.Properties["samAccountName"].Value = invoerinlognaamtextbox.Text; //Accountnaam
                childEntry.Properties["sn"].Value = invoernaamtextbox.Text; //Achternaam
                childEntry.Properties["givenName"].Value = invoervoornaamtextbox.Text; //Voornaam
                childEntry.Properties["l"].Value = invoerwoonplaatstextbox.Text; //Woonplaats
                childEntry.Properties["streetAddress"].Value = invoeradrestextbox.Text; //Adres
                childEntry.Properties["personalTitle"].Value = invoermanvrouwcombobox.Text; //Geslacht
                childEntry.Properties["info"].Value = invoergeboortedatummaskedtextbox.Text; // Geboortedatum
                childEntry.CommitChanges();
                directoryEntry.CommitChanges();
                childEntry.Invoke("SetPassword", new object[] { invoerwachtwoordtextbox });
                childEntry.CommitChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("De gebruiker kon niet toegevoegd worden. Foutmelding: " + ex.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            String[] groupNames = new string[2];
            groupNames[0] = invoermanvrouwcombobox.Text == "Man" ? "GL_Mannen" : "GL_Vrouwen";
            groupNames[1] = jaarVerschil >= 22 ? "GL_StudOuderdan22" : "GL_StudeJongerdan22";
            string userNamePath = invoerinlognaamtextbox.Text; // "CN=" + invoerinlognaamtextbox.Text +",OU=" + invoerstudierichtingcombobox.Text + ",DC=jonard,DC=prive";
            for (int i = 0; i < groupNames.Length; i++)
            {
                try
                {
                    using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "jonard.prive"))
                    {
                        GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupNames[i]);
                        group.Members.Add(pc, IdentityType.SamAccountName, userNamePath);
                        group.Save();
                    }
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    MessageBox.Show("De gebruiker kon niet toegevoegd worden aan groep: " +  groupNames[i] + ". Foutmelding: " + E.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            MessageBox.Show("Gebruiker toegevoegd", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        
        //Wachtwoord laten zien.
        private void invoerwachtwoordlatenzienbutton_Click(object sender, EventArgs e)
        {
            //Vraagt  of het wachtwoord getoond moet worden
            if (invoerwachtwoordtextbox.PasswordChar == '*')
            {
                DialogResult result;
                result = MessageBox.Show("Weet u zeker dat u het wachtwoord wil weergeven?", "Veiligheidsvraag", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
                invoerwachtwoordlatenzienbutton.Text = "Wachtwoord verbergen";
                invoerwachtwoordtextbox.PasswordChar = '\0'; //Geeft een null character aan PasswordChar
                wijzigenwachtwoordtextbox.PasswordChar = '\0';
                return;
            } else
            {
                invoerwachtwoordlatenzienbutton.Text = "Wachtwoord laten zien";
                invoerwachtwoordtextbox.PasswordChar = '*';
                wijzigenwachtwoordtextbox.PasswordChar = '*';
            }
        }

        //Ophalen gegevens
        private void wijzigeninlognaambutton_Click(object sender, EventArgs e)
        {
            if (wijzigeninlognaamtextbox.Text == "" || wijzigeninlognaamtextbox.Text == " ")
            {
                MessageBox.Show("Er is geen naam ingevuld", "Invoerfout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int i = 0;
            foreach (string OU in wijzigenstudierichtingcombobox.Items)
            {
                DirectoryEntry adUserFolder = new DirectoryEntry("LDAP://OU=" + OU + ",DC=jonard,DC=prive");
                DirectorySearcher searcher = new DirectorySearcher(adUserFolder)
                {
                    PageSize = int.MaxValue,
                    Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" + wijzigeninlognaamtextbox.Text + "))"
                };
                var gebruikerInfo = searcher.FindOne();
                if (gebruikerInfo == null)
                {
                    if (i == wijzigenstudierichtingcombobox.Items.Count - 1)
                    {
                        MessageBox.Show("Er is geen gebruiker met die inlognaam", "Melding", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    i++;
                    continue;
                }
                wijzigennaamtextbox.Text =  gebruikerInfo.Properties.Contains("sn") ? gebruikerInfo.Properties["sn"][0].ToString() : "";
                wijzigenvoornaamtextbox.Text = gebruikerInfo.Properties.Contains("givenName") ? gebruikerInfo.Properties["givenName"][0].ToString() : "";
                wijzigenwoonplaatstextbox.Text = gebruikerInfo.Properties.Contains("l") ? gebruikerInfo.Properties["l"][0].ToString() : "";
                wijzigenadrestextbox.Text = gebruikerInfo.Properties.Contains("streetAddress") ? gebruikerInfo.Properties["streetAddress"][0].ToString() : "";
                wijzigenmanvrouwcombobox.Text = gebruikerInfo.Properties.Contains("personalTitle") ? gebruikerInfo.Properties["personalTitle"][0].ToString() : "";
                wijzigengeboortedatummaskedtextbox.Text = gebruikerInfo.Properties.Contains("info") ? gebruikerInfo.Properties["info"][0].ToString() : "";
                wijzigenstudierichtingcombobox.Text = OU;
            }

        }

        private void wijzigenstudentverwijderenbutton_Click(object sender, EventArgs e)
        {
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, wijzigeninlognaamtextbox.Text);
            if (user != null)
            {
                DialogResult result;
                result = MessageBox.Show("Weet u zeker dat u deze gebruiker wil verwijderen? Dit is onomkeerbaar.", "Melding!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    user.Delete();
                    return;
                }
                
            }
        }

        private void wijzigenstudentopslaanbutton_Click(object sender, EventArgs e)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://DC=jonard,DC=prive");
            ldapConnection.Path = "LDAP://DC=jonard,DC=prive";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            string username = wijzigeninlognaamtextbox.Text;
            try
            {
                DirectoryEntry myLdapConnection = ldapConnection;
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                search.Filter = "(cn=" + username + ")";
                search.PropertiesToLoad.Add("title");
                SearchResult result = search.FindOne();
                if (result != null)
                {
                    MessageBox.Show("Resultaat gevonden!" + result.ToString());
                    return;
                    // create new object from search result
                    DirectoryEntry entryToUpdate = result.GetDirectoryEntry();
                    // show existing title
                    Console.WriteLine("Current title: " + entryToUpdate.Properties["title"][0].ToString());
                    Console.Write("\n\nEnter new title : ");
                    // get new title and write to AD
                    String newTitle = Console.ReadLine();
                    entryToUpdate.Properties["title"].Value = newTitle;
                    entryToUpdate.CommitChanges();
                    Console.WriteLine("\n\n…new title saved");
                }
                else
                {
                    MessageBox.Show("Geen resultaat gevonden!" + result.ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught:\n\n" + ex.ToString());
            }
        }
    }
}