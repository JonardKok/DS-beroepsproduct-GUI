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
            //Check of alles ingevuld is.
            String[] checkItems = new string[9];
            checkItems[0] = invoerinlognaamtextbox.Text; //Accountnaam
            checkItems[1] = invoernaamtextbox.Text; //Achternaam
            checkItems[2] = invoervoornaamtextbox.Text; //Voornaam
            checkItems[3] = invoerwoonplaatstextbox.Text; //Woonplaats
            checkItems[4] = invoeradrestextbox.Text; //Adres
            checkItems[5] = invoermanvrouwcombobox.Text; //Man/vrouw
            checkItems[6] = invoergeboortedatummaskedtextbox.Text; //Geboortedatum
            checkItems[7] = invoerwachtwoordtextbox.Text; //Wachtwoord
            checkItems[8] = invoerstudierichtingcombobox.Text; //Studierichting

            for (int i = 0; i < checkItems.Length; i++)
            {
                if (checkItems[i] == " " || checkItems[i] == "")
                {
                    MessageBox.Show("Niet alle gegevens zijn ingevuld.", "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

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
                childEntry.Properties["gender"].Value = invoermanvrouwcombobox.Text; //Geslacht
                childEntry.Properties["info"].Value = invoergeboortedatummaskedtextbox.Text; // Geboortedatum
                childEntry.CommitChanges();
                directoryEntry.CommitChanges();
                childEntry.Invoke("SetPassword", new object[] { invoerwachtwoordtextbox.Text });
                childEntry.CommitChanges();
            }
            catch (Exception ex)
            {
                MessageBox.Show("De gebruiker kon niet toegevoegd worden. Foutmelding: " + ex.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "jonard.prive");
                UserPrincipal userPrincipal = UserPrincipal.FindByIdentity(principalContext, invoerinlognaamtextbox.Text);
                userPrincipal.Enabled = true;
                userPrincipal.Save();

            }
            catch (Exception ex)
            {
                Console.WriteLine("De gebruiker kon niet toegevoegd worden. Fout:" + ex.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //Voeg gebruiker toe aan een groep
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
                wijzigenmanvrouwcombobox.Text = gebruikerInfo.Properties.Contains("gender") ? gebruikerInfo.Properties["gender"][0].ToString() : "";
                wijzigengeboortedatummaskedtextbox.Text = gebruikerInfo.Properties.Contains("info") ? gebruikerInfo.Properties["info"][0].ToString() : "";
                wijzigenstudierichtingcombobox.Text = OU;
                wijzigenoudestudierichtinglabel.Text = OU;
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
            //Check of alles ingevuld is.
            String[] checkItems = new string[9];
            checkItems[0] = wijzigeninlognaamtextbox.Text; //Accountnaam
            checkItems[1] = wijzigennaamtextbox.Text; //Achternaam
            checkItems[2] = wijzigenvoornaamtextbox.Text; //Voornaam
            checkItems[3] = wijzigenwoonplaatstextbox.Text; //Woonplaats
            checkItems[4] = wijzigenadrestextbox.Text; //Adres
            checkItems[5] = wijzigenmanvrouwcombobox.Text; //Man/vrouw
            checkItems[6] = wijzigengeboortedatummaskedtextbox.Text; //Geboortedatum
            checkItems[7] = wijzigenwachtwoordtextbox.Text; //Wachtwoord
            checkItems[8] = wijzigenstudierichtingcombobox.Text; //Studierichting

            for (int i = 0; i < checkItems.Length; i++)
            {
                if (checkItems[i] == " " || checkItems[i] == "")
                {
                    MessageBox.Show("Niet alle gegevens zijn ingevuld.", "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            //Checkt of de datum geldig is.
            string strDateTime = wijzigengeboortedatummaskedtextbox.Text;
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
            }
            else if (jaarVerschil > 150)
            {
                MessageBox.Show("Opgegeven geboortedatum geeft een te oude leeftijd, datum mag niet in het systeem.", "Datumfout", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Kijkt of de gebruiker in de goede groep zit na het aanpassen van de gegevens.
            String[] gewensteGroep = new string[2];
            String[] ongewensteGroep = new string[2];
            gewensteGroep[0] = wijzigenmanvrouwcombobox.Text == "Man" ? "GL_Mannen" : "GL_Vrouwen";
            gewensteGroep[1] = jaarVerschil >= 22 ? "GL_StudOuderdan22" : "GL_StudeJongerdan22";
            ongewensteGroep[0] = wijzigenmanvrouwcombobox.Text != "Man" ? "GL_Mannen" : "GL_Vrouwen";
            ongewensteGroep[1] = jaarVerschil < 22 ? "GL_StudOuderdan22" : "GL_StudeJongerdan22";
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, "jonard.prive");
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, wijzigeninlognaamtextbox.Text);
            for (int i = 0; i < gewensteGroep.Length; i++)
            {
                GroupPrincipal targetGroup = GroupPrincipal.FindByIdentity(ctx, gewensteGroep[i]);
                if (user != null)
                {
                    if (!user.IsMemberOf(targetGroup)) //Als user geen member van de targetGroup is
                    {
                        //Voeg gebruiker toe aan nieuwe groep
                        try
                        {
                            using (PrincipalContext addGroup = new PrincipalContext(ContextType.Domain, "jonard.prive"))
                            {
                                GroupPrincipal group = GroupPrincipal.FindByIdentity(addGroup, gewensteGroep[i]);
                                group.Members.Add(addGroup, IdentityType.SamAccountName, wijzigeninlognaamtextbox.Text);
                                group.Save();
                            }
                        }
                        catch (System.DirectoryServices.DirectoryServicesCOMException E)
                        {
                            MessageBox.Show("De gebruiker kon niet toegevoegd worden aan groep: " + gewensteGroep[i] + ". Foutmelding: " + E.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                //Verwijder gebruiker uit oude groep
                GroupPrincipal oldGroup = GroupPrincipal.FindByIdentity(ctx, gewensteGroep[i]);
                try
                {
                    if (user != null)
                    {
                        if (!user.IsMemberOf(targetGroup)) //Als user geen member van de targetGroup is
                        {
                            using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, "jonard.prive"))
                            {
                                GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, ongewensteGroep[i]);
                                group.Members.Remove(pc, IdentityType.SamAccountName, wijzigeninlognaamtextbox.Text);
                                group.Save();
                            }
                        }
                    }
                }
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    MessageBox.Show("De gebruiker kon niet verwijderd worden uit groep: " + ongewensteGroep[i] + ". Foutmelding: " + E.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://DC=jonard,DC=prive");
            ldapConnection.Path = "LDAP://DC=jonard,DC=prive";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            string username = wijzigeninlognaamtextbox.Text;
            DirectoryEntry myLdapConnection = ldapConnection;
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);
            search.Filter = "(cn=" + username + ")";
            search.PropertiesToLoad.Add("title");
            SearchResult result = search.FindOne();
            if (result != null)
            {
                try
                {
                    //Gevonden user editen
                    DirectoryEntry entryToUpdate = result.GetDirectoryEntry();
                    entryToUpdate.Properties["sn"].Value = wijzigennaamtextbox.Text;
                    entryToUpdate.Properties["givenName"].Value = wijzigenvoornaamtextbox.Text;
                    entryToUpdate.Properties["l"].Value = wijzigenwoonplaatstextbox.Text;
                    entryToUpdate.Properties["streetAddress"].Value = wijzigenadrestextbox.Text;
                    entryToUpdate.Properties["gender"].Value = wijzigenmanvrouwcombobox.Text;
                    entryToUpdate.Properties["info"].Value = wijzigengeboortedatummaskedtextbox.Text;
                    entryToUpdate.CommitChanges();
                    //Wachtwoord aanpassen
                    if (wijzigenwachtwoordtextbox.Text != " " && wijzigenwachtwoordtextbox.Text != "") {
                        using (var context = new PrincipalContext(ContextType.Domain, "jonard.prive"))
                        {
                            using (var gebruiker = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, wijzigeninlognaamtextbox.Text))
                            {
                                gebruiker.SetPassword(wijzigenwachtwoordtextbox.Text);
                                gebruiker.Save();
                            }
                        }
                        //User naar andere OU verplaatsen
                        if (wijzigenstudierichtingcombobox.Text != wijzigenoudestudierichtinglabel.Text)
                        {
                            DirectoryEntry oudeLocatie = new DirectoryEntry("LDAP://CN=" + wijzigeninlognaamtextbox.Text + ", OU=" + wijzigenoudestudierichtinglabel.Text + ", DC=jonard,DC=prive");
                            DirectoryEntry nieuweLocatie = new DirectoryEntry("LDAP://OU=" + wijzigenstudierichtingcombobox.Text + ", DC=jonard,DC=prive");
                            oudeLocatie.MoveTo(nieuweLocatie);
                            nieuweLocatie.Close();
                            oudeLocatie.Close();
                        }
                    }
                } 
                catch (System.DirectoryServices.DirectoryServicesCOMException E) {
                    MessageBox.Show("Het wachtwoord kon niet aangepast worden. Foutmelding: " + E.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Er is iets fout gegaan bij het wijzigen van de user, probeer opnieuw te zoeken." + result.ToString());
            }
            MessageBox.Show("Gebruiker gewijzigd", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
    }
}