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
            invoerwachtwoordtextbox.Text = ""; //Zorgt ervoor dat het wachtwoordvak leeg is.
            string allowedPasswordCharacters = "aAbBcCdDeEfFgGhHiIjJhHkKlLmMnNoOpPqQrRsStTuUvVwWxXyYzZ01234567890123456789,;:!*$@-_=,;:!*$@-_=";
            int length = random.Next(8, 15);
            for (int i = 0; i < length; i++)
            {
                int characterNumber = random.Next(1, allowedPasswordCharacters.Length);
                invoerwachtwoordtextbox.Text = invoerwachtwoordtextbox.Text + allowedPasswordCharacters[characterNumber];
            }
            return;
        }

        //Opslaan user
        private void invoeropslaanbutton_Click(object sender, EventArgs e)
        {
            try
            {
                DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://OU=" + invoerstudierichtingcombobox.Text + ",DC=jonard,DC=prive");
                DirectoryEntry childEntry = directoryEntry.Children.Add("CN=" + invoerinlognaamtextbox.Text, "user");
                childEntry.Properties["samAccountName"].Value = invoerinlognaamtextbox.Text; //Accountnaam
                childEntry.Properties["sn"].Value = invoernaamtextbox.Text; //Achternaam
                childEntry.Properties["givenName"].Value = invoervoornaamtextbox.Text; //Voornaam
                childEntry.Properties["l"].Value = invoerwoonplaatstextbox.Text; //Woonplaats
                childEntry.Properties["streetAddress"].Value = invoeradrestextbox.Text; //Adres
                childEntry.Properties["personalTitle"].Value = invoergeboortedatummaskedtextbox.Text; //Geslacht
                childEntry.Properties["info"].Value = invoervoornaamtextbox.Text; // Geboortedatum
                childEntry.CommitChanges();
                directoryEntry.CommitChanges();
                childEntry.Invoke("SetPassword", new object[] { invoerwachtwoordtextbox });
                childEntry.CommitChanges();
                MessageBox.Show("Gebruiker toegevoegd", "Succes", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {
              MessageBox.Show(ex.ToString(), "FOUT", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        
        //Wachtwoord laten zien.
        private void invoerwachtwoordlatenzienbutton_Click(object sender, EventArgs e)
        {
            //Vraagt voor of het wachtwoord getoond moet worden
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
                return;
            } else
            {
                invoerwachtwoordtextbox.PasswordChar = '*';
                invoerwachtwoordlatenzienbutton.Text = "Wachtwoord laten zien";
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
                wijzigenvoornaamtextbox.Text = gebruikerInfo.Properties.Contains("givenName") ? gebruikerInfo.Properties["sn"][0].ToString() : "";
                wijzigenwoonplaatstextbox.Text = gebruikerInfo.Properties.Contains("l") ? gebruikerInfo.Properties["l"][0].ToString() : "";
                wijzigenadrestextbox.Text = gebruikerInfo.Properties.Contains("streetAddress") ? gebruikerInfo.Properties["streetAddress"][0].ToString() : "";
                wijzigenmanvrouwtextbox.Text = gebruikerInfo.Properties.Contains("personalTitle") ? gebruikerInfo.Properties["personalTitle"][0].ToString() : "";
                wijzigengeboortedatummaskedtextbox.Text = gebruikerInfo.Properties.Contains("info") ? gebruikerInfo.Properties["info"][0].ToString() : "";
                wijzigenstudierichtingcombobox.Text = OU;
            }

        }
    }
}