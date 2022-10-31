using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DS_GUI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
        
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
                MessageBox.Show("Er is geen naam ingevuld, vul deze in om een Inlognaam te genereren.", "Invoerfout", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            for (int i = 0; i < 5; i++) {
                Console.WriteLine(i.ToString());
                int num = random.Next(1, 10);
                invoerinlognaamtextbox.Text = invoerinlognaamtextbox.Text + num.ToString(); 
            }
            return;
        }
    }
}
