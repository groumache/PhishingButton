using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PhishingButton
{
    public partial class Form1 : Form
    {
        public bool formSubmitted = false;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        { }

        // check if both questions have been answered before enabling the "Submit" button
        private void EnableButton1()
        {
            if (this.comboBox1.Text != "" && this.comboBox2.Text != "")
            {
                this.button1.Enabled = true;
            }
        }

        // when the user answers a question, the program will check if it should enable the "Submit" button
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.EnableButton1();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.EnableButton1();
        }

        // close the form and change the value of the public member "formSubmitted" which show the form wasn't cancelled
        private void button1_Click(object sender, EventArgs e)
        {
            this.formSubmitted = true;
            this.Close();
        }

        // simple functions to get the answers of the user
        public bool AttachmentOpened()
        {
            bool result = this.comboBox1.Text == "Yes";
            return result;
        }
        public bool CredentialsProvided()
        {
            bool result = this.comboBox2.Text == "Yes";
            return result;
        }
    }
}
