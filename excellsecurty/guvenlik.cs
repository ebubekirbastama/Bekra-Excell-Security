using System;
using System.Collections;
using System.IO;
using System.Windows.Forms;
using Ionic.Zip;

namespace excellsecurty
{
    public partial class guvenlik : Form
    {
        StreamReader sr;
        string cikartilacakyol= Application.StartupPath + @"\analiz";
        private string bekra = "Bekra Hayr Nester";
        public guvenlik()
        {
            InitializeComponent();
        }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK==openFileDialog1.ShowDialog())
            {
                textBoxX1.Text = openFileDialog1.FileName.ToString();
                MessageBox.Show("Dosya Seçildi",bekra);
            }
        }
        private void buttonX2_Click(object sender, EventArgs e)
        {
            try
            {
                zipcikar();
                kontrol();
                Directory.Delete(cikartilacakyol, true);
            }
            catch (Exception)
            {
                Directory.Delete(cikartilacakyol, true);
            }
        }
        void zipcikar()
        {
            ZipFile zipFile = new ZipFile(textBoxX1.Text);
            zipFile.ExtractAll(cikartilacakyol);
        }
        void kontrol()
        {
            //Directory.GetFileSystemEntries(cikartilacakyol)[0].ToString().Substring(cikartilacakyol.Length);
            var dz = Directory.GetDirectories(cikartilacakyol+@"\xl");
            for (int i = 0; i < dz.Length; i++)
            {
                string vrtlstring= dz[i].ToString().Substring(cikartilacakyol.Length).Replace("\\xl\\","").Trim();
                if (vrtlstring== "externalLinks")
                {
                    kontrolxml();
                    break;
                }
                else
                {
                    if (dz.Length-1==i)
                    {
                        richTextBox1.Text = "Excel'de Virüs Yok :)";
                    }
                }
            }
        }
        void kontrolxml()
        {                 //C:\Users\Ebubekir\source\repos\excellsecurty\excellsecurty\bin\Debug\analiz\xl\externalLinks
            string vrtlyol = cikartilacakyol + @"\xl\externalLinks\externalLink1.xml";
            ArrayList vrtlary = new ArrayList();
            try
            {
                using (sr = File.OpenText(vrtlyol))
                {
                    string telefon = sr.ReadLine();
                    while (telefon != null)
                    {
                        //sycc++;
                        vrtlary.Add(telefon);
                        telefon = sr.ReadLine();
                    }
                }
                richTextBox1.Text = "                                        !!! Warning Virus !!!";
                string s = "=\"";
                string s1 = "\"><ddeItems><ddeItem name=\"_xlbgnm.A1\" advise=\"1\"/><ddeItem name=\"StdDocumentName" +
    "\" ole=\"1\" advise=\"1\"/></ddeItems></ddeLink></externalLink>";
                richTextBox1.Text += "\r"; richTextBox1.Text += "\r";
                richTextBox1.Text += vrtlary[1].ToString().Substring(166, 16).Replace(s," = ");
                richTextBox1.Text += "\r";
                richTextBox1.Text += vrtlary[1].ToString().Substring(184).Replace(s1," ").Replace(s," = ");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}



