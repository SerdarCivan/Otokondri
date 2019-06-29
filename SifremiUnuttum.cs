using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;

namespace Otokondri
{
    public partial class SifremiUnuttum : Form
    {
        public SifremiUnuttum()
        {
            InitializeComponent();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
            
        }


        private void SifremiUnuttum_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void btn_SifremiUnuttumGerii_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        char karakter;
        private void btn_SifremiUnuttumGonder_Click(object sender, EventArgs e)
        {
            Random rastgele = new Random();
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 8; i++)
            {
                int ascii = rastgele.Next(32, 127);
                karakter = Convert.ToChar(ascii);
                sb.Append(karakter);
            }

            Random deneme = new Random();
            string harfler = "ABCDEFGHIJKLMNOPRSTUVYZabcdefghijklmnoprstuvyz1234567890";
            string uret = "";
            for (int i = 0; i < 6; i++)
            {
                uret += harfler[deneme.Next(harfler.Length)];
            }
            

            MailMessage ePosta = new MailMessage();
            ePosta.From = new MailAddress("otokondriinfo@gmail.com");
            ePosta.To.Add(textBox_Eposta.Text);
            ePosta.Subject = "parola yenile";
            ePosta.Body = "Parola yenileme kodunuz : "+uret.ToString();
            SmtpClient smtp = new SmtpClient();
            smtp.Credentials = new System.Net.NetworkCredential("otokondriinfo@gmail.com", "otokondri123");
            smtp.Port = 587;
            smtp.Host = "smtp.gmail.com";
            smtp.EnableSsl = true;
            object userState = ePosta;
            bool kontrol = true;
            try
            {
                smtp.SendAsync(ePosta, (object)ePosta);
                MessageBox.Show("Mail gönderildi.");
            }
            catch (SmtpException ex)
            {
                kontrol = false;
                System.Windows.Forms.MessageBox.Show(ex.Message, "Mail Gönderme Hatasi");
            }

            

            string islem = "select Id from tbl_Personel where Email = '" + textBox_Eposta.Text + "'";
            string id = SqlConn.Id_Deger(islem);
            

            string insrt = "insert into ParolaYenile(Random,Personel_Id)values('" + uret + "','"+ id + "')";
            SqlConn.Islemler(insrt);
                
            SifreYenileme ynl=new SifreYenileme();

            ynl.id = id;

            ynl.Show();
            this.Hide();


        }
    }
}
