using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.SqlClient;

namespace Otokondri
{
    public partial class PersonelGuncelle : Form
    {
        public PersonelGuncelle()
        {
            InitializeComponent();
        }
        
       

        private void btn_PersonelGuncelle_Geri_Click(object sender, EventArgs e)
        {
            PersonelEkleSil prsnl = new PersonelEkleSil();
            prsnl.Show();
            this.Hide();
        }


        public string ad;
        public  string soyad;
        public  string tc;
        public  string parola;
        public string cinsiyet;
        public string yetki;
        public string id;
        public string email;
        private void PersonelGuncelle_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            textBox_Email.Text = email;
            textBox_PersonelAd.Text = ad;
            textBox_PersonelSoyad.Text = soyad;
            textBox_TCKN.Text = tc;
            textBox_PersonelParola.Text = parola;

            

            if (cinsiyet == "E")
            {
                checkBox_Erkek.Checked = true;
            }
            else if (cinsiyet == "K")
            {
                checkBox_Kadin.Checked = true;
            }


            if (yetki == "Y")
            {
                checkBox_Yonetici.Checked=true;
            }
            else if (yetki == "AD")
            {
                checkBox_AlisDep.Checked = true;
            }
            else if (yetki == "SD")
            {
                checkBox_SatisDep.Checked = true;
            }
        }



        private void btn_Guncelle_Click(object sender, EventArgs e)
        {
            try
            {
                string cinsiyett = "";
                string yetkii = "";
                if (checkBox_Erkek.Checked == (true))
                {
                    cinsiyett = "E";
                }
                else if (checkBox_Kadin.Checked == (true))
                {
                    cinsiyett = "K";
                }

                if (checkBox_SatisDep.Checked == (true))
                {
                    yetkii = "SD";
                }
                else if (checkBox_AlisDep.Checked == (true))
                {
                    yetkii = "AD";
                }
                else if (checkBox_Yonetici.Checked == (true))
                {
                    yetkii = "Y";
                }

                string komut = " UPDATE tbl_Personel SET Ad = '" + textBox_PersonelAd.Text + "', Soyad ='" + textBox_PersonelSoyad.Text + "', TCKN = '" + textBox_TCKN.Text + "', Parola = '" + textBox_PersonelParola.Text + "',Yetki = '" + yetkii + "', Cinsiyet ='" + cinsiyett + "',Email='"+ email + "'  WHERE Id = '" + id + "'";
                SqlConn.Islemler(komut);
                MessageBox.Show("Personel bilgileri güncellendi.");
            }
            catch (Exception)
            {
                MessageBox.Show("Hata");
            }
            
        }

        private void btn_Sil_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult cikis = new DialogResult();
                cikis = MessageBox.Show("Silmek istediğinizden emin misiniz ?", "Uyarı",
                   MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (cikis == DialogResult.Yes)
                {
                    string a = textBox_TCKN.Text;
                    string komut = "delete from tbl_Personel WHERE TCKN='" + textBox_TCKN.Text + "'";
                    SqlConn.Islemler(komut);
                }
                else if (cikis == DialogResult.No)
                {
                    MessageBox.Show("iptal edildi.");
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hata!");
            }
        }

        private void checkBox_Erkek_CheckedChanged_1(object sender, EventArgs e)
        {
            checkBox_Kadin.Checked = false;
        }

        private void checkBox_Kadin_CheckedChanged(object sender, EventArgs e)
        {
            checkBox_Erkek.Checked = false;
        }

        private void checkBox_Yonetici_CheckedChanged(object sender, EventArgs e)
        {
            checkBox_SatisDep.Checked = false;
            checkBox_AlisDep.Checked = false;
        }

        private void checkBox_AlisDep_CheckedChanged(object sender, EventArgs e)
        {
            checkBox_Yonetici.Checked = false;
            checkBox_SatisDep.Checked = false;
        }

        private void checkBox_SatisDep_CheckedChanged(object sender, EventArgs e)
        {
            checkBox_AlisDep.Checked = false;
            checkBox_Yonetici.Checked = false;
        }

    }
}
