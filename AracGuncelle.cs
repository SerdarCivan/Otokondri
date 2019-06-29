using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otokondri
{
    public partial class AracGuncelle : Form
    {
        public AracGuncelle()
        {
            InitializeComponent();
        }
        public string marka;
        public string model;
        public string detay;
        public string alisfiyati;
        public string satisfiyati;
        public string id;
        public string alistarihi;
        public string satistarihi;
        public string plaka;
        public string satis_durum;

        public string personelid;
        

        public string uzanti1;
        public string uzanti2;
        public string uzanti3;
        public string uzanti4;

        public DataTable resimtable;
        private void AracGuncelle_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            if (personelid == "AD")
            {
                btn_Sil.Visible = false;

            }
            if (personelid == "SD")
            {
                btn_Sil.Visible = false;
            }

            try
            {
                if (resimtable.Rows.Count>0)
                {
                  uzanti1 = resimtable.Rows[0]["Fotograf_Uzanti"].ToString();
                  uzanti2 = resimtable.Rows[1]["Fotograf_Uzanti"].ToString();
                  uzanti3 = resimtable.Rows[2]["Fotograf_Uzanti"].ToString();
                  uzanti4 = resimtable.Rows[3]["Fotograf_Uzanti"].ToString();
                }
                textBox_FotoUzanti1.Text = uzanti1;
                textBox_FotoUzanti2.Text = uzanti2;
                textBox_FotoUzanti3.Text = uzanti3;
                textBox_FotoUzanti4.Text = uzanti4;
                
                textBox_Marka.Text = marka;
                textBox_Model.Text = model;
                textBox_Detay.Text = detay;
                textBox_AlisFiyati.Text = alisfiyati;
                textBox_SatisFiyati.Text = satisfiyati;
                textBox_Plaka.Text = plaka;
                textBox_id.Text = id;
                textBox_SatisDurum.Text = satis_durum;
                dateTime_AlisTarihi.Value = Convert.ToDateTime(alistarihi);
                dateTime_SatisTarihi.Value = Convert.ToDateTime(satistarihi);
                
            }
            catch (Exception)
            {
                
                if (resimtable.Rows.Count > 0)
                {
                  uzanti1 = resimtable.Rows[0]["Fotograf_Uzanti"].ToString();
                  uzanti2 = resimtable.Rows[1]["Fotograf_Uzanti"].ToString();
                  uzanti3 = resimtable.Rows[2]["Fotograf_Uzanti"].ToString();
                  uzanti4 = resimtable.Rows[3]["Fotograf_Uzanti"].ToString();
                }
                textBox_FotoUzanti1.Text = uzanti1;
                textBox_FotoUzanti2.Text = uzanti2;
                textBox_FotoUzanti3.Text = uzanti3;
                textBox_FotoUzanti4.Text = uzanti4;

                textBox_Marka.Text = marka;
                textBox_Model.Text = model;
                textBox_Detay.Text = detay;
                textBox_AlisFiyati.Text = alisfiyati;
                textBox_SatisFiyati.Text = satisfiyati;
                textBox_Plaka.Text = plaka;
                textBox_id.Text = id;
                textBox_SatisDurum.Text = satis_durum;
                dateTime_AlisTarihi.Value = Convert.ToDateTime(alistarihi);
            }

        }


        public string dosyayolu1;
        public string dosyayolu2;
        public string dosyayolu3;
        public string dosyayolu4;

        private void btn_FotoSec1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu1 = dosya.FileName;
            textBox_FotoUzanti1.Text = dosyayolu1;
            pictureBox_Foto.ImageLocation = dosyayolu1;
        }

        private void button_FotoSec2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu2 = dosya.FileName;
            textBox_FotoUzanti2.Text = dosyayolu2;
            pictureBox_Foto.ImageLocation = dosyayolu2;
        }

        private void button_FotoSec3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu3 = dosya.FileName;
            textBox_FotoUzanti3.Text = dosyayolu3;
            pictureBox_Foto.ImageLocation = dosyayolu3;
        }

        private void button_FotoSec4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu4 = dosya.FileName;
            textBox_FotoUzanti4.Text = dosyayolu4;
            pictureBox_Foto.ImageLocation = dosyayolu4;
        }

        string dosyaismi;
        string yeniad;
        private void btn_Guncelle_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox_FotoUzanti1.Text != uzanti1)
                {
                    //sildik
                    string hedef = Application.StartupPath + @"\photos\";
                    if (System.IO.File.Exists(@hedef + uzanti1))
                        {
                            System.IO.File.Delete(@hedef + uzanti1);
                        }                    
                    //ekledik
                    dosyaismi = Path.GetFileName(dosyayolu1);
                    string kaynak = dosyayolu1;
                    yeniad = textBox_Plaka.Text.ToString() + "_1.jpg";
                    string deneme = textBox_Plaka.Text;
                   

                    string fotografinsert = "update tbl_Fotograf set Fotograf_Uzanti='"+ yeniad + "',Arac_Id='"+textBox_id.Text+"',Sira='"+1+"',ProfilFotograf='"+1+ "' where Sira='"+1+ "' and Arac_Id='" + textBox_id.Text + "'";
                    SqlConn.Islemler(fotografinsert);
                    File.Copy(kaynak, hedef + yeniad);
                }
                //***************************************************************************************
                if (textBox_FotoUzanti2.Text != uzanti2)
                {
                    //sildik
                    string hedef = Application.StartupPath + @"\photos\";
                    if (System.IO.File.Exists(@hedef + uzanti2))
                    {
                        System.IO.File.Delete(@hedef + uzanti2);
                    }
                    //ekledik
                    dosyaismi = Path.GetFileName(dosyayolu2);
                    string kaynak = dosyayolu2;
                    yeniad = textBox_Plaka.Text + "_2.jpg";

                    string fotografinsert = "update tbl_Fotograf set Fotograf_Uzanti='" + yeniad + "',Arac_Id='" + textBox_id.Text + "',Sira='" + 2 + "' where Sira='" + 2 + "' and Arac_Id='" + textBox_id.Text + "'";
                    SqlConn.Islemler(fotografinsert);
                    File.Copy(kaynak, hedef + yeniad);
                }
                //***************************************************************************************
                if (textBox_FotoUzanti3.Text != uzanti3)
                {
                    //sildik
                    string hedef = Application.StartupPath + @"\photos\";
                    if (System.IO.File.Exists(@hedef + uzanti3))
                    {
                        System.IO.File.Delete(@hedef + uzanti3);
                    }


                    //ekledik
                    dosyaismi = Path.GetFileName(dosyayolu3);
                    string kaynak = dosyayolu3;
                    yeniad = textBox_Plaka.Text + "_3.jpg";

                    string fotografinsert = "update tbl_Fotograf set Fotograf_Uzanti='" + yeniad + "',Arac_Id='" + textBox_id.Text + "',Sira='" + 3 + "' where Sira='" + 3 + "' and Arac_Id='" + textBox_id.Text + "'";
                    SqlConn.Islemler(fotografinsert);
                    File.Copy(kaynak, hedef + yeniad);
                }
                //***************************************************************************************
                if (textBox_FotoUzanti4.Text != uzanti4)
                {

                    //sildik
                    string hedef = Application.StartupPath + @"\photos\";
                    if (System.IO.File.Exists(@hedef + uzanti4))
                    {
                        System.IO.File.Delete(@hedef + uzanti4);
                    }
                    //ekledik
                    dosyaismi = Path.GetFileName(dosyayolu4);
                    string kaynak = dosyayolu4;
                    yeniad = textBox_Plaka.Text + "_4.jpg";

                    string fotografinsert = "update tbl_Fotograf set Fotograf_Uzanti='" + yeniad + "',Arac_Id='" + textBox_id.Text + "',Sira='" + 4 + "'  where Sira='" + 4 + "' and Arac_Id='"+textBox_id.Text+"'";
                    SqlConn.Islemler(fotografinsert);
                    File.Copy(kaynak, hedef + yeniad);
                }

                
                string komut = " UPDATE tbl_Arac SET Marka = '" + textBox_Marka.Text + "', Model ='" + textBox_Model.Text + "', Detay_Ozellik = '" + textBox_Detay.Text + "', Alis_Fiyati = '" + textBox_AlisFiyati.Text + "',Satis_Fiyati = '" + textBox_SatisFiyati.Text + "',Alis_Tarihi = '" +dateTime_AlisTarihi.Text+ "',Satis_Tarihi = '" + dateTime_SatisTarihi.Text + "',Plaka ='"+textBox_Plaka.Text+"' WHERE Id = '" + textBox_id.Text + "'";
                SqlConn.Islemler(komut);
                MessageBox.Show("Güncelleme başarılı.");
            }
            catch (Exception hata)
            {

                MessageBox.Show("Hata!! "+ hata.Message);
            }
            
        }   

        private void btn_Sil_Click(object sender, EventArgs e)
        {
            DialogResult sil = new DialogResult();
            sil = MessageBox.Show("Silmek istediğinizden emin misiniz ?", "Uyarı",
               MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (sil == DialogResult.Yes)
            {
                string fotodosyasil = "select Fotograf_Uzanti from tbl_Fotograf where Arac_Id='"+textBox_id.Text+"'";
                DataTable dtbl = SqlConn.goster(fotodosyasil);

                string hedef = Application.StartupPath + @"\photos\";

                for (int i = 0; i < SqlConn.goster(fotodosyasil).Rows.Count; i++)
                {
                    string uzanti = dtbl.Rows[i]["Fotograf_Uzanti"].ToString();
                    if (System.IO.File.Exists(@hedef+uzanti))
                    {
                        System.IO.File.Delete(@hedef+uzanti);
                    }
                }
                
                string deletefoto = "delete from tbl_Fotograf WHERE Arac_Id='" + textBox_id.Text + "'";
                SqlConn.Islemler(deletefoto);

                try
                {
                    string deletearac = "delete from tbl_Arac WHERE Id='" + textBox_id.Text + "'";
                    SqlConn.Islemler(deletearac);
                    MessageBox.Show("Araç silindi.");
                }
                catch (Exception)
                {

                    MessageBox.Show("Hata!!");
                }
            }
            else if (sil == DialogResult.No)
            {
                MessageBox.Show("iptal edildi.");
            }
        }
        

        private void btn_Geri_Click(object sender, EventArgs e)
        {
            AracDetayDuzenle arcgnc = new AracDetayDuzenle();
            arcgnc.Show();
            this.Hide();
        }

        private void button_Satildi_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox_SatisFiyati.Text=="")
                {
                    MessageBox.Show("Araç satış bedeli girilmedi!");
                }
                else if (Convert.ToInt32(textBox_AlisFiyati.Text)>Convert.ToInt32(textBox_SatisFiyati.Text))
                {
                    int hesapla =  Convert.ToInt32(textBox_AlisFiyati.Text) - Convert.ToInt32(textBox_SatisFiyati.Text) ;
                    string insert="update  tbl_Arac set Zarar='"+hesapla+ "', Kar='"+0+"' where Id = '" + textBox_id.Text + "'";
                    SqlConn.Islemler(insert);

                    string updatesatis = "update tbl_Arac set Satis_Durumu = 1, Satis_Fiyati ='" + textBox_SatisFiyati.Text + "',Satis_Tarihi = '" + dateTime_SatisTarihi.Text + "' where Id = '" + textBox_id.Text + "'";
                    SqlConn.Islemler(updatesatis);
                    MessageBox.Show("Araç satıldı olarak işaretlendi.");
                }
                else if (Convert.ToInt32(textBox_AlisFiyati.Text) < Convert.ToInt32(textBox_SatisFiyati.Text))
                {
                    int hesapla =  Convert.ToInt32(textBox_SatisFiyati.Text) - Convert.ToInt32(textBox_AlisFiyati.Text)  ;
                    string insert = "update  tbl_Arac set Kar='" + hesapla + "', Zarar='" + 0 + "' where Id = '" + textBox_id.Text + "'";
                    SqlConn.Islemler(insert);

                    string updatesatis = "update tbl_Arac set Satis_Durumu = 1, Satis_Fiyati ='" + textBox_SatisFiyati.Text + "',Satis_Tarihi = '" + dateTime_SatisTarihi.Text + "' where Id = '" + textBox_id.Text + "'";
                    SqlConn.Islemler(updatesatis);
                    MessageBox.Show("Araç satıldı olarak işaretlendi.");
                }
                
            }
            catch (Exception hata)
            {

                MessageBox.Show(hata.Message);
            }
            
        }

        private void dateTime_AlisTarihi_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
