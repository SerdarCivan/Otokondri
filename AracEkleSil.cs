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
using System.IO;

namespace Otokondri
{
    public partial class AracEkleSil : Form
    {
        public AracEkleSil()
        {
            InitializeComponent();
        }

        private void AracEkleSil_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            SqlConnection con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter("Select Marka, Model,Detay_Ozellik,Alis_Tarihi,Alis_Fiyati,Satis_Tarihi,Satis_Fiyati from tbl_Arac", con);
            DataTable dtable = new DataTable();
            adapter.Fill(dtable);
            dataGridView_AracSilListe.DataSource = dtable;
            con.Close();
        }
        string dosyaismi;
        public string dosyayolu1 = "";
        public string dosyayolu2 = "";
        public string dosyayolu3 = "";
        public string dosyayolu4 = "";
        private void button_Fotograf1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu1 = dosya.FileName;
            textBox_FotoUzanti1.Text = dosyayolu1;
            pictureBox_AracEklePicBox.ImageLocation = dosyayolu1;
        }
        private void button_Fotograf2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu2 = dosya.FileName;
            textBox_FotoUzanti2.Text = dosyayolu2;
            pictureBox_AracEklePicBox.ImageLocation = dosyayolu2;
        }

        private void button_Fotograf3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu3 = dosya.FileName;
            textBox_FotoUzanti3.Text = dosyayolu3;
            pictureBox_AracEklePicBox.ImageLocation = dosyayolu3;
        }

        private void button_Fotograf4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dosya = new OpenFileDialog();
            dosya.Filter = "Resim Dosyası |*.jpg;*.nef;*.png;*.jpeg  |  Tüm Dosyalar |*.*";
            dosya.ShowDialog();
            dosyayolu4 = dosya.FileName;
            textBox_FotoUzanti4.Text = dosyayolu4;
            pictureBox_AracEklePicBox.ImageLocation = dosyayolu4;
        }
        string yeniad = "";
        string id = "";

        private void button_AracEkleme_Click(object sender, EventArgs e)
        {
            try
            {
                string aracinsert = "insert into tbl_Arac(Marka,Model,Detay_Ozellik,Alis_Tarihi,Alis_Fiyati,Plaka) values('" + textBox_Marka.Text + "','" + textBox_Model.Text + "','" + textBox_AracEkleDetay.Text + "','" + dateTime_AlisTarihi.Text + "','" + int.Parse(textBox_AlisFiyati.Text) + "','" + textBox_Plaka.Text + "')SELECT SCOPE_IDENTITY();";
                id = SqlConn.Id_Deger(aracinsert);
                MessageBox.Show("İşlem Başarılı.");
            }
            catch (Exception)
            {
                MessageBox.Show("Hata!!");
            }
            

            if (textBox_FotoUzanti1.Text != "")
            {
                dosyaismi = Path.GetFileName(dosyayolu1);
                string kaynak = dosyayolu1;
                string hedef = Application.StartupPath + @"\photos\";
                yeniad = textBox_Plaka.Text + "_1.jpg";
                
                string fotografinsert = "insert into tbl_Fotograf(Fotograf_Uzanti,Arac_Id,Sira,ProfilFotograf) values('" + yeniad + "','" + id + "','"+1+"','"+1+"')";
                SqlConn.Islemler(fotografinsert);
                File.Copy(kaynak, hedef + yeniad);
            }
            

            if (textBox_FotoUzanti2.Text!="")
            {
                dosyaismi = Path.GetFileName(dosyayolu2);
                string kaynak = dosyayolu2;
                string hedef = Application.StartupPath + @"\photos\";
                yeniad = textBox_Plaka.Text + "_2.jpg";

                string fotografinsert = "insert into tbl_Fotograf(Fotograf_Uzanti,Arac_Id,Sira) values('" + yeniad + "','" + id + "','" + 2 + "')";
                SqlConn.Islemler(fotografinsert);
                File.Copy(kaynak, hedef + yeniad);
            }

            if (textBox_FotoUzanti3.Text != "")
            {
                dosyaismi = Path.GetFileName(dosyayolu3);
                string kaynak = dosyayolu3;
                string hedef = Application.StartupPath + @"\photos\";
                yeniad = textBox_Plaka.Text + "_3.jpg";

                string fotografinsert = "insert into tbl_Fotograf(Fotograf_Uzanti,Arac_Id,Sira) values('" + yeniad + "','" + id + "','" + 3 + "')";
                SqlConn.Islemler(fotografinsert);
                File.Copy(kaynak, hedef + yeniad);
            }

            if (textBox_FotoUzanti4.Text != "")
            {
                dosyaismi = Path.GetFileName(dosyayolu4);
                string kaynak = dosyayolu4;
                string hedef = Application.StartupPath + @"\photos\";
                yeniad = textBox_Plaka.Text + "_4.jpg";

                string fotografinsert = "insert into tbl_Fotograf(Fotograf_Uzanti,Arac_Id,Sira) values('" + yeniad + "','" + id + "','" + 4 + "')";
                SqlConn.Islemler(fotografinsert);
                File.Copy(kaynak, hedef + yeniad);
            }

            SqlConnection con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter("Select Marka, Model,Detay_Ozellik,Alis_Tarihi,Alis_Fiyati,Satis_Tarihi,Satis_Fiyati from tbl_Arac", con);
            DataTable dtable = new DataTable();
            adapter.Fill(dtable);
            dataGridView_AracSilListe.DataSource = dtable;
            con.Close();
        }
        private void AracEkleSil_geri_Click(object sender, EventArgs e)
        {
            PersonelPanel PersonelPanel = new PersonelPanel();
            this.Hide();
        }
    }
}
