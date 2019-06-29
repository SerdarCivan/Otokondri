using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;


namespace Otokondri
{

    public partial class Login : Form
    {
        public static int s = 0;
        public Login()
        {

            Thread t = new Thread(new ThreadStart(SplashScreen));
            t.Start();
            Thread.Sleep(1500);
            InitializeComponent();
            t.Abort();
            
        }

        public void SplashScreen()
        {
            Application.Run(new SplashScreen());
        }

        private void Login_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        public string yetki;
        private void button_Loginn(object sender, EventArgs e)
        {
            SqlConnection con;
            SqlCommand cmd;
            SqlDataReader dr;

            string user = textBox_Personelid.Text;
            string pass = textBox_PersonelParola.Text;

            string islem = "select Id from tbl_Personel where Email = '" + textBox_Personelid.Text + "'";
            string id = SqlConn.Id_Deger(islem);

            try
            {
                string sorgu = "select Durum from ParolaYenile where Personel_Id='" + id + "' and Durum= '" + 'A' + "'";
                DataTable dtbl = SqlConn.goster(sorgu);
                string durum = dtbl.Rows[0]["Durum"].ToString(); ;

                if (durum == "A")
                {
                    MessageBox.Show("Daha önce parola yenileme isteği yolladınız, Lütfen mailinizi kontrol ediniz.");
                    SifreYenileme sfr = new SifreYenileme();
                    sfr.id = id;
                    sfr.Show();
                    //this.Hide();
                }

                else
                {
                    con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
                    cmd = new SqlCommand();
                    con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT * FROM tbl_Personel where Email='" + textBox_Personelid.Text + "' AND Parola='" + textBox_PersonelParola.Text + "'";
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        PersonelPanel PersonelPanelsec = new PersonelPanel();
                        PersonelPanelsec.personelid = textBox_Personelid.Text;
                        PersonelPanelsec.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Email veya şifrenizi kontrol ediniz.");
                    }
                    con.Close();
                }
            }
            catch
            {
                con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
                cmd = new SqlCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "SELECT * FROM tbl_Personel where Email ='" + textBox_Personelid.Text + "' AND Parola='" + textBox_PersonelParola.Text + "'";
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    PersonelPanel PersonelPanelsec = new PersonelPanel();
                    PersonelPanelsec.personelid = textBox_Personelid.Text;
                    PersonelPanelsec.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Email veya şifrenizi kontrol ediniz.");
                }
                con.Close();
            }
        }

        private void Sifremi_unuttum_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SifremiUnuttum Sfr = new SifremiUnuttum();
            Sfr.Show();
        }

        private void textBox_PersonelParola_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                SqlConnection con;
                SqlCommand cmd;
                SqlDataReader dr;

                string user = textBox_Personelid.Text;
                string pass = textBox_PersonelParola.Text;
                string islem = "select Id from tbl_Personel where Email = '" + textBox_Personelid.Text + "'";
                string id = SqlConn.Id_Deger(islem);

                try
                {
                    string sorgu = "select Durum from ParolaYenile where Personel_Id='" + id + "' and Durum= '" + 'A' + "'";
                    DataTable dtbl = SqlConn.goster(sorgu);
                    string durum = dtbl.Rows[0]["Durum"].ToString(); ;

                    if (durum == "A")
                    {
                        MessageBox.Show("Daha önce parola yenileme isteği yolladınız, Lütfen mailinizi kontrol ediniz.");
                        SifreYenileme sfr = new SifreYenileme();
                        sfr.id = id;
                        sfr.Show();
                        this.Hide();
                    }
                    else
                    {
                        con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
                        cmd = new SqlCommand();
                        con.Open();
                        cmd.Connection = con;
                        cmd.CommandText = "SELECT * FROM tbl_Personel where Email='" + textBox_Personelid.Text + "' AND Parola='" + textBox_PersonelParola.Text + "'";
                        dr = cmd.ExecuteReader();
                        if (dr.Read())
                        {
                            PersonelPanel PersonelPanelsec = new PersonelPanel();
                            PersonelPanelsec.personelid = textBox_Personelid.Text;
                            PersonelPanelsec.Show();
                            this.Hide();
                        }
                        else
                        {
                            MessageBox.Show("Email veya şifrenizi kontrol ediniz.");
                        }
                        con.Close();
                    }
                }
                catch
                {
                    con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
                    cmd = new SqlCommand();
                    con.Open();
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT * FROM tbl_Personel where Email='" + textBox_Personelid.Text + "' AND Parola='" + textBox_PersonelParola.Text + "'";
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        PersonelPanel PersonelPanelsec = new PersonelPanel();
                        PersonelPanelsec.personelid = textBox_Personelid.Text;
                        PersonelPanelsec.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Email veya şifrenizi kontrol ediniz.");
                    }
                    con.Close();
                }
            }
        }
    }
}
