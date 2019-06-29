using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otokondri
{
    public partial class SifreYenileme : Form
    {
        public SifreYenileme()
        {
            InitializeComponent();
        }
        public string id;
        private void SifreYenileme_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void button_Kaydet_Click(object sender, EventArgs e)
        {
            string sorgu = "select Random from ParolaYenile where Personel_Id='" + id + "' and Durum = '" + 'A' + "'";
            DataTable tbl = SqlConn.goster(sorgu);
            if (textBox_Random.Text==tbl.Rows[0]["Random"].ToString())
            {
                if (textBox_Parola.Text==textBox_ParolaTekrar.Text)
                {
                    //string sorguu = "select Durum from ParolaYenile where Personel_Id='"+id+"' and Durum = '"+'A'+"'";
                    //DataTable dtbl = SqlConn.goster(sorguu);

                    string guncelle = "Update tbl_Personel set  Parola='" + textBox_Parola.Text + "' where Id='" + id + "'";
                    try
                    {
                        SqlConn.Islemler(guncelle);
                        MessageBox.Show("Güncelleme Başarili.");
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Hata!!");
                    }
                    
                        string guncelleepsota = "update ParolaYenile set GuncellemeTarihi='"+DateTime.Now.ToShortDateString()+"',Durum ='"+"P"+"' where Personel_Id=" + id;
                        SqlConn.Islemler(guncelleepsota);
                    
                        this.Hide();
                }
                else
                {
                    MessageBox.Show("Şifreler Uyuşmuyor!!");
                }
                
            }
            else
            {
                MessageBox.Show("Parola yenileme kodu yanlış.");
            }
            
        }
    }
}
