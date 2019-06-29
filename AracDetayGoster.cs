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
    public partial class AracDetayGoster : Form
    {
        public AracDetayGoster()
        {
            InitializeComponent();
        }
        DataTable dtable;
        private void AracDetayGoster_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            SqlConnection con = new SqlConnection("Data Source=DESKTOP-7L6TOHU;Initial Catalog=OtokondriOtomasyon;Persist Security Info=True;User ID=sa;Password=123456789");
            con.Open();
            SqlDataAdapter adapter = new SqlDataAdapter("Select Id,Marka, Model,Detay_Ozellik from tbl_Arac", con);
             dtable= new DataTable();
            adapter.Fill(dtable);
            dataGridView_Goster.DataSource = dtable;
            con.Close();
            dataGridView_Goster.Columns[0].Visible = false;
        }

        private void AracDetayGoster_geri_Click(object sender, EventArgs e)
        {
            PersonelPanel PersonelPanel = new PersonelPanel();
            this.Hide();
        }

        int satirbilgisi;
        private void dataGridView_Goster_MouseClick(object sender, MouseEventArgs e)
        {
             satirbilgisi = dataGridView_Goster.HitTest(e.X, e.Y).RowIndex;

            if (satirbilgisi==null || satirbilgisi<0 || satirbilgisi >=dtable.Rows.Count)
            {
                return;
            }

            string islem = "select * from tbl_Arac where Id="+ dataGridView_Goster.Rows[satirbilgisi].Cells["Id"].Value.ToString();
            DataTable tbl = SqlConn.goster(islem);
            textBox1.Text = tbl.Rows[0]["Detay_Ozellik"].ToString();

            string sql = " select  ftg.Fotograf_Uzanti,ftg.Sira,ftg.ProfilFotograf from tbl_Arac arc,tbl_Fotograf ftg where arc.Id=ftg.Arac_Id and arc.Id='" + dataGridView_Goster.Rows[satirbilgisi].Cells["Id"].Value.ToString()+ "'and ftg.ProfilFotograf=1";
            DataTable  table = SqlConn.goster(sql);
            if (SqlConn.goster(sql).Rows.Count>0)
            {

            
                    if (table.Rows[0]["Fotograf_Uzanti"].ToString()!="")
                    {
                        pictureBox1.ImageLocation = Application.StartupPath + @"\photos\" + table.Rows[0]["Fotograf_Uzanti"].ToString();
                    }


            }
            else
            {
                pictureBox1.ImageLocation = "";
            }

        }
        int i = 1;
        private void button_right_Click(object sender, EventArgs e)
        {
            if (satirbilgisi == null || satirbilgisi < 0 || satirbilgisi >= dtable.Rows.Count)
            {
                return;
            }

            string sql = " select  ftg.Fotograf_Uzanti,ftg.Sira,ftg.ProfilFotograf from tbl_Arac arc,tbl_Fotograf ftg where arc.Id=ftg.Arac_Id and arc.Id=" + dataGridView_Goster.Rows[satirbilgisi].Cells["Id"].Value.ToString();
            DataTable table = SqlConn.goster(sql);
            if (SqlConn.goster(sql).Rows.Count > 0)
            {
                int resimcount = SqlConn.goster(sql).Rows.Count;
                
                if (i<=resimcount-1)
                {
                    if (table.Rows[i]["Fotograf_Uzanti"].ToString() != "")
                    {
                        pictureBox1.ImageLocation = Application.StartupPath + @"\photos\" + table.Rows[i]["Fotograf_Uzanti"].ToString();
                    }
                    i = i + 1;
                }
                else
                {
                    i = 0;
                }
            }
            else
            {
                pictureBox1.ImageLocation = "";
            }
        }
    }
}
