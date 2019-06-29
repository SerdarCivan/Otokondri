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
    public partial class Raporlar : Form
    {
        public Raporlar()
        {
            InitializeComponent();
        }

        private void button_karzarar_Click(object sender, EventArgs e)
        {
            KarZarar zarar = new KarZarar();
            zarar.Show();
            this.Hide();
        }
        

        private void button_markabaz_Click(object sender, EventArgs e)
        {
            MarkaBaz mrkbaz = new MarkaBaz();
            mrkbaz.Show();
            this.Hide();
        }

        private void Raporlar_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }
    }
}
