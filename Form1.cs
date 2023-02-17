using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace STH_Automation_22
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (var item in listBox1.SelectedItems)
            {
                if (!listBox2.Items.Contains(item))
                {
                    listBox2.Items.Add(item);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<object> objs = new List<object>();

            foreach (var item in listBox2.SelectedItems)
            {
                objs.Add(item);

            }
            foreach (var item in objs)
            {
                listBox2.Items.Remove(item);
            }
        }
    }
    
}
