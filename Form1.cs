using iwantedue;
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

namespace textFileReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void onDragEnter(object sender, DragEventArgs e)
        {

        }

        private void onDragDrop(object sender, DragEventArgs e)
        {
            string filename;
            bool validData = GetFilename(out filename, e);
        }

        private void onDragLeave(object sender, EventArgs e)
        {

        }

        private void onDragOver(object sender, DragEventArgs e)
        {
            string filename;
            bool validData = GetFilename(out filename, e);
        }

        protected bool GetFilename(out string filename, DragEventArgs e)
        {
            bool ret = false;
            filename = String.Empty;

            if ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy)
            {
                Array data = ((IDataObject)e.Data).GetData("FileName") as Array;
                if (data != null)
                {
                    if ((data.Length == 1) && (data.GetValue(0) is String))
                    {
                        filename = ((string[])data)[0];
                        string ext = Path.GetExtension(filename).ToLower();
                        if ((ext == ".jpg") || (ext == ".csv") || (ext == ".txt") || (ext == ".msg"))
                        {
                            ret = true;
                        }
                    }
                }
            }
            return ret;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (this.openFileDialog1.Multiselect)
            {
                foreach (string s in this.openFileDialog1.FileNames)
                {
                    string ext = Path.GetExtension(s).ToLower();
                    if( ext == ".msg" )
                        this.listView1.Items.Add(new ListViewItem(s));
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string textmessage = "";
            foreach(ListViewItem lvi in this.listView1.Items){
                OutlookStorage.Message outlookMsg = new OutlookStorage.Message(lvi.Text);

                //Console.WriteLine("Subject: {0}", outlookMsg.Subject);
                //Console.WriteLine("Body: {0}", outlookMsg.BodyText);
                textmessage += outlookMsg.BodyText;
            }

            this.outputText.Text = textmessage;

            using (System.IO.StreamWriter file = new System.IO.StreamWriter("output.txt", true))
            {
                file.Write(textmessage);
            }  
        }
    }
}
