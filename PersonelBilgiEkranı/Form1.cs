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
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Media.Imaging;
using System.Windows.Controls;
using System.Data.SQLite;

namespace PersonelBilgiEkranı
{
    public partial class Form1 : Form
    {
        String file, image, deger, app;
        double aranan;
        Excel.Application xlApp ;
        Excel.Workbook xlWorkBook ;
        Excel.Worksheet xlWorkSheet ;
        Excel.Range range ;
        public Form1()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            listView1.View = View.Details;
            app = "Database";
            pictureBox1.Image = System.Drawing.Image.FromFile("default-avatar.png");
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            
            SQLiteConnection con = new SQLiteConnection("Data Source="+app+"\\db2.db3;Version=3;Read Only=False;");
            SQLiteCommand q1 = new SQLiteCommand("select * from veri where id=1", con);
            SQLiteCommand q2 = new SQLiteCommand("select * from veri where id=2", con);
            con.Open();
            SQLiteDataReader r1 = q1.ExecuteReader();
                while (r1.Read())
                {
                    file = r1["yol"].ToString();
                }
                r1.Close();
                SQLiteDataReader r2 = q2.ExecuteReader();
                while (r2.Read())
                {
                    image = r2["yol"].ToString();
                }
                r2.Close();
            con.Close();
        }
        
                    private void releaseObject(object obj)
                    {
                        try
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                            obj = null;
                        }
                        catch (Exception ex)
                        {
                            obj = null;
                            MessageBox.Show("Unable to release the Object " + ex.ToString());
                        }
                        finally
                        {
                            GC.Collect();
                        }
            }

                    private void button1_Click(object sender, EventArgs e)
                    {
                        listView1.Items.Clear();
                                    string c, sondeger;
                                    SQLiteConnection cnn = new SQLiteConnection("Data Source=" + app + "\\db2.db3;Version=3;Read Only=False;");
                                    cnn.Open();
                                    if (textBox1.Text != "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox1.Text; c = "c1"; textBox1.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text != "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox2.Text; c = "c2"; textBox2.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text != "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox3.Text; c = "c3"; textBox3.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text != ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox4.Text; c = "c4"; textBox4.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text != "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox5.Text; c = "c5"; textBox5.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text != "" && textBox7.Text == "" && textBox8.Text == "")
                                    { deger = textBox6.Text; c = "c6"; textBox6.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text != "" && textBox8.Text == "")
                                    { deger = textBox7.Text; c = "c7"; textBox7.Clear(); }
                                    else if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == ""
                                    && textBox5.Text == "" && textBox6.Text == "" && textBox7.Text == "" && textBox8.Text != "")
                                    { deger = textBox8.Text; c = "c8"; textBox8.Clear(); }
                                    else
                                    {
                                        MessageBox.Show("Sadece bir kritere göre arama yapılabilir!"); c = "c1"; deger = "yok";
                                    }
                                        if(comboBox1.SelectedIndex == 0){
                                            sondeger ="'%" + deger + "%'";
                                        }
                                        else if (comboBox1.SelectedIndex == 1)
                                        {
                                            sondeger = "'" + deger + "%'";
                                        }
                                        else                                   {
                                            sondeger = "'%" + deger + "'";
                                        }
                                        SQLiteCommand q = new SQLiteCommand("select * from personal where " + c + " like "+sondeger, cnn);
                                        SQLiteDataReader r = q.ExecuteReader();
                                        while (r.Read())
                                        {
                                            string[] row = { r["c1"].ToString(), r["c2"].ToString(), r["c3"].ToString(),
                                                               r["c4"].ToString(), r["c5"].ToString(), r["c6"].ToString(),
                                                               r["c7"].ToString(), r["c8"].ToString(), r["c9"].ToString() };
                                            var listViewItem = new System.Windows.Forms.ListViewItem(row);
                                            listView1.Items.Add(listViewItem);
                                        }
                                        if (listView1.Items.Count > 0)
                                        {
                                            listView1.Items[0].Selected = true;
                                            listView1.Select();
                                        }
                                        else
                                        {
                                            MessageBox.Show("Aranılan sonuç bulunamadı");
                                        }
                                        cnn.Close();
                    }
                    private void resim(double deger) {
                        double bas = basamak(deger);
                                            try
                                            {
                                                if (bas == 2)
                                                {
                                                    if (File.Exists(image + "\\000" + deger + ".jpg"))
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\000" + deger + ".jpg");
                                                    else
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\YENİ ELEMAN FOTOĞRAFLARI\\000" + deger + ".jpg");
                                                }
                                                else if (bas == 3)
                                                {
                                                    if (File.Exists(image + "\\00" + deger + ".jpg"))
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\00" + deger + ".jpg");
                                                    else
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\YENİ ELEMAN FOTOĞRAFLARI\\00" + deger + ".jpg");
                                                }
                                                else if (bas == 4)
                                                {
                                                    if (File.Exists(image + "\\0" + deger + ".jpg"))
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\0" + deger + ".jpg");
                                                    else
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\YENİ ELEMAN FOTOĞRAFLARI\\0" + deger + ".jpg");
                                                }
                                                else
                                                {
                                                    if (File.Exists(image + "\\" + deger + ".jpg"))
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\" + deger + ".jpg");
                                                    else
                                                        pictureBox1.Image = System.Drawing.Image.FromFile(image + "\\YENİ ELEMAN FOTOĞRAFLARI\\" + deger + ".jpg");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                pictureBox1.Image = System.Drawing.Image.FromFile("default-avatar.png");
                                            }
                                            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
                    }

                    private void listView1_SelectedIndexChanged(object sender, EventArgs e)
                    {
                        if (listView1.SelectedItems.Count > 0)
                        {
                            SQLiteConnection cnn = new SQLiteConnection("Data Source=" + app + "\\db2.db3;Version=3;Read Only=False;");
                            cnn.Open();
                            SQLiteCommand q = new SQLiteCommand("select * from personal where c6 = " + listView1.SelectedItems[0].SubItems[5].Text, cnn);
                            SQLiteDataReader r = q.ExecuteReader();
                            while (r.Read())
                            {
                                label10.Text = r["c1"].ToString();
                                label11.Text = r["c2"].ToString();
                                label12.Text = r["c3"].ToString();
                                label13.Text = r["c4"].ToString();
                                label14.Text = r["c5"].ToString();
                                label15.Text = r["c6"].ToString();
                                label16.Text = r["c7"].ToString();
                                label17.Text = r["c8"].ToString();
                                label18.Text = r["c9"].ToString();
                                resim(Convert.ToDouble(r["c6"].ToString()));
                            }
                            cnn.Close();
                        }
                    }
                    public double basamak(double sayi)
                    {
                        double bolen = 10, j = 0;
                        while (true)
                        {
                            double bolum = sayi / bolen;
                            if (bolum > 1)
                                j++;
                            else
                                break;
                            bolen *= 10;
                        }
                        j += 1;
                        return j;
                    }

                    private void excelDosyasınıYenileToolStripMenuItem_Click(object sender, EventArgs e)
                    {
                        DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
                        if (result == DialogResult.OK) // Test result.
                        {
                            if (openFileDialog1.FileName != null)
                            {
                                try
                                {
                                    SQLiteConnection cnn2 = new SQLiteConnection("Data Source=" + app + "\\db2.db3;Version=3;Read Only=False;");
                                    cnn2.Open();
                                    SQLiteCommand com = new SQLiteCommand("DELETE FROM veri where id=1", cnn2);
                                    com.ExecuteNonQuery();
                                    cnn2.Close();
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }
                            }
                            try
                            {
                                SQLiteConnection cnn2 = new SQLiteConnection("Data Source=" + app + "\\db2.db3;Version=3;Read Only=False;");
                                cnn2.Open();
                                SQLiteCommand com = new SQLiteCommand("INSERT INTO veri(yol,id) VALUES(@file,@id)", cnn2);
                                com.Parameters.AddWithValue("@file", openFileDialog1.FileName);
                                com.Parameters.AddWithValue("@id", 1);
                                com.ExecuteNonQuery();
                                cnn2.Close();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }


                        }
                    }

                    private void resimKlasörünüYenileToolStripMenuItem_Click(object sender, EventArgs e)
                    {
                        folderBrowserDialog1.ShowNewFolderButton = true;
 
                        // Kontrolü göster
                        DialogResult result = folderBrowserDialog1.ShowDialog();      
                        if (result == DialogResult.OK)
                        {
                         image = folderBrowserDialog1.SelectedPath;
                         SQLiteConnection con3 = new SQLiteConnection("Data Source=" + app + "\\db2.db3;Version=3;Read Only=False;");
                         con3.Open();
                         SQLiteCommand com3 = new SQLiteCommand("select * from veri where id=2", con3);
                         SQLiteDataReader et = com3.ExecuteReader();
                         if (et != null)
                         {
                             SQLiteCommand comd = new SQLiteCommand("delete from veri where id=2", con3);
                             comd.ExecuteNonQuery();
                         }
                         et.Close();
                             SQLiteCommand comi = new SQLiteCommand("insert into veri(yol,id) values(@image,@id)", con3);
                             comi.Parameters.AddWithValue("@image", image);
                             comi.Parameters.AddWithValue("@id", 2);
                             comi.ExecuteNonQuery();
                         con3.Close();
                        }
        }
                    }
                    }