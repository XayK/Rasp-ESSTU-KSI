using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace Rasp
{
    public partial class Form1 : Form
    {
        class PeopleComparer : IComparer<Auditor>
        {
            public int Compare(Auditor p1, Auditor p2)
            {
                if (string.Compare(p1.name, p2.name) > 0)
                    return 1;
                else if (string.Compare(p1.name, p2.name) < 0)
                    return -1;
                else
                    return 0;
            }
        }


        class Teacher
        {
            public string name;
            string[,] subject;

            
            public Teacher()
            {
                name = "";
                subject = new string[12,6];//первое кол-во дней, второе кол-во пар
            }
            public void setname(string tmp)
            {
                name = tmp;
            }
            public string getname()
            {
                return name;
            }
            public void setsubject(string tmp,int day,int pair)
            {
                subject[day, pair] = tmp;
            }
            public string getsubject(int day, int pair)
            {
                return subject[day,pair];
            }
            
        }
        class Auditor:Teacher
        {
            static List<string> exists=new List<string>();
            public Auditor() : base()
            {
            }
            static public bool Contains(string tmp)
            {
                for (int i = 0; i < exists.Count; i++)
                    if (exists[i] == tmp)
                        return true;
                return false;
            }
            public new void setname(string tmp)
            {
                exists.Add(tmp);
                name = tmp;
            }
            static public void ClearExistc()
            {
                exists.Clear();
            }

        }
        public struct record
        {
            public string teacher;
            public string subject;
            public string aud;
            public string group;

            public string toTeacher()
            {
                return group+' '+ aud+' '+subject;
            }
            public string toAud()
            {
                return group + ' ' + teacher + ' ' + subject;
            }
        }

        public class Pair<T, K>
        {
            public T First { get; set; }
            public K Second { get; set; }
        }

        List<string> mass = new List<string>();
        string line;
        List<Teacher> Tmass = new List<Teacher>();
        List<Auditor> Aud = new List<Auditor>();
        Stack<Pair<int, int>> htmlopen = new Stack<Pair<int, int>>();
        bool colledge = true;

      
        public Form1()
        {
            InitializeComponent();
            this.KeyPreview = true;
            ///
            string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "configuration" + ".cfg";
            if (!File.Exists(fileName))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(fileName))
                {
                    //написать настройки
                    sw.WriteLine("[CONFIG]");

                    sw.WriteLine("BakLink = https://portal.esstu.ru/bakalavriat/Caf42.htm");

                    sw.WriteLine("MagLink = https://portal.esstu.ru/spezialitet/Caf44.htm");

                    sw.WriteLine("Mag/Col = On");
                }
            }

            // Open the file to read from.
            using (StreamReader sr = File.OpenText(fileName))
            {
                string s = "";
                while ((s = sr.ReadLine()) != null)
                {
                    if (s[0] != '[')
                    {
                        if (s.Substring(0, 7) == "BakLink")
                        {
                            linkb = s.Substring(10);
                        }
                        else if (s.Substring(0, 7) == "MagLink")
                        {
                            linkm = s.Substring(10);
                        }
                        else
                        {
                            if (s.Substring(10) == "On")
                            {
                                colledge = true;
                            }
                            else colledge = false;
                        }
                    }
                    //считать настройки
                }
            }

            ////
            //dataGridView1.Visible = false;///////
            start();
            for (int x = 0; x < 13; x++)
            {
                dataGridView1.Rows.Add();
            }
            dataGridView1.Rows[0].Cells[0].Value = "Пнд";
            dataGridView1.Rows[1].Cells[0].Value = "Втр";
            dataGridView1.Rows[2].Cells[0].Value = "Срд";
            dataGridView1.Rows[3].Cells[0].Value = "Чтв";
            dataGridView1.Rows[4].Cells[0].Value = "Птн";
            dataGridView1.Rows[5].Cells[0].Value = "Сбт";

            dataGridView1.Rows[7].Cells[0].Value = "Пнд";
            dataGridView1.Rows[8].Cells[0].Value = "Втр";
            dataGridView1.Rows[9].Cells[0].Value = "Срд";
            dataGridView1.Rows[10].Cells[0].Value = "Чтв";
            dataGridView1.Rows[11].Cells[0].Value = "Птн";
            dataGridView1.Rows[12].Cells[0].Value = "Сбт";
            comboBox1.SelectedIndex = 0;

            saveFileDialog1.Filter = "MS Office Excel(*.xlsx)|*.xlsx|All files(*.*)|*.*";


           
            string fileName2 = System.Windows.Forms.Application.StartupPath + "\\" + "Shablon" + ".xlsx";
            if (!File.Exists(fileName2))
            {
                
            }
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            show();
            
        }
        string linkb= "https://portal.esstu.ru/bakalavriat/Caf42.htm",
               linkm= "https://portal.esstu.ru/spezialitet/Caf44.htm";
        public void start()
        {
            try
            {
                GetHTML(linkb);//ввод ссылки
            }
            catch
            {
                MessageBox.Show("Невозможно скачать расписание с сайта esstu.ru", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }
            Parser();


            mass.RemoveAt(0); mass.RemoveAt(0);
            /* string[] newmas = new string[mass.Count];

             mass.CopyTo(newmas,0);
             newmas[mass.Count - 1] = "";
             listBox1.Items.AddRange(newmas);*/
            NextParser('b');
            //if (radioButton2.Checked)
            //{
            //    for (int i = 0; i < Aud.Count; i++)
            //    {
            //        comboBox1.Items.Add(Aud[i].getname());
            //    } 
            //}
            //else
            //    for (int i = 0; i < Tmass.Count; i++)
            //    {
            //        comboBox1.Items.Add(Tmass[i].getname());
            //    }

            /////////////////
            mass.Clear();
            try
            {
                GetHTML(linkm);//ввод ссылки
            }
            catch
            {
                MessageBox.Show("Невозможно скачать расписание с сайта esstu.ru", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }
            Parser();

            Aud.Sort(new PeopleComparer());

            mass.RemoveAt(0); mass.RemoveAt(0);

            NextParser('m');
            if (radioButton2.Checked)
            {
                for (int i = 0; i < Aud.Count; i++)
                {
                    comboBox1.Items.Add(Aud[i].getname());
                }
            }
            else
                for (int i = 0; i < Tmass.Count; i++)
                {
                    comboBox1.Items.Add(Tmass[i].getname());
                }

        }

        public void show()
        {
          

            int Z = comboBox1.SelectedIndex;
            int pl = 0;

            for (int x = 0; x < 12; x++)
            {

                for (int y = 1; y < 7; y++)
                {
                    
                    
                        if (x<=5) pl = 0;
                        else pl = 1;


                        if(radioButton1.Checked==true)
                            dataGridView1.Rows[x+pl].Cells[y].Value = Tmass[Z].getsubject(x , y-1);
                        else
                            dataGridView1.Rows[x+pl].Cells[y].Value = Aud[Z].getsubject(x, y - 1);
                    
                }
                
            }
        }

        public record takeapart(string str)
        {
            record temprecord = new record();
            string tempstr = ""; //int flag = 0;
            int i = 0;

            try
            {
                while (!((str[i+2] == 'a' && str[i + 3] == '.') || (str[i + 2] == 'а'  && str[i+3] == '.')))
                {
                    tempstr += str[i];
                    i++;
                }
                temprecord.group = tempstr;
                tempstr = "";


                i += 4;               
                while (str[i] != ' ')
                {
                     tempstr += str[i];
                    i++;
                }
                temprecord.aud = tempstr;
                tempstr = "";
                while (str[i] == ' ')
                {
                    i++;
                }

                while (!((str[i]<='Я' && str[i]>='А') && (str[i+1] <= 'я' && str[i+1] >= 'а')))
                {
                    tempstr += str[i];
                    i++;
                }
                while (str[i] == ' ')
                {
                    i++;
                }
                tempstr += ' ';
                while (i < str.Length)
                {
                    if (str[i - 1] == ' ' || str[i - 1] == '.' || str[i - 1] == '-') tempstr += str[i].ToString().ToUpper();
                    i++;
                }
                temprecord.subject = tempstr;
            }
            catch
            {
                return temprecord;
            }
            return temprecord;
        }

        public void NextParser(char type)
        {
            //type=='m'- колледж+мага
            //type=='b'- основное - бакалавры
            int flag = 0; int inputcounter = 0;
            Teacher temp=new Teacher();
            for(int i=0;i<mass.Count;i++)
            {
                if(mass[i]!=null && mass[i].Length>10)
                if(mass[i].Substring(0,10)=="Расписание")
                {
                    temp = new Teacher();
                    string tmpname= mass[i].Substring(mass[i].IndexOf(':')+2,mass[i].Length- mass[i].IndexOf(':')-2);
                    temp.setname(tmpname);
                    flag = 1;
                    inputcounter = 0;
                }
                if(flag>0)
                {
                    int plus;
                    if (flag == 1) plus = 0;
                    else plus = 6;
                    if (mass[i] == "Пнд")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 0 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' ||  temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 0 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if(checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 0 + plus, tmpi);
                                        inputcounter++;
                                    }


                                }
                                if (type == 'b' || inputcounter>0)
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 0 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 0 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                        //////////////////////////////////

                    }
                    else if(mass[i] == "Втр")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 1 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' || temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 1 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if (checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 1 + plus, tmpi);
                                        inputcounter++;
                                    }
                                }
                                if (type == 'b' || inputcounter >0)
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 1 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 1 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                        //////////////////////////////////
                    }
                    else if (mass[i] == "Срд")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 2 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' || temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 2 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if (checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 2 + plus, tmpi);
                                        inputcounter++;
                                    }
                                }
                                if (type == 'b' || inputcounter>0)  
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 2 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 2 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                    }
                    else if (mass[i] == "Чтв")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 3 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' || temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 3 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if (checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 3 + plus, tmpi);
                                        inputcounter++;
                                    }
                                }
                                if (type == 'b' || inputcounter>0)
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 3 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 3 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                    }
                    else if (mass[i] == "Птн")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 4 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' || temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 4 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if (checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 4 + plus, tmpi);
                                        inputcounter++;
                                    }
                                }
                                if (type == 'b' || inputcounter>0)
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 4 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 4 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                    }
                    else if (mass[i] == "Сбт")
                    {
                        for (int tmpi = 0; tmpi < 6; tmpi++)
                        {
                            i++;
                            if (mass[i].Contains('_'))
                                mass[i] = "       ";
                            else
                            {
                                record temprec = takeapart(mass[i]);
                                if ((temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K') && colledge == false)
                                    break;
                                temprec.teacher = temp.getname();
                                if (type == 'b')
                                    temp.setsubject(temprec.toTeacher(), 5 + plus, tmpi);
                                else if (temprec.group[0] == 'К' || temprec.group[0] == 'K' || temprec.group[1] == 'К' || temprec.group[1] == 'K' || temprec.group[0] == 'М' || temprec.group[1] == 'М')
                                {
                                    bool checkflag1 = true;


                                    for (int search = 0; search < Tmass.Count; search++)
                                    {
                                        if (temp.getname() == Tmass[search].getname())
                                        {
                                            Tmass[search].setsubject(temprec.toTeacher(), 5 + plus, tmpi);
                                            checkflag1 = false;
                                        }
                                    }

                                    if (checkflag1)
                                    {
                                        temp.setsubject(temprec.toTeacher(), 5 + plus, tmpi);
                                        inputcounter++;
                                    }
                                }
                                if (type == 'b' || inputcounter>0)
                                    if (Auditor.Contains(temprec.aud))
                                    {
                                        for (int tempi = 0; tempi < Aud.Count(); tempi++)
                                        {
                                            if (Aud[tempi].name == temprec.aud)
                                            {
                                                Aud[tempi].setsubject(temprec.toAud(), 5 + plus, tmpi);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Auditor tempaud = new Auditor();
                                        tempaud.setname(temprec.aud);
                                        tempaud.setsubject(temprec.toAud(), 5 + plus, tmpi);
                                        Aud.Add(tempaud);
                                    }
                            }
                        }
                        if (flag == 1) flag = 2;
                        //else if(type=='b')
                        else if(inputcounter>0 || type == 'b')
                        {
                            flag = 0;
                            Tmass.Add(temp);
                        }
                        //else
                        //{
                        //    bool ffl = true;
                        //    for(int ij1=0;ij1<Tmass.Count;ij1++)
                        //    {
                        //        if (Tmass[ij1].name == temp.name)
                        //        {
                        //            ffl = false;
                        //            break;
                        //        }
                        //    }
                        //    if (ffl)
                        //    {
                        //        flag = 0;
                        //        Tmass.Add(temp);
                        //    }
                        //}
                    }
                }
            }
        }  
        

        public void GetHTML(string link)
        {
            int i = 0;
            WebClient client = new WebClient();
            client.Encoding = Encoding.GetEncoding(1251);
            client.Headers.Add("user-agent", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            Stream data = client.OpenRead(link);//сайт скачивания расписания
            //для бакалавров - https://portal.esstu.ru/bakalavriat/Caf40.htm
            // для маги - https://portal.esstu.ru/spezialitet/Caf44.htm
            StreamReader reader = new StreamReader(data, Encoding.GetEncoding(1251));
            line = reader.ReadLine();
            mass.Add(line);
            while (line != null)
            {
                line = reader.ReadLine();
                mass.Add(line);
                i++;
            }
            data.Close();
            reader.Close();

        }
      
        public void Parser()
        {
 
                
            for (int i = 0; i < mass.Count()-1; i++)
            {
                if (i < 0) i = 0;
                for (int j = 0; j < mass[i].Length; j++)
                {
                   
                    if (mass[i][j] == '<')
                    {
                        Pair<int, int> temp = new Pair<int, int>();
                        temp.First = i;
                        temp.Second = j;
                        htmlopen.Push(temp);
                    }
                    else if (mass[i][j] == '>')
                    {
                        if (htmlopen.Peek().First != i)
                        {
                            mass[htmlopen.Peek().First] = mass[htmlopen.Peek().First].Remove(htmlopen.Peek().Second, mass[htmlopen.Peek().First].Length - htmlopen.Peek().Second);
                            mass[i] = mass[i].Remove(0, j + 1);
                            for (int cou = i; cou > htmlopen.Peek().First; cou--)
                            {
                                mass.RemoveAt(cou);
                                i--;
                            }
                            
                        }
                        else
                            mass[i] = mass[i].Remove(htmlopen.Peek().Second, j - htmlopen.Peek().Second+1);
                        htmlopen.Pop();
                        if (mass[i] == "")
                        {
                            mass.RemoveAt(i);
                            i--;
                           
                        }
                        i--;
                        break;

                    }
                    
                }
            }
            /////////////////////////
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void поПреподавателямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "Shablon" + ".xlsx";
            Cursor.Current = Cursors.WaitCursor;
            //fileName_new = System.Windows.Forms.Application.StartupPath + "\\" + "Shablon_new" + ".xlsx";
            try
            {
                //Приложение самого Excel
                var excelapp = new Excel.Application();
                excelapp.Visible = false;
                //Книга.
                var excelappworkbooks = excelapp.Workbooks;
                excelapp.Workbooks.Open(fileName,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing);
                var excelappworkbook = excelappworkbooks[1];
                //Получаем массив ссылок на листы выбранной книги
                var excelsheets = excelappworkbook.Worksheets;
                //Выбираем лист 1
                var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);


                for (int m = 3; m < 2 + Tmass.Count * 2; m += 2)
                {
                    var excelcells = (Excel.Range)excelworksheet.Cells[2, m];
                    excelcells.Value2 = Tmass[(m - 2) / 2].getname();
                }
                for (int m = 2; m < 2 + Tmass.Count * 2; m++)
                {
                    for (int n = 3; n < 50; n++)
                    {
                        var excelcells = (Excel.Range)excelworksheet.Cells[n, m + 1];
                        int plus = 0;
                        if ((m + 1) % 2 == 0) plus = 6;
                        if ((n - 3) % 8 < 6)
                            excelcells.Value2 = Tmass[(m - 2) / 2].getsubject((n - 3) / 8 + (plus), (n - 3) % 8);
                        //excelworksheet.Cells[m, n] = Tmass[(m - 2) / 2].getsubject((n - 2) / 8, (n - 2) % 8);
                    }
                }

                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                // получаем выбранный файл
                string filename_new = saveFileDialog1.FileName;
                // сохраняем текст в файл
                excelworksheet.SaveAs(filename_new);
                Cursor.Current = Cursors.Arrow;
                MessageBox.Show("Файл сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);


                excelappworkbook.Close();
                excelappworkbooks.Close();

                excelapp.Quit();
            }
            catch
            {
                MessageBox.Show("Нету доступа к Excel файлу шаблона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void обновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Tmass.Clear();
            Aud.Clear();
            Auditor.ClearExistc();
            comboBox1.Items.Clear();
            Cursor.Current = Cursors.WaitCursor;
            mass.Clear();
            start();
            comboBox1.SelectedIndex = 0;
            
            Cursor.Current = Cursors.Arrow;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)//аудит
        {
            label1.Text = "Аудитории";
            comboBox1.Items.Clear();
            for (int i = 0; i < Aud.Count; i++)
            {
                comboBox1.Items.Add(Aud[i].getname());
            }
            comboBox1.SelectedIndex = 0;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)//препода
        {
            label1.Text = "Преподаватели";
            comboBox1.Items.Clear();

            for (int i = 0; i < Tmass.Count; i++)
            {
                comboBox1.Items.Add(Tmass[i].getname());
            }
            comboBox1.SelectedIndex = 0;

        }

        private void поАудиториямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 Input=new Form2();
            for (int i = 0; i < Aud.Count; i++)
            {
                Input.comboBox1.Items.Add(Aud[i].getname());
            }
            Input.comboBox1.SelectedIndex = 0;
            Input.ShowDialog();
            int selected = Input.comboBox1.SelectedIndex;
            if (Input.flagger==false) return;
            ///////////
            Cursor.Current = Cursors.WaitCursor;

            
            string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "Shabaud" + ".xlsx";
            //fileName_new = System.Windows.Forms.Application.StartupPath + "\\" + "Shabaud_new" + ".xlsx";

            try
            {
                //Приложение самого Excel
                var excelapp = new Excel.Application();
                excelapp.Visible = false;
                //Книга.
                var excelappworkbooks = excelapp.Workbooks;
                excelapp.Workbooks.Open(fileName,
                                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                          Type.Missing, Type.Missing);
                var excelappworkbook = excelappworkbooks[1];
                //Получаем массив ссылок на листы выбранной книги
                var excelsheets = excelappworkbook.Worksheets;
                //Выбираем лист 1
                var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                ///ввод названия ячейки
                {
                    var excelcells = (Excel.Range)excelworksheet.Cells[2, 1];
                    excelcells.Value2 = Aud[selected].getname();
                }
                for (int m = 2; m < 8; m++)//столбцы
                {
                    for (int n = 3; n < 14; n++)
                    {
                        var excelcells = (Excel.Range)excelworksheet.Cells[n, m];
                        int plus = 0;
                        if (n  % 2 == 0) plus = 6;
                        if (Aud[selected].getsubject(m - 2 + plus, (n - 3) / 2)!=null)
                            excelcells.Value2 = Aud[selected].getsubject(m-2+plus,(n-3)/2);
                        //excelworksheet.Cells[m, n] = Tmass[(m - 2) / 2].getsubject((n - 2) / 8, (n - 2) % 8);
                    }
                }



                saveFileDialog1.FileName = Input.comboBox1.SelectedItem.ToString();
                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                // получаем выбранный файл
                string filename_new = saveFileDialog1.FileName;
                excelworksheet.SaveAs(filename_new);
                Cursor.Current = Cursors.Arrow;
                MessageBox.Show("Файл сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);


                excelappworkbook.Close();
                excelappworkbooks.Close();

                excelapp.Quit();



            }
            catch
            {
                MessageBox.Show("Нету доступа к Excel файлу шаблона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        private void помощьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Для нормальной работы - шаблон Excel файла берется из файлов #shablon# и #shabaud#","Окно Помощи",MessageBoxButtons.OK,MessageBoxIcon.Information);
        }

        private void настройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.textBox1.Text = linkb;
            form4.textBox2.Text = linkm;
            form4.checkBox1.Checked = colledge;
            form4.ShowDialog();
            if(form4.flagger==1)
            {
                colledge = form4.checkBox1.Checked;
                linkb = form4.textBox1.Text;
                linkm = form4.textBox2.Text;
                string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "configuration" + ".cfg";
                // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(fileName))
                    {
                        //написать настройки
                        sw.WriteLine("[CONFIG]");

                        sw.WriteLine("BakLink = "+ linkb);

                        sw.WriteLine("MagLink = "+ linkm);

                        if (colledge)
                            sw.WriteLine("Mag/Col = On");
                        else
                            sw.WriteLine("Mag/Col = Off");

                }


            }

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                Tmass.Clear();
                Aud.Clear();
                Auditor.ClearExistc();
                comboBox1.Items.Clear();
                Cursor.Current = Cursors.WaitCursor;
                mass.Clear();
                start();
                comboBox1.SelectedIndex = 0;

                Cursor.Current = Cursors.Arrow;
            }
            else if (e.KeyCode == Keys.F2)
            {
                string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "Shablon" + ".xlsx";
                Cursor.Current = Cursors.WaitCursor;
                //fileName_new = System.Windows.Forms.Application.StartupPath + "\\" + "Shablon_new" + ".xlsx";
                try
                {
                    //Приложение самого Excel
                    var excelapp = new Excel.Application();
                    excelapp.Visible = false;
                    //Книга.
                    var excelappworkbooks = excelapp.Workbooks;
                    excelapp.Workbooks.Open(fileName,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing);
                    var excelappworkbook = excelappworkbooks[1];
                    //Получаем массив ссылок на листы выбранной книги
                    var excelsheets = excelappworkbook.Worksheets;
                    //Выбираем лист 1
                    var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);


                    for (int m = 3; m < 2 + Tmass.Count * 2; m += 2)
                    {
                        var excelcells = (Excel.Range)excelworksheet.Cells[2, m];
                        excelcells.Value2 = Tmass[(m - 2) / 2].getname();
                    }
                    for (int m = 2; m < 2 + Tmass.Count * 2; m++)
                    {
                        for (int n = 3; n < 50; n++)
                        {
                            var excelcells = (Excel.Range)excelworksheet.Cells[n, m + 1];
                            int plus = 0;
                            if ((m + 1) % 2 == 0) plus = 6;
                            if ((n - 3) % 8 < 6)
                                excelcells.Value2 = Tmass[(m - 2) / 2].getsubject((n - 3) / 8 + (plus), (n - 3) % 8);
                            //excelworksheet.Cells[m, n] = Tmass[(m - 2) / 2].getsubject((n - 2) / 8, (n - 2) % 8);
                        }
                    }

                    if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                        return;
                    // получаем выбранный файл
                    string filename_new = saveFileDialog1.FileName;
                    // сохраняем текст в файл
                    excelworksheet.SaveAs(filename_new);
                    Cursor.Current = Cursors.Arrow;
                    MessageBox.Show("Файл сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);


                    excelappworkbook.Close();
                    excelappworkbooks.Close();

                    excelapp.Quit();
                }
                catch
                {
                    MessageBox.Show("Нету доступа к Excel файлу шаблона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (e.KeyCode == Keys.F3)
            {
                Form2 Input = new Form2();
                for (int i = 0; i < Aud.Count; i++)
                {
                    Input.comboBox1.Items.Add(Aud[i].getname());
                }
                Input.comboBox1.SelectedIndex = 0;
                Input.ShowDialog();
                int selected = Input.comboBox1.SelectedIndex;
                if (Input.flagger == false) return;
                ///////////
                Cursor.Current = Cursors.WaitCursor;


                string fileName = System.Windows.Forms.Application.StartupPath + "\\" + "Shabaud" + ".xlsx";
                //fileName_new = System.Windows.Forms.Application.StartupPath + "\\" + "Shabaud_new" + ".xlsx";

                try
                {
                    //Приложение самого Excel
                    var excelapp = new Excel.Application();
                    excelapp.Visible = false;
                    //Книга.
                    var excelappworkbooks = excelapp.Workbooks;
                    excelapp.Workbooks.Open(fileName,
                                            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                              Type.Missing, Type.Missing);
                    var excelappworkbook = excelappworkbooks[1];
                    //Получаем массив ссылок на листы выбранной книги
                    var excelsheets = excelappworkbook.Worksheets;
                    //Выбираем лист 1
                    var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);

                    ///ввод названия ячейки
                    {
                        var excelcells = (Excel.Range)excelworksheet.Cells[2, 1];
                        excelcells.Value2 = Aud[selected].getname();
                    }
                    for (int m = 2; m < 8; m++)//столбцы
                    {
                        for (int n = 3; n < 14; n++)
                        {
                            var excelcells = (Excel.Range)excelworksheet.Cells[n, m];
                            int plus = 0;
                            if (n % 2 == 0) plus = 6;
                            if (Aud[selected].getsubject(m - 2 + plus, (n - 3) / 2) != null)
                                excelcells.Value2 = Aud[selected].getsubject(m - 2 + plus, (n - 3) / 2);
                            //excelworksheet.Cells[m, n] = Tmass[(m - 2) / 2].getsubject((n - 2) / 8, (n - 2) % 8);
                        }
                    }



                    saveFileDialog1.FileName = Input.comboBox1.SelectedItem.ToString();
                    if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                        return;
                    // получаем выбранный файл
                    string filename_new = saveFileDialog1.FileName;
                    excelworksheet.SaveAs(filename_new);
                    Cursor.Current = Cursors.Arrow;
                    MessageBox.Show("Файл сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);


                    excelappworkbook.Close();
                    excelappworkbooks.Close();

                    excelapp.Quit();



                }
                catch
                {
                    MessageBox.Show("Нету доступа к Excel файлу шаблона.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            show();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //int index=e.RowIndex;
            if(dataGridView1.CurrentCell.Value!=null)
                textBox1.Text = dataGridView1.CurrentCell.Value.ToString();
        }
    }
}
