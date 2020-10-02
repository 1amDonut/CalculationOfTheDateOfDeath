using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Threading;
namespace TaiwanLunisolarCalendar

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("往生者姓名請勿空白!!!", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
               // MessageBox.Show("往生者姓名請問空白!!!");
            }
            else
            {
                SaveFileDialog sf = new SaveFileDialog();//使用存檔的對話框
                sf.DefaultExt = "doc";
                sf.Filter = "Word document|*.doc";
                sf.AddExtension = true;
                sf.RestoreDirectory = false;
                sf.Title = "另存新檔";
                sf.InitialDirectory = @"C:/";
                Word._Application word_app = new Microsoft.Office.Interop.Word.Application();
                Word._Document word_document;
                Object oEndOfDoc = "\\endofdoc";
                object path;//設定一些object宣告
                object oMissing = System.Reflection.Missing.Value;
                object oSaveChanges = Word.WdSaveOptions.wdSaveChanges;
                object oformat = Word.WdSaveFormat.wdFormatDocument97;//wdFormatDocument97為Word 97-2003 文件 (*.doc)
                object start = 0, end = 0;
                if (word_app == null)//若無office程式則無法使用
                    MessageBox.Show("無法建立word檔案!!");
                else
                {
                    if (sf.ShowDialog() == DialogResult.OK)
                    {
                        path = sf.FileName;
                        word_app.Visible = false;//不顯示word程式
                        word_app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;//不顯示警告或彈跳視窗。如果出現彈跳視窗，將選擇預設值繼續執行。
                        word_document = word_app.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);//新增檔案

                        word_document.PageSetup.TopMargin = word_app.CentimetersToPoints(float.Parse("1.5"));
                        word_document.PageSetup.BottomMargin = word_app.CentimetersToPoints(float.Parse("1.5"));
                        word_document.PageSetup.LeftMargin = word_app.CentimetersToPoints(float.Parse("2"));
                        word_document.PageSetup.RightMargin = word_app.CentimetersToPoints(float.Parse("2"));

                        Word.Range rng = word_document.Range(ref start, ref end);


                        Word.Paragraph oPara1;
                        oPara1 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara1.Range.Text = "三寶尊佛前引過";
                        oPara1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        oPara1.Range.Font.Size = 24;
                        // oPara1.Range.Font.Bold = 1;
                        // oPara1.Format.SpaceAfter = 36;    //24 pt spacing after paragraph.//在段落之後 24 pt 空格
                        oPara1.Range.InsertParagraphAfter();

                        //Insert a paragraph at the end of the document.
                        // ' 在文件的尾端插入一個段落。            
                        Word.Paragraph oPara2;
                        object oRng = word_document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                        oPara2 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara2.Range.Text = textBox1.Text+"府              末切日期";
                        oPara2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        oPara2.Range.Font.Size = 14;
                        // oPara2.Format.SpaceAfter = 6;
                        oPara2.Range.InsertParagraphAfter();

                        //Insert another paragraph.
                        //插入另外一個段落。
                      
                        /*------------------------------------*/
                        string dd = dateTimePicker1.Text;
                        dd = dd.Replace("年", ",").Replace("月", ",").Replace("日", ",");
                        string[] ddd = { " ", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十", "三十一" };
                        string[] ss = { "", "秦廣王", "楚江王", "宋帝王", "伍官王", "閻羅王", "變成王", "泰山王", "平等王", "都市王", "輪轉王" };
                        string[] arrday = { "", "首七", "二七", "三七", "四七", "五七", "六七", "滿七", "百日", "對年" };
                        string[] family = { "", "（兒子）", "", "（女兒）", "", "（孫輩）", "", "（圓滿）", "" };
                        string[] week = { "", "一", "二", "三", "四", "五", "六","日" };
                        Dictionary<int, int> lsCalendar = new Dictionary<int, int>() { { 2020, 4 } };
                        string[] arr = dd.Split(",".ToCharArray());
                        int a = (int)decimal.Parse(arr[0]);
                        int b = (int)decimal.Parse(arr[1]);
                        int c = (int)decimal.Parse(arr[2]);


                        Word.Paragraph oPara3;

                        oRng = word_document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                        oPara3 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara3.Range.Text = "年國曆     月      日";
                        oPara3.Range.Font.Size = 14;
                        oPara3.Range.Font.Bold = 0;
                        // oPara3.Format.SpaceAfter = 6;
                        Word.Paragraph oPara10;
                        oPara10 = word_document.Content.Paragraphs.Add(ref oMissing);

                        oPara10.Range.Text =  (a-1911)+"  年國曆  " + ddd[b] + "月" + ddd[c] + "日";
                        oPara10.Range.Font.Size = 14;
                        oPara10.Range.Font.Bold = 0;
                        // oPara10.Format.SpaceAfter = 6;

                        DateTime dtObj = new DateTime(2018, 6, 11);

                        //以 月份 日, 年 的格式輸出
                        string outputDate = dtObj.ToString("MMMM dd, yyyy", new CultureInfo("zh-tw"));
                       // label7.Text = outputDate;
                        DateTime dt = new DateTime(a, b, c);
                        Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-tw");
                        System.Globalization.TaiwanCalendar TC = new System.Globalization.TaiwanCalendar();
                        System.Globalization.TaiwanLunisolarCalendar TA = new System.Globalization.TaiwanLunisolarCalendar();
                        //   label1.Text = string.Format("明國:{0}/{1}/{2} <br>", TC.GetYear(dt), TC.GetMonth(dt), TC.GetDayOfMonth(dt)) + ddd[TC.GetDayOfMonth(dt)];
                        //  label2.Text = string.Format("農歷:{0}/{1}/{2} <br>", TA.GetYear(dt), TA.GetMonth(dt), TA.GetDayOfMonth(dt));
                        string  Date_of_death = "農歷"+ddd[TA.GetMonth(dt)]+"月"+ ddd[TA.GetDayOfMonth(dt)] + "日";
                        //   label2.Text = string.Format("農歷:{0}<br>", TA.AddDays(dt, 365));
                        int count = 4;

                        for (int i = 1; i <= 7; i++)
                        {
                            int d = 7 * i;
                            int dn;
                            //  label3.Text += arrday[i] + " " + dt.AddDays(d - 1).ToString("MMMM dd, yyyy", new CultureInfo("zh-tw")) + Environment.NewLine;
                            /*
                            Word.Paragraph oPara4;
                            oRng = word_document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                            oPara4 = word_document.Content.Paragraphs.Add(ref oRng);
                            Word.Paragraph oPara5;
                            oRng = word_document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                            oPara5 = word_document.Content.Paragraphs.Add(ref oRng);
                            oPara4.Range.Text += ss[i] ;
                            oPara4.Range.Font.Superscript = 0;
                            oPara4.Range.Font.Size = 14;
                            oPara4.Range.Font.Bold = 0;*/
                            Word.Paragraph oPara = word_document.Content.Paragraphs.Add(ref oMissing);
                            if (i == 2)
                            {
                                dn = int.Parse(dt.AddDays((d / 2) - 1).ToString("dd"));
                              //  label3.Text += "地方風俗 " + arrday[i] + " " + dt.AddDays((d / 2) - 2).ToString("MM 月dd", new CultureInfo("zh-tw")) + "日" + Environment.NewLine;
                                oPara.Range.Text += "　　　地方風俗 " + "國曆  " + dt.AddDays((d / 2) - 2).ToString("MMMM", new CultureInfo("zh-tw")) +ddd[dn].PadLeft(3,'　')+ "日　上午  八時至十七時之間（"+week[int.Parse(TC.GetDayOfWeek(dt.AddDays((d / 2) - 2)).ToString("d"))]+"）";
                            }
                            dn = int.Parse(dt.AddDays(d - 1).ToString("dd"));
                            /*  oPara5.Range.Text += ss[i]+" "+arrday[i] + " " + dt.AddDays(d - 1).ToString("MMMM dd", new CultureInfo("zh-tw"))+ddd[dn] ;
                              oPara5.Range.Font.Superscript = 1;
                              oPara5.Range.Font.Size = 14;
                              oPara5.Range.Font.Bold = 0;
                              // oPara4.Format.SpaceAfter = 6;
                              // oPara4.Range.InsertParagraphAfter();*/

                            string s7 = dt.AddDays(d - 1).ToString("MMMM", new CultureInfo("zh-tw")) + ddd[dn].PadLeft(3,'　') + "日";
                            
                            oPara.Range.Text += ss[i] + "   " + arrday[i] + "  國曆  " +s7.PadRight(7,'　') + "上午  八時至十七時之間（"+week[(int)decimal.Parse(TC.GetDayOfWeek(dt.AddDays(d-1)).ToString("d"))]+"）"+family[i];
                            // oPara.Range.InsertParagraphAfter();
                            /*  object oStart = oPara.Range.Start + 4;
                              object oEnd = oPara.Range.Start + oPara.Range.Text.Length;
                              Word.Range rSuperscript = word_document.Range(ref oStart, ref oEnd);
                              rSuperscript.Font.Superscript = 1;*/
                            count += 1;
                        }
                        /*
                        Word.Paragraph oPara14;
                        oRng = word_document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                        oPara14 = word_document.Content.Paragraphs.Add(ref oRng);

                        oPara14.Range.Font.Bold = 0; // 0 為非粗體， 1 為粗體
                        oPara14.Range.Font.Superscript = 0;
                        oPara14.Range.Font.Name = "標楷體"; // 字型
                        oPara14.Range.Font.Size = 14; // 字體大小
                                                     // oPara4.Range.Font.Color = WdColor.wdColorLime; // 顏色
                                                     // oPara4.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // 置中

                        oPara14.Range.Text = "作七時刻均為概定　請自行擬定時刻 以上時間參考用";
                        oPara14.Range.InsertParagraphAfter();
                        oPara14.Range.Text = "一般做七『例如星期三逝世-每星期二為七』";
                        oPara14.Range.Font.Name = "標楷體";
                        oPara14.Range.Font.Superscript = 0;
                        oPara14.Range.InsertParagraphAfter();
                        oPara14.Range.Text = "民間風俗做法頭七可作前一日晚上21-23時『交時』";
                        oPara14.Range.Font.Name = "標楷體";
                        oPara14.Range.Font.Superscript = 0;
                        oPara14.Range.InsertParagraphAfter();
                        oPara14.Range.Text = "出殯後之做七-可於自宅.（靈骨安置地點或寺廟祭祀）";
                        oPara14.Range.Font.Name = "標楷體";
                        oPara14.Range.Font.Superscript = 0;
                        oPara14.Range.InsertParagraphAfter();
                       */
                        //百日
                        int dn99 = int.Parse(dt.AddDays(99).ToString("dd"));
                        //  dt = dt.AddDays(99);
                        string s99 = dt.AddDays(99).ToString("MMMM", new CultureInfo("zh-tw")) + ddd[dn99].PadLeft(3,'　') + "日";
                        Word.Paragraph oPara4 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara4.Range.Text += ss[8] + "   " + arrday[8] + "  國曆  " + s99.PadRight(7,'　')+ "上午  八時至十七時之間（"+week[int.Parse(TC.GetDayOfWeek(dt.AddDays(99)).ToString("d"))]+"）";
                        oPara4.Range.InsertParagraphAfter();
                        //對年
                        string result = string.Empty;
                        System.Globalization.TaiwanLunisolarCalendar tls = new System.Globalization.TaiwanLunisolarCalendar();
                        DateTime begin = tls.AddDays(DateTime.Now, 0);
                        Boolean leap = tls.IsLeapYear(TA.GetYear(dt.AddYears(1)));
                        DateTime dtt = tls.ToDateTime(TA.GetYear(dt), TA.GetMonth(dt), TA.GetDayOfMonth(dt), 0, 0, 0, 0);
                        
                        if (leap)
                        {
                            if (TA.GetMonth(dtt) > lsCalendar[(int)a])
                            {

                                dtt = tls.AddMonths(dtt, -1);
                            }
                            //label5.Text = "適逢" + TA.GetYear(dt.AddYears(1)) + "年閏月，因而農曆月份須提前一個月";
                            dtt = tls.AddMonths(dtt, -1);
                        }
                        dtt = tls.AddYears(dtt, 1);

                        TimeSpan tss = dtt - begin;
                        int day = tls.GetDayOfMonth(dtt);
                        int month = tls.GetMonth(dtt);
                        int year = tls.GetYear(dtt);

                       // label4.Text = string.Format("國歷{0}\n農曆{0}年{1}月{2}日", year, month, day, DateTime.Now.Add(tss).ToString("yyyy/MM/dd"));
                        int dn365 = int.Parse(DateTime.Now.Add(tss).ToString("dd"));
                        string s365 = string.Format("{0}", DateTime.Now.Add(tss).ToString("MMMM")) + ddd[dn365].PadLeft(3,'　') + "日";
                        Word.Paragraph oPara5 = word_document.Content.Paragraphs.Add(ref oMissing);
                        int year365 = int.Parse(DateTime.Now.Add(tss).ToString("yyyy"))-1911;
                        oPara5.Range.Text = ss[9] + "   " + arrday[9] + "  國曆  " + s365.PadRight(7,'　')+ "上午  八時至十七時之間（"+week[int.Parse(TC.GetDayOfWeek(DateTime.Now.Add(tss)).ToString("d"))]+"）"+"（"+ year365 + "）";
                        oPara5.Range.InsertParagraphAfter();

                        Word.Paragraph oPara6 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara6.Range.Text = "民間風俗-逝者（對年-足一年）不閏月-閏月-正常計算";
                        Word.Paragraph oPara7 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara6.Range.Text = "忌日祭祀均以農曆計算-"+ Date_of_death;
                        Word.Paragraph oPara3y = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara6.Range.Text = "輪轉王　 三年　國曆　　月　　　日　　　　　時至　　時";
                        oPara6.Range.InsertParagraphAfter();
                        Word.Paragraph oPara8 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara8.Range.Text = "拾殿冥王判超昇";
                        oPara8.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        oPara8.Range.Font.Size = 24;
                        Word.Paragraph oPara9 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara8.Range.Text = "作七時刻均為概定　請自行擬定時刻 以上時間參考用";
                        oPara8.Range.Font.Size = 14;
                        Word.Paragraph oPara11 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara8.Range.Text = "一般做七『例如星期三逝世-每星期二為七』";
                        oPara8.Range.Font.Size = 14;
                        Word.Paragraph oPara12 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara8.Range.Text = "民間風俗做法頭七可作前一日晚上21-23時『交時』";
                        oPara8.Range.Font.Size = 14;
                        Word.Paragraph oPara13 = word_document.Content.Paragraphs.Add(ref oMissing);
                        oPara8.Range.Text = "出殯後之做七-可於自宅.（靈骨安置地點或寺廟祭祀）";
                        oPara8.Range.Font.Size = 14;
                        //  label6.Text = string.Format("農歷:{0}/{1}/{2} <br>", TA.GetYear(dt), TA.GetMonth(dt), TA.GetDayOfMonth(dt));

                        /*    Word.Table table = word_document.Tables.Add(rng, 10, 6, ref oMissing, ref oMissing);//設定表格
                           table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;//內框線
                           table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;//外框線
                           table.Select();//選取指令
                           word_app.Selection.Font.Name = "標楷體";//設定選取的資料字型
                           word_app.Selection.Font.Size = 10;//設定文字大小
                           word_app.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter; //將其設為靠中間
                            for (int i = 1; i <= 10; i++)//將表格的資料寫入word檔案裡,第一列的值為1,第一行的值為1
                                 for (int j = 1; j <= 6; j++)
                                 {
                                     table.Cell(i, j).Range.Text = i + "," + j;
                                     table.Cell(i, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                     table.Columns[j].Width = 70f;
                                     table.Rows[i].Height = 70f;
                                 }*/

                        word_document.SaveAs(ref path, ref oformat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                        , ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);//存檔
                        word_document.Close(ref oMissing, ref oMissing, ref oMissing);//關閉
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(word_document);//釋放
                        word_document = null;
                        word_app.Quit(ref oMissing, ref oMissing, ref oMissing);//結束
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(word_app);//釋放
                        word_app = null;
                        MessageBox.Show("寫入檔案，儲存成功", " 提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    sf.Dispose();//釋放
                    sf = null;
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            label3.Text = "";
            /*test*/
            System.Globalization.CultureInfo tc = new System.Globalization.CultureInfo("zh-TW");
            //tc.DateTimeFormat.Calendar = new System.Globalization.TaiwanLunisolarCalendar();
            DateTime dt2 = DateTime.Now;
            //label5.Text = dt2.ToString(tc); ;
            /*test2*/
            CultureInfo m_ciTaiwan = new CultureInfo("zh-TW");
            m_ciTaiwan.DateTimeFormat.Calendar = m_ciTaiwan.OptionalCalendars[2];

            string strDate = DateTime.Now.Date.ToString("yyyyMMdd", m_ciTaiwan);

           // label4.Text=strDate;//output:1000512

            DateTime dtNow = DateTime.ParseExact(strDate.PadLeft(8, '0'), "yyyyMMdd", m_ciTaiwan);


            /*test3*/
           
            /**/
            string dd = dateTimePicker1.Text;
            dd = dd.Replace("年", ",").Replace("月", ",").Replace("日", ",");
            string[] ddd = { " ", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", "二十一", "二十二", "二十三", "二十四", "二十五", "二十六", "二十七", "二十八", "二十九", "三十", "三十一" };
            string[] ss = { "", "秦廣王", "楚江王", "宋帝王", "伍官王", "閻羅王", "變成王", "泰山王", "平等王", "都市王", "輪轉王" };
            string[] arrday = { "", "首七", "二七", "三七", "四七", "五七", "六七", "滿七", "百日", "對年" };
            string[] arr = dd.Split(",".ToCharArray());
            Dictionary<int,int> lsCalendar = new Dictionary<int,int>() { {2020,4 } }
            int a = (int)decimal.Parse(arr[0]);
            int b = (int)decimal.Parse(arr[1]);
            int c = (int)decimal.Parse(arr[2]);

            DateTime dtObj = new DateTime(2018, 6, 11);

            //以 月份 日, 年 的格式輸出
            string outputDate = dtObj.ToString("MMMM dd, yyyy", new CultureInfo("zh-tw"));
            //label7.Text = outputDate;
            DateTime dt = new DateTime(a, b, c);
            Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-tw");
            System.Globalization.TaiwanCalendar TC = new System.Globalization.TaiwanCalendar();
            System.Globalization.TaiwanLunisolarCalendar TA = new System.Globalization.TaiwanLunisolarCalendar();
            label1.Text = string.Format("民國:{0}/{1}/{2}", TC.GetYear(dt), TC.GetMonth(dt), TC.GetDayOfMonth(dt)) ;
            
            // label2.Text = string.Format("農歷:{0}<br>", TA.AddDays(dt, 99));
            //label4.Text = string.Format("民國:{0}/{1}/{2}",TA.GetYear(), )

            int count = 4;

            for (int i = 1; i <= 7; i++)
            {
                //做七
                int d = 7 * i;
                int dn;
                //  label3.Text += arrday[i] + " " + dt.AddDays(d - 1).ToString("MMMM dd, yyyy", new CultureInfo("zh-tw")) + Environment.NewLine;
                if (i == 2) {
                    dn = int.Parse(dt.AddDays((d/2) - 1).ToString("dd"));
                    label3.Text += "地方風俗 " +"      " + dt.AddDays((d/2) - 2).ToString("MM 月dd", new CultureInfo("zh-tw")) + "日" + Environment.NewLine;
                }
                dn = int.Parse(dt.AddDays(d - 1).ToString("dd"));
                label3.Text += ss[i].PadRight(4,' ') + " " + arrday[i] + " " + dt.AddDays(d - 1).ToString("MM 月dd", new CultureInfo("zh-tw")) +"日"+Environment.NewLine;
               
                count += 1;
            }
            //百日
            label6.Text = dt.AddDays(99).ToString("yyyy年MM月dd日 ", new CultureInfo("zh-tw"));
            //對年
            string result = string.Empty;
            System.Globalization.TaiwanLunisolarCalendar tls = new System.Globalization.TaiwanLunisolarCalendar();
            DateTime begin = tls.AddDays(DateTime.Now, 0);

            Boolean leap = tls.IsLeapYear(TA.GetYear(dt));
            DateTime dtt = tls.ToDateTime(TA.GetYear(dt), TA.GetMonth(dt), TA.GetDayOfMonth(dt), 0, 0, 0, 0);
            
            if (leap)
              {
                //label5.Text = "適逢"+ TA.GetYear(dt.AddYears(1) )+ "年閏月，因而農曆月份須提前一個月";
                System.Diagnostics.Debug.WriteLine(a);
                if (TA.GetMonth(dtt) > lsCalendar[(int)a])
                {
                    
                    dtt = tls.AddMonths(dtt,-1);
                }

            }
            label2.Text = string.Format("農曆:{0}/{1}/{2}", TA.GetYear(dtt), TA.GetMonth(dtt), TA.GetDayOfMonth(dtt));
            dtt = tls.AddYears(dtt, 1);

            TimeSpan tss = dtt - begin;
            int day = tls.GetDayOfMonth(dtt);
            int month = tls.GetMonth(dtt);
            int year = tls.GetYear(dtt);
           
            label4.Text = string.Format("國歷{3}\n農曆{0}年{1}月{2}日", year, month, day, DateTime.Now.Add(tss).ToString("yyyy/MM/dd"));
            

        }
    }
}
