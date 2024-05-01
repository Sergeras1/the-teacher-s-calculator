using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace Алгоритмизация
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            //поиск файла Excel
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Title = "Выберите документ Excel";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string xlFileName = ofd.FileName; //имя нашего Excel файла

            //рабоата с Excel
            Excel.Range Rng1;
            Excel.Range Rng2;
            Excel.Range Rng3;
            Excel.Range Rng4;
            Excel.Range Rng5;
            Excel.Range Rng6;
            Excel.Range Rng7;
            Excel.Range Rng8;
            Excel.Range Rng9;
            Excel.Range Rng10;
            Excel.Range Rng11;
            Excel.Range Rng12;
            Excel.Range Rng13;
            Excel.Range Rng14;
            Excel.Range Rng15;
            Excel.Range Rng16;
            Excel.Range Rng17;
            Excel.Range Rng18;
            Excel.Range Rng19;
            Excel.Range Rng20;
            Excel.Range Rng21;
            Excel.Range Rng22;
            Excel.Range Rng23;
            Excel.Range Rng24;
            Excel.Range Rng25;
            Excel.Range Rng26;
            Excel.Range Rng27;
            Excel.Range Rng28;
            Excel.Range Rng29;
            Excel.Range Rng30;
            Excel.Range Rng31;
            Excel.Range Rng32;
            Excel.Range Rng33;
            Excel.Range Rng34;
            Excel.Range Rng35;
            Excel.Range Rng36;
            Excel.Range Rng37;
            Excel.Range Rng38;
            Excel.Range Rng39;
            Excel.Range Rng40;
            Excel.Range Rpg;
            Excel.Range Rpg2;


            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;
            int iLastRow, iLastCol;

            Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
            xlWB = xlApp.Workbooks.Open(xlFileName); //открываем наш файл           
            xlSht = xlWB.Worksheets["Лист1"]; //или так xlSht = xlWB.ActiveSheet //активный лист

            //iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А
            //iLastCol = xlSht.Cells[1, xlSht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; //последний заполненный столбец в 1-й строке
            //Rng = (Excel.Range)xlSht.Range["A1", xlSht.Cells[iLastRow, iLastCol]]; //пример записи диапазона ячеек в переменную Rng

            Rng1 = xlSht.get_Range("C6:AE6"); //берём СТРОКУ в переменную Rng
            Rng2 = xlSht.get_Range("C7:AE7");
            Rng3 = xlSht.get_Range("C8:AE8");
            Rng4 = xlSht.get_Range("C9:AE9");
            Rng5 = xlSht.get_Range("C10:AE10");
            Rng6 = xlSht.get_Range("C11:AE11");
            Rng7 = xlSht.get_Range("C12:AE12");
            Rng8 = xlSht.get_Range("C13:AE13");
            Rng9 = xlSht.get_Range("C14:AE14");
            Rng10 = xlSht.get_Range("C15:AE15");
            Rng11 = xlSht.get_Range("C16:AE16");
            Rng12 = xlSht.get_Range("C17:AE17");
            Rng13 = xlSht.get_Range("C18:AE18");
            Rng14 = xlSht.get_Range("C19:AE19");
            Rng15 = xlSht.get_Range("C20:AE20");
            Rng16 = xlSht.get_Range("C21:AE21");
            Rng17 = xlSht.get_Range("C22:AE22");
            Rng18 = xlSht.get_Range("C23:AE23");
            Rng19 = xlSht.get_Range("C24:AE24");
            Rng20 = xlSht.get_Range("C25:AE25");
            Rng21 = xlSht.get_Range("C26:AE26");
            Rng22 = xlSht.get_Range("C27:AE27");
            Rng23 = xlSht.get_Range("C28:AE28");
            Rng24 = xlSht.get_Range("C29:AE29");
            Rng25 = xlSht.get_Range("C30:AE30");
            Rng26 = xlSht.get_Range("C31:AE31");
            Rng27 = xlSht.get_Range("C32:AE32");
            Rng28 = xlSht.get_Range("C33:AE33");
            Rng29 = xlSht.get_Range("C34:AE34");
            Rng30 = xlSht.get_Range("C35:AE35");
            Rng31 = xlSht.get_Range("C36:AE36");
            Rng32 = xlSht.get_Range("C37:AE37");
            Rng33 = xlSht.get_Range("C38:AE38");
            Rng34 = xlSht.get_Range("C39:AE39");
            Rng35 = xlSht.get_Range("C40:AE40");
            Rng36 = xlSht.get_Range("C41:AE41");
            Rng37 = xlSht.get_Range("C42:AE42");
            Rng38 = xlSht.get_Range("C43:AE43");
            Rng39 = xlSht.get_Range("C44:AE44");
            Rng40 = xlSht.get_Range("C45:AE45");
            Rpg = xlSht.get_Range("B6:B40");
            Rpg2 = xlSht.get_Range("C6:AE40");



            double SumCell_1 = 0;
            for (int y = 1; y < Rng1.Columns.Count; y++)
            {
                if (Rng1[y].Value != null)
                {
                    SumCell_1++;
                }
            }
            double SumCell_2 = 0;
            for (int y = 1; y < Rng2.Columns.Count; y++)
            {
                if (Rng2[y].Value != null)
                {
                    SumCell_2++;
                }
            }
            double SumCell_3 = 0;
            for (int y = 1; y < Rng3.Count; y++)
            {
                if (Rng3[y].Value != null)
                {
                    SumCell_3++;
                }
            }
            double SumCell_4 = 0;
            for (int y = 1; y < Rng4.Columns.Count; y++)
            {
                if (Rng4[y].Value != null)
                {
                    SumCell_4++;
                }
            }
            double SumCell_5 = 0;
            for (int y = 1; y < Rng5.Columns.Count; y++)
            {
                if (Rng5[y].Value != null)
                {
                    SumCell_5++;
                }
            }
            double SumCell_6 = 0;
            for (int y = 1; y < Rng6.Columns.Count; y++)
            {
                if (Rng6[y].Value != null)
                {
                    SumCell_6++;
                }
            }
            double SumCell_7 = 0;
            for (int y = 1; y < Rng7.Columns.Count; y++)
            {
                if (Rng7[y].Value != null)
                {
                    SumCell_7++;
                }
            }
            double SumCell_8 = 0;
            for (int y = 1; y < Rng8.Columns.Count; y++)
            {
                if (Rng8[y].Value != null)
                {
                    SumCell_8++;
                }
            }
            double SumCell_9 = 0;
            for (int y = 1; y < Rng9.Columns.Count; y++)
            {
                if (Rng9[y].Value != null)
                {
                    SumCell_9++;
                }
            }
            double SumCell_10 = 0;
            for (int y = 1; y < Rng10.Columns.Count; y++)
            {
                if (Rng10[y].Value != null)
                {
                    SumCell_10++;
                }
            }
            double SumCell_11 = 0;
            for (int y = 1; y < Rng11.Columns.Count; y++)
            {
                if (Rng11[y].Value != null)
                {
                    SumCell_11++;
                }
            }
            double SumCell_12 = 0;
            for (int y = 1; y < Rng12.Columns.Count; y++)
            {
                if (Rng12[y].Value != null)
                {
                    SumCell_12++;
                }
            }
            double SumCell_13 = 0;
            for (int y = 1; y < Rng13.Columns.Count; y++)
            {
                if (Rng13[y].Value != null)
                {
                    SumCell_13++;
                }
            }
            double SumCell_14 = 0;
            for (int y = 1; y < Rng14.Columns.Count; y++)
            {
                if (Rng14[y].Value != null)
                {
                    SumCell_14++;
                }
            }
            double SumCell_15 = 0;
            for (int y = 1; y < Rng15.Columns.Count; y++)
            {
                if (Rng15[y].Value != null)
                {
                    SumCell_15++;
                }
            }
            double SumCell_16 = 0;
            for (int y = 1; y < Rng16.Columns.Count; y++)
            {
                if (Rng16[y].Value != null)
                {
                    SumCell_16++;
                }
            }
            double SumCell_17 = 0;
            for (int y = 1; y < Rng17.Columns.Count; y++)
            {
                if (Rng17[y].Value != null)
                {
                    SumCell_17++;
                }
            }
            double SumCell_18 = 0;
            for (int y = 1; y < Rng18.Columns.Count; y++)
            {
                if (Rng18[y].Value != null)
                {
                    SumCell_18++;
                }
            }
            double SumCell_19 = 0;
            for (int y = 1; y < Rng19.Columns.Count; y++)
            {
                if (Rng19[y].Value != null)
                {
                    SumCell_19++;
                }
            }
            double SumCell_20 = 0;
            for (int y = 1; y < Rng20.Columns.Count; y++)
            {
                if (Rng20[y].Value != null)
                {
                    SumCell_20++;
                }
            }
            double SumCell_21 = 0;
            for (int y = 1; y < Rng21.Columns.Count; y++)
            {
                if (Rng21[y].Value != null)
                {
                    SumCell_21++;
                }
            }
            double SumCell_22 = 0;
            for (int y = 1; y < Rng22.Columns.Count; y++)
            {
                if (Rng22[y].Value != null)
                {
                    SumCell_22++;
                }
            }
            double SumCell_23 = 0;
            for (int y = 1; y < Rng23.Columns.Count; y++)
            {
                if (Rng23[y].Value != null)
                {
                    SumCell_23++;
                }
            }
            double SumCell_24 = 0;
            for (int y = 1; y < Rng24.Columns.Count; y++)
            {
                if (Rng24[y].Value != null)
                {
                    SumCell_24++;
                }
            }
            double SumCell_25 = 0;
            for (int y = 1; y < Rng25.Columns.Count; y++)
            {
                if (Rng25[y].Value != null)
                {
                    SumCell_25++;
                }
            }
            double SumCell_26 = 0;
            for (int y = 1; y < Rng26.Columns.Count; y++)
            {
                if (Rng26[y].Value != null)
                {
                    SumCell_26++;
                }
            }
            double SumCell_27 = 0;
            for (int y = 1; y < Rng27.Columns.Count; y++)
            {
                if (Rng27[y].Value != null)
                {
                    SumCell_27++;
                }
            }
            double SumCell_28 = 0;
            for (int y = 1; y < Rng28.Columns.Count; y++)
            {
                if (Rng28[y].Value != null)
                {
                    SumCell_28++;
                }
            }
            double SumCell_29 = 0;
            for (int y = 1; y < Rng29.Columns.Count; y++)
            {
                if (Rng29[y].Value != null)
                {
                    SumCell_29++;
                }
            }
            double SumCell_30 = 0;
            for (int y = 1; y < Rng30.Columns.Count; y++)
            {
                if (Rng30[y].Value != null)
                {
                    SumCell_30++;
                }
            }
            double SumCell_31 = 0;
            for (int y = 1; y < Rng31.Columns.Count; y++)
            {
                if (Rng31[y].Value != null)
                {
                    SumCell_31++;
                }
            }
            double SumCell_32 = 0;
            for (int y = 1; y < Rng32.Columns.Count; y++)
            {
                if (Rng32[y].Value != null)
                {
                    SumCell_32++;
                }
            }
            double SumCell_33 = 0;
            for (int y = 1; y < Rng33.Columns.Count; y++)
            {
                if (Rng33[y].Value != null)
                {
                    SumCell_33++;
                }
            }
            double SumCell_34 = 0;
            for (int y = 1; y < Rng34.Columns.Count; y++)
            {
                if (Rng34[y].Value != null)
                {
                    SumCell_34++;
                }
            }
            double SumCell_35 = 0;
            for (int y = 1; y < Rng35.Columns.Count; y++)
            {
                if (Rng35[y].Value != null)
                {
                    SumCell_35++;
                }
            }
            double SumCell_36 = 0;
            for (int y = 1; y < Rng36.Columns.Count; y++)
            {
                if (Rng36[y].Value != null)
                {
                    SumCell_36++;
                }
            }
            double SumCell_37 = 0;
            for (int y = 1; y < Rng37.Columns.Count; y++)
            {
                if (Rng37[y].Value != null)
                {
                    SumCell_37++;
                }
            }
            double SumCell_38 = 0;
            for (int y = 1; y < Rng38.Columns.Count; y++)
            {
                if (Rng38[y].Value != null)
                {
                    SumCell_38++;
                }
            }
            double SumCell_39 = 0;
            for (int y = 1; y < Rng39.Columns.Count; y++)
            {
                if (Rng39[y].Value != null)
                {
                    SumCell_39++;
                }
            }
            double SumCell_40 = 0;
            for (int y = 1; y < Rng40.Columns.Count; y++)
            {
                if (Rng40[y].Value != null)
                {
                    SumCell_40++;
                }
            }


            int SumRpg = 0;
            for (int i = 1; i < Rpg.Rows.Count; i++)
            {

                if (Rpg[i].Value != null)
                {
                    SumRpg++;
                }

            }



            double sum1 = xlApp.WorksheetFunction.Sum(Rng1); //вычисляем сумму ячеек
            double sum2 = xlApp.WorksheetFunction.Sum(Rng2);
            double sum3 = xlApp.WorksheetFunction.Sum(Rng3);
            double sum4 = xlApp.WorksheetFunction.Sum(Rng4);
            double sum5 = xlApp.WorksheetFunction.Sum(Rng5);
            double sum6 = xlApp.WorksheetFunction.Sum(Rng6);
            double sum7 = xlApp.WorksheetFunction.Sum(Rng7);
            double sum8 = xlApp.WorksheetFunction.Sum(Rng8);
            double sum9 = xlApp.WorksheetFunction.Sum(Rng9);
            double sum10 = xlApp.WorksheetFunction.Sum(Rng10);
            double sum11 = xlApp.WorksheetFunction.Sum(Rng11);
            double sum12 = xlApp.WorksheetFunction.Sum(Rng12);
            double sum13 = xlApp.WorksheetFunction.Sum(Rng13);
            double sum14 = xlApp.WorksheetFunction.Sum(Rng14);
            double sum15 = xlApp.WorksheetFunction.Sum(Rng15);
            double sum16 = xlApp.WorksheetFunction.Sum(Rng16);
            double sum17 = xlApp.WorksheetFunction.Sum(Rng17);
            double sum18 = xlApp.WorksheetFunction.Sum(Rng18);
            double sum19 = xlApp.WorksheetFunction.Sum(Rng19);
            double sum20 = xlApp.WorksheetFunction.Sum(Rng20);
            double sum21 = xlApp.WorksheetFunction.Sum(Rng21);
            double sum22 = xlApp.WorksheetFunction.Sum(Rng22);
            double sum23 = xlApp.WorksheetFunction.Sum(Rng23);
            double sum24 = xlApp.WorksheetFunction.Sum(Rng24);
            double sum25 = xlApp.WorksheetFunction.Sum(Rng25);
            double sum26 = xlApp.WorksheetFunction.Sum(Rng26);
            double sum27 = xlApp.WorksheetFunction.Sum(Rng27);
            double sum28 = xlApp.WorksheetFunction.Sum(Rng28);
            double sum29 = xlApp.WorksheetFunction.Sum(Rng29);
            double sum30 = xlApp.WorksheetFunction.Sum(Rng30);
            double sum31 = xlApp.WorksheetFunction.Sum(Rng31);
            double sum32 = xlApp.WorksheetFunction.Sum(Rng32);
            double sum33 = xlApp.WorksheetFunction.Sum(Rng33);
            double sum34 = xlApp.WorksheetFunction.Sum(Rng34);
            double sum35 = xlApp.WorksheetFunction.Sum(Rng35);
            double sum36 = xlApp.WorksheetFunction.Sum(Rng36);
            double sum37 = xlApp.WorksheetFunction.Sum(Rng37);
            double sum38 = xlApp.WorksheetFunction.Sum(Rng38);
            double sum39 = xlApp.WorksheetFunction.Sum(Rng39);
            double sum40 = xlApp.WorksheetFunction.Sum(Rng40);


            double sr1 = sum1 / SumCell_1;
            double sr2 = sum2 / SumCell_2;
            double sr3 = sum3 / SumCell_3;
            double sr4 = sum4 / SumCell_4;
            double sr5 = sum5 / SumCell_5;
            double sr6 = sum6 / SumCell_6;
            double sr7 = sum7 / SumCell_7;
            double sr8 = sum8 / SumCell_8;
            double sr9 = sum9 / SumCell_9;
            double sr10 = sum10 / SumCell_10;
            double sr11 = sum11 / SumCell_11;
            double sr12 = sum12 / SumCell_12;
            double sr13 = sum13 / SumCell_13;
            double sr14 = sum14 / SumCell_14;
            double sr15 = sum15 / SumCell_15;
            double sr16 = sum16 / SumCell_16;
            double sr17 = sum17 / SumCell_17;
            double sr18 = sum18 / SumCell_18;
            double sr19 = sum19 / SumCell_19;
            double sr20 = sum20 / SumCell_20;
            double sr21 = sum21 / SumCell_21;
            double sr22 = sum22 / SumCell_22;
            double sr23 = sum23 / SumCell_23;
            double sr24 = sum24 / SumCell_24;
            double sr25 = sum25 / SumCell_25;
            double sr26 = sum26 / SumCell_26;
            double sr27 = sum27 / SumCell_27;
            double sr28 = sum28 / SumCell_28;
            double sr29 = sum29 / SumCell_29;
            double sr30 = sum30 / SumCell_30;
            double sr31 = sum31 / SumCell_31;
            double sr32 = sum32 / SumCell_32;
            double sr33 = sum33 / SumCell_33;
            double sr34 = sum34 / SumCell_34;
            double sr35 = sum35 / SumCell_35;
            double sr36 = sum36 / SumCell_36;
            double sr37 = sum37 / SumCell_37;
            double sr38 = sum38 / SumCell_38;
            double sr39 = sum39 / SumCell_39;
            double sr40 = sum40 / SumCell_40;


            string a1 = xlSht.Range["B6"].Value; //берём значение ячейки в переменную
            string a2 = xlSht.Range["B7"].Value;
            string a3 = xlSht.Range["B8"].Value;
            string a4 = xlSht.Range["B9"].Value;
            string a5 = xlSht.Range["B10"].Value;
            string a6 = xlSht.Range["B11"].Value;
            string a7 = xlSht.Range["B12"].Value;
            string a8 = xlSht.Range["B13"].Value;
            string a9 = xlSht.Range["B14"].Value;
            string a10 = xlSht.Range["B15"].Value;
            string a11 = xlSht.Range["B16"].Value;
            string a12 = xlSht.Range["B17"].Value;
            string a13 = xlSht.Range["B18"].Value;
            string a14 = xlSht.Range["B19"].Value;
            string a15 = xlSht.Range["B20"].Value;
            string a16 = xlSht.Range["B21"].Value;
            string a17 = xlSht.Range["B22"].Value;
            string a18 = xlSht.Range["B23"].Value;
            string a19 = xlSht.Range["B24"].Value;
            string a20 = xlSht.Range["B25"].Value;
            string a21 = xlSht.Range["B26"].Value; //берём значение ячейки в переменную
            string a22 = xlSht.Range["B27"].Value;
            string a23 = xlSht.Range["B28"].Value;
            string a24 = xlSht.Range["B29"].Value;
            string a25 = xlSht.Range["B30"].Value;
            string a26 = xlSht.Range["B31"].Value;
            string a27 = xlSht.Range["B32"].Value;
            string a28 = xlSht.Range["B33"].Value;
            string a29 = xlSht.Range["B34"].Value;
            string a30 = xlSht.Range["B35"].Value;
            string a31 = xlSht.Range["B36"].Value; //берём значение ячейки в переменную
            string a32 = xlSht.Range["B37"].Value;
            string a33 = xlSht.Range["B38"].Value;
            string a34 = xlSht.Range["B39"].Value;
            string a35 = xlSht.Range["B40"].Value;
            /*string a36 = xlSht.Range["B41"].Value;
            string a37 = xlSht.Range["B42"].Value;
            string a38 = xlSht.Range["B43"].Value;
            string a39 = xlSht.Range["B44"].Value;
            string a40 = xlSht.Range["B45"].Value;*/


            listBox1.Items.Add(a1 + sr1);
            listBox1.Items.Add(a2 + sr2);
            listBox1.Items.Add(a3 + sr3);
            listBox1.Items.Add(a4 + sr4);
            listBox1.Items.Add(a5 + sr5);
            listBox1.Items.Add(a6 + sr6);
            listBox1.Items.Add(a7 + sr7);
            listBox1.Items.Add(a8 + sr8);
            listBox1.Items.Add(a9 + sr9);
            listBox1.Items.Add(a10 + sr10);
            listBox1.Items.Add(a11 + sr11);
            listBox1.Items.Add(a12 + sr12);
            listBox1.Items.Add(a13 + sr13);
            listBox1.Items.Add(a14 + sr14);
            listBox1.Items.Add(a15 + sr15);
            listBox1.Items.Add(a16 + sr16);
            listBox1.Items.Add(a17 + sr17);
            listBox1.Items.Add(a18 + sr18);
            listBox1.Items.Add(a19 + sr19);
            listBox1.Items.Add(a20 + sr20);
            listBox1.Items.Add(a21 + sr21);
            listBox1.Items.Add(a22 + sr22);
            listBox1.Items.Add(a23 + sr23);
            listBox1.Items.Add(a24 + sr24);
            listBox1.Items.Add(a25 + sr25);
            listBox1.Items.Add(a26 + sr26);
            listBox1.Items.Add(a27 + sr27);
            listBox1.Items.Add(a28 + sr28);
            listBox1.Items.Add(a29 + sr29);
            listBox1.Items.Add(a30 + sr30);
            listBox1.Items.Add(a31 + sr31);
            listBox1.Items.Add(a32 + sr32);
            listBox1.Items.Add(a33 + sr33);
            listBox1.Items.Add(a34 + sr34);
            listBox1.Items.Add(a35 + sr35);
            /* listBox1.Items.Add(a36 + sr36);
             listBox1.Items.Add(a37 + sr37);
             listBox1.Items.Add(a38 + sr38);
             listBox1.Items.Add(a39 + sr39);
             listBox1.Items.Add(a40 + sr40);*/


            double Summa1 = (sr1 + sr2 + sr3 + sr4 + sr5 + sr6 + sr7 + sr8 + sr9 +
                sr10 + sr11 + sr12 + sr13 + sr14 + sr15 + sr16 + sr17 + sr19 + sr19 +
                sr20 + sr21 + sr22 + sr23 + sr24 + sr25 + sr26 + sr27 + sr28 + sr29 + sr30 +
                sr31 + sr32 + sr33 + sr34 + sr35) / SumRpg;


            double students;
            students = Convert.ToDouble(textBox4.Text);

            textBox1.Text = Summa1.ToString();




            double countSymbols = 0;
            for (int y = 1; y < Rpg2.Count; y++)
            {
                if (Rpg2[y].Value != null && Rpg2[y].Value != 2)
                {
                    countSymbols++;
                }
            }
            double abcolt = countSymbols * 100 / students;

            textBox2.Text = abcolt.ToString();



            double countSymbols2 = 0;
            for (int y = 1; y < Rpg2.Count; y++)
            {
                if (Rpg2[y].Value != null && Rpg2[y].Value != 2 && Rpg2[y].Value != 3)
                {
                    countSymbols2++;
                }
            }
           double tb3 = countSymbols2 / SumRpg;

            textBox3.Text = tb3.ToString();






            //закрытие Excel
            xlWB.Close(true); //сохраняем и закрываем файл
            xlApp.Quit();
            releaseObject(xlSht);
            releaseObject(xlWB);
            releaseObject(xlApp);


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

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
