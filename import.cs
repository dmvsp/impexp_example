// Демо импорт  
// Дмитрий Синельников
// https://github.com/sinelnikovdm
// 2020

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace export_client
{
    public partial class import : Form
    {
        public import()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            textBox1.Text = openFileDialog1.FileName;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0)
            {
                MessageBox.Show("Выберите файл обмена данными для загрузки", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            enable_controls(false);

            try
            {
                var exp = new export();

                ExchangeExportData ced = exp.Deserialize(textBox1.Text);

                string file_info = string.Empty;

                /*
                       file_info = textBox1.Text + "\r\n";
                       file_info += "Версия формата: " + ced.export_version.ToString() + "\r\n";
                       file_info += "Обмен из центра: " + ced.from_center.ToString() + "\r\n";
                       file_info += "Дата выгрузки данных: " + ced.export_data + "\r\n";
                       file_info += "Источник данных: " + ced.SOURCE_ARCHIVE.Rows[0]["NAME_ARCHIVE"].ToString() + "\r\n";
                 */
                       
                if (ced.from_center == 0)
                {                    
                    do_export_fromlocal(ced);
                }
                else
                {                    
                    do_export_fromcenter(ced);
                }                
            }
            catch (Exception ex)
            {
                add_log(textBox1.Text + ex.Message);
                add_log("Произошла ошибка во время загрузки данных ! Данные файла обмена не загружены" + ex.Message);
            }

            enable_controls(true);
        }

        void do_export_fromcenter(ExchangeExportData ced)
        {                                                          
            string file_info = string.Empty;

            file_info = textBox1.Text + "\r\n";
            file_info += "Версия формата: " + ced.export_version.ToString() + "\r\n";
            file_info += "Обмен из центра: " + ced.from_center.ToString() + "\r\n";
            file_info += "Дата выгрузки данных: " + ced.export_data + "\r\n";
            file_info += "Источник данных: Центральная БД\r\n";

            List<string> l_ID_ARC = new List<string>();

            foreach (DataRow dr in ced.ALL_ORGANIZATION.Rows)
            {
                if (!l_ID_ARC.Contains(dr["ID_ARC"].ToString()))
                {
                    l_ID_ARC.Add(dr["ID_ARC"].ToString());
                }
            }

            string ID_ARC = string.Empty;

            foreach (var l in l_ID_ARC)
            {
                ID_ARC += l;
                ID_ARC += ",";
            }

            if (ID_ARC.Length > 0)
            {
                ID_ARC = ID_ARC.Substring(0, ID_ARC.Length - 1);
            }

            if (ID_ARC.Length == 0)
            {
                add_log("Значения ID_ARC в файле обмена не обнаружены");
                add_log("Файл обмена данными не загружен");

                return;
            }

            file_info += "Обнаруженные ID_ARC: " + ID_ARC+"\r\n";            

            string sql_text = "";
            DataTable dt = null;

            if (MessageBox.Show(file_info + "Загрузить файл обмена данными, полученный из центральной БД  ?",  Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            add_log(file_info);

            add_log("Обнаруженные ID_ARC: " + ID_ARC);
            add_log("Кол-во ALL_TOWN: " + ced.TOWN.Rows.Count.ToString() );
            add_log("Кол-во ALL_ORG_DOCS: " + ced.ORG_DOCS.Rows.Count.ToString());
            add_log("Кол-во ALL_ORGANIZATION: " + ced.ORGANIZATION.Rows.Count.ToString() );
            add_log("Кол-во ALL_REBUILD_ORG: " + ced.REBUILD_ORG.Rows.Count.ToString() );
            add_log("*****************************************************************");
            add_log("Загрузка ALL_TOWN");

            sql_text = "delete from ALL_TOWN where ID_ARC in ( :ID_ARC )";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ALL_TOWN.Rows)
            {
                sql_text = "insert into ALL_TOWN (ID_REC, ID_PREF_TOWN, ID_COUNTRY, ID_REGION, ID_DISTRICT, TOWN, DT_CHANGE, ID_ARC) " +
                            "values (:ID_REC, :ID_PREF_TOWN, :ID_COUNTRY, :ID_REGION, :ID_DISTRICT, :TOWN, :DT_CHANGE, :ID_ARC)";

                sql_text = sql_text.Replace(":ID_REC", dr["ID_REC"].ToString());
                sql_text = sql_text.Replace(":ID_PREF_TOWN", dr["ID_PREF_TOWN"].ToString().Length == 0 ? "NULL" :  dr["ID_PREF_TOWN"].ToString());
                sql_text = sql_text.Replace(":ID_COUNTRY", dr["ID_COUNTRY"].ToString().Length == 0 ? "NULL" : dr["ID_COUNTRY"].ToString());
                sql_text = sql_text.Replace(":ID_REGION", dr["ID_REGION"].ToString().Length == 0 ? "NULL" : dr["ID_REGION"].ToString());
                sql_text = sql_text.Replace(":ID_DISTRICT",dr["ID_DISTRICT"].ToString().Length == 0 ? "NULL" : dr["ID_DISTRICT"].ToString());
                sql_text = sql_text.Replace(":TOWN", dr["TOWN"].ToString().Length == 0 ? "NULL" : "'" + dr["TOWN"].ToString()  + "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":ID_ARC", dr["ID_ARC"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка ALL_ORG_DOCS");

            sql_text = "delete from ALL_ORG_DOCS where ID_ARC in ( :ID_ARC )";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ALL_ORG_DOCS.Rows)
            {
                sql_text = "insert into ALL_ORG_DOCS (ID_ORG, ID_DOCS, ID_CHILD_ORG, BEGIN_DATE, END_DATE, DT_CHANGE, DT_IMPORT, ADD_INFO, CHAR_DATE, ID_ARC, ID_REC) " +
                    "values (:ID_ORG, :ID_DOCS, :ID_CHILD_ORG, :BEGIN_DATE, :END_DATE, :DT_CHANGE, :DT_IMPORT, :ADD_INFO, :CHAR_DATE, :ID_ARC, :ID_REC)";

                sql_text = sql_text.Replace(":ID_ORG", dr["ID_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_DOCS", dr["ID_DOCS"].ToString().Length == 0 ? "NULL" : dr["ID_DOCS"].ToString());
                sql_text = sql_text.Replace(":ID_CHILD_ORG", dr["ID_CHILD_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_CHILD_ORG"].ToString());
                sql_text = sql_text.Replace(":BEGIN_DATE", dr["BEGIN_DATE"].ToString().Length == 0 ? "NULL" :"'" + dr["BEGIN_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":END_DATE", dr["END_DATE"].ToString().Length == 0 ? "NULL" : "'" + dr["END_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":DT_CHANGE",  "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT",  "GETDATE()");
                sql_text = sql_text.Replace(":ADD_INFO", dr["ADD_INFO"].ToString().Length == 0 ? "NULL" : "'" + dr["ADD_INFO"].ToString() + "'");
                sql_text = sql_text.Replace(":CHAR_DATE", dr["CHAR_DATE"].ToString().Length == 0 ? "NULL" : "'" + dr["CHAR_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":ID_ARC", dr["ID_ARC"].ToString());
                sql_text = sql_text.Replace(":ID_REC", dr["ID_REC"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка ALL_ORGANIZATION");

            sql_text = "delete from ALL_ORGANIZATION where ID_ARC in ( :ID_ARC )";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ALL_ORGANIZATION.Rows)
            {
                sql_text = "insert into ALL_ORGANIZATION (ID_STAT_ORG, ID_DISTRICT, ID_TOWN, ID_PREF_TOWN, FL_OBJECT_NAME, SH_OBJECT_NAME, POSTINDEX, STREET, HS_NUMB, CORPUS, FLAT, PHONE, EMAIL, LEADER, LEADER_POS, DT_CHANGE, DT_IMPORT, ORG_NOTE, FOND_NUMBER, ID_ARC, ID_REC, IS_DELETED) " +
                    "values (:ID_STAT_ORG, :ID_DISTRICT, :ID_TOWN, :ID_PREF_TOWN, :FL_OBJECT_NAME, :SH_OBJECT_NAME, :POSTINDEX, :STREET, :HS_NUMB, :CORPUS, :FLAT, :PHONE, :EMAIL, :LEADER, :LEADER_POS, :DT_CHANGE, :DT_IMPORT, :ORG_NOTE, :FOND_NUMBER, :ID_ARC, :ID_REC, :IS_DELETED)";

                sql_text = sql_text.Replace(":ID_STAT_ORG", dr["ID_STAT_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_STAT_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_DISTRICT", dr["ID_DISTRICT"].ToString().Length == 0 ? "NULL" : dr["ID_DISTRICT"].ToString());
                sql_text = sql_text.Replace(":ID_TOWN", dr["ID_TOWN"].ToString().Length == 0 ? "NULL" : dr["ID_TOWN"].ToString());
                sql_text = sql_text.Replace(":ID_PREF_TOWN", dr["ID_PREF_TOWN"].ToString().Length == 0 ? "NULL" : dr["ID_PREF_TOWN"].ToString());
                sql_text = sql_text.Replace(":FL_OBJECT_NAME", dr["FL_OBJECT_NAME"].ToString().Length == 0 ? "NULL" : "'" + dr["FL_OBJECT_NAME"].ToString() + "'");
                sql_text = sql_text.Replace(":SH_OBJECT_NAME", dr["SH_OBJECT_NAME"].ToString().Length == 0 ? "NULL" : "'" + dr["SH_OBJECT_NAME"].ToString() + "'");
                sql_text = sql_text.Replace(":POSTINDEX", dr["POSTINDEX"].ToString().Length == 0 ? "NULL" : "'" + dr["POSTINDEX"].ToString() + "'");
                sql_text = sql_text.Replace(":STREET", dr["STREET"].ToString().Length == 0 ? "NULL" : "'" + dr["STREET"].ToString() + "'");
                sql_text = sql_text.Replace(":HS_NUMB", dr["HS_NUMB"].ToString().Length == 0 ? "NULL" : "'" + dr["HS_NUMB"].ToString() + "'");
                sql_text = sql_text.Replace(":CORPUS", dr["CORPUS"].ToString().Length == 0 ? "NULL" : "'" + dr["CORPUS"].ToString() + "'");
                sql_text = sql_text.Replace(":FLAT", dr["FLAT"].ToString().Length == 0 ? "NULL" : "'" + dr["FLAT"].ToString() + "'");
                sql_text = sql_text.Replace(":PHONE", dr["PHONE"].ToString().Length == 0 ? "NULL" : "'" + dr["PHONE"].ToString() + "'");
                sql_text = sql_text.Replace(":EMAIL", dr["EMAIL"].ToString().Length == 0 ? "NULL" : "'" + dr["EMAIL"].ToString() + "'");
                sql_text = sql_text.Replace(":LEADER_POS", dr["LEADER_POS"].ToString().Length == 0 ? "NULL" : "'" + dr["LEADER_POS"].ToString() + "'");
                sql_text = sql_text.Replace(":LEADER", dr["LEADER"].ToString().Length == 0 ? "NULL" : "'" + dr["LEADER"].ToString() + "'");                
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT", "GETDATE()");
                sql_text = sql_text.Replace(":ORG_NOTE", dr["ORG_NOTE"].ToString().Length == 0 ? "NULL" : "'" + dr["ORG_NOTE"].ToString() + "'");
                sql_text = sql_text.Replace(":FOND_NUMBER", dr["FOND_NUMBER"].ToString().Length == 0 ? "NULL" : "'"+ dr["FOND_NUMBER"].ToString() + "'");
                sql_text = sql_text.Replace(":ID_ARC", dr["ID_ARC"].ToString());
                sql_text = sql_text.Replace(":ID_REC", dr["ID_REC"].ToString());
                sql_text = sql_text.Replace(":IS_DELETED", dr["IS_DELETED"].ToString().Length == 0 ? "NULL" : dr["IS_DELETED"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка ALL_REBUILD_ORG");

            sql_text = "delete from ALL_REBUILD_ORG where ID_ARC in ( :ID_ARC )";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ALL_REBUILD_ORG.Rows)
            {
                sql_text = "insert into ALL_REBUILD_ORG (ID_ORG, ID_REBUILD, DT_REBUILD, ADD_INFO, DT_CHANGE, DT_IMPORT, ID_ARC, ID_REC ) " +
                    "values (:ID_ORG, :ID_REBUILD, :DT_REBUILD, :ADD_INFO, :DT_CHANGE, :DT_IMPORT, :ID_ARC, :ID_REC)";

                sql_text = sql_text.Replace(":ID_ORG", dr["ID_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_REBUILD", dr["ID_REBUILD"].ToString().Length == 0 ? "NULL" : dr["ID_REBUILD"].ToString());
                sql_text = sql_text.Replace(":DT_REBUILD", dr["DT_REBUILD"].ToString().Length == 0 ? "NULL" : "'" + dr["DT_REBUILD"].ToString() + "'");
                sql_text = sql_text.Replace(":ADD_INFO", dr["ADD_INFO"].ToString().Length == 0 ? "NULL" : "'" + dr["ADD_INFO"].ToString()+ "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT", "GETDATE()");
                sql_text = sql_text.Replace(":ID_ARC", dr["ID_ARC"].ToString());
                sql_text = sql_text.Replace(":ID_REC", dr["ID_REC"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("*****************************************************************");
            add_log("Загрузка успешно завершена !");
        }

        void do_export_fromlocal(ExchangeExportData ced)
        {
            var exp = new export();
           
            ced = exp.Deserialize(textBox1.Text);

            string file_info = string.Empty;

            file_info = textBox1.Text + "\r\n";
            file_info += "Версия формата: " + ced.export_version.ToString() + "\r\n";
            file_info += "Дата выгрузки данных: " + ced.export_data + "\r\n";
            file_info += "Источник данных: " + ced.SOURCE_ARCHIVE.Rows[0]["NAME_ARCHIVE"].ToString() + "\r\n";
            file_info += "Код источника данных (ID_ARC): " + ced.SOURCE_ARCHIVE.Rows[0]["ID_ARC"].ToString() + "\r\n\r\n";

            string ID_ARC = ced.SOURCE_ARCHIVE.Rows[0]["ID_ARC"].ToString();

            string sql_text = "";
            DataTable dt = null;

            if (MessageBox.Show(file_info + "Загрузить файл обмена данными ?", Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }

            add_log(file_info);

            add_log("Кол-во TOWN: " + ced.TOWN.Rows.Count.ToString());
            add_log("Кол-во ORG_DOCS: " + ced.ORG_DOCS.Rows.Count.ToString());
            add_log("Кол-во ORGANIZATION: " + ced.ORGANIZATION.Rows.Count.ToString());
            add_log("Кол-во REBUILD_ORG: " + ced.REBUILD_ORG.Rows.Count.ToString());
            add_log("*****************************************************************");
            add_log("Загрузка TOWN");

            sql_text = "delete from ALL_TOWN where ID_ARC = :ID_ARC";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.TOWN.Rows)
            {
                sql_text = "insert into ALL_TOWN (ID_REC, ID_PREF_TOWN, ID_COUNTRY, ID_REGION, ID_DISTRICT, TOWN, DT_CHANGE, ID_ARC) " +
                            "values (:ID_REC, :ID_PREF_TOWN, :ID_COUNTRY, :ID_REGION, :ID_DISTRICT, :TOWN, :DT_CHANGE, :ID_ARC)";

                sql_text = sql_text.Replace(":ID_REC", dr["ID"].ToString());
                sql_text = sql_text.Replace(":ID_PREF_TOWN", dr["ID_PREF_TOWN"].ToString().Length == 0 ? "NULL" : dr["ID_PREF_TOWN"].ToString());
                sql_text = sql_text.Replace(":ID_COUNTRY", dr["ID_COUNTRY"].ToString().Length == 0 ? "NULL" : dr["ID_COUNTRY"].ToString());
                sql_text = sql_text.Replace(":ID_REGION", dr["ID_REGION"].ToString().Length == 0 ? "NULL" : dr["ID_REGION"].ToString());
                sql_text = sql_text.Replace(":ID_DISTRICT", dr["ID_DISTRICT"].ToString().Length == 0 ? "NULL" : dr["ID_DISTRICT"].ToString());
                sql_text = sql_text.Replace(":TOWN", dr["TOWN"].ToString().Length == 0 ? "NULL" : "'" + dr["TOWN"].ToString() + "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка ORG_DOCS");

            sql_text = "delete from ALL_ORG_DOCS where ID_ARC = :ID_ARC";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ORG_DOCS.Rows)
            {
                sql_text = "insert into ALL_ORG_DOCS (ID_ORG, ID_DOCS, ID_CHILD_ORG, BEGIN_DATE, END_DATE, DT_CHANGE, DT_IMPORT, ADD_INFO, CHAR_DATE, ID_ARC, ID_REC) " +
                    "values (:ID_ORG, :ID_DOCS, :ID_CHILD_ORG, :BEGIN_DATE, :END_DATE, :DT_CHANGE, :DT_IMPORT, :ADD_INFO, :CHAR_DATE, :ID_ARC, :ID_REC)";

                sql_text = sql_text.Replace(":ID_ORG", dr["ID_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_DOCS", dr["ID_DOCS"].ToString().Length == 0 ? "NULL" : dr["ID_DOCS"].ToString());
                sql_text = sql_text.Replace(":ID_CHILD_ORG", dr["ID_CHILD_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_CHILD_ORG"].ToString());
                sql_text = sql_text.Replace(":BEGIN_DATE", dr["BEGIN_DATE"].ToString().Length == 0 ? "NULL" : "'" + dr["BEGIN_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":END_DATE", dr["END_DATE"].ToString().Length == 0 ? "NULL" : "'" + dr["END_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT", "GETDATE()");
                sql_text = sql_text.Replace(":ADD_INFO", dr["ADD_INFO"].ToString().Length == 0 ? "NULL" : "'" + dr["ADD_INFO"].ToString() + "'");
                sql_text = sql_text.Replace(":CHAR_DATE", dr["CHAR_DATE"].ToString().Length == 0 ? "NULL" : "'" + dr["CHAR_DATE"].ToString() + "'");
                sql_text = sql_text.Replace(":ID_ARC", ID_ARC);
                sql_text = sql_text.Replace(":ID_REC", dr["ID"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка ORGANIZATION");

            sql_text = "delete from ALL_ORGANIZATION where ID_ARC = :ID_ARC";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.ORGANIZATION.Rows)
            {
                sql_text = "insert into ALL_ORGANIZATION (ID_STAT_ORG, ID_DISTRICT, ID_TOWN, ID_PREF_TOWN, FL_OBJECT_NAME, SH_OBJECT_NAME, POSTINDEX, STREET, HS_NUMB, CORPUS, FLAT, PHONE, EMAIL, LEADER, LEADER_POS, DT_CHANGE, DT_IMPORT, ORG_NOTE, FOND_NUMBER, ID_ARC, ID_REC, IS_DELETED) " +
                    "values (:ID_STAT_ORG, :ID_DISTRICT, :ID_TOWN, :ID_PREF_TOWN, :FL_OBJECT_NAME, :SH_OBJECT_NAME, :POSTINDEX, :STREET, :HS_NUMB, :CORPUS, :FLAT, :PHONE, :EMAIL, :LEADER, :LEADER_POS, :DT_CHANGE, :DT_IMPORT, :ORG_NOTE, :FOND_NUMBER, :ID_ARC, :ID_REC, :IS_DELETED)";

                sql_text = sql_text.Replace(":ID_STAT_ORG", dr["ID_STAT_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_STAT_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_DISTRICT", dr["ID_DISTRICT"].ToString().Length == 0 ? "NULL" : dr["ID_DISTRICT"].ToString());
                sql_text = sql_text.Replace(":ID_TOWN", dr["ID_TOWN"].ToString().Length == 0 ? "NULL" : dr["ID_TOWN"].ToString());
                sql_text = sql_text.Replace(":ID_PREF_TOWN", dr["ID_PREF_TOWN"].ToString().Length == 0 ? "NULL" : dr["ID_PREF_TOWN"].ToString());
                sql_text = sql_text.Replace(":FL_OBJECT_NAME", dr["FL_OBJECT_NAME"].ToString().Length == 0 ? "NULL" : "'" + dr["FL_OBJECT_NAME"].ToString() + "'");
                sql_text = sql_text.Replace(":SH_OBJECT_NAME", dr["SH_OBJECT_NAME"].ToString().Length == 0 ? "NULL" : "'" + dr["SH_OBJECT_NAME"].ToString() + "'");
                sql_text = sql_text.Replace(":POSTINDEX", dr["POSTINDEX"].ToString().Length == 0 ? "NULL" : "'" + dr["POSTINDEX"].ToString() + "'");
                sql_text = sql_text.Replace(":STREET", dr["STREET"].ToString().Length == 0 ? "NULL" : "'" + dr["STREET"].ToString() + "'");
                sql_text = sql_text.Replace(":HS_NUMB", dr["HS_NUMB"].ToString().Length == 0 ? "NULL" : "'" + dr["HS_NUMB"].ToString() + "'");
                sql_text = sql_text.Replace(":CORPUS", dr["CORPUS"].ToString().Length == 0 ? "NULL" : "'" + dr["CORPUS"].ToString() + "'");
                sql_text = sql_text.Replace(":FLAT", dr["FLAT"].ToString().Length == 0 ? "NULL" : "'" + dr["FLAT"].ToString() + "'");
                sql_text = sql_text.Replace(":PHONE", dr["PHONE"].ToString().Length == 0 ? "NULL" : "'" + dr["PHONE"].ToString() + "'");
                sql_text = sql_text.Replace(":EMAIL", dr["EMAIL"].ToString().Length == 0 ? "NULL" : "'" + dr["EMAIL"].ToString() + "'");
                sql_text = sql_text.Replace(":LEADER_POS", dr["LEADER_POS"].ToString().Length == 0 ? "NULL" : "'" + dr["LEADER_POS"].ToString() + "'");
                sql_text = sql_text.Replace(":LEADER", dr["LEADER"].ToString().Length == 0 ? "NULL" : "'" + dr["LEADER"].ToString() + "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT", "GETDATE()");
                sql_text = sql_text.Replace(":ORG_NOTE", dr["ORG_NOTE"].ToString().Length == 0 ? "NULL" : "'" + dr["ORG_NOTE"].ToString() + "'");
                sql_text = sql_text.Replace(":FOND_NUMBER", dr["FOND_NUMBER"].ToString().Length == 0 ? "NULL" : "'" + dr["FOND_NUMBER"].ToString() + "'");
                sql_text = sql_text.Replace(":ID_ARC", ID_ARC);
                sql_text = sql_text.Replace(":ID_REC", dr["ID"].ToString());
                sql_text = sql_text.Replace(":IS_DELETED", dr["IS_DELETED"].ToString().Length == 0 ? "NULL" : dr["IS_DELETED"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("Загрузка REBUILD_ORG");

            sql_text = "delete from ALL_REBUILD_ORG where ID_ARC = :ID_ARC";
            sql_text = sql_text.Replace(":ID_ARC", ID_ARC);

            new work_with_oledb().ExecSQL_nonquery(sql_text);

            foreach (DataRow dr in ced.REBUILD_ORG.Rows)
            {
                sql_text = "insert into ALL_REBUILD_ORG (ID_ORG, ID_REBUILD, DT_REBUILD, ADD_INFO, DT_CHANGE, DT_IMPORT, ID_ARC, ID_REC ) " +
                    "values (:ID_ORG, :ID_REBUILD, :DT_REBUILD, :ADD_INFO, :DT_CHANGE, :DT_IMPORT, :ID_ARC, :ID_REC)";

                sql_text = sql_text.Replace(":ID_ORG", dr["ID_ORG"].ToString().Length == 0 ? "NULL" : dr["ID_ORG"].ToString());
                sql_text = sql_text.Replace(":ID_REBUILD", dr["ID_REBUILD"].ToString().Length == 0 ? "NULL" : dr["ID_REBUILD"].ToString());
                sql_text = sql_text.Replace(":DT_REBUILD", dr["DT_REBUILD"].ToString().Length == 0 ? "NULL" : "'" + dr["DT_REBUILD"].ToString() + "'");
                sql_text = sql_text.Replace(":ADD_INFO", dr["ADD_INFO"].ToString().Length == 0 ? "NULL" : "'" + dr["ADD_INFO"].ToString() + "'");
                sql_text = sql_text.Replace(":DT_CHANGE", "GETDATE()");
                sql_text = sql_text.Replace(":DT_IMPORT", "GETDATE()");
                sql_text = sql_text.Replace(":ID_ARC", ID_ARC);
                sql_text = sql_text.Replace(":ID_REC", dr["ID"].ToString());

                new work_with_oledb().ExecSQL_nonquery(sql_text);
            }

            add_log("*****************************************************************");
            add_log("Загрузка успешно завершена !");
        }


        void add_log(string text)
        {
            textBox2.Text = textBox2.Text + "[" + DateTime.Now.ToShortDateString()+ " " + DateTime.Now.ToLongTimeString() + "] "+ text + "\r\n";

            textBox2.SelectionStart = textBox2.TextLength;
            textBox2.ScrollToCaret();

            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
            Application.DoEvents();
        }

        void enable_controls(bool enable)
        {
            button1.Enabled = enable;
            button2.Enabled = enable;
            button3.Enabled = enable;
            button4.Enabled = enable;
        }
    }
}