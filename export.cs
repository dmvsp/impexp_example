// Демо экспорт  
// Дмитрий Синельников
// https://github.com/sinelnikovdm
// 2020

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace import_client
{
    public class export
    {
        public void do_export(string fn)
        {
            ExchangeExportData ced = new ExchangeExportData();

            ced.from_center = 0;

            if (new utils().GetCENTER())
            {
                ced.from_center = 1;
            }

            string sql_text;

            sql_text = "select * from ORG_DOCS";

            try
            {
                ced.ORG_DOCS = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////
            sql_text = "select * from REBUILD_ORG";

            try
            {
                ced.REBUILD_ORG = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from ORGANIZATION";

            try
            {
                ced.ORGANIZATION = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            //////////////////////////////////////\
            sql_text = "select * from COUNTRY";

            try
            {
                ced.COUNTRY = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from DISTRICT";

            try
            {
                ced.DISTRICT = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from KIND_DOCS";

            try
            {
                ced.KIND_DOCS = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from PREF_TOWN";

            try
            {
                ced.PREF_TOWN = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from REGION";

            try
            {
                ced.REGION = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from SOURCE_ARCHIVE";

            try
            {
                ced.SOURCE_ARCHIVE = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from STAT_ORG";

            try
            {
                ced.STAT_ORG = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from TOWN";

            try
            {
                ced.TOWN = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from O_USER";

            try
            {
                ced.O_USER = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from ALL_ORG_DOCS";

            try
            {
                ced.ALL_ORG_DOCS = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from ALL_ORGANIZATION";

            try
            {
                ced.ALL_ORGANIZATION = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from ALL_REBUILD_ORG";

            try
            {
                ced.ALL_REBUILD_ORG = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);
            }

            //////////////////////////////////////\
            sql_text = "select * from ALL_TOWN";

            try
            {
                ced.ALL_TOWN = new work_with_oledb().ExecSQL_query(sql_text);
            }
            catch (Exception ex)
            {
                new work_with_oledb().show_err_exception(ex);

            }

            string tmp = SerializeString(ced);

            File.WriteAllText(fn, tmp);        
        }

        public  string SerializeString(ExchangeExportData input)
        {
            XmlSerializer serializer = new XmlSerializer(input.GetType());

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UnicodeEncoding(false, false); // no BOM in a .NET string
            settings.Indent = false;
            settings.OmitXmlDeclaration = false;

            using (StringWriter textWriter = new StringWriter())
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(textWriter, settings))
                {
                    serializer.Serialize(xmlWriter, input);
                }
                return textWriter.ToString();
            }
        }

        public  ExchangeExportData Deserialize(string fileName)
        {
            var stream = new StreamReader(fileName);
            var ser = new XmlSerializer(typeof(ExchangeExportData));
            object obj = ser.Deserialize(stream);
            stream.Close();
            return (ExchangeExportData)obj;
        }
    }

    [Serializable]
    public class ExchangeExportData
    {
        public int export_version = 1;
        public int from_center = 0;

        public DateTime export_data = DateTime.Now;

        public DataTable ORG_DOCS;
        public DataTable ORGANIZATION;
        public DataTable REBUILD_ORG;
        public DataTable COUNTRY;
        public DataTable DISTRICT;
        public DataTable KIND_DOCS;
        public DataTable PREF_TOWN;
        public DataTable REGION;
        public DataTable SOURCE_ARCHIVE;
        public DataTable STAT_ORG;
        public DataTable TOWN;
        public DataTable O_USER;
        public DataTable ALL_ORG_DOCS;
        public DataTable ALL_ORGANIZATION;
        public DataTable ALL_REBUILD_ORG;
        public DataTable ALL_TOWN;        
    }
}