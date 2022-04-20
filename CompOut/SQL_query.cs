using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompOut
{
    public class SQL_query
    {
        //строка подключения к базе данных
        public static string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename="+Form1.desktopPath+ @"\CompOutDB.mdf;  Connect Timeout=30";
        //SQL запросы к БД
        public static string sql = "SELECT * FROM division";
        public static string sql_div = "SELECT * FROM division WHERE div_id= {0}";
        public static string sql_div_comp = "SELECT comps.comp_name, comps.comp_ip, comps.comp_inv FROM comps WHERE comps.comp_div_id = {0}";
        public static string sql_all = "SELECT * FROM {0} WHERE {0}.{1}_id = {2}";
        public static string sql_master = "SELECT * FROM master";
        public static string sql_manuf = "SELECT * FROM manuf";
        public static string sql_comps_all = "SELECT * FROM comps";
        public static string sql_comps_delete = "DELETE comps WHERE comp_id = {0}";
        public static string sql_where = "SELECT * FROM division WHERE div_id = {0}";
        public static string sql_tables = "SELECT * FROM [{0}]";
        public static string sql_tables_where = "SELECT [{0}].*, manuf.manuf_name  FROM [{0}] JOIN manuf ON manuf.manuf_id = [{0}].{1}_manuf_id WHERE [{0}].{1}_id = {2}";
        public static string sql_tables_components = "SELECT [{0}].*, manuf.manuf_name  FROM [{0}] JOIN manuf ON manuf.manuf_id = [{0}].{1}_manuf_id";
        public static string sql_comps = "SELECT comp_id, comp_div_id, comp_name FROM comps WHERE comp_div_id ={0}";
        public static string Save_button = "INSERT INTO division (div_name, div_chief, div_phone, div_note) VALUES (N'{0}', N'{1}', N'{2}', N'{3}')";
        public static string Delete_button = "DELETE division WHERE div_id = {0}";
        public static string Delete_complect = "DELETE {0} WHERE {1}_id = {2}";
        public static string Delete_soft = "DELETE comp_soft WHERE soft_id = {0}";
        public static string Delete_service = "DELETE service WHERE service_id = {0}";
        public static string Delete_user = "DELETE [user] WHERE user_id = {0}";
        public static string Delete_usb = "DELETE [comp_usb] WHERE comp_usb_id = {0}";
        public static string Delete_other = "DELETE [comp_other] WHERE comp_other_id = {0}";
        public static string Delete_master = "DELETE master WHERE master_id = {0}";
        public static string Delete_manuf = "DELETE manuf WHERE manuf_id = {0}";
        public static string Edit_button = "UPDATE division SET div_name = N'{0}', div_chief = N'{1}', div_phone = N'{2}', div_note = N'{3}' WHERE div_id = {4}";
        public static string Comps_info = "SELECT * FROM comps WHERE comp_name = N'{0}'";
        public static string other_info = "SELECT comp_other.comp_other_id, other.other_model, manuf.manuf_name, other.other_note FROM other JOIN manuf ON other.other_manuf_id = manuf.manuf_id  JOIN comp_other ON comp_other.other_id = other.other_id JOIN comps ON comps.comp_Id= comp_other.comp_id WHERE comp_name = N'{0}'";
        public static string usb_info = "SELECT comp_usb.comp_usb_id, usb.usb_model, manuf.manuf_name, usb.usb_note FROM usb JOIN manuf ON usb.usb_manuf_id = manuf.manuf_id  JOIN comp_usb ON comp_usb.usb_id = usb.usb_id JOIN comps ON comps.comp_Id= comp_usb.comp_id WHERE comp_name = N'{0}'";
        public static string soft_info = "SELECT soft.soft_id, soft.soft_name, manuf.manuf_name, soft.soft_note FROM soft JOIN manuf ON soft.soft_manuf_id = manuf.manuf_id  JOIN comp_soft ON comp_soft.soft_id = soft.soft_id JOIN comps ON comps.comp_Id= comp_soft.comp_id WHERE comp_name = N'{0}'";
        public static string service_info = "SELECT service.service_id, master.master_name, service.service_date, service.service_note FROM  service JOIN master ON master.master_id = service.master_id JOIN comps ON comps.comp_Id = service.comp_id WHERE comps.comp_name = N'{0}'";
        public static string user_info = "SELECT [user].user_id, [user].user_name, [user].user_login, [user].user_pass, [user].user_note FROM [user] JOIN comps ON comps.comp_Id = [user].comp_id WHERE comps.comp_name = N'{0}'";
        public static string New_polzovatel_button = "INSERT INTO comps VALUES (55, {0}, N'{1}', N'{2}', N'{3}', N'{4}', {5}, N'{6}'";
        public static string New_comp_soft = "INSERT INTO comp_soft VALUES ({0}, {1})";
        public static string New_comp_serves = "INSERT INTO comp_soft VALUES ({0}, {1}, {2}, N'{3}')";
        public static string New_user = "INSERT INTO [user] VALUES ({0}, N'{1}', N'{2}', N'{3}', N'{4}')";
        public static string New_usb = "INSERT INTO [comp_usb] VALUES ({0}, {1})";
        public static string New_other = "INSERT INTO [comp_other] VALUES ({0}, {1})";
        public static string New_master = "INSERT INTO [master] VALUES (N'{0}', N'{1}')";
        public static string New_manuf = "INSERT INTO [manuf] VALUES (N'{0}', N'{1}')";



        //названия таблиц и таблиц подключаемых к компонентам формы, префиксы столбцов в БД
        public static List<string> Tables = new List<string>() {"division", "user", "cpu", "mb", "memory", "video", "sound", "cases", "hdd", "lan", "fdd", "cdr", "cdrw", "dvd", "display", "printer", "scaner", "modem", "key", "mouse", "ups", "soft",  "master",    "other",  "service",    "usb"};
        public static List<string> Tables2 = new List<string>() { "division", "user", "cpu", "cpu", "mb", "memory", "memory", "memory", "memory", "video", "sound", "cases", "hdd", "hdd", "hdd", "lan", "fdd", "cdr", "cdrw", "dvd", "display", "display", "printer", "printer", "printer", "printer", "scaner", "modem", "key", "mouse", "ups", "soft", "master", "other", "service", "usb" };
        public static List<string> Item = new List<string>()         { "div", "user", "cpu", "cpu", "mb", "memory", "memory", "memory", "memory", "video", "sound", "case", "hdd", "hdd", "hdd", "lan", "fdd", "cdr", "cdrw", "dvd", "display", "display", "printer", "printer", "printer", "printer", "scaner", "modem", "key", "mouse", "ups", "soft", "master", "other", "service", "usb" };
        public static List<string> Item2 = new List<string>() {"cpu", "mb", "memory", "video", "sound", "case", "hdd", "lan", "fdd", "cdr", "cdrw", "dvd", "display", "printer", "scaner", "modem", "key", "mouse", "ups", "soft", "master", "other", "service", "usb" };

    }

    
}
