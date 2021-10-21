using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;

namespace VigzkAcc_MYSQL
{
    class Program
    {
        static OleDbConnection connect = new OleDbConnection();
        static MySqlConnection ms;
        static void Main(string[] args)
        {
            while (true)
            {
                string path_mane_BD = @"\\10.0.8.204\Crown2021\crownBD.mdb";
                var builder = new MySqlConnectionStringBuilder
                {
                    Server = "127.0.0.1",
                    Database = "v",
                    UserID = "mysql",
                    Password = "mysql",
                    SslMode = MySqlSslMode.None,
                };
                ms = new MySqlConnection(builder.ConnectionString);
                ms.Open();
                connect.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_mane_BD + ";Persist Security Info=False;";
                connect.Open();
                var command = connect.CreateCommand();
                command.CommandText = "SELECT * FROM `humans`";
                var commands = ms.CreateCommand();
                commands.CommandText = "DELETE FROM `humans`";
                commands.ExecuteNonQuery();

                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        commands = ms.CreateCommand();
                        commands.CommandText = "INSERT INTO `humans` (`dnow`, `id`, `fam`, `im`, `ot`, `dr`, `pol`, `adres`, `phone`, `ORG`, `cel`, `DataP`, `DataV`, `resultCROWN`, `login`, `status`, `numPart`, `dedit`, `UNIK`, `prim`, `cito`, `polis`, `unicod`, `statusOMS`, `docsign`, `mysql`, `testsystems`, `level`, `workplace`, `DataZ`, `Cod`, `id_mazok`, `statusMIAC`, `DataD`, `DataDnow`, `vac_type`, `vac_data`, `vac_name`) VALUES ('" + reader[0] + "', NULL, '" + reader[2] + "', '" + reader[3] + "', '" + reader[4] + "', '" + reader[5] + "', '" + reader[6] + "', '" + reader[7].ToString().Replace('\'','"') + "', '" + reader[8] + "', '" + reader[9] + "', '" + reader[10] + "', '" + reader[11].ToString().Split(' ')[0] + "', '" + reader[12].ToString().Split(' ')[0] + "', '" + reader[13] + "', '" + reader[14] + "', '" + reader[15] + "', '" + reader[16] + "', '" + reader[17] + "', '" + reader[18] + "', '" + reader[19] + "', '" + reader[20] + "', '" + reader[21] + "', '" + reader[22] + "', '" + reader[23] + "', '" + reader[24] + "', '" + reader[25] + "', '" + reader[26] + "', '" + reader[27] + "', '" + reader[28] + "', '" + reader[29] + "', '" + reader[30] + "', '" + reader[31] + "', '" + reader[32] + "', '" + reader[33] + "', '" + reader[34] + "', '" + reader[35] + "', '" + reader[36] + "', '" + reader[37] + "')";
                        commands.ExecuteNonQuery();
                        Console.WriteLine("+");
                    }
                }
                connect.Close();
                ms.Close();
                Thread.Sleep(1000*60*60);
            }
        }
    }
}
