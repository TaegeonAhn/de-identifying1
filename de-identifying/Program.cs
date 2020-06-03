using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections;
using System.ComponentModel;

namespace MdbReader
{
    class Program
    {
        static List<string> errorFiles = new List<string>();

        static void Main(string[] args)
        {
            Console.WriteLine("A"); // A 출력 테스트

            Stopwatch stopWatch = new Stopwatch(); // 스탑워치

            stopWatch.Start();
            //string[] file_lists = Directory.GetFiles(@"D:\TEST\챠트프로", "*.mdb",SearchOption.AllDirectories);
            //Console.WriteLine("---MDB FILES---");
            //foreach (string name in file_lists)
            //{
            //    Console.WriteLine(name);
            //}


            string folderName = @"d:\챠트프로";
            string pathString = Path.Combine(folderName, "ExportFiles");
            Directory.CreateDirectory(pathString);  // 폴더 만들기

            try
            {
                var filesPathList = GetFiles();

                foreach (string filePath in filesPathList)
                {
                    var mdbPath = filePath;
                    var fileName = Path.GetFileName(mdbPath); //파일명 불러오기
                    var tables = GetTables(mdbPath); // 테이블명 불러오기
                                                     //var columns = GetColumns(mdbPath); // 컬럼명 불러오기


                    string[] readlines = File.ReadAllLines(@"D:\TEST\인명사전.txt");
                    //string[] FileLines = File.ReadAllLines("D:\TEST\인명사전.txt"); 동일
                    /*
                    foreach (DictionaryEntry entry in hash)
                    {
                        Console.WriteLine("{0}-{1}", entry.Key, entry.Value);
                    }*/


                    //System.IO.StreamReader file = new System.IO.StreamReader(@"D:\TEST\인명사전.txt");


                    //hash.Add("a", "de!identifying test");
                    /*
                    var sz = readlines.Length;
                    for (var i = 0; i < sz; ++i)
                    {
                        Console.WriteLine("{0} => {1}", i, readlines[i]);
                    }
                    */
                    var hash = new Hashtable(); // hash table 만들기
                    var i = 0;
                    foreach (string line in readlines)
                    {
                        hash.Add(i, line);
                        i += 1;
                    }                                      
                    PrintKeysAndValues(hash);
                    /*
                    foreach (DictionaryEntry de in hash)
                    {
                        Console.WriteLine("Key = {0}, Value ={1}", de.Key, de.Value);
                    }
                    */



                }

                foreach (var file in errorFiles)
                {
                    Console.WriteLine(file);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                Console.WriteLine("완료");
            }

            stopWatch.Stop();

            TimeSpan ts = stopWatch.Elapsed;
            Console.WriteLine("Elapsed Time is {0:00}:{1:00}:{2:00}",
                                ts.Hours, ts.Minutes, ts.Seconds);

            Console.ReadLine();
        }

        // public static bool isAccessAble(string path)

        public static string[] GetFiles()
        {
            string[] filesPathList = Directory.GetFiles(@"D:\TEST\챠트프로2", "*.mdb", SearchOption.AllDirectories);

            return filesPathList;
        }

        public static string[] GetTables(string mdbPath) //테이블명 얻기
        {
            try
            {
                var results = new List<string>();
                var connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbPath};User ID=admin;JET OLEDB:Database Password=;";
                var connection = new OleDbConnection(connString);
                using (connection)
                {
                    connection.Open();
                    var restrictions = new string[4];
                    restrictions[3] = "Table";
                    var tables = connection.GetSchema("Tables", restrictions); // 테이블명 
                    foreach (DataRow row in tables.Rows)
                    {
                        results.Add(row[2].ToString());
                    }
                }
                return results.ToArray();
            }
            catch (Exception e)
            {
                errorFiles.Add(mdbPath);
                Console.WriteLine(e);
                return new string[0];
            }
}
        /*
        public static string[] GetColumns(string mdbPath) // 컬럼명 얻기
        {
            try
            {
                var results = new List<string>();
                var connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbPath};User ID=admin;JET OLEDB:Database Password=;";
                var connection = new OleDbConnection(connString);
                using (connection)
                {
                    connection.Open();
                    var restrictions = new string[4];
                    restrictions[3] = "Column";
                    var tables = connection.GetSchema("Columns", restrictions); // 테이블명 
                    foreach (DataRow row in tables.Columns)
                    {
                        results.Add(row[2].ToString());
                    }
                }
                return results.ToArray();
            }
            catch (Exception e)
            {
                errorFiles.Add(mdbPath);
                Console.WriteLine(e);
                return new string[0];
            }
        }
        //{DateTime.Now.ToString("yyyyMMddHHmmss")}

        public static void UpdatingData(string mdbPath, string tableName, string pathString, string fileName, string columnName)
        {
            var connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbPath};User ID=admin;JET OLEDB:Database Password=;";
            var connection = new OleDbConnection(connString);
            using (connection)
            {
                var command = new OleDbCommand($"SELECT * FROM [{tableName}]", connection);
                connection.Open();
                using (var reader = command.ExecuteReader())
                using (writer)
                {
                    var Tablecolumns = new DataTable();
                    var encryptedColumnIndexSet = new HashSet<int>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        Tablecolumns.Columns.Add(reader.GetName(i));
                    }
                    writer.WriteLine(string.Join("|", Tablecolumns.Columns.Cast<DataColumn>().Select(column => column.ColumnName)));
                    while (reader.Read())
                    {
                        var data = "";
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            var tempData = reader[i].ToString();
                            if (i == 0)
                            {
                                data = tempData;
                            }
                            else
                            {
                                data += "|" + tempData;
                            }
                        }
                        writer.WriteLine(data);
                    }
                }
            }
        }
    
        private void HashTable()
        {
            Hashtable hash = new Hashtable();
            {
                { 1, "안태건"},
                { 2, "김개똥"},
                { 3, "황정원"},
                { 4, "이현준"}
            };

            foreach (var key in hash.Keys)
                Console.WriteLine("Key:{0}, Value:{1}", key, hash[key]);
            
            Console.WriteLine("*** All Values***");

            foreach (var value in hash.Values)
                Console.WriteLine("Value : {0} ", value);
 
        */
        public static void PrintKeysAndValues (Hashtable hash)
        {
            Console.WriteLine("\t-KEY-\t-VALUE-");
            foreach (DictionaryEntry de in hash)
                Console.WriteLine($"\t{de.Key}:\t{de.Value}");
            Console.WriteLine();
        }       
        
    }
}

