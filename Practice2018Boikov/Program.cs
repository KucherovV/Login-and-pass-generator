using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;
using Newtonsoft.Json;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Practice2018Boikov
{
    class Program
    {
        private static string filePath;
        static List<StudentMark> markList = new List<StudentMark>();
        static List<string> loginList = new List<string>();
        static List<string> loginAndPassList = new List<string>();

        static void Main(string[] args)
        {          
            Console.OutputEncoding = Encoding.Default;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WindowHeight = Console.LargestWindowHeight;
            Console.WindowWidth = Console.LargestWindowWidth;
            Menu();
            GenerateLoginsAndPass();
            SaveMunu();


        }

        private static void Menu()
        {
            Console.Clear();
            Console.WriteLine("Программа для чтения .xls-файлов.");
            Console.Write("Введите путь к файлу: ");

            filePath = Console.ReadLine();
            Console.Clear();

            Application ObjWorkEcxel;
            Workbook ObjWorkBook;
            Worksheet ObjWorkSheet;
            Range range;

            try
            {
                ObjWorkEcxel = new Application();
                ObjWorkBook = ObjWorkEcxel.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
                range = ObjWorkSheet.UsedRange;
                int rows = range.Rows.Count;

                int id;
                string lastName, nameUkr, groupNumber, shortName, namec, checkForm, name1, lastNameUkr1, nameUkr1, chairnumber, chairnumber1 ;

                //string

                for(int i = 2; i <= rows; i++)
                {
                    id = (int)(range.Cells[i, 1] as Range).Value;
                    lastName = (string)(range.Cells[i, 2] as Range).Value;
                    nameUkr = (string)(range.Cells[i, 3] as Range).Value;
                    groupNumber = (string)(range.Cells[i, 4] as Range).Value;
                    shortName = (string)(range.Cells[i, 5] as Range).Value;
                    namec = (string)(range.Cells[i, 6] as Range).Value;
                    checkForm = (string)(range.Cells[i, 7] as Range).Value;
                    name1 = (string)(range.Cells[i, 8] as Range).Value;
                    lastNameUkr1 = (string)(range.Cells[i, 9] as Range).Value;
                    nameUkr1 = (string)(range.Cells[i, 10] as Range).Value;
                    chairnumber = (string)(range.Cells[i, 11] as Range).Value;
                    chairnumber1 = (string)(range.Cells[i, 12] as Range).Value;

                    markList.Add(new StudentMark(id, lastName, nameUkr, groupNumber, shortName, namec, checkForm, name1,
                        lastNameUkr1, nameUkr1, chairnumber, chairnumber1));

                    if(lastName.Length > 11)
                        lastName = lastName.Remove(11);
                    else
                    {
                        lastName = lastName + new string(' ', 11 - lastName.Length);
                    }

                    if (nameUkr.Length > 11)
                        nameUkr = nameUkr.Remove(11);
                    else
                    {
                        nameUkr = nameUkr + new string(' ', 11 - nameUkr.Length);
                    }

                    if (groupNumber.Length > 8)
                        groupNumber = groupNumber.Remove(8);
                    else
                    {
                        groupNumber = groupNumber + new string(' ', 8 - groupNumber.Length);
                    }

                    if (shortName.Length > 40)
                        shortName = shortName.Remove(40);
                    else
                    {
                        shortName = shortName + new string(' ', 40 - shortName.Length);
                    }

                    if (namec.Length > 3)
                        namec = namec.Remove(3);
                    else
                    {
                        namec = namec + new string(' ', 3 - namec.Length);
                    }

                    if (namec.Length > 3)
                        namec = namec.Remove(3);
                    else
                    {
                        namec = namec + new string(' ', 3 - namec.Length);
                    }

                    if (checkForm.Length > 10)
                        checkForm = checkForm.Remove(10);
                    else
                    {
                        checkForm = checkForm + new string(' ', 10 - checkForm.Length);
                    }

                    if (name1.Length > 8)
                        name1 = name1.Remove(8);
                    else
                    {
                        name1 = name1 + new string(' ', 8 - name1.Length);
                    }

                    if (lastNameUkr1.Length > 11)
                        lastNameUkr1 = lastNameUkr1.Remove(11);
                    else
                    {
                        lastNameUkr1 = lastNameUkr1 + new string(' ', 11 - lastNameUkr1.Length);
                    }

                    if (nameUkr1.Length > 11)
                        nameUkr1 = nameUkr1.Remove(11);
                    else
                    {
                        nameUkr1 = nameUkr1 + new string(' ', 11 - nameUkr1.Length);
                    }

                    Console.WriteLine(id.ToString() + " | " + lastName + " | " + nameUkr + " | " + groupNumber + " | " + shortName + 
                        " | " + namec + " | " + checkForm + " | " + name1 + " | " + lastNameUkr1 + " | " + nameUkr1 + 
                        " | " + chairnumber + " | " + chairnumber1);
                }

                ObjWorkEcxel.Quit();
                GC.Collect();

                Console.WriteLine(new string('=', 150));
                Console.WriteLine(new string('=', 150));

            }
            catch (IOException)
            {
                Console.WriteLine("Файл не найден.");
                Console.ReadKey();
                Menu();
            }

            catch(System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Файл не найден.");
                Console.ReadKey();
                Menu();
            }
        }

        private static void GenerateLoginsAndPass()
        {
            string newLogin;
            string newPass;
            string prevLastName = "";
            string currName;
            

            foreach(var note in markList)
            {
                currName = note.last_name_ukr;
                if (prevLastName == currName)
                    continue;
                prevLastName = currName;

                newLogin = GenerateLogin(note.last_name_ukr, note.name_ukr);
                if (loginList.Contains(newLogin))
                    continue;

                Thread.Sleep(20);
                newPass = GeneratePass(note.last_name_ukr, note.name_ukr);

                string surnameName = note.last_name_ukr + " " + note.name_ukr;               

                if (newLogin.Length > 20)
                    newLogin = newLogin.Remove(20);
                else
                {
                    newLogin = newLogin + new string(' ', 20 - newLogin.Length);
                }

                if (newPass.Length > 20)
                    newPass = newPass.Remove(20);
                else
                {
                    newPass = newPass + new string(' ', 20 - newPass.Length);
                }

                if (surnameName.Length > 30)
                    surnameName = surnameName.Remove(30);
                else
                {
                    surnameName = surnameName + new string(' ', 30 - surnameName.Length);
                }
                loginAndPassList.Add(newLogin + " " + newPass + " " + surnameName);

                Console.WriteLine(newLogin + " | " + newPass + " | " + surnameName);
            }
        }

        private static string GenerateLogin(string lastName, string firstName)
        {
            Random rand = new Random();
            string randNumber = rand.Next(0, 999).ToString();

            string login = Transliteration.Front(lastName) + "_" +Transliteration.Front(firstName[0].ToString().ToUpper());
            login += randNumber;

            return login;
        }

        private static string GeneratePass(string lastName, string firstName)
        {
            const string valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890фбвгдеёжзийклмнросптфхцчшщьыэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЬЭЮЯ";
            Random rand = new Random();
            string pass = "";
            string finalPass = "";
            int randNumber;
            int charsCount = valid.Length;
            string chunkLName = lastName[0].ToString() + lastName[1].ToString() + lastName[2].ToString();
            chunkLName = Transliteration.Front(chunkLName);
            string chunkName = firstName[0].ToString() + firstName[1].ToString() + firstName[2].ToString();
            chunkName = Transliteration.Front(chunkName);

            for (int i = 0; i < 7; i++)
            {
                randNumber = rand.Next(1, 99999);
                pass = pass + valid[randNumber % charsCount].ToString();

            }

            randNumber = rand.Next(1, 99999);
            int stringBuildChoise = randNumber % 4;

            switch(stringBuildChoise)
            {
                case 0:
                    {
                        finalPass = chunkName + pass;
                    }
                    break;

                case 1:
                    {
                        finalPass = pass + chunkName;
                    }
                    break;

                case 2:
                    {
                        finalPass = chunkName + pass[0].ToString() + pass[1].ToString() + pass[2].ToString() + pass[3].ToString() + chunkLName;
                    }
                    break;

                case 3:
                    {
                        finalPass = chunkLName + pass[0].ToString() + pass[1].ToString() + pass[2].ToString() + pass[3].ToString() + chunkName;
                    }
                    break;
            }

            int capitalChoise;
            string chosenCapital;
            string finalFinalPass = "";

            for(int i = 0; i < finalPass.Length; i++)
            {
                if(Char.IsLetter(finalPass[i]))
                {
                    capitalChoise = rand.Next() % 2;

                    if(capitalChoise == 0)
                    {
                        finalFinalPass = finalFinalPass + finalPass[i].ToString().ToUpper();
                    }
                    else
                    {
                        finalFinalPass = finalFinalPass + finalPass[i].ToString().ToLower();
                    }
                    continue;
                }

                finalFinalPass = finalFinalPass + finalPass[i].ToString();


            }

            return finalFinalPass;
        }

        private static void SaveMunu()
        {
            Console.WriteLine(new string('=', 150));
            Console.WriteLine(new string('=', 150));
            Console.Write("Введите путь для сохранения результатов в txt-файл: ");
            string filePath = Console.ReadLine();

            try
            {
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                using (StreamWriter sw = new StreamWriter(fs, Encoding.Default))
                {
                    foreach (string str in loginAndPassList)
                    {
                        sw.WriteLine(str);
                    }
                }
            }
            catch(ArgumentException)
            {
                Console.WriteLine("Введите путь.");
            }

        }

    }
}


