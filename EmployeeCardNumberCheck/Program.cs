using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace EmployeeCardNumberCheck
{
    class Program
    {
        static Logger _logger = LogManager.GetCurrentClassLogger();
        static string FileNameEmployee = "員工.csv";
        static string FileNameDeparment = "部門.csv";
        static DateTime dateTime = System.DateTime.Now;
        static string DateTime = dateTime.ToString("yyyyMMddhhMMss");
        static string DataFileFromDir = @ConfigurationManager.AppSettings["DataFileFromDir"];
        static string CopyFileSaveDir = @ConfigurationManager.AppSettings["CopyFileSaveDir"];
        static string ReceiveMailUser = @ConfigurationManager.AppSettings["ReceiveMailUser"];
        static string CCEmailUser = @ConfigurationManager.AppSettings["CCEmailUser"];
        static string SendEmailUser = @ConfigurationManager.AppSettings["SendEmailUser"];
        static string SendEmailAccount = @ConfigurationManager.AppSettings["SendEmailAccount"];
        static string SendEmailPassword = @ConfigurationManager.AppSettings["SendEmailPassword"];
        static string SMPTServer = @ConfigurationManager.AppSettings["SMPTServer"];
        static int Port = Convert.ToInt32(@ConfigurationManager.AppSettings["Port"]);
        static string Subject= @ConfigurationManager.AppSettings["Subject"];
        static string EmailBody = @ConfigurationManager.AppSettings["EmailBody"];
        static string line="";
        static string[] values ;
        static List<string> CardNum = new List<string>();
        static List<string> EmpNO = new List<string>();
        static List<string> Name = new List<string>();
        static List<string> FireDate = new List<string>();
        static List<string> Dep = new List<string>();
        static List<string> Area = new List<string>();
        static List<string> Title = new List<string>();
        static List<string> caterogy = new List<string>();
        static List<string> DistinctCardNum = new List<string>();
        static List<string> DistinctEmpNo = new List<string>();


        static void Main(string[] args)
        {
            EmpNO.Clear();Name.Clear();FireDate.Clear();CardNum.Clear();Dep.Clear();Title.Clear();Area.Clear();caterogy.Clear();
            CopyFileSaveToFile();
            ReadCSV();
            DistintOrNot();
            SaveNewCSV();
            _logger.Info(dateTime.ToString("yyyyMMddHHmm")+"排程執行完畢。");
            Console.Read();
        }

        static void CopyFileSaveToFile()
        {
            try
            {
                string[] EmployeeList = Directory.GetFiles(DataFileFromDir, "*.csv");

                foreach (string item in EmployeeList)
                {
                    if (item == Path.Combine(DataFileFromDir, FileNameEmployee))
                    {
                        File.Copy(Path.Combine(DataFileFromDir, FileNameEmployee), Path.Combine(CopyFileSaveDir, "員工備份" + DateTime + ".csv"));
                        _logger.Info(dateTime.ToString("yyyyMMddHHmmss") + "員工.CSVCopy Success。");
                    }
                    else
                    {
                        File.Copy(Path.Combine(DataFileFromDir, FileNameDeparment), Path.Combine(CopyFileSaveDir, "部門備份" + DateTime + ".csv"));
                        _logger.Info(dateTime.ToString("yyyyMMddHHmmss") + "部門.CSVCopy Success。");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(" CopyFileSaveToFile Error" + ex);
                Console.WriteLine(" CopyFileSaveToFile Error" + ex);
            }
        }
        static void ReadCSV()
        {
            try
            {
                using (var sr = new System.IO.StreamReader(Path.Combine(DataFileFromDir, FileNameEmployee)))
                {
                    while (!sr.EndOfStream)
                    {
                       line = sr.ReadLine();
                       values = line.Split(',');
                        EmpNO.Add(values[0]);
                        Name.Add(values[1]);
                        CardNum.Add(values[3]);
                        FireDate.Add(values[2]);
                       Dep.Add(values[4]);
                       Title.Add(values[5]);
                       Area.Add(values[6]);
                       caterogy.Add(values[7]);
                     }
                    _logger.Info("ReadCSV Success!");
                }
                File.Delete(Path.Combine(DataFileFromDir, FileNameEmployee));
            }
            catch (Exception ex)
            {
                _logger.Error(" CReadCSV" + ex);
                Console.WriteLine(" ReadCSV" + ex);
            }
        }
        static void DistintOrNot() 
        {
            try
            {
                foreach (var item in CardNum.GroupBy(s => s))
                {
                    _logger.Info(dateTime.ToString("yyyyMMddHHmmss")+"{0}:{1}次", item.Key, item.Count());
                    if (item.Count()>1) 
                    {
                        for (int i=0; i<CardNum.Count;i++) {
                            if (CardNum[i] == item.Key.ToString() && CardNum[i] != "NULL")
                            {
                                DistinctCardNum.Add(CardNum[i]);
                                DistinctEmpNo.Add(EmpNO[i]);
                                CardNum[i] = "NULL";
                                Console.WriteLine("已經重複卡號改為NULL");
                            }    
                        }
                        string SendAllDistinctEmpNo = "";
                        foreach (string DistinctEmpNo in DistinctEmpNo) 
                        {
                            Console.WriteLine(DistinctEmpNo);
                            SendAllDistinctEmpNo += DistinctEmpNo+"、";
                        }
                        SendAllDistinctEmpNo=SendAllDistinctEmpNo.Remove(SendAllDistinctEmpNo.Length-1,1);
                        SendAllDistinctEmpNo += "。";
                        string DistinctCardNumber = "";
                        foreach (string number in DistinctCardNum)
                        {
                            Console.WriteLine(number);
                            DistinctCardNumber += number ;
                        }
                        SendAutomatedEmail(ReceiveMailUser,item.Key.ToString(),SendAllDistinctEmpNo);
                        Console.WriteLine("有重複卡號"+ DistinctCardNumber +"，員工工號為"+SendAllDistinctEmpNo+"已寄發信件");
                        _logger.Info("有重複卡號" + item.Key.ToString() + "，員工工號為" + SendAllDistinctEmpNo + "已寄發信件");
                        DistinctEmpNo.Clear();DistinctCardNum.Clear();
                    }
                }

            } catch (Exception ex)
            {
                _logger.Error(" DistintOrNo" + ex);
                Console.WriteLine(" DistintOrNo" + ex);
            }
        }
        static void SaveNewCSV() 
        {
            try
            {
                FileStream FS = new FileStream(Path.Combine(DataFileFromDir, FileNameEmployee), FileMode.OpenOrCreate);
                StreamWriter streamWriter = new StreamWriter(FS, Encoding.UTF8);
                for (int i = 0; i < EmpNO.Count; i++)
                {
                    Console.WriteLine(EmpNO[i]+ "," + Name[i] + "," + FireDate[i].ToString() + "," + CardNum[i] + "," + Dep[i] + "," + Title[i] + "," + Area[i] + "," + caterogy[i] + "\r\n");
                    streamWriter.Write(EmpNO[i] + "," + Name[i] + "," + FireDate[i].ToString() + "," + CardNum[i] + "," + Dep[i] + "," + Title[i] + "," + Area[i] + "," + caterogy[i] + "\r\n");
                }
                streamWriter.Flush();
                streamWriter.Close();

                Console.WriteLine("已完成另存檔案");
            }
            catch (Exception ex) 
            {
                Console.WriteLine("SaveNewCSV()"+ex);
                _logger.Info("SaveNewCSV()" + ex);
            }
            
        }
         static void SendAutomatedEmail(string ReceiveMail,string DistinctCardNumber,string EmpNo)
        {
            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress(SendEmailUser);
            MyMail.To.Add(ReceiveMail); //設定收件者Email
            MyMail.Bcc.Add(CCEmailUser); //加入密件副本的Mail          
            MyMail.SubjectEncoding = System.Text.Encoding.UTF8;
            MyMail.Subject = Subject+"卡號"+DistinctCardNumber+"工號為"+EmpNo;
            MyMail.BodyEncoding = System.Text.Encoding.UTF8;
            MyMail.Body = EmailBody; //設定信件內容
            MyMail.IsBodyHtml = true; //是否使用html格式
            
            SmtpClient client = new SmtpClient();
            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient();
            MySMTP.Credentials = new System.Net.NetworkCredential(SendEmailAccount, SendEmailPassword);
            MySMTP.Host = SMPTServer; //設定smtp Server
            MySMTP.Port =Port; //設定Port
            MySMTP.EnableSsl = true;
            MySMTP.Send(MyMail); //寄出信件
            MySMTP.Dispose();
            
            try
            {
                MySMTP.Send(MyMail);
                MyMail.Dispose(); //釋放資源
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }
    }
}

