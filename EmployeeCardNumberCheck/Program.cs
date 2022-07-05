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
        static string ReceiveMailUser = @ConfigurationManager.AppSettings["Mail_TO"];
        static string CCEmailUser = @ConfigurationManager.AppSettings["Mail_CC"];
        static string SendEmailUser = @ConfigurationManager.AppSettings["Mail_Serder"];
        static string SendEmailAccount = @ConfigurationManager.AppSettings["Sender_Account"];
        static string SendEmailPassword = @ConfigurationManager.AppSettings["Sender_Password"];
        static string SMPTServer = @ConfigurationManager.AppSettings["SMPT_Server"];
        static int Port = Convert.ToInt32(@ConfigurationManager.AppSettings["SMPT_Port"]);
        static string Subject= @ConfigurationManager.AppSettings["Mail_Subject"];
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
        static List<string> Banbie = new List<string>();
        static List<string> DistinctCardNum = new List<string>();
        static List<string> DistinctEmpNo = new List<string>();

        static void Main(string[] args)
        {
            try
            {
                EmpNO.Clear(); Name.Clear(); FireDate.Clear(); CardNum.Clear(); Dep.Clear(); Title.Clear(); Area.Clear(); caterogy.Clear();Banbie.Clear();
                CopyFileSaveToFile();
                ReadCSV();
                DistintOrNot();
                SaveNewCSV();
                _logger.Info("排程執行完畢。");
                //Console.Read();
            }
            catch (Exception ex) 
            {
                _logger.Error("Main"+ex);
            }
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
                    else if (item == Path.Combine(DataFileFromDir, FileNameDeparment))
                    {
                        File.Copy(Path.Combine(DataFileFromDir, FileNameDeparment), Path.Combine(CopyFileSaveDir, "部門備份" + DateTime + ".csv"));
                        _logger.Info(dateTime.ToString("yyyyMMddHHmmss") + "部門.CSVCopy Success。");
                    }
                    else 
                    {
                        
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
                using (var sr = new System.IO.StreamReader(Path.Combine(DataFileFromDir, FileNameEmployee),System.Text.Encoding.UTF8))
                {
                    while (!sr.EndOfStream)
                    {
                       line = sr.ReadLine();
                        char[] delimiterChars = { ' ', ',', '\t' };
                        values = line.Split(delimiterChars);
                        EmpNO.Add(values[0]);
                        Name.Add(values[1]);
                        CardNum.Add(values[3]);
                        FireDate.Add(values[2]);
                       Dep.Add(values[4]);
                       Title.Add(values[5]);
                       Area.Add(values[6]);
                       caterogy.Add(values[7]);
                        Banbie.Add(values[8]);
                     }
                    _logger.Info("ReadCSV Success!");
                }
                File.Delete(Path.Combine(DataFileFromDir, FileNameEmployee));
            }
            catch (Exception ex)
            {
                _logger.Error(" CReadCSV" + ex);
            }
        }
        static void DistintOrNot() 
        {
            try
            {
                foreach (var item in CardNum.GroupBy(s => s))
                {
                    if (item.Count() > 1 && item.Key.ToString() != "")
                    {
                        for (int i = 0; i < CardNum.Count; i++) {
                            if (CardNum[i] == item.Key.ToString())
                            {
                                DistinctCardNum.Add(CardNum[i]);
                                DistinctEmpNo.Add(EmpNO[i]);
                                CardNum[i] = null;
                            }
                        }
                        string SendAllDistinctEmpNo = "";
                        foreach (string DistinctEmpNo in DistinctEmpNo)
                        {
                            SendAllDistinctEmpNo += DistinctEmpNo + "、";
                        }
                        SendAllDistinctEmpNo = SendAllDistinctEmpNo.Remove(SendAllDistinctEmpNo.Length - 1, 1);
                        SendAllDistinctEmpNo += "。";
                        string DistinctCardNumber = "";
                        foreach (string number in DistinctCardNum)
                        {
                            DistinctCardNumber = number;
                        }
                        SendAutomatedEmail(item.Key.ToString(), SendAllDistinctEmpNo);
                        _logger.Info("有重複卡號為" + item.Key.ToString() + "，員工工號為" + SendAllDistinctEmpNo + "已寄發信件");
                        DistinctEmpNo.Clear(); DistinctCardNum.Clear();
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
                    streamWriter.Write(EmpNO[i] + "," + Name[i] + "," + FireDate[i].ToString() + "," + CardNum[i] + "," + Dep[i] + "," + Title[i] + "," + Area[i] + "," + caterogy[i] + "," + Banbie[i] + "\r\n");
                }
                streamWriter.Flush();
                streamWriter.Close();
                _logger.Info("已完成另存檔案");
            }
            catch (Exception ex) 
            {
                _logger.Error("SaveNewCSV()" + ex);
            }
            
        }
         static void SendAutomatedEmail(string DistinctCardNumber,string EmpNo)
        {
            try
            {
                System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
                MyMail.From = new System.Net.Mail.MailAddress(SendEmailUser);

                string[] ToMail = ReceiveMailUser.Split(';');
                foreach (string temp in ToMail)
                {
                    string str = temp.Trim();
                    if (!string.IsNullOrEmpty(str))
                    {
                        MyMail.To.Add(str); 
                    }
                }
                string[] ToCC = CCEmailUser.Split(';');
                foreach (string temp in ToCC)
                {
                    string str = temp.Trim();
                    if (!string.IsNullOrEmpty(str))
                    {
                        MyMail.CC.Add(str);
                    }
                }   
                MyMail.SubjectEncoding = System.Text.Encoding.UTF8;
                MyMail.Subject = Subject +"卡號為" +DistinctCardNumber +"/工號為" +EmpNo;
                MyMail.BodyEncoding = System.Text.Encoding.UTF8;
                MyMail.Priority = MailPriority.High;
                MyMail.IsBodyHtml = true;
                MyMail.Body = "<body style=\"background-color:	#FFFCEC;padding=20px\"><br>系統排成檢查在貴司的員工.CSV檔案中有發現重複的卡號，<p><h3 style=\"color: red\">重複卡號為: " + DistinctCardNumber+"，<br> 員工工號為:"+EmpNo+"</h3>" +
                    " 已先將重複的卡號寫入空值，以利讓系統順利運作。<p> 再麻煩相關單位人員盡速處理。<br><p>此為系統發送，請勿回覆。<p></body>";
                SmtpClient client = new SmtpClient();
                System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient();
                MySMTP.Credentials = new System.Net.NetworkCredential(SendEmailAccount, SendEmailPassword);
                MySMTP.Host = SMPTServer; 
                MySMTP.Port = Port; 
                MySMTP.EnableSsl = true;
                MySMTP.Send(MyMail); 
                MySMTP.Dispose();
                MyMail.Dispose(); 
            }
            catch (Exception ex)
            {
                _logger.Error("SendAutomatedEmail" + ex);
            }
        }
    }
}

