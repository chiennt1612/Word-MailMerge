using Microsoft.AspNetCore.Http;
using System;
using System.IO;
using System.Web;
namespace Utils
{
    public enum LogTarget
    {
        API, Mail, SMS, Invoice, Search
    }

    public static class LogHelper
    {
        private static LogBase logger = null;

        public static byte[] FileReadAllBytes(string srcPath)
        {
            logger = new APILogger();
            return logger.FileReadAllBytes(srcPath);
        }
        public static string FileReadAllText(string srcPath)
        {
            logger = new APILogger();
            return logger.FileReadAllText(srcPath);
        }
        public static void Log(LogTarget target, string message)
        {
            switch (target)
            {
                case LogTarget.API:
                    logger = new APILogger();
                    logger.Log(message);
                    break;
                case LogTarget.Mail:
                    logger = new MailLogger();
                    logger.Log(message);
                    break;
                case LogTarget.SMS:
                    logger = new SMSLogger();
                    logger.Log(message);
                    break;
                case LogTarget.Invoice:
                    logger = new InvoiceLogger();
                    logger.Log(message);
                    break;
                case LogTarget.Search:
                    logger = new SearchLogger();
                    logger.Log(message);
                    break;
                default:
                    return;
            }
        }
    }

    public abstract class LogBase
    {
        protected readonly object lockObj = new object();
        public abstract void Log(string message);

        public string FileReadAllText(string srcPath)
        {
            if (File.Exists(srcPath))
                return File.ReadAllText(srcPath);
            else
                return "";
        }
        public byte[] FileReadAllBytes(string srcPath)
        {
            if (File.Exists(srcPath))
                return File.ReadAllBytes(srcPath);
            else
                return null;
        }
    }

    public class APILogger : LogBase
    {
        public string filePath = Directory.GetCurrentDirectory() + "\\LOG\\API";

        public override void Log(string message)
        {
            DateTime d = DateTime.Now;
            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            var fileName = filePath + "\\log_" + d.ToString("yyyyMMddHHmm").Substring(0, 11) + ".log";
            lock (lockObj)
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.Write("\r\nLog Entry {0} {1}: ", d.ToLongTimeString(), d.ToLongDateString());
                        sw.Write("\n" + message);
                        sw.Close();
                    }
                    fs.Close();
                }
            }
        }
    }

    public class MailLogger : LogBase
    {
        public string filePath = Directory.GetCurrentDirectory() + "\\LOG\\MAIL";
        public override void Log(string message)
        {
            DateTime d = DateTime.Now;
            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            var fileName = filePath + "\\log_" + d.ToString("yyyyMMddHHmm").Substring(0, 11) + ".log";
            lock (lockObj)
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.Write("\r\nLog Entry {0} {1}: ", d.ToLongTimeString(), d.ToLongDateString());
                        sw.Write("\n" + message);
                        sw.Close();
                    }
                    fs.Close();
                }
            }
        }
    }

    public class SMSLogger : LogBase
    {
        public string filePath = Directory.GetCurrentDirectory() + "\\LOG\\SMS";
        public override void Log(string message)
        {
            DateTime d = DateTime.Now;
            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            var fileName = filePath + "\\log_" + d.ToString("yyyyMMddHHmm").Substring(0, 11) + ".log";
            lock (lockObj)
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.Write("\r\nLog Entry {0} {1}: ", d.ToLongTimeString(), d.ToLongDateString());
                        sw.Write("\n" + message);
                        sw.Close();
                    }
                    fs.Close();
                }
            }
        }
    }

    public class InvoiceLogger : LogBase
    {
        public string filePath = Directory.GetCurrentDirectory() + "\\LOG\\Invoice";
        public override void Log(string message)
        {
            DateTime d = DateTime.Now;
            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            var fileName = filePath + "\\log_" + d.ToString("yyyyMMddHHmm").Substring(0, 11) + ".log";
            lock (lockObj)
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.Write("\r\nLog Entry {0} {1}: ", d.ToLongTimeString(), d.ToLongDateString());
                        sw.Write("\n" + message);
                        sw.Close();
                    }
                    fs.Close();
                }
            }
        }
    }

    public class SearchLogger : LogBase
    {
        public string filePath = Directory.GetCurrentDirectory() + "\\LOG\\Search";
        public override void Log(string message)
        {
            DateTime d = DateTime.Now;
            if (!Directory.Exists(filePath)) Directory.CreateDirectory(filePath);
            var fileName = filePath + "\\log_" + d.ToString("yyyyMMddHHmm").Substring(0, 11) + ".log";
            lock (lockObj)
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                    using (var sw = new StreamWriter(fs))
                    {
                        sw.Write("\r\nLog Entry {0} {1}: ", d.ToLongTimeString(), d.ToLongDateString());
                        sw.Write("\n" + message);
                        sw.Close();
                    }
                    fs.Close();
                }
            }
        }
    }

}