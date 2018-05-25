using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Danhmuc27lvl
{
    class layfileoutlook
    {
        static string duongdangoc = Application.StartupPath + @"\filedanhmuc\";
        private static layfileoutlook _instance = null;
        public static layfileoutlook Instance()
        {
            if (_instance == null)
                _instance = new layfileoutlook();
            return _instance;
        }
        public string laydiachimail(Outlook.Account account)
        {
            try
            {
                if (string.IsNullOrEmpty(account.SmtpAddress) || string.IsNullOrEmpty(account.UserName))
                {
                    Outlook.AddressEntry oAE = account.CurrentUser.AddressEntry as Outlook.AddressEntry;
                    if (oAE.Type == "EX")
                    {
                        Outlook.ExchangeUser oEU = oAE.GetExchangeUser() as Outlook.ExchangeUser;
                        return oEU.PrimarySmtpAddress;
                    }
                    else
                    {
                        return oAE.Address;
                    }
                }
                else
                {
                    return account.SmtpAddress;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "";
            }
        }
        public Outlook.Folder GetFolder(string folderPath)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                String[] folders = folderPath.Split(backslash.ToCharArray());
                Outlook.Application Application = new Outlook.Application();
                folder = Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            return null;
                        }
                    }
                }
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        public void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    if (childFolder.FolderPath.Contains("Inbox"))
                    {
                        EnumerateFolders(childFolder);
                    }
                }
            }
             IterateMessages(folder);
        }
        public void IterateMessages(Outlook.Folder folder)
        {
            // attachment extensions to save
            string[] extensionsArray = {  ".xls" };
            string mau = "^Danh muc treo ban hang";
            var fi = folder.Items;
            if (fi != null)
            {
                try
                {
                    foreach (Object item in fi)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mi = (Outlook.MailItem)item;
                            if (mi != null && mi.UnRead == true)
                            {
                                var attachments = mi.Attachments;
                                
                                if (attachments.Count != 0)
                                {

                                    for (int i = 1; i <= mi.Attachments.Count; i++)
                                    {
                                        var fn = mi.Attachments[i].FileName.ToLower();
                                        if (extensionsArray.Any(fn.Contains))
                                        {
                                            if (Regex.IsMatch(mi.Attachments[i].FileName, mau))
                                            {

                                                if (!File.Exists(duongdangoc + mi.Attachments[i].FileName))
                                                {
                                                    mi.Attachments[i].SaveAsFile(duongdangoc + mi.Attachments[i].FileName);

                                                }
                                                else
                                                {
                                                    Console.WriteLine("Already saved " + mi.Attachments[i].FileName);

                                                }
                                            }
                                        }

                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("loi \n"+e.Message);
                }
            }
        }

        public void xuly()
        {
            Outlook.Application Application = new Outlook.Application();
            Outlook.Accounts accounts = Application.Session.Accounts;
            foreach (Outlook.Account taikhoan in accounts)
            {
                if (laydiachimail(taikhoan)=="nvhoang.hts@gmail.com")
                {
                    Outlook.Folder selectedFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
                    selectedFolder = GetFolder(@"\\" + taikhoan.DisplayName);
                    EnumerateFolders(selectedFolder);
                }
            }
        }
    }
}
