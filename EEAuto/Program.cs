using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using System.Net;
using System.Threading;
using OfficeTagLib;
using System.Diagnostics;
using CCtool;
using System.Configuration;
using System.Text.RegularExpressions;

namespace EEAuto
{
    class Program
    {

        /// <summary>
        /// Certificate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="certificate"></param>
        /// <param name="chain"></param>
        /// <param name="sslPolicyErrors"></param>
        /// <returns></returns>
        private static bool CertificateValidationCallBack(
         object sender,
         System.Security.Cryptography.X509Certificates.X509Certificate certificate,
         System.Security.Cryptography.X509Certificates.X509Chain chain,
         System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
            
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

		public string currentPath = System.IO.Directory.GetCurrentDirectory();
        public string TestCase = @"C:\EE\Test Case.xml";
        //public string TestCase = @"C:\EE\failed__2020-9-21-2-39-36.xml";
        public string Configfile = @"C:\EE\EE.xml";
        public int ThinkTime = 0;
        public int Wait = 0; //waiting for email
        public string URL = null;
        public string Domain = null;
        public string FQDN = null;
        public string Password = null;


        //get exchange parameters for a certain Domain
        public bool Initialize(string Domain)
        {
            XmlNode root;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(Configfile);
            root = xmlDoc.DocumentElement;

            //read exchange info from EE.xml file
            if (root.HasChildNodes) //root is EE
            {
                    foreach (XmlNode node in root.ChildNodes)
                    {
                        if ((node.Name.ToLower() + ".com" == Domain.ToLower() )||node.Name.ToLower() == Domain.ToLower())
                        {
                            URL = node.SelectSingleNode("//" + node.Name + "//URL").InnerText;
                            FQDN = node.SelectSingleNode("//" + node.Name + "//FQDN").InnerText;
                            Password = node.SelectSingleNode("//" + node.Name + "//Password").InnerText;
                        }
                        if (node.Name.ToLower() == "qapf1" && (Domain.ToLower() == "qapf1.qalab01.nextlabs.com" || Domain.ToLower() == "qapf1"))
                        {
                            URL = node.SelectSingleNode("//" + node.Name +"//URL").InnerText;
                            FQDN = node.SelectSingleNode("//" + node.Name + "//FQDN").InnerText;
                            Password = node.SelectSingleNode("//" + node.Name + "//Password").InnerText;
                        }
                    }
            }
            return true;

        }

        /// <summary>
        /// Get test case parameters, sender, recipients, attachments
        /// </summary>
        /// <param name="caseNO"></param>
        /// <returns>return CaseList, it is dictionary, key is caseID, value is case parameters. and value is a dictionary as well </returns>
        public Dictionary<string, Dictionary<string, string>> GetCase()
        {
            Dictionary<string, Dictionary<string, string>> CaseList = new Dictionary<string, Dictionary<string, string>>();
            XmlNode root;
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(TestCase);
                root = xmlDoc.DocumentElement;
                if (root.HasChildNodes) //root is EE-Auto
                {
                    foreach (XmlNode caseID in root.ChildNodes)
                    {
                        Dictionary<string, string> myCase = new Dictionary<string, string>();
                        if (caseID.HasChildNodes) //caseID
                        {
                            foreach (XmlNode subNode in caseID.ChildNodes)
                            {
                                myCase.Add(subNode.Name.ToString(), subNode.InnerText.ToString()); //case parameters
                            }
                        }
                        CaseList.Add(caseID.Name.ToString(), myCase);
                    }
                    
                }
            }
            catch(Exception e)
            {
                Console.Write(e.ToString());
            }
            Console.WriteLine(CaseList);
            return CaseList;
        }

        public List<string> ClearAllEmails(Dictionary<string, string> CasePar)
        {
            

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            //exchange2019 update the support api transfer protocol version, so need point out the support version is TLS1.2
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ExchangeService servicePremise = new ExchangeService(ExchangeVersion.Exchange2013);
            List<string> Tos = new List<string>();
            List<string> CCs = new List<string>();
            List<string> BCCs = new List<string>();

            /**
           try
           {              
               Tos = CasePar["To"].Split(';').ToList();
               CCs = CasePar["CC"].Split(';').ToList();
               BCCs = CasePar["BCC"].Split(';').ToList();
           }
           catch
           { }
           **/
            try
            {
                Tos = CasePar["To"].Split(';').ToList();
            }
            catch
            { }

            try
            {
                CCs = CasePar["CC"].Split(';').ToList();
            }
            catch
            { }

            try
            {
                BCCs = CasePar["BCC"].Split(';').ToList();
            }
            catch
            { }
            
                List<String> UserList = new List<string>();

                UserList = Tos.Union(CCs).ToList<string>();
                UserList = UserList.Union(BCCs).ToList<string>();
                UserList.Add(CasePar["Sender"]);
            
            UserList.Remove("");

            foreach(string user in UserList)
            {
                Initialize(user.Split('@')[1].ToLower());

                servicePremise.Credentials = new NetworkCredential(user.Split('@')[0], Password, user.Split('@')[1]);
                servicePremise.Url = new Uri(URL);

                ItemView view = new ItemView(int.MaxValue);
                FindItemsResults<Item> findResults = servicePremise.FindItems(WellKnownFolderName.Inbox, SetFilter(), view);

                foreach (Item item in findResults)
                {
                    Dictionary<string, object> receivedEmail = new Dictionary<string, object>();
                    item.Load(); //have to load email before reading.
                    item.Delete(DeleteMode.HardDelete);
                    //item.Move(WellKnownFolderName.DeletedItems);

                }
            }
            return UserList;

        }

        public void SendEmail(Dictionary<string, string> CasePar)
        {
            Initialize(CasePar["Sender"].Split('@')[1].ToLower());

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService servicePremise = new ExchangeService(ExchangeVersion.Exchange2013);
            servicePremise.Credentials = new NetworkCredential(CasePar["Sender"].Split('@')[0], Password, CasePar["Sender"].Split('@')[1]);
            servicePremise.Url = new Uri(URL);
            List<string> Tos = new List<string>();
            List<string> CCs = new List<string>();
            List<string> BCCs = new List<string>();
            List<string> Attachments = new List<string>();

            if (CasePar.ContainsKey("To"))
            {
                if (CasePar["To"].Contains(';'))
                    Tos = CasePar["To"].Split(';').ToList();
                else
                    Tos.Add(CasePar["To"]);
            }
            if (CasePar.ContainsKey("CC"))
            {
                if (CasePar["CC"].Contains(';'))
                    CCs = CasePar["CC"].Split(';').ToList();
                else
                    CCs.Add(CasePar["CC"]);
            }
            if (CasePar.ContainsKey("BCC"))
            {
                if (CasePar["BCC"].Contains(';'))
                    BCCs = CasePar["BCC"].Split(';').ToList();
                else
                    BCCs.Add(CasePar["BCC"]);
            }
            if (CasePar.ContainsKey("Attachment"))
            {
                if (CasePar["Attachment"].Contains(';'))
                    Attachments = CasePar["Attachment"].Split(';').ToList();
                else
                    Attachments.Add(CasePar["Attachment"]);
            }
            

            EmailMessage msg = new EmailMessage(servicePremise); //new email
            // Add x-header for email
            /*
             * ExtendedPropertyDefinition eExperimentalHeader_1 = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "ITAR", MapiPropertyType.String);
            msg.SetExtendedProperty(eExperimentalHeader_1, "Yes");

            ExtendedPropertyDefinition eExperimentalHeader_2 = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.InternetHeaders, "EAR", MapiPropertyType.String);
            msg.SetExtendedProperty(eExperimentalHeader_2, "Yes");
            */

            msg.Subject = CasePar["Subject"];
            msg.Body = new MessageBody(CasePar["Body"]);

                for (int i = 0; i < Tos.Count; i++)
                {
                    if(Tos[i].Length > 0)
                        msg.ToRecipients.Add(Tos[i]); //To recipients

                }
           
                for (int i = 0; i < CCs.Count; i++)
                {
                    if (CCs[i].Length > 0)
                        msg.CcRecipients.Add(CCs[i]);//CC recipients

                }
            
                for (int i = 0; i < BCCs.Count; i++)
                {
                    if(BCCs[i].Length >0)
                        msg.BccRecipients.Add(BCCs[i]);//BCC recipients

                }

                if (Attachments.Count > 0)
                {
                    foreach (string attachment in Attachments)
                    {
                        if (attachment != "")
                        {
                            msg.Attachments.AddFileAttachment(attachment);//attach files
                        }
                    }
                }
                
                Thread.Sleep(ThinkTime);
                //msg.SendAndSaveCopy();
                msg.SendAndSaveCopy(WellKnownFolderName.SentItems);
        }


        /// <summary>
        /// get all expected result from test case, who will receive email, and email content
        /// </summary>
        /// <param name="caseID">caseID</param>
        /// <returns>ExpectedResult is a dictiionary like this: {[recipient1, {[result, allow], [To, XXXX], [CC,XXXX], [BCC, XXXX],......}], [recipient2, {[result, allow], [To, XXXX], [CC, XXXX], [BCC, XXXX]....}]}</returns>
        public Dictionary<string, Dictionary<string, string>> GetExpectedResult(string caseID)
        {
            Dictionary<string, Dictionary<string, string>> ExpectedResult = new Dictionary<string, Dictionary<string, string>>();
            XmlNode root;
            
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(TestCase);
                root = xmlDoc.DocumentElement;
                root = xmlDoc.SelectSingleNode("//"+ caseID+ "/Assertion");  //root is caseID/assertion.
                if (root.HasChildNodes) //root is is caseID/assertion
                {
                    foreach (XmlNode recipient in root.ChildNodes) //recipient is "domain-XXX" this node
                    {
                        Dictionary<string, string> myResult = new Dictionary<string, string>();
                        if (recipient.HasChildNodes) //recipient is "domain-XXX" this node
                        {
                           
            
                            foreach (XmlNode subNode in recipient.ChildNodes) //subNode is email infos
                            {
                                myResult.Add(subNode.Name.ToString(), subNode.InnerText.ToString());
                            }

                            ExpectedResult.Add(recipient.Name, myResult);
                        }

                        //Console.WriteLine(recipient);
                    }
                }
            }
            catch (Exception e)
            {
                Console.Write(e.ToString());
            }
            return ExpectedResult;
 
        }

        public Dictionary<string, string> ReadTag(string filePath)
        {
            Dictionary<string, string> tags = new Dictionary<string, string>();
            TagOpera tagOpera = new TagOpera();
            tagOpera.SetOfficeFilePath(filePath);
            tagOpera.ExecuteReadCustomTag();
            for (int i = 0; i < tagOpera.GetTagKeyCount(); i++)
            { 
                tags.Add(tagOpera.GetTagKey(i), tagOpera.GetTagValue(tagOpera.GetTagKey(i), 0));
            }
            Console.WriteLine("these are tags:");
            foreach(KeyValuePair<string, string> KVP in tags)
            {
                Console.WriteLine(KVP.ToString());
            }
            return tags;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="recipient"></param>
        /// <returns>ReceiveEmail is a dictionary like {[user1, {[TO, XXXX], [CC, XXXX], [Subject, XXXX], [body, XXXX]...}]}
        /// it means, in user1's inbox, there is one email, email to XX, CC XXX, Subject is XXX, body is XXX</returns>
        public Dictionary<string, Dictionary<string, string>> ReadEmail(string  caseID, string domain, string recipient)
        {
            
            Dictionary<string, Dictionary<string, string>> ReceivedEmail = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, string> realTag = new Dictionary<string, string>();
            

            //currently recipient is domain-user
            //List<string> user = recipient.Split('-').ToList(); //get recipient of email domain, user[0] is domain
            //Initialize(user[0].ToLower()); //Initialize exchange info for the domain
            Initialize(domain); //Initialize exchange info for the domain

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            ExchangeService servicePremise = new ExchangeService(ExchangeVersion.Exchange2013);

            servicePremise.Credentials = new NetworkCredential(recipient, Password, domain);
            servicePremise.Url = new Uri(URL);

            ItemView view = new ItemView(int.MaxValue);
            FindItemsResults<Item> findResults = servicePremise.FindItems(WellKnownFolderName.Inbox, SetFilter(), view);

            PropertySet PS = new PropertySet();
            ExtendedPropertyDefinition PD_Header = new ExtendedPropertyDefinition(0x007D, MapiPropertyType.String);
            PS.Add(PD_Header);
            string Headers;


            foreach (Item item in findResults)
            {


                Dictionary<string, string> EmailContent = new Dictionary<string, string>();
                item.Load(); //have to load email before reading.

                string TO = null;
                string CC = null;


               EmailMessage message = EmailMessage.Bind(servicePremise, item.Id, new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Attachments));
               
               
                if (item.DisplayCc != null)
                {
                    foreach (EmailAddress eAdd in message.CcRecipients)
                    {
                        CC = CC + eAdd.Address + ";";
                    }

                    EmailContent.Add("CC", CC.TrimEnd(';')); 
                }
                else
                {
                    EmailContent.Add("CC", "");
                }

                if (item.DisplayTo != null)
                {
                    foreach (EmailAddress eAdd in message.ToRecipients)
                    {
                        TO = TO + eAdd.Address + ";";
                    }

                    EmailContent.Add("To", TO.TrimEnd(';')); 
                }
                else
                {
                    EmailContent.Add("To", "");
                }

                if(item.Subject!= null)
                {
                    EmailContent.Add("Subject", item.Subject);
                }
                if (item.Body != null)
                {
                    EmailContent.Add("Body", item.Body);
                }
                else
                {
                    EmailContent.Add("Body", "");
                }
                 if (item.HasAttachments)
                {
                    string attachmentName = null;
                    foreach (Attachment attachment in item.Attachments)
                    {
                        string tags = null;
                        if (attachment is FileAttachment)
                        {
                            FileAttachment fileAttachment = attachment as FileAttachment;
                            string tag = null;
                            
                            CreateFolder(@"C:\EE\" + caseID + @"\" + domain + "-" + recipient);
                            fileAttachment.Load(@"C:\EE\" + caseID + @"\" + domain + "-" + recipient + @"\" + attachment.Name); //download attachment
                            realTag = ReadTag(@"C:\EE\" + caseID + @"\" + domain + "-" + recipient + @"\" + attachment.Name);
                            if (!EmailContent.Keys.Contains("Tag"))
                            {
                                EmailContent.Add("Tag", ""); 
                            }

                            if (realTag.Count != 0)
                            {
                                foreach (KeyValuePair<string, string> kvp in realTag)
                                {
                                    tag += kvp.Key.ToString() + ":" + kvp.Value.ToString() + "/";
                                }

                                tags = attachment.Name + "*" + tag.TrimEnd('/') + ";";
                            }
                            
                            if (attachmentName == null)
                                attachmentName = attachment.Name;
                            else
                                attachmentName = attachmentName + ";" + attachment.Name;
                            if (attachment.Name.Contains(".nxl"))
                                File.Copy(@"C:\EE\" + caseID + @"\" + domain + "-" + recipient + @"\" + attachment.Name, @"C:\EE\" + caseID + @"\" + domain + "-" + recipient + @"\" + attachment.Name + ".txt", true);
                        }
                        else
                        {
                            ItemAttachment itemAttachment = attachment as ItemAttachment;
                            itemAttachment.Load();
                            if (attachmentName == null)
                                attachmentName = itemAttachment.Name;
                            else
                                attachmentName = itemAttachment + ";" + itemAttachment.Name;
                        }

                        if (EmailContent.Keys.Contains("Tag"))
                            EmailContent["Tag"] = EmailContent["Tag"] +  tags;
                        
                    }
                    if (EmailContent.Keys.Contains("Tag"))
                        EmailContent["Tag"] = EmailContent["Tag"].TrimEnd(';');
                    EmailContent.Add("Attachment", attachmentName);
                
                }
                else
                {
                    EmailContent.Add("Attachment", "");
                }

                servicePremise.LoadPropertiesForItems(findResults, PS);
                if (item.TryGetProperty(PD_Header, out Headers))
                {
                    EmailContent.Add("Header", Headers.ToLower());
                }
                

                ReceivedEmail.Add(recipient, EmailContent);
            }
            return ReceivedEmail;
        }

        
        public SearchFilter SetFilter()
        {
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();

            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true));
            //SearchFilter searchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
            //SearchFilter searchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, true); 
            //searchFilterCollection.Add(searchFilter);
            SearchFilter s = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, searchFilterCollection.ToArray());
            return s;
        }

        public void CreateFolder(string dir)
        {
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
        }

        //compare tag or header
        public static bool CompareTwoDictionaries(Dictionary<string, string> Dic1, Dictionary<string, string> Dic2)
        {
            bool result = true;
            Dictionary<string, string>.KeyCollection keys1 = Dic1.Keys; // expected
            Dictionary<string, string>.KeyCollection keys2 = Dic2.Keys; // real
            foreach(string k1 in keys1)
            {
                if (!keys2.Contains(k1))
                {
                    Console.WriteLine("Assertion Fail: can't find expected key/header: " + k1.ToString());
                    return result = false;
                }
                if (Dic1[k1].ToLower() != Dic2[k1].ToLower())
                {
                    return result = false;
                }
                
            }
             
            return result;
        }

        //sting format: aaa;bbb;ccc; or aaa;bbb;ccc    s1 is expected, s2 is real
        public static bool ConverAndCompareTwoLists(string s1, string s2)
        {
            bool result = true;
            List<string> list1 = new List<string>();
            List<string> list2 = new List<string>();

            s1 = s1.TrimEnd(';').ToLower();
            s2 = s2.TrimEnd(';').ToLower();

            if (s1.Contains(';'))
                list1 = s1.Split(';').ToList();
            else
                list1.Add(s1);
            if (s2.Contains(';'))
                list2 = s2.Split(';').ToList();
            else
                list2.Add(s2);

            foreach(string item in list1)
            {
                if (!list2.Contains(item))
                {
                    Console.WriteLine("Assertion Fail: can't find expected: " + item);
                    result = false;
                }
            }

            foreach(string item in list2)
            {
                if (!list1.Contains(item))
                {
                    Console.WriteLine("Assetion Fail: " + item + " is unexcepted");
                    result = false;
                }
            }

            return result;
        }

        /// <summary>
        /// Doing assert
        /// </summary>
        /// <param name="expectedResult"></param>
        /// <returns></returns>
        public bool Assertion(string caseID, string domain, string recipient, string ExpectedReciver, Dictionary<string, string> Expected, Dictionary<string, Dictionary<string, string>> Real)
        {
            bool assertion = true;
            Dictionary<string, string> myReal = new Dictionary<string, string>();  //real email info
            List<string> attachmentList = new List<string>();

            foreach (KeyValuePair<string, Dictionary<string, string>> kvp in Real)
            {
                myReal = kvp.Value;

            }

            Console.WriteLine("Assertion: verify email for " + ExpectedReciver);
            if (Expected.ContainsKey("Attachment"))
            {
                if (Expected["Attachment"].Contains(';'))
                {
                    attachmentList = Expected["Attachment"].Split(';').ToList();
                }
                else
                    attachmentList.Add(Expected["Attachment"]);
            }
            try
            {
                //how many emails in real
                int emailNumber = Real.Count;

                switch (Expected["Result"].ToLower())
                {
                    case "allow":
                        if (emailNumber == 0)
                        {
                            Console.WriteLine("Assertion Fail: Expected to recive email but no email recived");
                            assertion = false;
                        }
                        else
                        {
                            foreach (string key in Expected.Keys)
                            {
                                switch (key)
                                {
                                    case "Result":
                                        break;
                                    case "To":
                                        {
                                            Console.WriteLine("verify To");
                                            if (!ConverAndCompareTwoLists(Expected["To"], myReal["To"]))
                                                assertion = false;
                                            break;
                                        }
                                    case "CC":
                                        {
                                            Console.WriteLine("verify CC");
                                            //if (!ConverAndCompareTwoLists(Expected["To"], myReal["To"]))
                                            if (!ConverAndCompareTwoLists(Expected["CC"], myReal["CC"]))
                                                assertion = false;
                                            break;
                                        }
                                    case "Subject":
                                        {
                                            Console.WriteLine("verify Subject");
                                            if (Expected["Subject"].ToLower() == myReal["Subject"].ToLower())
                                                break;
                                            else
                                            {
                                                Console.WriteLine("Assertion Fail: Expected Subject: " + Expected["Subject"].ToLower() + " But Real Subject is: " + myReal["Subject"]);
                                                assertion = false;
                                                break;
                                            }
                                        }
                                    case "Body":
                                        {
                                            Console.WriteLine("verify Body");
                                            //if (myReal["Body"].ToLower() == Expected["Body"].ToLower() || myReal["Body"].ToLower().Contains(Expected["Body"].ToLower()))
                                            //break;

                                            string realbody = myReal["Body"].ToLower();
                                            string[] myrealbody = Regex.Split(realbody, "<body>\r\n", RegexOptions.IgnoreCase);
                                            myrealbody = Regex.Split(myrealbody[1].ToLower(), "\r\n</body>", RegexOptions.IgnoreCase);
                                            if (Expected["Body"].ToLower().Equals(myrealbody[0].ToLower()))
                                                break;


                                            else
                                            {
                                                //Console.WriteLine("Assertion Fail: Expected Body: " + Expected["Body"].ToLower() + " But Real Body is: " + myReal["Body"].ToLower());
                                                Console.WriteLine("Assertion Fail: Expected Body: " + Expected["Body"].ToLower() + " But Real Body is: " + myrealbody[0].ToString());
                                                assertion = false;
                                                break;
                                            }
                                        }
                                    case "Attachment":
                                        {
                                            Console.WriteLine("verify Attachment");
                                            if (!Expected.Keys.Contains("Attachment") || Expected["Attachment"] == "") // it means don't need to verify attachment.
                                            {
                                                break;
                                            }
                                            else
                                            {
                                                if (Expected["Attachment"].ToLower() == "noattachment" && myReal["Attachment"] == "") // it means no attachment.
                                                    break;
                                                if (!ConverAndCompareTwoLists(Expected["Attachment"], myReal["Attachment"]))
                                                    assertion = false;
                                                break;
                                            }
                                        }
                                    case "Header":
                                        {
                                            Console.WriteLine("verify Header");
                                            
                                            Dictionary<string, string> Expected_Header = HeaderStringToDictionary(Expected["Header"].ToLower());
                                            string Real_Header = myReal["Header"].ToLower();
                                            
                                            foreach(KeyValuePair<string, string> Header in Expected_Header)
                                            {
                                                if (!Real_Header.Contains(Header.Key + ": " + Header.Value)) // add a space after : to match the real value in Header
                                                    assertion = false;
                                            }
                                                
                                            break;
                                        }
                                    case "Tag":
                                        {
                                            Console.WriteLine("verify Tag");
                                            List<bool> result = new List<bool>();
                                           
                                            //Expected_FileTag and Real_FileTags: filename1:{tagname:value,tagname:value};filename*tagname:value
                                            Dictionary<string, Dictionary<string, string>> Expected_FileTags = TagStringToDictionary(Expected["Tag"]);
                                            Dictionary<string, Dictionary<string, string>> Real_FileTags = TagStringToDictionary(myReal["Tag"]);
                                            
                                            foreach (KeyValuePair<string, Dictionary<string, string>> kvp in Expected_FileTags)
                                            {
                                                //verify tag for NXL file
                                                if (kvp.Key.Contains(".nxl"))
                                                {
                                                    List<string> NXLtag = TagDictionaryToList(kvp.Value);
                                                    string fileName = @"C:\EE\" + caseID + @"\" + domain + "-" + recipient + @"\" + kvp.Key;
                                                    
                                                    StreamReader sr = new StreamReader(fileName + ".txt");
                                                    string line = sr.ReadToEnd().ToLower();

                                                    foreach (string tag in NXLtag)
                                                    {
                                                        if (!line.Contains(tag))
                                                            assertion = false;
                                                    }

                                                }
                                                else
                                                {
                                                    //verify tag for non-nxl file
                                                    try
                                                    {
                                                        result.Add(CompareTwoDictionaries(Expected_FileTags[kvp.Key], Real_FileTags[kvp.Key]));   
                                                    }
                                                    catch
                                                    {
                                                        Console.WriteLine("Assertion Fail: please check file name in tag section.");
                                                        result.Add(false);
                                                    }
                                                    if (result.Contains(false))
                                                    {
                                                        assertion = false;
                                                    }
 
                                                }
                                            }
                                            break;
                                        }
                                }
                            }
                        }
                        return assertion;
                    case "deny":
                        {
                            if (emailNumber != 0)
                            {
                                Console.WriteLine("Assertion Fail: expected result is denied, but " + ExpectedReciver + " gets email");
                                assertion = false;
                            }
                            return assertion;
                        }
                    default:
                        return assertion;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                assertion = false;
            }
            return assertion;
            
        }

        //myString formate: filename1*tag1:value1/tag2:value2;filename2*tag1:value1/tag2:value2;
        static Dictionary<string, Dictionary<string, string>> TagStringToDictionary(string myString)
        {
            Dictionary<string, Dictionary<string, string>> myDic = new Dictionary<string, Dictionary<string, string>>();
            List<string> FileTagList = new List<string>(); //FileTagList: "filename1*tag1:value1,tag2:value2","filename2*tag1:value1,tag2:value2"
            
            myString = myString.ToLower().TrimEnd(';');
            if (myString.Contains(';')) //multiple files
            {
                FileTagList = myString.Split(';').ToList();
            }
            else
            {
                FileTagList.Add(myString); //only one file
            }
            FileTagList.Remove("");

            foreach (string FileTag in FileTagList)
            {
                string FileName = null;
                List<string> TagList = new List<string>();
                Dictionary<string, string> TagsDic = new Dictionary<string, string>();

                FileName = FileTag.Split('*').ToList()[0];
                TagList = FileTag.Split('*').ToList()[1].Split('/').ToList();   //TagList: tagname:value, tagname:value

                foreach (string Tag in TagList)
                {

                    TagsDic.Add(Tag.Split(':').ToList()[0], Tag.Split(':').ToList()[1]);
                }

                myDic.Add(FileName, TagsDic);
            }
            return myDic;
        }

        //
        static List<string> TagDictionaryToList(Dictionary<string, string> Tags)
        {
            List<string> myList = new List<string>();
            string myString = null;
            foreach (KeyValuePair<string, string> kvp in Tags)
            {
                    myString = myString + "\"" + kvp.Key + "\":[\"" + kvp.Value + "\"],";
            }
           
            myString = myString.TrimEnd(',').ToLower();
            myList = myString.Split(',').ToList() ;
            return myList;
        }


        //myString formate: tag1:value1/tag2:value2/tag3:value3
        static Dictionary<string, string> HeaderStringToDictionary(string myString)
        {
            Dictionary<string, string> myDic = new Dictionary<string, string>();
            List<string> Tags = new List<string>();
            if (myString.Contains('/'))
            {
                Tags = myString.ToLower().Split('/').ToList();
            }
            else
            {
                Tags.Add(myString);
            }
            foreach (string tag in Tags)
            {
                string TagName = tag.Split(':').ToList()[0];
                string TagValue = tag.Split(':').ToList()[1];
                myDic.Add(TagName, TagValue); //add a space in front of tag value because in real email Header, the value is begin with space.
            }
            return myDic;
        }


        

        public bool CopyFolder(string sourcePath, string destPath)
        {
            if (Directory.Exists(sourcePath))
            {
                if (!Directory.Exists(destPath))
                    try
                    {
                        Directory.CreateDirectory(destPath);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Fail to create the directory: " + destPath + " " + ex.Message);
                    }

                List<string> files = new List<string>(Directory.GetFiles(sourcePath));
                files.ForEach(c =>
                    {
                        string destFile = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                        File.Copy(c, destFile, true);
                    });

                List<string> folders = new List<string>(Directory.GetDirectories(sourcePath));
                folders.ForEach(c =>
                    {
                        string destDir = Path.Combine(new string[] { destPath, Path.GetFileName(c) });
                        CopyFolder(c, destDir);
                    }
                    );
                return true;
            }
            else
            {
                throw new DirectoryNotFoundException("No source directory: " + sourcePath);

            }
        }
        
        
        public bool ClearUp()
        {
            DateTime now = new DateTime();
            now = DateTime.Now;
            string name = now.Year.ToString() + '-' + now.Month.ToString() + '-' + now.Day.ToString() + '-' + now.Hour.ToString() + '-' + now.Minute.ToString() + '-' + now.Millisecond.ToString();
            CopyFolder(@"C:\EE", @"C:\\EE-Backup-" + name);
            return true;
        }

		public Dictionary<string, Dictionary<string, string>> EEFlow(Dictionary<string, Dictionary<string, string>> CaseList)
        
        {
            string path = ConfigurationManager.AppSettings["CCToolPath"];
            string host = ConfigurationManager.AppSettings["CCHost"];
            
            
            bool result = true;

            Dictionary<string, Dictionary<string, string>> failedCase = new Dictionary<string, Dictionary<string, string>>();
            Dictionary<string, Dictionary<string, string>> ExceptedResult = new Dictionary<string, Dictionary<string, string>>();
            
            foreach (KeyValuePair<string, Dictionary<string, string>> MyCase in CaseList)//go through all cases, MyCase.key is caseID.
            {//MyCase.Key is caseID, MyCase.Value is case parameters
                List<bool> lastResult = new List<bool>();

                
                Policy policy = new Policy();
                Dictionary<string, string> caseInfo = new Dictionary<string, string>();
                caseInfo=MyCase.Value;
                List<string> policyName=caseInfo["Policy"].Split(';').ToList();
				Console.WriteLine("******This case get start " + MyCase.Key + ":");
                Console.WriteLine("Prepare the policy for case " + MyCase.Key);
                policy.deploy(path, host, policyName[0]);
                //wait the policy heartbeat, wait 1min
                Thread.Sleep(60000);
                Console.WriteLine("Waitting the policy send to the PEP");
                Console.WriteLine("Prepare start: Begin to remove all emails in Inbox for all related users");
                ClearAllEmails(MyCase.Value);
                Console.WriteLine("Prepare completed: Remove all emails in Inbox");

                Console.WriteLine("Execute testing start: Begin to send email");
                SendEmail(MyCase.Value);
                Console.WriteLine("Execute testing completed: completed send action");
                Console.WriteLine("waiting for email...");
				//after testing completed, inactive the current policy 
                Console.WriteLine("begin to undeploy...");
                policy.unDeploy(path, host, policyName[0]);
                Thread.Sleep(20000);     
                          
                Console.WriteLine("Assert start: Begin to read email");
                ExceptedResult = GetExpectedResult(MyCase.Key.ToString());
                Console.WriteLine("read all real emails info");
                foreach (KeyValuePair<string, Dictionary<string, string>> CasePars in ExceptedResult) //go through all assertions in XML file for current Case
                {//casePars.key is user who receive email, casePars.key is domain-user


                    Dictionary<string, Dictionary<string, string>> RealEmail = new Dictionary<string, Dictionary<string, string>>();
                    List<string> user = CasePars.Key.Split('-').ToList();
                    RealEmail = ReadEmail(MyCase.Key, user[0], user[1]);
                    result = Assertion(MyCase.Key, user[0], user[1], CasePars.Key, CasePars.Value, RealEmail);
                    lastResult.Add(result);
                    Console.WriteLine(result);
                    Console.WriteLine(lastResult);
                    
                }
                Console.WriteLine("Assert completed");
                if (lastResult.Contains(false))
                {
					failedCase.Add(MyCase.Key, MyCase.Value);
                    Console.WriteLine("******" + MyCase.Key + ": False");
                }
                else
                {
                    Console.WriteLine("******" + MyCase.Key + ": Pass");
                }
                Console.WriteLine("-----------------------------------");
			 }
            return failedCase;
        }
        public void collectionFailedCase(XmlNode xmlNode)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes");
            XmlNode rootNode = xmlDoc.CreateElement("EE-Auto");
            XmlNode caseNoNode = xmlDoc.CreateElement("Auto-1");
            XmlNode policyNode = xmlDoc.CreateElement("Policy");
            policyNode.InnerText= "auto test1";
            XmlNode senderNode = xmlDoc.CreateElement("sender");
            senderNode.InnerText = "auto1@auto.com";
            caseNoNode.AppendChild(senderNode);
            caseNoNode.AppendChild(policyNode);
            rootNode.AppendChild(caseNoNode);
            xmlDoc.AppendChild(rootNode);
            xmlDoc.Save(@"C:\EE\failed.xml");
                Console.ReadKey();
		}
        static void Main(string[] args)
        {
            //string path = ConfigurationManager.AppSettings["CCToolPath"];
            //string host = ConfigurationManager.AppSettings["CCHost"];
            //Dictionary<string, Dictionary<string, string>> failedCase = new Dictionary<string, Dictionary<string, string>>();
            Program MyPro = new Program();
            Dictionary<string, Dictionary<string, string>> CaseList =MyPro.GetCase();
            Dictionary<string, Dictionary<string, string>> result=MyPro.EEFlow(CaseList);
            if (result.Count > 0)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.CreateXmlDeclaration("1.0", "utf-8", "yes");
                XmlNode rootNode = xmlDoc.CreateElement("EE-Auto");
                xmlDoc.AppendChild(rootNode);

                foreach (string failedName in result.Keys)
                {
                                       
                    Dictionary<string, string> values = result[failedName];
                    XmlNode caseNoNode = xmlDoc.CreateElement(failedName);
                    rootNode.AppendChild(caseNoNode);
                    foreach (var key in values)
                    {
                        if (key.Key == "Policy")
                        {
                            XmlNode policyNode = xmlDoc.CreateElement("Policy");
                            policyNode.InnerText = key.Value;
                            caseNoNode.AppendChild(policyNode);
                        }
                        if (key.Key == "Sender")
                        {
                            XmlNode senderNode = xmlDoc.CreateElement("Sender");
                            senderNode.InnerText = key.Value;
                            caseNoNode.AppendChild(senderNode);
                        }
                        if (key.Key == "To")
                        {
                            XmlNode toNode = xmlDoc.CreateElement("To");
                            toNode.InnerText = key.Value;
                            caseNoNode.AppendChild(toNode);
                        }
                        if (key.Key == "CC")
                        {
                            XmlNode ccNode = xmlDoc.CreateElement("CC");
                            ccNode.InnerText = key.Value;
                            caseNoNode.AppendChild(ccNode);
                        }
                        if (key.Key == "BCC")
                        {
                            XmlNode bccNode = xmlDoc.CreateElement("BCC");
                            bccNode.InnerText = key.Value;
                            caseNoNode.AppendChild(bccNode);
                        }
                        if (key.Key == "Subject")
                        {
                            XmlNode subjectNode = xmlDoc.CreateElement("Subject");
                            subjectNode.InnerText = key.Value;
                            caseNoNode.AppendChild(subjectNode);
                        }
                        if (key.Key == "Body")
                        {
                            XmlNode bodyNode = xmlDoc.CreateElement("Body");
                            bodyNode.InnerText = key.Value;
                            caseNoNode.AppendChild(bodyNode);
                        }
                        if (key.Key == "Header")
                        {
                            XmlNode headerNode = xmlDoc.CreateElement("Header");
                            headerNode.InnerText = key.Value;
                            caseNoNode.AppendChild(headerNode);
                        }
                        if (key.Key == "Attachment")
                        {
                            XmlNode attNode = xmlDoc.CreateElement("AttAchment");
                            attNode.InnerText = key.Value;
                            caseNoNode.AppendChild(attNode);
                        }
                        if (key.Key == "Tag")
                        {
                            XmlNode tagNode = xmlDoc.CreateElement("Tag");
                            tagNode.InnerText = key.Value;
                            caseNoNode.AppendChild(tagNode);
                        }
                        if (key.Key == "Assertion")
                        {
                            XmlNode assNode = xmlDoc.CreateElement("Assertion");
                            caseNoNode.AppendChild(assNode);
                            Dictionary<string, Dictionary<string, string>> ExceptedResult = new Dictionary<string, Dictionary<string, string>>();
                            ExceptedResult = MyPro.GetExpectedResult(failedName);
                            foreach (string recipient in ExceptedResult.Keys)
                            {
                                Dictionary<string, string> assvalues = ExceptedResult[recipient];
                                XmlNode AssNoNode = xmlDoc.CreateElement(recipient);
                                assNode.AppendChild(AssNoNode);
                                foreach (var assvalue in assvalues)
                                {
                                    if (assvalue.Key == "Result")
                                    {
                                        XmlNode resultNode = xmlDoc.CreateElement("Result");
                                        resultNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(resultNode);

                                    }                                   
                                    if (assvalue.Key == "To")
                                    {
                                        XmlNode toNode = xmlDoc.CreateElement("To");
                                        toNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(toNode);
                                    }
                                    if (assvalue.Key == "CC")
                                    {
                                        XmlNode ccNode = xmlDoc.CreateElement("CC");
                                        ccNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(ccNode);
                                    }
                                    if (assvalue.Key == "BCC")
                                    {
                                        XmlNode bccNode = xmlDoc.CreateElement("BCC");
                                        bccNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(bccNode);
                                    }
                                    if (assvalue.Key == "Subject")
                                    {
                                        XmlNode subjectNode = xmlDoc.CreateElement("Subject");
                                        subjectNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(subjectNode);
                                    }
                                    if (assvalue.Key == "Body")
                                    {
                                        XmlNode bodyNode = xmlDoc.CreateElement("Body");
                                        bodyNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(bodyNode);
                                    }
                                    if (assvalue.Key == "Header")
                                    {
                                        XmlNode headerNode = xmlDoc.CreateElement("Header");
                                        headerNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(headerNode);
                                    }
                                    if (assvalue.Key == "Attachment")
                                    {
                                        XmlNode attNode = xmlDoc.CreateElement("AttAchment");
                                        attNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(attNode);
                                    }
                                    if (assvalue.Key == "Tag")
                                    {
                                        XmlNode tagNode = xmlDoc.CreateElement("Tag");
                                        tagNode.InnerText = assvalue.Value;
                                        AssNoNode.AppendChild(tagNode);
                                    }
                                }
                            }

                        }
                    }


                }
                DateTime now = new DateTime();
                now = DateTime.Now;
                string name = now.Year.ToString() + '-' + now.Month.ToString() + '-' + now.Day.ToString() + '-' + now.Hour.ToString() + '-' + now.Minute.ToString() + '-' + now.Second.ToString();
                xmlDoc.Save(@"C:\EE\failed__"+name+".xml");
                Console.WriteLine("*************Failed case has been save in the file ---Failed" + name + ".xml*************");
                Console.ReadKey();

            }
			else
            
            {
                Console.WriteLine("*************Test case run completed without failed*************");
           	 	Console.ReadKey();
            	MyPro.ClearUp();
                 
			}
        }

        
    }
}
