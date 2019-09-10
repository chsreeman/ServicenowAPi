using Newtonsoft.Json.Linq;
using Microsoft.Bot.Sample.LuisBot.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using Newtonsoft.Json;
using ClosedXML.Excel;

namespace ServicenowAPi
{
    public class RestAPIService
    {
        string adminName = ConfigurationManager.AppSettings["IncidentAdmin"];
        string password = ConfigurationManager.AppSettings["IncidentPassword"];
        string knEndPoint = ConfigurationManager.AppSettings["knArticleEndPoint"];
        string catEndPoint = ConfigurationManager.AppSettings["catEndPoint"];
        string knowledgeUserViewUrl = ConfigurationManager.AppSettings["knowledgeUserViewUrl"];
        string assignmentEndPoint = ConfigurationManager.AppSettings["AssignmentEndPoint"];

        /// <summary>
        /// Get Category list from Service now
        /// </summary>
        /// <returns></returns>
        public string[] GetCategoryList()
        {
            var catend = catEndPoint + "element=category^inactive=false";
            var categoryList = new List<CategoryModel>();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(catend);
            request.ContentType = "application/json";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;
            try
            {
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    if (webStream != null)
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            var response = responseReader.ReadToEnd();
                            var selectedNode = JObject.Parse(response).SelectToken("result").ToList();
                            foreach (var item in selectedNode)
                            {
                                var catModel = new CategoryModel();
                                //catModel.SubCateogry = item["label"].ToString();
                                catModel.Category = item["label"].ToString();
                                categoryList.Add(catModel);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
            }
            return categoryList.Select(s => s.Category).ToArray();
        }

        /// <summary>
        /// Get Sub Category list from Service now
        /// </summary>
        /// <param name="categoryName"></param>
        /// <returns></returns>
        public List<string> GetSubCategoryCategoryList(string categoryName)
        {
            var categoryEndPoint = catEndPoint + "element=subcategory^inactive=false";
            var categoryList = new List<CategoryModel>();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(categoryEndPoint);
            request.ContentType = "application/json";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;
            var subCat = new List<string>();
            try
            {
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    if (webStream != null)
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            var response = responseReader.ReadToEnd();
                            var selectedNode = JObject.Parse(response).SelectToken("result").ToList();
                            foreach (var item in selectedNode)
                            {
                                var catModel = new CategoryModel();
                                catModel.SubCateogry = item["label"].ToString();
                                //catModel.Category = item["dependent_value"].ToString();
                                categoryList.Add(catModel);
                            }
                        }
                    }
                }
                subCat = categoryList.Where(s => s.SubCateogry.ToLower().StartsWith("acc")).Select(se => se.SubCateogry).Take(4).ToList();
                List<string> subCatco = categoryList.Where(s => s.SubCateogry.ToLower().StartsWith("co")).Select(se => se.SubCateogry).Take(4).ToList();
                subCat.AddRange(subCatco);
                List<string> subCatd = categoryList.Where(s => s.SubCateogry.ToLower().StartsWith("d")).Select(se => se.SubCateogry).Take(4).ToList();
                subCat.AddRange(subCatd);
                List<string> subCatLocal = categoryList.Where(s => s.SubCateogry.ToLower().StartsWith("local")).Select(se => se.SubCateogry).Take(4).ToList();
                subCat.AddRange(subCatLocal);
            }
            catch (Exception e)
            {
            }
            return subCat.Distinct().ToList();
        }

        /// <summary>
        /// Get Knowledge information from service now
        /// </summary>
        /// <param name="category"></param>
        /// <param name="topicName"></param>
        /// <param name="shortDesc"></param>
        /// <returns></returns>
        internal List<KnowledgeArticleModel> GetKnowledgeInformation(string category, string topicName, string shortDesc)
        {
            var knendpoint = knEndPoint + $"category={category}^topic={topicName}^short_descriptionLIKE{shortDesc}";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(knendpoint);
            request.ContentType = "application/json";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;

            var knArticle = new List<KnowledgeArticleModel>();
            try
            {
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    if (webStream != null)
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            var response = responseReader.ReadToEnd();
                            var selectedNode = JObject.Parse(response).SelectToken("result").Take(2).ToList();
                            foreach (var item in selectedNode)
                            {
                                var knModel = new KnowledgeArticleModel();
                                knModel.ShortDescription = item["short_description"].ToString();
                                knModel.KnowledgeBaseURL = knowledgeUserViewUrl + item["sys_id"].ToString();
                                knArticle.Add(knModel);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return knArticle;
        }

        internal KnowledgeArticleModel GetKnowledgeInformation(string category, string shortDesc)
        {
            var knendpoint = knEndPoint + $"category={category}^short_descriptionLIKE{shortDesc}";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(knendpoint);
            request.ContentType = "application/json";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;

            var knArticle = new KnowledgeArticleModel();
            try
            {
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    if (webStream != null)
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            var response = responseReader.ReadToEnd();
                            var selectedNode = JObject.Parse(response).SelectToken("result").FirstOrDefault();
                            if (selectedNode != null)
                            {
                                knArticle.ShortDescription = selectedNode["short_description"].ToString();
                                knArticle.Category = selectedNode["category"].ToString();
                                knArticle.KnowledgeBaseURL = knowledgeUserViewUrl + selectedNode["sys_id"].ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return knArticle;
        }

        internal KnowledgeArticleModel GetKnowledgeForAllType(string shortDesc)
        {
            var knend = knEndPoint + $"short_descriptionLIKE{shortDesc}";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(knend);
            request.ContentType = "application/xml";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            request.Timeout = 999999;
            request.KeepAlive = true;



            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;

            var knArticle = new KnowledgeArticleModel();
            try
            {
                using (Stream webStream = webResponse.GetResponseStream())
                {
                    if (webStream != null)
                    {
                        using (StreamReader responseReader = new StreamReader(webStream))
                        {
                            var response = responseReader.ReadToEnd();
                            var selectedNode = JObject.Parse(response).SelectToken("result").FirstOrDefault();
                            if (selectedNode != null)
                            {
                                knArticle.ShortDescription = selectedNode["short_description"].ToString();
                                knArticle.Category = selectedNode["category"].ToString();
                                knArticle.KnowledgeBaseURL = knowledgeUserViewUrl + selectedNode["sys_id"].ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return knArticle;
        }

        public string GetIncidentbyAssigmentgroup()
        {
            var catend = assignmentEndPoint;
            var categoryList = new List<CategoryModel>();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(catend);
            request.ContentType = "application/JSON";
            request.ContentType = "application/json; charset=utf-8";

            // request.Headers.Add("Authorization", "Basic " + encoded);
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            request.Timeout = 999999;
            request.KeepAlive = true;
            string responseBody = "";



            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    responseBody = reader.ReadToEnd();
                }
            }
            //HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;
            //try
            //{
            //    using (Stream webStream = webResponse.GetResponseStream())
            //    {
            //        if (webStream != null)
            //        {
            //            using (StreamReader responseReader = new StreamReader(webStream))
            //            {
            //                var response = responseReader.ReadToEnd();
            //                var selectedNode = JObject.Parse(response).SelectToken("result").ToList();
            //                foreach (var item in selectedNode)
            //                {
            //                    var catModel = new CategoryModel();
            //                    //catModel.SubCateogry = item["label"].ToString();
            //                    catModel.Category = item["label"].ToString();
            //                    categoryList.Add(catModel);
            //                }
            //            }
            //        }
            //    }
            //}
            //catch (Exception e)
            //{
            //}
            //return categoryList.Select(s => s.Category).ToArray();
            return responseBody;
        }

        public string SubmitData()
        {
            string responseBody = "";
            string retval = "0";
            try
            {


                //string url = @"https://altriadev.service-now.com/api/now/v1/table/incident?sysparm_query=assignment_group=5925dc406f3625007dad9e0cbb3ee453^active=true";

                string url = @"https://altriadev.service-now.com/api/now/v1/table/incident?sysparm_query=active=true^assignment_group=5925dc406f3625007dad9e0cbb3ee453^ORDERBYpriority^ORDERBYnumber";

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                       | SecurityProtocolType.Tls11
                       | SecurityProtocolType.Tls12
                       | SecurityProtocolType.Ssl3;

                WebRequest request = WebRequest.Create(url);
                request.Credentials = GetCredential(url);
                request.PreAuthenticate = true;


                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    {
                        responseBody = reader.ReadToEnd();
                         retval = SaveExcel(responseBody);

                        retval = responseBody;
                    }
                }
                return retval;

            }

            catch (Exception ex)
            {
                return retval;
                //MessageBox.Show("Error : " + ex.Message);

            }

        }

        public string GetdataForCRTApplications()
        {
            string responseBody = "";
            string retval = "0";
            try
            {


                //string url = @"https://altriadev.service-now.com/api/now/v1/table/incident?sysparm_query=assignment_group=166af5d5db8de344352450fadc96199b^active=true";

                string url = @"https://altriadev.service-now.com/api/now/v1/table/incident?sysparm_query=active=true^assignment_group=166af5d5db8de344352450fadc96199b^ORDERBYpriority^ORDERBYnumber";

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                       | SecurityProtocolType.Tls11
                       | SecurityProtocolType.Tls12
                       | SecurityProtocolType.Ssl3;

                WebRequest request = WebRequest.Create(url);
                request.Credentials = GetCredential(url);
                request.PreAuthenticate = true;
                

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                    {
                        responseBody = reader.ReadToEnd();
                        retval = SaveExcel(responseBody);

                        retval = responseBody;
                    }
                }
                return retval;

            }

            catch (Exception ex)
            {
                return retval;
                //MessageBox.Show("Error : " + ex.Message);

            }

        }

        private CredentialCache GetCredential(string url)
        {

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;
            CredentialCache credentialCache = new CredentialCache();
            credentialCache.Add(new System.Uri(url), "Basic", new NetworkCredential(ConfigurationManager.AppSettings["IncidentAdmin"], ConfigurationManager.AppSettings["IncidentPassword"]));
            return credentialCache;
        }

        public string SaveExcel(string responseBody)
        {
            try
            {

                dynamic dynJson = JsonConvert.DeserializeObject(responseBody);
                using (XLWorkbook workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Tickets");
                    worksheet.Rows(1, 3000).Style.Font.FontName = "Arial";
                    worksheet.Rows(1, 3000).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.BlueGray;
                    worksheet.Range("A1:D1").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A1:D1").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A1:D1").Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A1:D1").Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    worksheet.Range("A1:D1").Style.Alignment.WrapText = true;
                    worksheet.Cell("A1").Value = "INCNumber";
                    worksheet.Cell("B1").Value = "ShortDescription";
                    worksheet.Cell("C1").Value = "state";
                    worksheet.Cell("D1").Value = "priority";
                    
                    int row = 2;
                    foreach (var item in dynJson.result)
                    {
                        //Console.WriteLine("{0} {1} {2} \n", item.number, item.short_description,
                        //    item.state);
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Font.FontSize = 8;
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        worksheet.Range("A" + row.ToString() + ":D" + row.ToString()).Style.Alignment.WrapText = true;
                        worksheet.Cell("A" + row.ToString()).Value = (string)item.number;
                        worksheet.Cell("B" + row.ToString()).Value = (string)item.short_description;
                        worksheet.Cell("C" + row.ToString()).Value = (string)item.state;
                        worksheet.Cell("D" + row.ToString()).Value = (string)item.priority;
                        row = row + 1;
                    }
                    string dateAppendToFileName = DateTime.Now.ToString("hhmmss");

                    string fileName = string.Format("Incident" + dateAppendToFileName + ".xlsx", DateTime.Now.ToString("MMddyyyy"));
                    workbook.SaveAs("D:\\" + fileName);

                }
                return "1";
            }
            catch (Exception ex)
            {
                return "0";
            }

        }
        public string createIncident()
        {
            try
            {

                string username = ConfigurationManager.AppSettings["IncidentAdmin"];
                string password = ConfigurationManager.AppSettings["IncidentPassword"];
                string url = ConfigurationManager.AppSettings["ServiceNowUrl"];

                var auth = "Basic " + Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(username + ":" + password));

                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Headers.Add("Authorization", auth);
                request.Method = "Post";

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    string json = JsonConvert.SerializeObject(new
                    {
                        description = "New Ticket from Service1",
                        short_description = "New Ticket from Service1",
                        category = "Application",
                        subcategory = "Configuration Issue",
                        impact = "2",
                        priority = "3",
                        urgency = "2",
                        reported_person = "14ede437db0563c8c92dfba31d961914",
                        affected_person = "14ede437db0563c8c92dfba31d961914",
                        phone_number = "9246850143",
                        opened_by_group = "SN CRT Application Support",
                        assignment_group = "SN CRT Application Support"


                        // contact_type = ConfigurationManager.AppSettings["ServiceNowContactType"],
                        //caller_id = ConfigurationManager.AppSettings["ServiceNowCallerId"],
                        //cmdb_ci = ConfigurationManager.AppSettings["ServiceNowCatalogueName"],
                        //comments = ConfigurationManager.AppSettings["ServiceNowTicketShortDescription"],

                    });

                    streamWriter.Write(json);
                }

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    var res = new StreamReader(response.GetResponseStream()).ReadToEnd();

                    JObject joResponse = JObject.Parse(res.ToString());
                    JObject ojObject = (JObject)joResponse["result"];
                    string incNumber = ((JValue)ojObject.SelectToken("number")).Value.ToString();

                    return incNumber;
                }
            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string updateIncident(string sysid1)
        {
            try
            {
                string sysid = "4edb12cedb7dbf0088197709af9619b4";
                string url = @"https://altriadev.service-now.com/api/now/table/incident/4edb12cedb7dbf0088197709af9619b4";
                //NetworkCredential credentials = new NetworkCredential
                //{
                //    UserName = ConfigurationManager.AppSettings["IncidentAdmin"],
                //    Password = ConfigurationManager.AppSettings["IncidentPassword"]
                //};
                WebClient ServiceNowClient = new WebClient();
               // ServiceNowClient.Credentials = credentials;

                var auth = "Basic " + Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(ConfigurationManager.AppSettings["IncidentAdmin"] + ":" + ConfigurationManager.AppSettings["IncidentPassword"]));
                ServiceNowClient.Headers.Add("Authorization", auth);                

                string json = JsonConvert.SerializeObject(new
                {
                    short_description = "MMS- need pallet set to loss in MMS #1101154171463801",
                    state = 3
                   // sys_id= "4edb12cedb7dbf0088197709af9619b4"
                });
                string data = JsonConvert.SerializeObject(json, new ResourceLinkConverter());
                string str = ServiceNowClient.UploadString(url, "PUT", json);

                return str;


            }

            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }


}