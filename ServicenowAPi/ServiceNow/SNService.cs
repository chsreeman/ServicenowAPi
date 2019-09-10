using Newtonsoft.Json.Linq;
using Microsoft.Bot.Sample.LuisBot.Model;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;

namespace ServicenowAPi
{
    [Serializable]
    public class SNService
    {
        string assignedGroupEndPoint = ConfigurationManager.AppSettings["assignedGroupEndPoint"];
        string adminName = ConfigurationManager.AppSettings["IncidentAdmin"];
        string password = ConfigurationManager.AppSettings["IncidentPassword"];
        string incidentDetailApi = ConfigurationManager.AppSettings["incidentDetail"];
        /// <summary>
        /// 
        /// </summary>
        /// <param name="category"></param>
        /// <param name="subCategory"></param>
        /// <param name="incidentDesc"></param>
        /// <param name="urgency"></param>
        /// <param name="assignmnetGroup"></param>
        /// <returns></returns>
        public string CallPOSTService(string category, string subCategory, string incidentDesc, string urgency, string assignmnetGroup)
        {
            try
            {
//                  using (var soapClient = new ServiceNowSoapClient())
//                  {
//                      soapClient.ClientCredentials.UserName.UserName = ConfigurationManager.AppSettings["IncidentAdmin"];
//                      soapClient.ClientCredentials.UserName.Password = ConfigurationManager.AppSettings["IncidentPassword"];
//  
//                      var insert = new insert();
//                      var response = new insertResponse();
//  
//                      insert.category = category;
//                      insert.subcategory = subCategory;
//                      insert.impact = urgency;
//                      insert.urgency = urgency;
//                      insert.assignment_group = assignmnetGroup;
//                      insert.short_description = incidentDesc;
//                      response = soapClient.insert(insert);
//                      return response.number;
//                  }
return "NoNumber";
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public IncidentDetailModel GETIncidenServiceInfo(string incNum)
        {
            var incDetail = incidentDetailApi + incNum + $"&sysparm_limit=1";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(incDetail);
            request.ContentType = "application/json";
            request.Method = "GET";
            NetworkCredential cred = new NetworkCredential(adminName, password);
            request.Credentials = cred;
            HttpWebResponse webResponse = request.GetResponse() as HttpWebResponse;

            var incidentDetail = new IncidentDetailModel();
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
                                incidentDetail.Active = selectedNode["active"].ToString();
                                incidentDetail.Category = selectedNode["category"].ToString();
                                incidentDetail.IncidentState = GetStatus(selectedNode["incident_state"].ToString());
                                incidentDetail.Comments = selectedNode["comments_and_work_notes"].ToString();
                                incidentDetail.ShortDescription = selectedNode["short_description"].ToString();
                                incidentDetail.Duedate = string.Empty;
                                incidentDetail.Opendate = string.Empty;
                                incidentDetail.CloseDate = string.Empty;
                                var dueDate = TryParseDate(selectedNode["due_date"].ToString());
                                if (dueDate != null)
                                {
                                    incidentDetail.Duedate = ConvertToEST(Convert.ToDateTime(dueDate)).ToString();
                                }
                                var openDate = TryParseDate(selectedNode["opened_at"].ToString());
                                if (openDate != null)
                                {
                                    incidentDetail.Opendate = ConvertToEST(Convert.ToDateTime(openDate)).ToString();
                                }
                                var closeDate = TryParseDate(selectedNode["closed_at"].ToString());
                                if (closeDate != null)
                                {
                                    incidentDetail.CloseDate = ConvertToEST(Convert.ToDateTime(closeDate)).ToString();
                                }
                                incidentDetail.Category = selectedNode["category"].ToString();
                                incidentDetail.SubCategory = selectedNode["subcategory"].ToString();
                                var impact = selectedNode["impact"].ToString();
                                incidentDetail.Impact = impact == "1" ? "High" : impact == "2" ? "Medium" : impact == "3" ? "Low" : "NA";
                                var urgency = selectedNode["urgency"].ToString();
                                incidentDetail.Urgency = urgency == "1" ? "High" : urgency == "2" ? "Medium" : urgency == "3" ? "Low" : "NA";
                                var priority = selectedNode["priority"].ToString();
                                incidentDetail.Priority = priority == "1" ? "Critical" : priority == "2" ? "High" : priority == "3" ? "Moderate" : priority == "4" ? "Low" : "NA";

                                // selectedNode["u_opened_by_group"].ToString()
                                var openByGroup = selectedNode["u_opened_by_group"];
                                if (openByGroup.HasValues)
                                {
                                    var sysId = openByGroup["value"].ToString();
                                    incidentDetail.OpenedByGroup = GetGroupName(assignedGroupEndPoint + $"{sysId}", sysId);
                                }
                                var closedBy = selectedNode["closed_by"];
                                if (closedBy.HasValues)
                                {
                                    var closedSysId = closedBy["value"].ToString();
                                    incidentDetail.ClosedByGroup = GetGroupName(assignedGroupEndPoint + $"{closedSysId}", closedSysId);
                                }
                                var assignedGroup = selectedNode["assignment_group"];
                                if (assignedGroup.HasValues)
                                {
                                    var assignedSysId = assignedGroup["value"].ToString();
                                    incidentDetail.AssignedToGroup = GetGroupName(assignedGroupEndPoint + $"{assignedSysId}", assignedSysId);
                                }
                            }
                            else
                            {
                                incidentDetail = null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return incidentDetail;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="incNum"></param>
        /// <returns></returns>
        //public IncidentDetailModel GETIncidenServiceInfo(string incNum)
        //{
        //    var incidentDetail = new IncidentDetailModel();
        //    try
        //    {
        //        using (var soapClient = new ServiceNowSoapClient())
        //        {
        //            soapClient.ClientCredentials.UserName.UserName = ConfigurationManager.AppSettings["IncidentAdmin"];
        //            soapClient.ClientCredentials.UserName.Password = ConfigurationManager.AppSettings["IncidentPassword"];

        //            var get = new getRecords();
        //            get.number = incNum;
        //            var response = soapClient.getRecords(get).FirstOrDefault();
        //            if (response != null)
        //            {
        //                incidentDetail.Active = response.active;
        //                incidentDetail.Category = response.category;
        //                incidentDetail.IncidentState = GetStatus(response.incident_state);
        //                incidentDetail.Comments = response.comments_and_work_notes;
        //                incidentDetail.ShortDescription = response.short_description;
        //                incidentDetail.Duedate = string.Empty;
        //                incidentDetail.Opendate = string.Empty;
        //                incidentDetail.CloseDate = string.Empty;
        //                var dueDate = TryParseDate(response.due_date);
        //                if (dueDate != null)
        //                {
        //                    incidentDetail.Duedate = ConvertToEST(Convert.ToDateTime(dueDate)).ToString();
        //                }
        //                var openDate = TryParseDate(response.opened_at);
        //                if (openDate != null)
        //                {
        //                    incidentDetail.Opendate = ConvertToEST(Convert.ToDateTime(openDate)).ToString();
        //                }
        //                var closeDate = TryParseDate(response.closed_at);
        //                if (closeDate != null)
        //                {
        //                    incidentDetail.CloseDate = ConvertToEST(Convert.ToDateTime(closeDate)).ToString();
        //                }
        //                incidentDetail.Category = response.category;
        //                incidentDetail.SubCategory = response.subcategory;
        //                incidentDetail.Impact = response.impact == "1" ? "High" : response.impact == "2" ? "Medium" : response.impact == "3" ? "Low" : "NA";
        //                incidentDetail.Urgency = response.severity == "1" ? "High" : response.severity == "2" ? "Medium" : response.severity == "3" ? "Low" : "NA";
        //                incidentDetail.Priority = response.priority == "1" ? "Critical" : response.priority == "2" ? "High" : response.priority == "3" ? "Moderate" : response.priority == "4" ? "Low" : "NA";
        //                incidentDetail.OpenedByGroup = GetGroupName(assignedGroupEndPoint + $"{response.u_opened_by_group}", response.u_opened_by_group);
        //                incidentDetail.ClosedByGroup = GetGroupName(assignedGroupEndPoint + $"{ response.closed_by}", response.closed_by);
        //                incidentDetail.AssignedToGroup = GetGroupName(assignedGroupEndPoint + $"{ response.assignment_group}", response.assignment_group);
        //            }
        //            else
        //            {
        //                incidentDetail = null;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //        throw;
        //    }
        //    return incidentDetail;
        //}

        /// <summary>
        /// Get status detail
        /// </summary>
        /// <param name="incident_state"></param>
        /// <returns></returns>
        private string GetStatus(string incident_state)
        {
            switch (incident_state)
            {
                case "1": return "New";
                case "2": return "Active";
                case "3": return "Awaiting Problem";
                case "4": return "Awaiting User Info";
                case "5": return "Awaiting Change";
                case "6": return "Resolved";
                case "7": return "Closed";
                case "8": return "Canceled";
                default:
                    break;
            }
            return string.Empty;
        }

        /// <summary>
        /// convert to est 
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static DateTime ConvertToEST(DateTime date)
        {
            TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            DateTime easternTime = TimeZoneInfo.ConvertTimeFromUtc(date, easternZone);
            return easternTime;
        }

        /// <summary>
        /// Custom tryparse date
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static DateTime? TryParseDate(string text)
        {
            DateTime date;
            return DateTime.TryParse(text, out date) ? date : (DateTime?)null;
        }



        internal string GetGroupName(string urlEndPoint, string sysID)
        {
            if (string.IsNullOrWhiteSpace(sysID))
                return string.Empty;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlEndPoint);
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
                                var groupName = selectedNode["name"].ToString();
                                return groupName;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
            return string.Empty;
        }
    }


}