using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ServicenowAPi.Controllers
{
  //  [Authorize]
    public class ValuesController : ApiController
    {
        // GET api/values
        public string Get()
        {
            RestAPIService api = new RestAPIService();
            //List<string> str= api.GetSubCategoryCategoryList("Application");
            //return str;
            //return api.GetIncidentbyAssigmentgroup();
             return api.SubmitData();
            // var abc = api.GetKnowledgeForAllType("MMS - Delete");
            //  api.GetKnowledgeForAllType("MMS - Delete ");
            // return "avb";

            //string str=api.createIncident();
            //return str;
           // return api.updateIncident();
        }

        // GET api/values/5
        public string Get(string  id)
        {
            RestAPIService api = new RestAPIService();
            return api.updateIncident(id);
            return "value";
        }

        // POST api/values
        public void Post([FromBody]string value)
        {
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
