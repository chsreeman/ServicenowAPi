using System;

namespace Microsoft.Bot.Sample.LuisBot.Model
{
    [Serializable]
    public class IncidentDetailModel
    {
        public string Active { get; set; }
        public string Category { get; set; }
        public string Comments { get; set; }
        public string IncidentState { get; set; }
        public string ShortDescription { get; set; }
        public string Opendate { get; set; }
        public string Duedate { get; set; }
        public string AssignedToGroup { get; set; }
        public string CloseDate { get; set; }
        public string SubCategory { get; set; }
        public string Impact { get; set; }
        public string Priority { get; set; }
        public string Urgency { get; set; }
        public string OpenedByGroup { get; set; }
        public string ClosedByGroup { get; set; }
    }
}