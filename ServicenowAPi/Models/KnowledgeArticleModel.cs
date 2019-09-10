using System;

namespace Microsoft.Bot.Sample.LuisBot.Model
{
    [Serializable]
    internal class KnowledgeArticleModel
    {
        public string Category { get; set; }
        public string KnowledgeBaseURL { get; set; }
        public string ShortDescription { get; set; }
    }
}