using System;

namespace Microsoft.Bot.Sample.LuisBot.Model
{
    [Serializable]
    public class CategoryModel
    {
        public string Category { get; set; }
        public string SubCateogry { get; set; }
    }
}