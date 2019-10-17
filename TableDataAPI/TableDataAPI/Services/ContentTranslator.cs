using System;
using System.Collections.Generic;
using TableDataAPI.Models;

namespace TableDataAPI.Services
{
    /// <summary>
    /// Implements Excel Table Data to POCO Translation
    /// </summary>
    public class ContentTranslator
    {
        public IList<string> Headers { get; private set; }

        /// <summary>
        /// Default Constructor
        /// </summary>
        /// <param name="translationMode">The string containing the header data for parsing mode</param>
        public ContentTranslator(string[] headers)
        {
            Headers = new List<string>(headers);
        }

        public List<Content> Translate(IList<string[]> input)
        {
            List<Content> translatedContent = new List<Content>();

            foreach (string[] entry in input)
            {
                translatedContent.Add(new Content()
                {
                    Name = entry[Headers.IndexOf("Name")],
                    Value = Convert.ToInt32(entry[Headers.IndexOf("Value")]),
                    Comment = entry[Headers.IndexOf("Comment")]
                });
            }

            return translatedContent;
        }
    }
}
