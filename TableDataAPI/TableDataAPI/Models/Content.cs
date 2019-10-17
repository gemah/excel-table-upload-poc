namespace TableDataAPI.Models
{
    /// <summary>
    /// Model for the entity expected in the table files (CSV or XLS(X))
    /// </summary>
    public class Content
    {
        /// <summary>
        /// The entity Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The entity Value.
        /// </summary>
        public int Value { get; set; }

        /// <summary>
        /// A comment about the entity Name and Value.
        /// </summary>
        public string Comment { get; set; }
    }
}
