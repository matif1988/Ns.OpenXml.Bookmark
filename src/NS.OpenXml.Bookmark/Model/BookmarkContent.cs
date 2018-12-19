/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.bookmark - All Rights Reserved */
namespace NS.OpenXml.Bookmark.Model
{
    /// <summary>
    /// The bookmark content
    /// </summary>
    public class BookmarkContent
    {
        /// <summary>
        /// Gets or sets the bookmark content key
        /// </summary>
        public string BookmarkContentKey { get; set; }

        /// <summary>
        /// Gets or sets the bookmark content type
        /// </summary>
        public BookmarkContentType BookmarkContentType { get; set; }

        /// <summary>
        /// Gets or sets the bookmark content value
        /// </summary>
        public object BookmarkContentValue { get; set; }
    }
}
