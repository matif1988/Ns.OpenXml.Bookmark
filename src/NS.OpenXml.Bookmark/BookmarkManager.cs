/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.bookmark - All Rights Reserved */
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using NS.OpenXml.Bookmark.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NS.OpenXml.Bookmark
{
    /// <summary>
    /// The bookmark manager class
    /// </summary>
    public static class BookmarkManager
    {
        #region public methods

        /// <summary>
        /// Opens and fills bookmark content to word document
        /// </summary>
        /// <param name="fileStream">The file stream</param>
        /// <param name="bookmarkContents">The bookmark content list</param>
        /// <returns>The stream file with filled bookmarks</returns>
        public static Stream OpenAndFillBookmarkOnWordDocument(Stream fileStream, List<BookmarkContent> bookmarkContents)
        {
            if (fileStream == null)
                return null;

            // Open a WordprocessingDocument for editing using the file path.
            using (var wordDoc = WordprocessingDocument.Open(fileStream, true))
            {
                // marker for main document part
                var mainDocPart = wordDoc.MainDocumentPart;

                // Fill all text, html and table bookmark content
                bookmarkContents.ForEach(b => FillBookmark(mainDocPart, b));

                // Close the handle explicitly
                mainDocPart.Document.Save();
                wordDoc.Close();

                fileStream.Seek(0, SeekOrigin.Begin);
            }

            return fileStream;
        }

        /// <summary>
        /// Opens and fills bookmark content to word document
        /// </summary>
        /// <param name="fileInfo">The file info</param>
        /// <param name="bookmarkContents">The bookmark content list</param>
        public static void OpenAndFillBookmarkOnWordDocument(FileInfo fileInfo, List<BookmarkContent> bookmarkContents)
        {
            // Open a WordprocessingDocument for editing using the file path.
            using (var wordDoc = WordprocessingDocument.Open(fileInfo.FullName, true))
            {
                // marker for main document part
                var mainDocPart = wordDoc.MainDocumentPart;

                // Fill all text, html and table bookmark content
                bookmarkContents.ForEach(b => FillBookmark(mainDocPart, b));

                // Close the handle explicitly
                mainDocPart.Document.Save();
                wordDoc.Close();
            }
        }

        /// <summary>
        /// Get bookmark list
        /// </summary>
        /// <param name="fileInfo">the file info</param>
        /// <returns>The list of bookmarks</returns>
        public static List<string> GetBookmarkList(FileInfo fileInfo)
        {
            List<string> bookmarks = new List<string>();
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileInfo.FullName, false))
            {
                if (wordDoc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().Any(bs => bs.Name != "_GoBack"))
                {
                    var list = wordDoc.MainDocumentPart.Document.Body.Descendants<BookmarkStart>().Where(bs => bs.Name != "_GoBack").Select(i => i.Name.ToString());
                    bookmarks.AddRange(list);
                }

                wordDoc.Close();
            }

            return bookmarks;
        }

        #endregion

        #region private methods

        /// <summary>
        /// Fill book mark
        /// </summary>
        /// <param name="mainDocPart">The main document part</param>
        /// <param name="bookmarkContent">The bookmark content</param>
        static void FillBookmark(MainDocumentPart mainDocPart, BookmarkContent bookmarkContent)
        {
            BookmarkStart bookmark = mainDocPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkContent.BookmarkContentKey);
            if (bookmark != null)
            {
                switch (bookmarkContent.BookmarkContentType)
                {
                    case BookmarkContentType.Text:
                        InsertIntoBookmark(bookmark, bookmarkContent.BookmarkContentValue?.ToString());
                        break;
                    case BookmarkContentType.Html:
                        InsertHtmlContentIntoBookmark(mainDocPart, bookmark, bookmarkContent.BookmarkContentValue?.ToString());
                        break;
                    case BookmarkContentType.Table:
                        FillTableBookmark(mainDocPart, bookmarkContent);
                        break;
                }
            }
        }

        /// <summary>
        /// Fills the table bookmark
        /// </summary>
        /// <param name="mainDocPart">The main document part</param>
        /// <param name="bookmarkContent">The bookmark content</param>
        static void FillTableBookmark(MainDocumentPart mainDocPart, BookmarkContent bookmarkContent)
        {
            // find our bookmark
            var bookmark = mainDocPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkContent.BookmarkContentKey);
            if (bookmark != null && bookmarkContent.BookmarkContentValue != null
                && bookmarkContent.BookmarkContentValue is ListOfList<BookmarkContent>)
            {
                // get to first element
                OpenXmlElement elem = bookmark.Parent;

                // isolate the table to be worked on
                while (!(elem is Table))
                    elem = elem.Parent;
                OpenXmlElement targetTable = elem;

                // save the row you want to copy each time you have data
                TableRow oldRow = elem.Elements<TableRow>().Last();
                TableRow rowTemplate = (TableRow)oldRow.Clone();

                // whack old row
                elem.RemoveChild(oldRow);

                // Time to slap our data into the table
                foreach (var item in bookmarkContent.BookmarkContentValue
                                    as ListOfList<BookmarkContent>)
                {
                    TableRow newRow = (TableRow)rowTemplate.Clone();
                    IEnumerable<TableCell> cells = newRow.Elements<TableCell>();

                    for (int i = 0; i < cells.Count(); i++)
                    {
                        TableCell cell = cells.ElementAt(i);
                        FillBookmarkListOnTableCell(mainDocPart, cell.Descendants<BookmarkStart>().ToList(), item);
                    }

                    targetTable.AppendChild(newRow);
                }
            }
        }

        /// <summary>
        /// Fills bookmark list on table cell
        /// </summary>
        /// <param name="mainDocPart">The main document part</param>
        /// <param name="bookmarksOnCellList">The bookmark on cell list</param>
        /// <param name="bookmarkContentList">The bookmark content list</param>
        static void FillBookmarkListOnTableCell(MainDocumentPart mainDocPart, List<BookmarkStart> bookmarksOnCellList, List<BookmarkContent> bookmarkContentList)
        {
            if (bookmarksOnCellList != null && bookmarksOnCellList.Count > 0)
            {
                bookmarksOnCellList.ForEach(b =>
                {
                    BookmarkContent content = bookmarkContentList.FirstOrDefault(c => c.BookmarkContentKey == b.Name);
                    if (content != null)
                    {
                        if (content.BookmarkContentType == BookmarkContentType.Html && !string.IsNullOrEmpty(content.BookmarkContentValue?.ToString()))
                        {
                            InsertHtmlContentIntoBookmark(mainDocPart, b, content.BookmarkContentValue.ToString());
                        }
                        else
                        {
                            InsertIntoBookmark(b, content.BookmarkContentValue?.ToString());
                        }
                    }
                });
            }
        }

        /// <summary>
        /// Insert into bookmark
        /// </summary>
        /// <param name="bookmarkStart">The bookmark start</param>
        /// <param name="text">the text to insert</param>
        static void InsertIntoBookmark(BookmarkStart bookmarkStart, string text)
        {
            Run bookmarkText = bookmarkStart.NextSibling<Run>();
            if (bookmarkText != null)
            {
                // if the bookmark has text replace it
                bookmarkText.GetFirstChild<Text>().Text = text;
            }
            else
            {
                // insert after bookmark parent
                bookmarkStart.Parent.Append(new Run(new Text(text)));
            }
        }

        /// <summary>
        /// Insert Html Content Into Bookmark
        /// </summary>
        /// <param name="mainDocPart">The Main Document Part</param>
        /// <param name="bookmarkStart">The Bookmark Start</param>
        /// <param name="html">The BookMark Value</param>
        static void InsertHtmlContentIntoBookmark(MainDocumentPart mainDocPart, BookmarkStart bookmarkStart, string html)
        {
            HtmlConverter converter = new HtmlConverter(mainDocPart);
            IList<OpenXmlCompositeElement> paragraphs = converter.Parse(html);
            if (paragraphs != null && paragraphs.Count > 0
                && paragraphs[0].FirstChild != null)
            {
                OpenXmlElement bookmarkText = bookmarkStart.NextSibling<Run>();
                OpenXmlElement element = paragraphs[0].FirstChild.Clone() as OpenXmlElement;
                if (bookmarkText != null)
                {
                    bookmarkText.RemoveAllChildren();
                    bookmarkText.AppendChild(element);
                }
            }
        }

        #endregion

    }
}
