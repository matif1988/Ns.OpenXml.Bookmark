/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.bookmark - All Rights Reserved */
using System.Collections.Generic;
using System.Web.Mvc;

namespace NS.OpenXml.Bookmark.Helper
{
    /// <summary>
    /// The html render helper class
    /// </summary>
    public static class HtmlRenderHelper
    {
        /// <summary>
        /// Render Image as HTML
        /// </summary>
        /// <param name="src">The image source</param>
        /// <param name="width">The image width</param>
        /// <param name="height">The image height</param>
        /// <returns>Image as HTML</returns>
        public static string RenderImageAsHtml(string src, int? width = null, int? height = null)
        {
            // Create tag builder
            var builder = new TagBuilder("img");

            builder.MergeAttribute(nameof(src), src);

            if (width.HasValue)
                builder.MergeAttribute(nameof(width), width.Value.ToString());

            if (height.HasValue)
                builder.MergeAttribute(nameof(height), height.Value.ToString());

            return builder.ToString(TagRenderMode.Normal);
        }

        /// <summary>
        /// Render Image as HTML
        /// </summary>
        /// <param name="src">The image source</param>
        /// <param name="attributes">The image attributes</param>
        /// <returns>Image as HTML</returns>
        public static string RenderImageAsHtml(string src, Dictionary<string, string> attributes)
        {
            // Create tag builder
            var builder = new TagBuilder("img");

            if (attributes != null)
            {
                foreach (var item in attributes)
                {
                    builder.MergeAttribute(item.Key, item.Value);
                }
            }

            return builder.ToString(TagRenderMode.Normal);
        }
    }
}
