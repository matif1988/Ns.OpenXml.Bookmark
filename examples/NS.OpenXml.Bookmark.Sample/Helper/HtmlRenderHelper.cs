#if NET_CORE
using Microsoft.AspNetCore.Html;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Text.Encodings.Web;
#else
using System.Web.Mvc;
#endif
using System.Collections.Generic;


namespace NS.OpenXml.Bookmark.Sample.Helper
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
#if NET_CORE
            var builder = new TagBuilder("img")
            {
                TagRenderMode = TagRenderMode.Normal
            };
#else
            var builder = new TagBuilder("img");
#endif
            builder.MergeAttribute(nameof(src), src);

            if (width.HasValue)
                builder.MergeAttribute(nameof(width), width.Value.ToString());

            if (height.HasValue)
                builder.MergeAttribute(nameof(height), height.Value.ToString());
#if NET_CORE
            return GetString(builder);
#else
            return builder.ToString(TagRenderMode.Normal);
#endif
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
#if NET_CORE
            var builder = new TagBuilder("img")
            {
                TagRenderMode = TagRenderMode.Normal
            };
#else
            var builder = new TagBuilder("img");
#endif

            if (attributes != null)
            {
                foreach (var item in attributes)
                {
                    builder.MergeAttribute(item.Key, item.Value);
                }
            }

#if NET_CORE
            return GetString(builder);
#else
            return builder.ToString(TagRenderMode.Normal);
#endif
        }

#if NET_CORE
        /// <summary>
        /// Gets html content as string
        /// </summary>
        /// <param name="content">The html content.</param>
        /// <returns>The html as string.</returns>
        public static string GetString(IHtmlContent content)
        {
            var writer = new System.IO.StringWriter();
            content.WriteTo(writer, HtmlEncoder.Default);
            return writer.ToString();
        }
#endif
    }
}
