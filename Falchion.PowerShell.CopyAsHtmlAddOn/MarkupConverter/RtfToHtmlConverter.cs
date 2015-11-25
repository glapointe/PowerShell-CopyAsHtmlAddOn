﻿using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

// Original Source From: http://code.msdn.microsoft.com/windowsdesktop/Converting-between-RTF-and-aaa02a6e/view/SourceCode

namespace Falchion.PowerShell.CopyAsHtmlAddOn.MarkupConverter
{
    public static class RtfToHtmlConverter
    {
        private const string FlowDocumentFormat = "<FlowDocument>{0}</FlowDocument>";

        public static string ConvertRtfToHtml(string rtfText)
        {
            var xamlText = string.Format(FlowDocumentFormat, ConvertRtfToXaml(rtfText));

            return HtmlFromXamlConverter.ConvertXamlToHtml(xamlText, false);
        }

        private static string ConvertRtfToXaml(string rtfText)
        {
            var richTextBox = new RichTextBox();
            if (string.IsNullOrEmpty(rtfText)) return "";

            var textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);

            //Create a MemoryStream of the Rtf content

            using (var rtfMemoryStream = new MemoryStream())
            {
                using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
                {
                    rtfStreamWriter.Write(rtfText);
                    rtfStreamWriter.Flush();
                    rtfMemoryStream.Seek(0, SeekOrigin.Begin);

                    //Load the MemoryStream into TextRange ranging from start to end of RichTextBox.
                    textRange.Load(rtfMemoryStream, DataFormats.Rtf);
                }
            }

            using (var rtfMemoryStream = new MemoryStream())
            {

                textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                textRange.Save(rtfMemoryStream, DataFormats.Xaml);
                rtfMemoryStream.Seek(0, SeekOrigin.Begin);
                using (var rtfStreamReader = new StreamReader(rtfMemoryStream))
                {
                    return rtfStreamReader.ReadToEnd();
                }
            }

        }




    }
}