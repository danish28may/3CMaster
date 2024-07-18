using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
namespace PublishingSWordtoHTML
{
    class GlobalMethods
    {
        public static string StrInFolder = null;
        public static string StrOutFolder = null;
        public static string StrProcessFolder = null;       
        public static string strDocumentType = null;
        public static string strJobTransID = null;

        public static void ReadCustomPropertiesXML(string strXMLPath)
        {
            var customProperties = from custProp in XElement.Load(strXMLPath).Elements() select custProp;

            foreach (var property in customProperties)
            {
                if (property.Name.LocalName == "DocumentType")
                {
                    strDocumentType = property.Value;
                    continue;
                }
                else if (property.Name.LocalName == "TransactionID")
                {
                    strJobTransID = property.Value;
                    continue;
                }
            }
           
        }
        public static void MSOCommentText(string strDoc)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(strDoc, Encoding.UTF8);

            doc.DocumentNode.Descendants("div").Where(c => c.GetAttributeValue("class", "").ToLower() == "msocomtxt").ToList().ForEach(x =>
            {

                x.Descendants("p").Where(l => l.GetAttributeValue("class", "").ToLower() == "msonormal").ToList().ForEach(k =>
                {

                    k.SetAttributeValue("class", "MsoCommentText");

                });

            });

            doc.Save(strDoc, Encoding.UTF8);
        }
        //Developer name:Priyanka Vishwakarma, Date:23_09_2019 ,Requirement:Add Author Query heading, Integrated by:Vikas sir.
        public static void ADDAuthorQueryHeading(string strDoc)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(strDoc, Encoding.UTF8);
            HtmlAgilityPack.HtmlNode AQ_heading = HtmlAgilityPack.HtmlNode.CreateNode("<p class=H1><b>Author Queries</b></p> ");


            foreach (HtmlAgilityPack.HtmlNode hr in doc.DocumentNode.Descendants("hr").Where(c => c.GetAttributeValue("class", "").ToLower() == "msocomoff").ToList())
            {

                hr.ParentNode.InsertAfter(AQ_heading, hr);
                break;
            }


            doc.Save(strDoc, Encoding.UTF8);
        }
        public static void MSOCommentListtoBodyA(string strDoc)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(strDoc, Encoding.UTF8);

            doc.DocumentNode.Descendants("div").Where(c => c.GetAttributeValue("class", "").ToLower() == "msocomtxt").ToList().ForEach(x =>
            {

                x.Descendants("p").Where(l => l.GetAttributeValue("class", "").ToLower().Contains("mso") && l.GetAttributeValue("class", "").ToLower().Contains("list")).ToList().ForEach(k =>
                {

                    k.SetAttributeValue("class", "BodyA");

                });

            });

            doc.Save(strDoc, Encoding.UTF8);
        }

    }
}
