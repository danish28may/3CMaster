using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace PublishingSWordtoHTML
{
    class ExportWord2HTML
    {
        public static object ExportHTML(string strDocPath)
        {
            try
            {
                object outputFileName = null;
                object oMissing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                try
                {
                    word.Visible = false;
                    word.ScreenUpdating = false;

                    Object filename = (Object)strDocPath;
                    Document doc = word.Documents.Open(ref filename, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                                       ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    try
                    {
                        WdSaveFormat formatpros = WdSaveFormat.wdFormatHTML;
                        doc.Activate();
                        outputFileName = strDocPath.Replace(".docx", ".html");
                        object prosfileFormat = formatpros;
                        doc.WebOptions.RelyOnCSS = false;
                        doc.WebOptions.OptimizeForBrowser = true;
                        doc.WebOptions.OrganizeInFolder = true;
                        doc.WebOptions.UseLongFileNames = true;
                        doc.WebOptions.RelyOnVML = false;
                        doc.WebOptions.AllowPNG = false;
                        doc.WebOptions.ScreenSize = Microsoft.Office.Core.MsoScreenSize.msoScreenSize1024x768;
                        doc.WebOptions.PixelsPerInch = 96;
                        doc.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

                        word.DefaultWebOptions().UpdateLinksOnSave = true;
                        word.DefaultWebOptions().CheckIfOfficeIsHTMLEditor = true;
                        word.DefaultWebOptions().CheckIfWordIsDefaultHTMLEditor = true;
                        word.DefaultWebOptions().AlwaysSaveInDefaultEncoding = false;
                        word.DefaultWebOptions().SaveNewWebPagesAsWebArchives = false;

                        doc.SaveAs(ref outputFileName,
                                   ref prosfileFormat, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                   ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    }
                    catch (Exception ex)
                    {
                        string strErrorFileName = null;

                        if (GlobalMethods.StrOutFolder != null || GlobalMethods.StrOutFolder != "")
                        {
                            strErrorFileName = GlobalMethods.StrOutFolder + "\\" + GlobalMethods.strJobTransID;

                            if (strErrorFileName.EndsWith("\\") == false)
                                strErrorFileName = strErrorFileName + "\\";

                            strErrorFileName = strErrorFileName + "Error.log";

                            // Generate Error log and the move the error log in the out folder
                            StreamWriter sw = new StreamWriter(strErrorFileName);
                            sw.WriteLine(ex.ToString());
                            sw.WriteLine("Exception in Processing the document. Please consult 3CM Administrator.");
                            sw.Close();
                        }
                        return null;
                    }
                    finally
                    {
                        object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                        doc.Close(ref saveChanges, ref oMissing, ref oMissing);
                        if (doc != null) Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
                catch (Exception ex)
                {
                    string strErrorFileName = null;

                    if (GlobalMethods.StrOutFolder != null || GlobalMethods.StrOutFolder != "")
                    {
                        strErrorFileName = GlobalMethods.StrOutFolder + "\\" + GlobalMethods.strJobTransID;

                        if (strErrorFileName.EndsWith("\\") == false)
                            strErrorFileName = strErrorFileName + "\\";

                        strErrorFileName = strErrorFileName + "Error.log";

                        // Generate Error log and the move the error log in the out folder
                        StreamWriter sw = new StreamWriter(strErrorFileName);
                        sw.WriteLine(ex.ToString());
                        sw.WriteLine("Exception in Processing the document. Please consult 3CM Administrator.");
                        sw.Close();
                    }
                    return null;
                }
                finally
                {
                    if (word != null)
                    {
                        word.Quit(ref oMissing, ref oMissing, ref oMissing);
                        if (word != null) Marshal.ReleaseComObject(word);
                        word = null;
                    }
                }

                return outputFileName;
            }
            catch (Exception ex)
            {
                string strErrorFileName = null;

                if (GlobalMethods.StrOutFolder != null || GlobalMethods.StrOutFolder != "")
                {
                    strErrorFileName = GlobalMethods.StrOutFolder + "\\" + GlobalMethods.strJobTransID;

                    if (strErrorFileName.EndsWith("\\") == false)
                        strErrorFileName = strErrorFileName + "\\";

                    strErrorFileName = strErrorFileName + "Error.log";

                    // Generate Error log and the move the error log in the out folder
                    StreamWriter sw = new StreamWriter(strErrorFileName);
                    sw.WriteLine(ex.ToString());
                    sw.WriteLine("Exception in Processing the document. Please consult 3CM Administrator.");
                    sw.Close();
                }
                return null;
            }
        }
        
    }
}
