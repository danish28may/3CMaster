using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PublishingSWordtoHTML
{
    class Program
    {
        static void Main(string[] args)
        {


            //string filepath = @"C:\3CMAutomation\WORD2HTML\Process\12675";

            string filepath = args[0];

            try
            {
                Start(filepath);
            }
            catch (Exception ex)
            {
                string folderpath = filepath.Substring(filepath.LastIndexOf("\\"));

                //folderpath = folderpath.Substring(folderpath.LastIndexOf("\\"));

                string strErrorFileName = null;
                string strProcessFolder = null;

                if (GlobalMethods.StrOutFolder != null || GlobalMethods.StrOutFolder != "")
                {
                    strErrorFileName = GlobalMethods.StrOutFolder;
                    if (folderpath.StartsWith("\\"))
                        strErrorFileName = strErrorFileName + folderpath;
                    else
                        strErrorFileName = strErrorFileName + "\\" + folderpath;

                    if (Directory.Exists(strErrorFileName) == false)
                        Directory.CreateDirectory(strErrorFileName);

                    if (strErrorFileName.EndsWith("\\") == false)
                        strErrorFileName = strErrorFileName + "\\";

                    strErrorFileName = strErrorFileName + "Error.log";

                    // Generate Error log and the move the error log in the out folder
                    StreamWriter sw = new StreamWriter(strErrorFileName);
                    sw.WriteLine(ex.ToString());
                    sw.WriteLine("Publishing Structured Word2HTML");
                    sw.WriteLine("Exception in Processing the document. Please consult 3CM Administrator.");
                    sw.Close();
                }

                strProcessFolder = GlobalMethods.StrProcessFolder;

                if (folderpath.StartsWith("\\"))
                    strProcessFolder = strProcessFolder + folderpath;
                else
                    strProcessFolder = strProcessFolder + "\\" + folderpath;
                if (Directory.Exists(strProcessFolder))
                    Directory.Delete(strProcessFolder, true);
            }

        }
        static void Start(string strDocumentName)
        {

            GlobalMethods.StrInFolder = ConfigurationManager.AppSettings.Get("Word2HTMLIN");
            GlobalMethods.StrOutFolder = ConfigurationManager.AppSettings.Get("Word2HTMLOUT");
            GlobalMethods.StrProcessFolder = ConfigurationManager.AppSettings.Get("Word2HTMLPROCESS");
            object OutputFilename = null;
            FileInfo ff = null;
            GlobalMethods.strJobTransID = null;           
            
            DirectoryInfo Dir = new DirectoryInfo(strDocumentName);

            FileInfo[] filesForProcess = Dir.GetFiles();

            foreach (var files in filesForProcess)
            {
                if (files.Extension == ".xml")
                {
                    ///Added by Manish on 07-05-2018 to read XML file for document custom properties
                    GlobalMethods.ReadCustomPropertiesXML(files.FullName);
                }
                if (files.Extension == ".docx")
                {
                    if (!files.Name.StartsWith("~$"))
                    {
                        ff = files;

                    }
                }
            }
            ///Added by Manish on 07-05-2018 to get files from process folder end

            if (ff == null)
            {
                Console.WriteLine("Job name: " + ff.Name + " ," + "Job ID: " + ff.Directory.Name + " ," + "Processing start time:" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss"));

                string strErrorFileName = null;

                if (GlobalMethods.StrOutFolder != null && GlobalMethods.StrOutFolder != "")
                {
                    strErrorFileName = GlobalMethods.StrOutFolder + "\\" + Dir.Name;

                    if (Directory.Exists(strErrorFileName) == false)
                        Directory.CreateDirectory(strErrorFileName);

                    if (strErrorFileName.EndsWith("\\") == false)
                        strErrorFileName = strErrorFileName + "\\";

                    strErrorFileName = strErrorFileName + "Error.log";

                    // Generate Error log and the move the error log in the out folder
                    StreamWriter sw = new StreamWriter(strErrorFileName);
                    sw.WriteLine("Publishing WordtoHTML");
                    sw.WriteLine("Document not found in Input folder. Please consult 3CM Administrator.");
                    sw.Close();
                }

                // Remove the document from Process folder

                goto EndProcess;
            }


            if (GlobalMethods.strJobTransID == null || GlobalMethods.strJobTransID == "")
            {
                string strErrorFileName = null;

                if (GlobalMethods.StrOutFolder != null && GlobalMethods.StrOutFolder != "")
                {
                    strErrorFileName = GlobalMethods.StrOutFolder + "\\" + ff.Directory.Name;

                    if (Directory.Exists(strErrorFileName) == false)
                        Directory.CreateDirectory(strErrorFileName);

                    if (strErrorFileName.EndsWith("\\") == false)
                        strErrorFileName = strErrorFileName + "\\";

                    strErrorFileName = strErrorFileName + "Error.log";

                    // Generate Error log and the move the error log in the out folder
                    StreamWriter sw = new StreamWriter(strErrorFileName);
                    sw.WriteLine("Publishing WordtoHTML");
                    sw.WriteLine("Either the Job Transaction ID is missing in the document properties or is empty. Please consult 3CM Administrator.");
                    sw.Close();
                }

                // Remove the document from Process folder

                goto EndProcess;
            }

            if (GlobalMethods.strDocumentType == null || GlobalMethods.strDocumentType == "")
            {
                string strErrorFileName = null;

                if (GlobalMethods.StrOutFolder != null && GlobalMethods.StrOutFolder != "")
                {
                    strErrorFileName = GlobalMethods.StrOutFolder + "\\" + GlobalMethods.strJobTransID;

                    if (Directory.Exists(strErrorFileName) == false)
                        Directory.CreateDirectory(strErrorFileName);

                    if (strErrorFileName.EndsWith("\\") == false)
                        strErrorFileName = strErrorFileName + "\\";

                    strErrorFileName = strErrorFileName + "Error.log";

                    // Generate Error log and the move the error log in the out folder
                    StreamWriter sw = new StreamWriter(strErrorFileName);
                    sw.WriteLine("Publishing WordtoHTML");
                    sw.WriteLine("Either the DocumentType is missing in the document properties or is empty. Please consult 3CM Administrator.");
                    sw.Close();
                }

                // Remove the document from Process folder

                goto EndProcess;
            }

            System.Threading.Thread.Sleep(3000);

            //Added by Karan on 11-09-2018 for equation Start 

            var sourceDoc = new FileInfo(ff.FullName);//In process folder
            
            OutputFilename = ExportWord2HTML.ExportHTML(ff.FullName);
            //added by vikas on 13-03-2021 for  content wordtoword conversion images come as its in word
            //if (GlobalMethods.strImagePath.ToString().Contains("_W_")&& (GlobalMethods.strImagePath.Contains("\\CSU\\")|| GlobalMethods.strImagePath.Contains("\\UTS\\")|| GlobalMethods.strImagePath.Contains("\\QUT\\"))&&GlobalMethods.strServiceType=="Content")
            //{
            //    ExportWord2HTML.CopyHtmlImagestoloresfolder(OutputFilename.ToString());
            //}



            if (OutputFilename != null)
            {
                GlobalMethods.MSOCommentText(OutputFilename.ToString());

                GlobalMethods.MSOCommentListtoBodyA(OutputFilename.ToString());

                GlobalMethods.ADDAuthorQueryHeading(OutputFilename.ToString());  //Developer name:Priyanka Vishwakarma ,Date:23_09_2019 ,Requirement:Add Author Query heading ,Integrated by:Vikas sir.


                if (File.Exists(OutputFilename.ToString()))
                {
                    // Move the processed document to out folder
                    FileInfo fout = new FileInfo(OutputFilename.ToString());

                    string strParentDirectory = fout.Directory.Name;
                    string outfile = null;

                    outfile = fout.Name;

                    // On completion copy the output file in out folder

                    if (Directory.Exists(GlobalMethods.StrOutFolder + "\\" + strParentDirectory) == false)
                    {
                        Directory.CreateDirectory(GlobalMethods.StrOutFolder + "\\" + strParentDirectory);
                    }

                    outfile = GlobalMethods.StrOutFolder + "\\" + strParentDirectory + "\\" + outfile;

                    if (Directory.Exists(GlobalMethods.StrOutFolder + "\\" + strParentDirectory))
                        File.Copy(OutputFilename.ToString(), outfile, true);
                    else
                    {
                        Directory.CreateDirectory(GlobalMethods.StrOutFolder + "\\" + strParentDirectory);
                        File.Copy(OutputFilename.ToString(), outfile, true);
                    }

                    // Delete the files from the Process folder

                    fout.Delete();

                    // Remove unnecessary files and folders from out folder

                    Directory.Delete(fout.Directory.FullName, true);

                    System.Threading.Thread.Sleep(3000);

                }
            }

        EndProcess:
            {
                if (Directory.Exists(GlobalMethods.StrProcessFolder + "\\" + Dir.Name))
                {
                    Directory.Delete(GlobalMethods.StrProcessFolder + "\\" + Dir.Name, true);
                }
            }
        }

    }
}
