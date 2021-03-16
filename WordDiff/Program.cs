// MIT License
// 
// Copyright (c) 2021 Norbert Schulz
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

using System;
using System.IO;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

/// <summary>
/// Small utility for using Word as a doc diff tool that can be called from
/// from a console.
/// </summary>
namespace WordDiff
{
    class Program
    {
        /// <summary>
        /// Application Entry point
        /// 
        /// usage: WordDiff base-document modified-document
        /// </summary>
        /// <param name="args">function arguments, see usage above</param>
        /// <returns>0 success, otherwise error</returns>
        public static int Main(string[] args)
        {
            try
            {
                PrintBanner();

                if (args.Length != 2)
                {
                    DieWithMessage("usage: WordDiff <base document> <modified document>");
                }

                // Make relative path absolute, word otherwise will not find them 
                String baseFile = Path.GetFullPath(args[0]);
                String modifiedFile = Path.GetFullPath(args[1]);

                if (!File.Exists(baseFile))
                {
                    DieWithMessage("base document '" + baseFile + "' does not exist");
                }
                if (!File.Exists(modifiedFile))
                {
                    DieWithMessage("modified document '" + modifiedFile + "' does not exist");
                }

                // Launch word in diff mode
                RunWordDiff(baseFile, modifiedFile);
            }
            catch (Exception e)
            {
                DieWithMessage(e.Message);
            }

            return 0;
        }

        /// <summary>
        /// Terminate with message
        /// </summary>
        /// <param name="msg">message to display</param>
        private static void DieWithMessage(string msg)
        {
            Console.WriteLine(msg);
            Environment.Exit(1);
        }

        /// <summary>
        /// Start word via COM interop API's and show the diff result of 2 documents
        /// </summary>
        /// <param name="baseFile">Word document which is assumed to be the initial version.</param>
        /// <param name="editedFile">Word document which is assumed to contain changes to the initial version</param>
        private static void RunWordDiff(string baseFile, string editedFile)
        {
            // COM interop constants
            object interOpmissing = Type.Missing;
            object interOpTRUE = true;
            object interOpFALSE = false;
            object interOpBase = baseFile;
            object interOpEditedFile = editedFile;

            // Create a hidden word instance
            Console.WriteLine("Starting hidden Word instance.");
            Word.Application wordInstance = new Word.Application();
            wordInstance.Visible = false;

            Word.Document baseDoc = null;
            Word.Document editedeDoc = null;
            Word.Document diffDoc = null;
            try
            {
                // Open base and edited doc files
                Console.WriteLine("Loading base document " + baseFile);
                baseDoc = wordInstance.Documents.Open(ref interOpBase,
                       ref interOpmissing, ref interOpFALSE, ref interOpFALSE, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpmissing, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpTRUE, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpmissing);

                Console.WriteLine("Loading modified document " + editedFile);
                editedeDoc = wordInstance.Documents.Open(ref interOpEditedFile,
                       ref interOpmissing, ref interOpFALSE, ref interOpFALSE, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpmissing, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpTRUE, ref interOpmissing,
                       ref interOpmissing, ref interOpmissing, ref interOpmissing);

                // Create the diff document
                Console.WriteLine("Creating comparsion document ..." + editedFile);
                diffDoc = wordInstance.CompareDocuments(
                        baseDoc, editedeDoc,
                        Word.WdCompareDestination.wdCompareDestinationNew,
                        Word.WdGranularity.wdGranularityWordLevel,
                        true, true, true, true, true, true, true, true, true, true,
                        "",
                        true);
            }
            finally
            {
                // Close input documents
                if (null != baseDoc)
                {
                    baseDoc.Close(false);
                    Marshal.ReleaseComObject(baseDoc);
                    baseDoc = null;
                }

                if (null != editedeDoc)
                {
                    editedeDoc.Close(false);
                    Marshal.ReleaseComObject(editedeDoc);
                    editedeDoc = null;
                }
            }

            // show the diff result
            if (null != diffDoc)
            {
                wordInstance.Visible = true;
                diffDoc.Activate();
                SetForegroundWindow(wordInstance.Application.ActiveWindow.Hwnd);
            }
            else
            {
                wordInstance.Quit();
                Marshal.ReleaseComObject(wordInstance);
                DieWithMessage("unknown error during creation of diff document.");
            }
        }
        /// <summary>
        /// Display program banner to stdout
        /// </summary>
        private static void PrintBanner()
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var versionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);

            string banner= string.Format(
                "{0} - {1}  {2}\n",
                assembly.GetName().Name,
                assembly.GetName().Version,
                versionInfo.LegalCopyright
            );

            Console.WriteLine(banner);
        }

        /// <summary>
        /// C# wrapper for WIN32  SetForegroundWindow
        /// </summary>
        /// <param name="hwnd">window handle </param>
        /// <returns>Win32 BOOL</returns>
        [DllImport("User32.dll")]
        [return: MarshalAs(UnmanagedType.U4)]
        private static extern int SetForegroundWindow(int hwnd);
    }
}
