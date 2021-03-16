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
        static int Main(string[] args)
        {
            try
            {
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
        static private void DieWithMessage(string msg)
        {
            Console.WriteLine(msg);
            Environment.Exit(1);
        }

        /// <summary>
        /// Start word via COM interop API's and show the diff result of 2 documents
        /// </summary>
        /// <param name="baseFile">Word document which is assumed to be the initial version.</param>
        /// <param name="editedFile">Word document which is assumed to contain changes to the initial version</param>
        static void RunWordDiff(string baseFile, string editedFile)
        {
            // COM interop constants
            object interOpmissing = Type.Missing;
            object interOpTRUE = true;
            object interOpFALSE = false;
            object interOpBase = baseFile;
            object interOpEditedFile = editedFile;

            // Create a hidden word instance
            Word.Application wordInstance = new Word.Application();
            wordInstance.Visible = false;

            // open base and edited files
            Word.Document baseDoc = wordInstance.Documents.Open(ref interOpBase,
                   ref interOpmissing, ref interOpFALSE, ref interOpFALSE, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpmissing, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpTRUE, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpmissing);
            Word.Document editedeDoc = wordInstance.Documents.Open(ref interOpEditedFile,
                   ref interOpmissing, ref interOpFALSE, ref interOpFALSE, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpmissing, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpTRUE, ref interOpmissing,
                   ref interOpmissing, ref interOpmissing, ref interOpmissing);

            // create diff document
            Word.Document doc = wordInstance.CompareDocuments(
                    baseDoc, editedeDoc,
                    Word.WdCompareDestination.wdCompareDestinationNew,
                    Word.WdGranularity.wdGranularityWordLevel,
                    true, true, true, true, true, true, true, true, true, true,
                    "WordDiff",
                    true);

            //close inputs
            baseDoc.Close();
            editedeDoc.Close();

            // show diff result
            wordInstance.Visible = true;
            wordInstance.Activate();
        }
    }
}
