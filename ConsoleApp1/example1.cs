using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConsoleApp1 {
    class example1 {

        public void test() {
            using (RichEditDocumentServer parentWordProcessor = new RichEditDocumentServer()) {
                                
                parentWordProcessor.Document.BeginUpdate();
                parentWordProcessor.LoadDocument("..\\..\\documents\\main_document.docx");                
                string documentTemplate = Path.Combine("..\\..\\documents\\template_part.docx");

                List<string> comments = new List<string>();

                //Loops max times to replicate the problem
                int max = 4;
                for (int i=0; i < max; i++) {
                    //create child wordprocessor
                    using (RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer()) {
                        //load document to child wordprocessor
                        childWordPrecessor.LoadDocumentTemplate(documentTemplate);

                        //Comments can be accessed here
                        string comment = this.getCommentText(childWordPrecessor.Document.Comments[0]);
                        comments.Add(comment);
                        //childWordPrecessor.Document.Comments.Remove(childWordPrecessor.Document.Comments[0]);

                        //but are lost here
                        parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                        //parentWordProcessor.Document.AppendDocumentContent(childWordPrecessor.Document.Range);
                    }

                    
                    //creating a page break at a particular place where the string appears
                    Regex r = new Regex("{PBR}");
                    if (max < 4) {                        
                        parentWordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
                    } else {
                        parentWordProcessor.Document.ReplaceAll(r, "");
                    }
                }
                                              

                parentWordProcessor.Document.EndUpdate();
                parentWordProcessor.SaveDocument("..\\..\\documents\\document_generated.docx", DocumentFormat.OpenXml);
                Process.Start(new ProcessStartInfo("..\\..\\documents\\document_generated.docx") { UseShellExecute = true });
            }

            
        }

        public DocumentRange getTextRange(string search, RichEditDocumentServer wp = null) {          
                Regex myRegEx = new Regex(search);                
                return wp.Document.FindAll(myRegEx).First();          
        }

        public string getCommentText(Comment comment) {
            SubDocument doc = comment.BeginUpdate();
            string commentText = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
            comment.EndUpdate(doc);

            return commentText;
        }
    }
}
