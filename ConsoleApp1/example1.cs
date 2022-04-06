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

                //create child wordprocessor
                string documentTemplate = Path.Combine("..\\..\\documents\\template_part.docx");
                using (RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer()) {
                    //load document to child wordprocessor
                    childWordPrecessor.LoadDocumentTemplate(documentTemplate);

                    childWordPrecessor.Document.BeginUpdate();

                    //Appears ok
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();
                    //Appears ok
                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();
                    //Doesnt Appears
                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();

                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();

                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);


                    childWordPrecessor.Document.EndUpdate();
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
    }
}
