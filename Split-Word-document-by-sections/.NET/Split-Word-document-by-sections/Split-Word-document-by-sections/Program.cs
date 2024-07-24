using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Load the template document as stream
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        int fileId = 1;
        //Iterate each section from Word document
        foreach (WSection section in document.Sections)
        {
            //Create new Word document
            using (WordDocument newDocument = new WordDocument())
            {
                //Add cloned section into new Word document
                newDocument.Sections.Add(section.Clone());
                //Save the Word document to MemoryStream
                using (FileStream outputStream = new FileStream(@"../../../Section" + fileId + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    newDocument.Save(outputStream, FormatType.Docx);
                }
            }
            fileId++;
        }
    }
}