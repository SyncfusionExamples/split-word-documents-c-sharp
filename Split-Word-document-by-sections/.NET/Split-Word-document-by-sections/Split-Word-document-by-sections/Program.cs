using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing Word document
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        int sectionNumber = 1;
        //Iterate each section from Word document
        foreach (WSection section in document.Sections)
        {
            //Create new Word document
            using (WordDocument newDocument = new WordDocument())
            {
                //Clone and add section from one Word document to another
                newDocument.Sections.Add(section.Clone());
                //Save the Word document
                using (FileStream outputStream = new FileStream(@"../../../Section" + sectionNumber + ".docx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    newDocument.Save(outputStream, FormatType.Docx);
                }
            }
            sectionNumber++;
        }
    }
}