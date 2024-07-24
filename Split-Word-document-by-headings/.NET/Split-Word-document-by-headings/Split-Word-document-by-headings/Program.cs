using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

using (FileStream inputStream = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read))
{
    //Open an existing Word document
    using (WordDocument document = new WordDocument(inputStream, FormatType.Docx))
    {
        WordDocument newDocument = null;
        WSection newSection = null;
        int headingIndex = 0;
        //Iterate each section in the Word document
        foreach (WSection section in document.Sections)
        {
            //Clone the section without items and add into new document
            if (newDocument != null)
                newSection = AddSection(newDocument, section);
            //Iterate each child entity in the Word document
            foreach (TextBodyItem item in section.Body.ChildEntities)
            {
                //If item is paragraph, then check for heading style and split
                //else, add the item into new document
                if (item is WParagraph)
                {
                    WParagraph paragraph = item as WParagraph;
                    //If paragraph has Heading 1 style, then save the traversed content as separate document
                    //And create new document for new heading content
                    if (paragraph.StyleName == "Heading 1")
                    {
                        if (newDocument != null)
                        {
                            //Save the Word document
                            string fileName = @"../../../Document" + (headingIndex + 1) + ".docx";
                            SaveWordDocument(newDocument, fileName);
                            headingIndex++;
                        }
                        //Create new document for new heading content
                        newDocument = new WordDocument();
                        newSection = AddSection(newDocument, section);
                        AddEntity(newSection, paragraph);
                    }
                    else if (newDocument != null)
                        AddEntity(newSection, paragraph);
                }
                else
                    AddEntity(newSection, item);
            }
        }
        //Save the remaining content as separate document
        if (newDocument != null)
        {
            //Save the Word document
            string fileName = @"../../../Document" + (headingIndex + 1) + ".docx";
            SaveWordDocument(newDocument, fileName);
        }
    }
}
/// <summary>
/// Clone and add the section without content into new Word document
/// </summary>
static WSection AddSection(WordDocument newDocument, WSection section)
{
    //Create new section based on original document
    WSection newSection = section.Clone();
    //Remove body items from section
    newSection.Body.ChildEntities.Clear();
    //Remove headers and footers
    newSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
    newSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
    newSection.HeadersFooters.OddFooter.ChildEntities.Clear();
    newSection.HeadersFooters.OddHeader.ChildEntities.Clear();
    newSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
    newSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
    //Add cloned section into new document
    newDocument.Sections.Add(newSection);
    return newSection;
}
/// <summary>
/// Add item into the section.
/// </summary>
static void AddEntity(WSection newSection, Entity entity)
{
    //Add cloned item into the newly created section
    newSection.Body.ChildEntities.Add(entity.Clone());
}
/// <summary>
/// Save the Word document.
/// </summary>
static void SaveWordDocument(WordDocument newDocument, string fileName)
{
    using (FileStream outputStream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
    {
        //Save file stream as Word document
        newDocument.Save(outputStream, FormatType.Docx);
        //Close the document
        newDocument.Close();
        newDocument = null;
    }
}