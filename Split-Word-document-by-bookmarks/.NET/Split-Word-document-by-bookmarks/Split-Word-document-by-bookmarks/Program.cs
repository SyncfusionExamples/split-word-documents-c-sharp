using Syncfusion.DocIO.DLS;
using Syncfusion.DocIO;

//Load an existing Word document
using (FileStream fileStreamPath = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Create the bookmark navigator instance to access the bookmark
        BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
        BookmarkCollection bookmarkCollection = document.Bookmarks;

        //Iterate each bookmark in Word document
        foreach (Bookmark bookmark in bookmarkCollection)
        {
            //Move the virtual cursor to the location before the end of the bookmark
            bookmarksNavigator.MoveToBookmark(bookmark.Name);

            //Get the bookmark content as WordDocumentPart
            WordDocumentPart documentPart = bookmarksNavigator.GetContent();

            //Save the WordDocumentPart as separate Word document
            using (WordDocument newDocument = documentPart.GetAsWordDocument())
            {
                //Save the Word document to file stream
                using (FileStream outputFileStream = new FileStream(@"../../../" + bookmark.Name + ".docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    newDocument.Save(outputFileStream, FormatType.Docx);
                }
            }
        }
    }
}