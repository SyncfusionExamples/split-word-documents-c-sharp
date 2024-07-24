using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Text.RegularExpressions;

using (FileStream fileStreamPath = new FileStream(@"../../../Data/Template.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    //Open an existing Word document
    using (WordDocument document = new WordDocument(fileStreamPath, FormatType.Docx))
    {
        //Find all the placeholder text in the Word document
        TextSelection[] textSelections = document.FindAll(new Regex("<<(.*)>>"));
        if (textSelections != null)
        {
            //Unique ID for each bookmark
            int bkmkId = 1;
            //Collection to hold the newly inserted bookmarks
            List<string> bookmarks = new List<string>();

            #region Add bookmark start and end in the place of placeholder text
            //Iterate each text selection
            for (int i = 0; i < textSelections.Length; i++)
            {
                //Get the placeholder as WTextRange
                WTextRange textRange = textSelections[i].GetAsOneRange();
                //Get the index of the placeholder text
                WParagraph startParagraph = textRange.OwnerParagraph;
                int index = startParagraph.ChildEntities.IndexOf(textRange);
                string bookmarkName = "Bookmark_" + bkmkId;

                //Add new bookmark to bookmarks collection
                bookmarks.Add(bookmarkName);
                //Create bookmark start
                BookmarkStart bkmkStart = new BookmarkStart(document, bookmarkName);
                //Insert the bookmark start before the start placeholder
                startParagraph.ChildEntities.Insert(index, bkmkStart);
                //Remove the placeholder text
                textRange.Text = string.Empty;

                i++;
                //Get the placeholder as WTextRange
                textRange = textSelections[i].GetAsOneRange();
                //Get the index of the placeholder text
                WParagraph endParagraph = textRange.OwnerParagraph;
                index = endParagraph.ChildEntities.IndexOf(textRange);

                //Create bookmark end
                BookmarkEnd bkmkEnd = new BookmarkEnd(document, bookmarkName);
                //Insert the bookmark end after the end placeholder
                endParagraph.ChildEntities.Insert(index + 1, bkmkEnd);
                bkmkId++;
                //Remove the placeholder text
                textRange.Text = string.Empty;
            }
            #endregion

            #region Split document based on newly inserted bookmarks
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            int fileIndex = 1;
            foreach (string bookmark in bookmarks)
            {
                //Move the virtual cursor to the location before the end of the bookmark
                bookmarksNavigator.MoveToBookmark(bookmark);
                //Get the bookmark content as WordDocumentPart
                WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                //Save the WordDocumentPart as separate Word document
                using (WordDocument newDocument = wordDocumentPart.GetAsWordDocument())
                {
                    //Save the Word document to file stream
                    using (FileStream outputFileStream = new FileStream(@"../../../Placeholder_" + fileIndex + ".docx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        newDocument.Save(outputFileStream, FormatType.Docx);
                    }
                }
                fileIndex++;
            }
            #endregion
        }
    }
}