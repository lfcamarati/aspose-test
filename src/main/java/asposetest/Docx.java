package asposetest;

import java.util.function.Supplier;

public interface Docx {

    Docx replace(String tag, String newText);

    Docx replace(String tag, Supplier<String> fnNewText);

    Docx clear(String tag);

    Docx replace(KeyValueReplace keyValueReplace);

    Docx replace(String tag, Docx docx);

    Docx remove(String tag);

    Docx insertAtBookmark(String bookmarkName, Docx doc);

    Docx insertAtBookmark(String bookmarkName, Docx doc, boolean removeBookmarkContent);

    Docx removeBookmark(String bookmarkName);

    Docx removeBookmark(String bookmarkName, boolean removeBookmarkContent);

    Docx removeAllBookmarks();

    Docx append(Docx doc, boolean insertLineBreak);

    Docx useTagResolver(TagResolver tagResolver);

    Docx append(Docx docx);

    Docx lineBreak();

    Dataset dataset();

    PageSetup pageSetup();

    byte[] toByteArray();

    String getBookmarkText(String bookmarkName);
}
