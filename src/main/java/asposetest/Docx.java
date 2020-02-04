package asposetest;

public interface Docx {

    Docx replace(String tag, String newText) throws Exception;

    Docx replace(KeyValueReplace keyValueReplace) throws Exception;

    Docx replace(String tag, Docx docx) throws Exception;

    Docx insertAtBookmark(String bookmarkName, Docx doc) throws Exception;

    Docx insertAtBookmark(String bookmarkName, Docx doc, boolean removeBookmarkContent) throws Exception;

    Docx append(Docx doc, boolean insertLineBreak);

    Docx useTagResolver(TagResolver tagResolver);

    Docx append(Docx docx);

    Dataset dataset();

    Docx lineBreak();

    Docx paragraph();

    void save(String newPath) throws Exception;
}
