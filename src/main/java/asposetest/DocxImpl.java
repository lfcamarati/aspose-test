package asposetest;

import com.aspose.words.Bookmark;
import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ImportFormatMode;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * https://docs.aspose.com/display/wordsjava/Find+and+Replace
 * https://docs.aspose.com/display/wordsjava/How+to++Insert+a+Document+into+another+Document
 * https://docs.aspose.com/display/wordsjava/Use+DocumentBuilder+to+Insert+Document+Elements
 */

class DocxImpl implements Docx {

    private Document document;
    private DocumentBuilder builder;
    private TagResolver tagResolver;

    DocxImpl() throws Exception {
        document = new Document();
        builder = new DocumentBuilder(document);
    }

    DocxImpl(InputStream stream, TagResolver tagResolver) throws Exception {
        document = new Document(stream);
        builder = new DocumentBuilder(document);
        this.tagResolver = tagResolver;
    }

    @Override
    public Docx replace(String tag, String newText) throws Exception {
        document.getRange().replace(resolveTag(tag), newText, getFindReplaceOptions());
        return this;
    }

    @Override
    public Docx replace(KeyValueReplace keyValueReplace) throws Exception {
        for(Map.Entry<String, String> entry : keyValueReplace.getValues().entrySet()) {
            replace(entry.getKey(), entry.getValue());
        }

        if (keyValueReplace.hasDatasetValues()) {
            final Dataset dataset = dataset();

            for (Map.Entry<String, List<String>> entry : keyValueReplace.getDatasetValues().entrySet()) {
                dataset.add(entry.getKey(), entry.getValue());
            }
        }

        return this;
    }

    @Override
    public Docx replace(String tag, Docx doc) throws Exception {
        DocxImpl documento = (DocxImpl) doc;
        document.getRange().replace(resolveTag(tag), "", getFindReplaceOptions(new InsertDocumentAtReplaceTextHandler(documento)));

        return this;
    }

    @Override
    public Docx insertAtBookmark(String bookmarkName, Docx doc) throws Exception {
        return insertAtBookmark(bookmarkName, doc, true);
    }

    @Override
    public Docx insertAtBookmark(String bookmarkName, Docx doc, boolean removeBookmarkContent) throws Exception {
        DocxImpl documento = (DocxImpl) doc;
        Bookmark bookmark = getBookmark(bookmarkName);

        if(removeBookmarkContent) {
            bookmark.setText("");
        }

        DocumentBuilder builder = getBuilder();
        builder.moveToBookmark(bookmarkName, false, false);
        bookmark.remove();

        builder.startBookmark(bookmarkName);
        builder.insertDocument(documento.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.endBookmark(bookmarkName);

        return this;
    }

    private String resolveTag(String tag) {
        if(Objects.isNull(tagResolver)) {
            return tag;
        }

        return tagResolver.apply(tag);
    }

    private static FindReplaceOptions getFindReplaceOptions(IReplacingCallback callback) {
        FindReplaceOptions findReplaceOptions = getFindReplaceOptions();
        findReplaceOptions.setReplacingCallback(callback);

        return findReplaceOptions;
    }

    private static FindReplaceOptions getFindReplaceOptions() {
        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.setDirection(FindReplaceDirection.FORWARD);
        findReplaceOptions.setMatchCase(true);

        return findReplaceOptions;
    }

    @Override
    public Dataset dataset() {
        return new DatasetImpl(this);
    }

    @Override
    public Docx append(Docx doc) {
        builder.moveToDocumentEnd();
        DocxImpl documento = (DocxImpl) doc;
        builder.insertDocument(documento.document, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        return this;
    }

    @Override
    public Docx append(Docx doc, boolean insertLineBreak) {
        return lineBreak().append(doc);
    }

    @Override
    public Docx useTagResolver(TagResolver tagResolver) {
        this.tagResolver = tagResolver;
        return this;
    }

    @Override
    public Docx lineBreak() {
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.LINE_BREAK);
        return this;
    }

    @Override
    public Docx paragraph() {
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        return this;
    }

    @Override
    public void save(String newPath) throws Exception {
        document.save(newPath);
    }

    private Bookmark getBookmark(String bookmarkName) throws Exception {
        return document.getRange().getBookmarks().get(bookmarkName);
    }

    public DocumentBuilder getBuilder() {
        return builder;
    }

    public Document getDocument() {
        return document;
    }
}
