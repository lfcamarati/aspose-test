package asposetest;

import com.aspose.words.*;

/**
 * https://docs.aspose.com/display/wordsjava/Find+and+Replace
 * https://docs.aspose.com/display/wordsjava/How+to++Insert+a+Document+into+another+Document
 * https://docs.aspose.com/display/wordsjava/Use+DocumentBuilder+to+Insert+Document+Elements
 */

class DocumentoImpl implements Documento {

    private Document document;
    private DocumentBuilder builder;

    DocumentoImpl() throws Exception {
        document = new Document();
        builder = new DocumentBuilder(document);
    }

    DocumentoImpl(String fullPath) throws Exception {
        document = new Document(fullPath);
        builder = new DocumentBuilder(document);
    }

    @Override
    public Documento replace(String text, String newText) throws Exception {
        document.getRange().replace(text, newText, getFindReplaceOptions());
        return this;
    }

    @Override
    public Documento replace(String replaceText, Documento doc) throws Exception {
        DocumentoImpl documento = (DocumentoImpl) doc;

        FindReplaceOptions findReplaceOptions = getFindReplaceOptions();
        findReplaceOptions.setReplacingCallback(new InsertDocumentAtReplaceTextHandler(documento));

        document.getRange().replace(replaceText, "", findReplaceOptions);

        return this;
    }

    private static FindReplaceOptions getFindReplaceOptions() {
        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.setDirection(FindReplaceDirection.FORWARD);
        findReplaceOptions.setFindWholeWordsOnly(true);
        findReplaceOptions.setMatchCase(true);

        return findReplaceOptions;
    }

    @Override
    public Dataset dataset() {
        return new DatasetImpl(this);
    }

    @Override
    public Documento append(Documento doc) {
        lineBreak();
        DocumentoImpl documento = (DocumentoImpl) doc;
        builder.insertDocument(documento.document, ImportFormatMode.KEEP_SOURCE_FORMATTING);
        return this;
    }

    @Override
    public Documento lineBreak() {
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.LINE_BREAK);
        return this;
    }

    @Override
    public String getBookmarkText(String bookmarkName) throws Exception {
        Bookmark bookmark = document.getRange().getBookmarks().get(bookmarkName);
        return bookmark.getText();
    }

    @Override
    public void save(String newPath) throws Exception {
        document.save(newPath);
    }

    public DocumentBuilder getBuilder() {
        return builder;
    }

    public Document getDocument() {
        return document;
    }
}
