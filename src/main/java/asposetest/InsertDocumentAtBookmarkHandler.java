package asposetest;

import com.aspose.words.*;

class InsertDocumentAtBookmarkHandler implements IReplacingCallback {

    private DocumentoImpl documento;
    private DocumentoImpl documentoInserir;
    private Bookmark bookmark;

    public InsertDocumentAtBookmarkHandler(DocumentoImpl documento, DocumentoImpl documentoInserir, Bookmark bookmark) {
        this.documento = documento;
        this.documentoInserir = documentoInserir;
        this.bookmark = bookmark;
    }

    public int replacing(ReplacingArgs replacer) throws Exception {
        DocumentBuilder builder = documento.getBuilder();
        builder.moveToBookmark(bookmark.getName());

        builder.startBookmark(bookmark.getName());
        builder.insertDocument(documentoInserir.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);
        builder.endBookmark(bookmark.getName());

        return ReplaceAction.REPLACE;
    }
}