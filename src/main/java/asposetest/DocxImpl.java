package asposetest;

import com.aspose.words.Bookmark;
import com.aspose.words.BookmarkCollection;
import com.aspose.words.BreakType;
import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceDirection;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SaveFormat;
import com.aspose.words.SaveOptions;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;

class DocxImpl implements Docx, DocxInternal {

    private static final String EMPTY = "";

    private Document document;
    private DocumentBuilder builder;
    private PageSetup pageSetup;
    private TagResolver tagResolver;

    DocxImpl() {
        init(null, null);
    }

    DocxImpl(InputStream stream, TagResolver tagResolver) {
        init(stream, tagResolver);
        initPageSetup();
    }

    private void init(InputStream stream, TagResolver tagResolver) {
        try {
            this.document = (stream == null) ? new Document() : new Document(stream);
            this.builder = new DocumentBuilder(document);
            this.tagResolver = tagResolver;
            initPageSetup();
        } catch(Exception ex) {
            throw new DocxException(ex);
        }
    }

    private void initPageSetup() {
        this.pageSetup = new PageSetupImpl(this);
    }

    @Override
    public Docx replace(String tag, String newText) {
        try {
            document.getRange().replace(resolveTag(tag), newText, getFindReplaceOptions());
            return this;
        } catch (Exception ex) {
            throw new DocxException("Erro ao substituir tag [tag=" + tag + ", newText=" + newText + "]", ex);
        }
    }

    @Override
    public Docx replace(String tag, Supplier<String> fnNewText) {
        return replace(tag, fnNewText.get());
    }

    @Override
    public Docx clear(String tag) {
        return replace(tag, EMPTY);
    }

    @Override
    public Docx replace(KeyValueReplace keyValueReplace) {
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
    public Docx replace(String tag, Docx doc) {
        try {
            DocxImpl documento = (DocxImpl) doc;
            document.getRange().replace(resolveTag(tag), EMPTY, getFindReplaceOptions(new InsertDocumentAtReplaceTextHandler(documento)));
            return this;
        } catch (Exception ex) {
            throw new DocxException("Erro ao substituir tag [tag=" + tag + "]", ex);
        }
    }

    @Override
    public Docx remove(String tag) {
        try {
            document.getRange().replace(resolveTag(tag), EMPTY, getFindReplaceOptions(new RemoveNodeHandler()));
            return this;
        } catch (Exception ex) {
            throw new DocxException("Erro ao remover tag [tag=" + tag + "]", ex);
        }
    }

    @Override
    public Docx insertAtBookmark(String bookmarkName, Docx doc) {
        return insertAtBookmark(bookmarkName, doc, true);
    }

    @Override
    public Docx insertAtBookmark(String bookmarkName, Docx doc, boolean removeBookmarkContent) {
        try {
            DocxImpl documento = (DocxImpl) doc;
            Bookmark bookmark = getBookmark(bookmarkName);

            if(removeBookmarkContent) {
                bookmark.setText(EMPTY);
            }

            DocumentBuilder builder = getBuilder();
            builder.moveToBookmark(bookmarkName, false, false);
            bookmark.remove();

            builder.startBookmark(bookmarkName);
            builder.insertDocument(documento.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);
            builder.endBookmark(bookmarkName);

            return this;
        } catch (Exception ex) {
            throw new DocxException("Erro ao inserir conteúdo em bookmark [bookmarkName=" + bookmarkName + "]", ex);
        }
    }

    @Override
    public Docx removeBookmark(String bookmarkName) {
        return removeBookmark(bookmarkName, false);
    }

    @Override
    public Docx removeBookmark(String bookmarkName, boolean removeBookmarkContent) {
        try {
            Bookmark bookmark = getBookmark(bookmarkName);
            CompositeNode parentNode = bookmark.getBookmarkStart().getParentNode();

            if (removeBookmarkContent) {
                bookmark.setText(EMPTY);
            }

            bookmark.remove();

            if (removeBookmarkContent) {
                parentNode.removeAllChildren();
                parentNode.remove();
            }

            return this;
        } catch (Exception ex) {
            throw new DocxException("Erro ao remover bookmark [bookmarkName=" + bookmarkName + "]", ex);
        }
    }

    @Override
    public Docx removeAllBookmarks() {
        try {
            BookmarkCollection bookmarks = document.getRange().getBookmarks();
            Iterator<Bookmark> iterator = bookmarks.iterator();

            while (iterator.hasNext()) {
                removeBookmark(iterator.next().getName());
            }

            return null;
        } catch (Exception ex) {
            throw new DocxException("Erro ao remover bookmarks", ex);
        }
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
    public PageSetup pageSetup() {
        return pageSetup;
    }

    @Override
    public byte[] toByteArray() {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            SaveOptions s = SaveOptions.createSaveOptions(SaveFormat.DOCX);
            document.getBuiltInDocumentProperties().clear();
            document.getCustomDocumentProperties().clear();
            document.save(outputStream, s);

            return outputStream.toByteArray();
        } catch (Exception e) {
            throw new DocxException("Erro ao salvar docx", e);
        }
    }

    @Override
    public String getBookmarkText(String bookmarkName) {
        try {
            Optional<Bookmark> bookmark = Optional.ofNullable(getBookmark(bookmarkName));

            if (bookmark.isPresent()) {
                byte[] bytes = bookmark.get().getText().getBytes(StandardCharsets.UTF_8);
                return new String(bytes, StandardCharsets.UTF_8);
            }

            return null;
        } catch (Exception ex) {
            throw new DocxException("Erro ao obter texto de bookmark [bookmarkName=" + bookmarkName + "]", ex);
        }
    }

    private Bookmark getBookmark(String bookmarkName) {
        try {
            return document.getRange().getBookmarks().get(bookmarkName);
        } catch (Exception ex) {
            throw new DocxException("Erro ao obter bookmark [bookmarkName=" + bookmarkName + "]", ex);
        }
    }

    @Override
    public DocumentBuilder getBuilder() {
        return builder;
    }

    @Override
    public Document getDocument() {
        return document;
    }

    @Override
    public Docx getDocx() {
        return this;
    }
}
