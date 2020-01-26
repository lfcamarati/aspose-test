package asposetest;

import java.net.URL;

public interface Documento {

    Documento replace(String tag, String newText) throws Exception;

    Documento replace(String tag, Documento documento) throws Exception;

    Documento append(Documento documento);

    Dataset dataset();

    Documento lineBreak();

    String getBookmarkText(String bookmarkName) throws Exception;

    void save(String newPath) throws Exception;

    static Documento create() throws Exception {
        return DocumentoFactory.create();
    }

    static Documento create(URL fullPath) throws Exception {
        return create(fullPath.getPath());
    }

    static Documento create(String fullPath) throws Exception {
        return DocumentoFactory.create(fullPath);
    }
}
