package asposetest;

import java.io.InputStream;

public class DocumentoFactory {

    public static Documento create() throws Exception {
        return new DocumentoImpl();
    }

    public static Documento create(InputStream stream) throws Exception {
        return new DocumentoImpl(stream);
    }
}
