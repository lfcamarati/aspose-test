package asposetest;

public class DocumentoFactory {

    public static Documento create() throws Exception {
        return new DocumentoImpl();
    }

    public static Documento create(String fullPath) throws Exception {
        return new DocumentoImpl(fullPath);
    }
}
