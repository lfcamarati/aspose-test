package asposetest;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Base64;

public class DocumentoFactory {

    public static Documento create() throws Exception {
        return new DocumentoImpl();
    }

    public static Documento create(InputStream stream) throws Exception {
        return new DocumentoImpl(stream);
    }

    public static Documento fromBase64String(String contentBase64) throws Exception {
        byte[] bytes = Base64.getDecoder().decode(contentBase64);
        return create(new ByteArrayInputStream(bytes));
    }
}
