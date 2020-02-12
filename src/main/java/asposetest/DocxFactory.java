package asposetest;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Base64;

public class DocxFactory {

    public static Docx createEmpty() {
        return new DocxImpl();
    }

    public static Docx fromInputStream(InputStream stream, TagResolver tagResolver) {
        return new DocxImpl(stream, tagResolver);
    }

    public static Docx fromInputStream(InputStream stream) {
        return fromInputStream(stream, null);
    }

    public static Docx fromBase64String(String contentBase64) {
        byte[] bytes = Base64.getDecoder().decode(contentBase64);
        return fromInputStream(new ByteArrayInputStream(bytes));
    }

    public static Docx fromBase64String(String contentBase64, TagResolver tagResolver) {
        byte[] bytes = Base64.getDecoder().decode(contentBase64);
        return fromInputStream(new ByteArrayInputStream(bytes), tagResolver);
    }
}
