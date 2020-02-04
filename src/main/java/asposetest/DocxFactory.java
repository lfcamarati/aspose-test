package asposetest;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.net.URL;
import java.util.Base64;

public class DocxFactory {

    public static Docx createEmpty() throws Exception {
        return new DocxImpl();
    }

    public static Docx fromInputStream(InputStream stream, TagResolver tagResolver) throws Exception {
        return new DocxImpl(stream, tagResolver);
    }

    public static Docx fromInputStream(InputStream stream) throws Exception {
        return fromInputStream(stream, null);
    }

    public static Docx fromUrl(URL fullPath) throws Exception {
        return fromInputStream(fullPath.openStream());
    }

    public static Docx fromUrl(URL fullPath, TagResolver tagResolver) throws Exception {
        return fromInputStream(fullPath.openStream(), tagResolver);
    }

    public static Docx fromBase64String(String contentBase64) throws Exception {
        byte[] bytes = Base64.getDecoder().decode(contentBase64);
        return fromInputStream(new ByteArrayInputStream(bytes));
    }

    public static Docx fromBase64String(String contentBase64, TagResolver tagResolver) throws Exception {
        byte[] bytes = Base64.getDecoder().decode(contentBase64);
        return fromInputStream(new ByteArrayInputStream(bytes), tagResolver);
    }
}
