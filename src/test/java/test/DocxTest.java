package test;

import asposetest.Docx;
import asposetest.DocxFactory;
import asposetest.KeyValueReplace;
import asposetest.TagResolver;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import org.junit.jupiter.api.Test;

public class DocxTest {

    @Test
    void deveGerarDocumento() throws Exception {
        final TagResolver tagResolver = (tag) -> "<<" + tag + ">>";

        InputStream urlTemplate = this.getClass().getResourceAsStream("/template.docx");
        InputStream urlDocA = this.getClass().getResourceAsStream("/a.docx");
        InputStream urlDocB = this.getClass().getResourceAsStream("/b.docx");
        InputStream urlDocC = this.getClass().getResourceAsStream("/c.docx");

        Docx docA = DocxFactory.fromInputStream(urlDocA)
            .useTagResolver(tagResolver)
            .replace(new KeyValueReplace()
                .add("textoA", "Texto A do capítulo A")
                .add("textoB", "Texto B do capítulo A")
                .add("textoC", "Texto C do capítulo A"));

        Docx docB = DocxFactory.fromInputStream(urlDocB, tagResolver)
            .replace("texto", "Texto do capítulo B");

        Docx docC = DocxFactory.fromInputStream(urlDocC, tagResolver)
            .dataset()
                .add("ds", Arrays.asList("qqq", "www", "eee"))
                .add("ds2", Arrays.asList("zzz", "xxx", "ccc"))
                .apply();

        Docx newDoc = DocxFactory.fromInputStream(urlTemplate, tagResolver)
            .insertAtBookmark("docA", docA)
            .replace("docB", docB)
            .lineBreak()
            .append(docC, true);

        Path path = Paths.get("teste-novo.docx");

        try (OutputStream fos = Files.newOutputStream(path)) {
            fos.write(newDoc.toByteArray());
        }
    }
}
