package test;

import asposetest.Docx;
import asposetest.DocxFactory;
import asposetest.KeyValueReplace;
import asposetest.TagResolver;
import java.net.URL;
import java.util.Arrays;
import org.junit.jupiter.api.Test;

public class DocxTest {

    @Test
    void deveGerarDocumento() throws Exception {
        final TagResolver tagResolver = (tag) -> "<<" + tag + ">>";

        URL urlTemplate = this.getClass().getResource("/template.docx");
        URL urlDocA = this.getClass().getResource("/a.docx");
        URL urlDocB = this.getClass().getResource("/b.docx");
        URL urlDocC = this.getClass().getResource("/c.docx");

        Docx docA = DocxFactory.fromUrl(urlDocA)
            .useTagResolver(tagResolver)
            .replace(new KeyValueReplace()
                .add("textoA", "Texto A do capítulo A")
                .add("textoB", "Texto B do capítulo A")
                .add("textoC", "Texto C do capítulo A"));

        Docx docB = DocxFactory.fromUrl(urlDocB, tagResolver)
            .replace("texto", "Texto do capítulo B");

        Docx docC = DocxFactory.fromUrl(urlDocC, tagResolver)
            .dataset()
                .add("ds", Arrays.asList("qqq", "www", "eee"))
                .add("ds2", Arrays.asList("zzz", "xxx", "ccc"))
                .apply();

        Docx newDoc = DocxFactory.fromUrl(urlTemplate, tagResolver)
            .insertAtBookmark("docA", docA)
            .replace("docB", docB)
            .paragraph()
            .append(docC, true);

        String path = this.getClass().getResource("/").getPath();
        newDoc.save(path + "teste-novo.docx");
    }

}
