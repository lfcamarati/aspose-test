package test;

import asposetest.Documento;
import org.junit.jupiter.api.Test;

import java.net.URL;
import java.util.Arrays;

public class DocumentoTest {

    @Test
    void deveGerarDocumento() throws Exception {
        URL urlTemplate = this.getClass().getResource("/template.docx");
        URL urlDocA = this.getClass().getResource("/a.docx");
        URL urlDocB = this.getClass().getResource("/b.docx");
        URL urlDocC = this.getClass().getResource("/c.docx");

        Documento docA = Documento.fromUrl(urlDocA)
            .replace("<<texto>>", "Texto do capítulo A");

        Documento docB = Documento.fromUrl(urlDocB)
            .replace("<<texto>>", "Texto do capítulo B");

        Documento docC = Documento.fromUrl(urlDocC)
            .dataset()
                .add("ds", Arrays.asList("qqq", "www", "eee"))
                .add("ds2", Arrays.asList("zzz", "xxx", "ccc"))
                .apply();

        Documento newDoc = Documento.fromUrl(urlTemplate)
            .insertAtBookmark("docA", docA)
            .replace("<<docB>>", docB)
            .append(docC);

        String path = this.getClass().getResource("/").getPath();
        newDoc.save(path + "teste-novo.docx");
    }

}
