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

        Documento docA = Documento.create(urlDocA)
            .replace("<<texto>>", "Texto do capítulo A");

        Documento docB = Documento.create(urlDocB)
            .replace("<<texto>>", "Texto do capítulo B");

        Documento docC = Documento.create(urlDocC)
            .dataset()
                .add("ds", Arrays.asList("qqq", "www", "eee"))
                .add("ds2", Arrays.asList("zzz", "xxx", "ccc"))
                .apply();

        Documento newDoc = Documento.create(urlTemplate)
            .replace("<<docA>>", docA)
            .lineBreak()
            .append(docB)
            .append(docC);

        String path = this.getClass().getResource("/").getPath();
        newDoc.save(path + "teste-novo.docx");
    }

}
