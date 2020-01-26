package test;

import asposetest.Documento;
import org.junit.jupiter.api.Test;

import java.util.Arrays;

public class DocumentoTest {

    @Test
    void deveGerarDocumento() throws Exception {
        Documento docA = Documento.create("/home/lucas/Downloads/a.docx")
            .replace("<<texto>>", "Texto do capítulo A");

        Documento docB = Documento.create("/home/lucas/Downloads/b.docx")
            .replace("<<texto>>", "Texto do capítulo B");

        Documento docC = Documento.create("/home/lucas/Downloads/c.docx")
            .dataset()
                .add("ds", Arrays.asList("qqq", "www", "eee"))
                .add("ds2", Arrays.asList("zzz", "xxx", "ccc"))
                .apply();

        Documento newDoc = Documento.create("/home/lucas/Downloads/template.docx")
            .replace("<<docA>>", docA)
            .lineBreak()
            .append(docB)
            .append(docC);

        newDoc.save("/home/lucas/Downloads/teste-novo.docx");
    }

}
