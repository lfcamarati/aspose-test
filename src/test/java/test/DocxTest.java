package test;

import asposetest.*;
import com.aspose.words.Font;
import com.aspose.words.*;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;

public class DocxTest {

    @Test
    void deveCriarDocumento() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Run hello = new Run(doc, "Hello");
        Run world = new Run(doc, "World");

        Font f = hello.getFont();
        f.setName("Courier New");
        f.setSize(36);
        f.setHighlightColor(Color.YELLOW);

        Font f2 = world.getFont();
        f2.setName("Courier New");
        f2.setSize(12);
        f2.setHighlightColor(Color.BLUE);

        Paragraph paragraph = builder.insertParagraph();
        paragraph.appendChild(hello);
        paragraph.appendField(" ");
        paragraph.appendChild(world);

        Path path = Paths.get("/home/lucas/Downloads/teste.docx");

        try (OutputStream fos = Files.newOutputStream(path)) {
            SaveOptions s = SaveOptions.createSaveOptions(SaveFormat.DOCX);
            doc.save(fos, s);
        }
    }

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

    @Test
    void deveFormatarDocumento() throws Exception {
//        InputStream isAcordao = this.getClass().getResourceAsStream("/acordaoRelacao.docx");
//        InputStream isAcordao = this.getClass().getResourceAsStream("/acordaoUnitario.docx");
        InputStream isAcordao = this.getClass().getResourceAsStream("/acordao.docx");

        InputStream isTplAcordao = this.getClass().getResourceAsStream("/templateAcordao.docx");

        Document srcDoc = new Document(isAcordao);
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        Document dstDoc = new Document(isTplAcordao);
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

        NodeCollection<Node> childNodes = srcDoc.getFirstSection().getBody().getChildNodes();

        for (Node node : childNodes) {
            if (node.getNodeType() != NodeType.PARAGRAPH && node.getNodeType() != NodeType.TABLE) {
                continue;
            }

            if (node.getText().contains("Assinado Eletronicamente") ||
                        node.getText().contains("Dados da Sessão:")) {
                break;
            }

            if (node.getNodeType() == NodeType.PARAGRAPH) {
                ParagraphFormat paragraphFormat = dstBuilder.getParagraphFormat();
                Font font = paragraphFormat.getStyle().getFont();

                paragraphFormat.setLeftIndent(convertCentimetersToPoints(0));
                paragraphFormat.setFirstLineIndent(convertCentimetersToPoints(1));
                paragraphFormat.setAlignment((node.getText().length() > 90) ? ParagraphAlignment.JUSTIFY : ParagraphAlignment.LEFT);
                font.setSize(12.0);
                font.setColor(Color.BLACK);
                font.setBold(false);

                dstBuilder.writeln(node.getText().trim());
                dstBuilder.getFont().clearFormatting();
                dstBuilder.getParagraphFormat().clearFormatting();
            } else {
                Node newNode = dstDoc.importNode(node, true, ImportFormatMode.USE_DESTINATION_STYLES);
                dstDoc.getFirstSection().getBody().appendChild(newNode);
            }

            dstBuilder.moveToDocumentEnd();
        }

        Path path = Paths.get("/home/lucas/Downloads/novoAcordao.docx");

        try (OutputStream outputStream = Files.newOutputStream(path)) {
            SaveOptions saveOptions = SaveOptions.createSaveOptions(SaveFormat.DOCX);

            dstDoc.getBuiltInDocumentProperties().clear();
            dstDoc.getCustomDocumentProperties().clear();
            dstDoc.save(outputStream, saveOptions);
        } catch (Exception e) {
            throw new DocxException("Erro ao salvar docx", e);
        }
    }

    @Test
    void deveSubstituirTags() throws Exception {
        final TagResolver tagResolver = (tag) -> "<<" + tag + ">>";

        InputStream inputStream = this.getClass().getResourceAsStream("/template_portaria_substituicao.docx");

        Docx docA = DocxFactory.fromInputStream(inputStream)
            .useTagResolver(tagResolver)
            .replace(new KeyValueReplace()
                 .add("NUMERO_PORTARIA", "Aaaaa")
                 .add("DATA_ASSINATURA_DOCUMENTO", "Bbbbb")
                 .add("ARTIGO_CARGO_ASSINATURA", "Ccccc")
                 .add("ARTIGO_CARGO_NOME_MINISTRO_SUBSTITUTO", "Ddddd"));

        Path path = Paths.get("/home/lucas/Downloads/teste-novo.docx");

        try (OutputStream fos = Files.newOutputStream(path)) {
            fos.write(docA.toByteArray());
        }
    }

    private static double convertCentimetersToPoints(int centimeters) {
        return centimeters * 72.0 * 0.3937;
    }
}
