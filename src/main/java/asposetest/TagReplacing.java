package asposetest;

import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Node;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.Section;
import java.util.ArrayList;
import java.util.List;

public class TagReplacing implements IReplacingCallback {

    private Document srcDoc;

    public TagReplacing(Document srcDoc) {
        this.srcDoc = srcDoc;
    }

    @Override
    public int replacing(ReplacingArgs replacer) {
        int retorno = ReplaceAction.SKIP;

        Node matchNode = replacer.getMatchNode();
        Node insertAfterNode = matchNode.getParentNode();

        // Make sure that the node is either a paragraph or table.
        if (insertAfterNode.getNodeType() != NodeType.PARAGRAPH && insertAfterNode.getNodeType() != NodeType.TABLE){
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }

        // We will be inserting into the parent of the destination paragraph.
        CompositeNode dstStory = insertAfterNode.getParentNode();
        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        final List<Node> nodes = obterNodesParagrafos(srcDoc);

        for (Node node: nodes) {
            // This creates a clone of the node, suitable for insertion into the destination document.
            Node newNode = importer.importNode(node, true);
            // Insert new node after the reference node.
            dstStory.insertAfter(newNode, insertAfterNode);
            insertAfterNode = newNode;

            retorno = ReplaceAction.REPLACE;
        }

        return retorno;
    }

    private static List<Node> obterNodesParagrafos(Document document){
        List<Node> retorno = new ArrayList<>();

        for (Section srcSection : document.getSections()) {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            for (Node srcNode : (Iterable<Node>) srcSection.getBody()) {
                // Let's skip the node if it is a last empty paragraph in a section.
                if(!verificaSeNodeEhParagrafoValido(srcNode)){
                    retorno.add(srcNode);
                }
            }
        }
        return retorno;
    }

    private static boolean verificaSeNodeEhParagrafoValido(Node srcNode) {
        boolean retorno = false;
        if (srcNode.getNodeType() == NodeType.PARAGRAPH) {
            Paragraph para = (Paragraph) srcNode;
            retorno = para.isEndOfSection() && !para.hasChildNodes();
        }
        return retorno;
    }
}
