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

class InsertDocumentAtReplaceTextHandler implements IReplacingCallback {

    private final DocxInternal docxInternal;

    InsertDocumentAtReplaceTextHandler(DocxInternal docxInternal) {
        this.docxInternal = docxInternal;
    }

    public int replacing(ReplacingArgs e) throws Exception {
        Document srcDoc = docxInternal.getDocument();

        // Insert a document after the paragraph, containing the match text.
        Node matchNode = e.getMatchNode();
        Paragraph para = (Paragraph) matchNode.getParentNode();
        insertDocument(para, srcDoc);

        // Remove the paragraph with the match text.
        para.remove();

        return ReplaceAction.SKIP;
    }

    private static void insertDocument(Node insertAfterNode, Document srcDoc) {
        // Make sure that the node is either a paragraph or table.
        if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH) && (insertAfterNode.getNodeType() != NodeType.TABLE)) {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }

        // We will be inserting into the parent of the destination paragraph.
        CompositeNode dstStory = insertAfterNode.getParentNode();

        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Loop through all sections in the source document.
        for (Section srcSection : srcDoc.getSections()) {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section.
            insertAfterNode = getNode(insertAfterNode, dstStory, importer, srcSection);
        }
    }

    private static Node getNode(Node insertAfterNode, CompositeNode dstStory, NodeImporter importer, Section srcSection) {
        for (Node srcNode : srcSection.getBody()) {
            // Let's skip the node if it is a last empty paragraph in a section.
            if (srcNode.getNodeType() == (NodeType.PARAGRAPH)) {
                Paragraph para = (Paragraph) srcNode;

                if (para.isEndOfSection() && !para.hasChildNodes()) {
                    continue;
                }
            }

            // This creates a clone of the node, suitable for insertion into the destination document.
            Node newNode = importer.importNode(srcNode, true);

            // Insert new node after the reference node.
            dstStory.insertAfter(newNode, insertAfterNode);
            insertAfterNode = newNode;
        }

        return insertAfterNode;
    }
}
