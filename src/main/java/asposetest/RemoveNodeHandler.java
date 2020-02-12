package asposetest;

import com.aspose.words.IReplacingCallback;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;

class RemoveNodeHandler implements IReplacingCallback {

    public int replacing(ReplacingArgs e) {
        Node matchNode = e.getMatchNode();
        Paragraph para = (Paragraph) matchNode.getParentNode();
        para.remove();

        return ReplaceAction.SKIP;
    }
}
