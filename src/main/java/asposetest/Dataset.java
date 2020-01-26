package asposetest;

import java.util.Collection;

public interface Dataset {

    Dataset add(String dsName, Collection<String> list);

    Documento apply() throws Exception;
}
