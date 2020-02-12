package asposetest;

import java.util.List;

public interface Dataset {

    Dataset add(String dsName, List<String> list);

    Docx apply();
}
