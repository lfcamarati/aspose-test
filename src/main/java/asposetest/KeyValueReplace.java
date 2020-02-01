package asposetest;

import java.util.HashMap;
import java.util.Map;

public class KeyValueReplace {

    private final Map<String, String> values;

    public KeyValueReplace() {
        this.values = new HashMap<>();
    }

    public KeyValueReplace add(String text, String newText) {
        values.put(text, newText);
        return this;
    }

    public Map<String, String> getValues() {
        return values;
    }
}
