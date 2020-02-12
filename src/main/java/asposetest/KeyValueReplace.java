package asposetest;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;

public class KeyValueReplace {

    private final Map<String, String> values;
    private final Map<String, List<String>> datasetValues;

    public KeyValueReplace() {
        this.values = new HashMap<>();
        this.datasetValues = new HashMap<>();
    }

    public KeyValueReplace addIf(String tag, String newText, Supplier<Boolean> fn) {
        if (fn != null && fn.get()) {
            return add(tag, newText);
        }

        return this;
    }

    public KeyValueReplace add(String tag, String newText) {
        values.put(tag, newText);
        return this;
    }

    public KeyValueReplace add(String datasetName, List<String> datasetList) {
        datasetValues.put(datasetName, datasetList);
        return this;
    }

    public Map<String, String> getValues() {
        return values;
    }

    public Map<String, List<String>> getDatasetValues() {
        return datasetValues;
    }

    public boolean hasDatasetValues() {
        return !datasetValues.isEmpty();
    }
}
