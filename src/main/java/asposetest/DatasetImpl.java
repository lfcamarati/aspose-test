package asposetest;

import com.aspose.words.ReportingEngine;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

class DatasetImpl implements Dataset {

    private final DocxInternal docxInternal;
    private final List<Object> dataSources = new ArrayList<>();
    private final List<String> dataSourcesNames = new ArrayList<>();

    public DatasetImpl(DocxInternal docxInternal) {
        this.docxInternal = docxInternal;
    }

    public DatasetImpl add(String dsName, List<String> list) {
        dataSources.add(list);
        dataSourcesNames.add(dsName);

        return this;
    }

    public Docx apply() {
        String[] arrayDataSourcesNames = dataSourcesNames.toArray(new String[0]);

        try {
            ReportingEngine engine = new ReportingEngine();
            engine.buildReport(docxInternal.getDocument(), dataSources.toArray(), arrayDataSourcesNames);
            return docxInternal.getDocx();
        } catch (Exception e) {
            throw new DocxException("Erro ao aplicar datasets ["+ Arrays.toString(arrayDataSourcesNames) +"]", e);
        }
    }
}
