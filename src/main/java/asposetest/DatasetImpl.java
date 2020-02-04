package asposetest;

import com.aspose.words.ReportingEngine;

import java.util.ArrayList;
import java.util.List;

class DatasetImpl implements Dataset {

    private final DocxImpl documento;
    private List<Object> dataSources = new ArrayList<>();
    private List<String> dataSourcesNames = new ArrayList<>();

    public DatasetImpl(DocxImpl documento) {
        this.documento = documento;
    }

    public DatasetImpl add(String dsName, List<String> list) {
        dataSources.add(list);
        dataSourcesNames.add(dsName);

        return this;
    }

    public Docx apply() throws Exception {
        String[] arrayDataSourcesNames = dataSourcesNames.toArray(new String[0]);

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(documento.getDocument(), dataSources.toArray(), arrayDataSourcesNames);

        return documento;
    }
}
