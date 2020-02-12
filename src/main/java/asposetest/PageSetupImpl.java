package asposetest;

import com.aspose.words.PaperSize;

import static com.aspose.words.ConvertUtil.inchToPoint;

class PageSetupImpl implements PageSetup {

    private final DocxInternal docxInternal;

    public PageSetupImpl(DocxInternal docxInternal) {
        this.docxInternal = docxInternal;
    }

    @Override
    public PageSetup a4() {
        docxInternal.getBuilder().getPageSetup().setPaperSize(PaperSize.A4);
        return this;
    }

    @Override
    public PageSetup margins(double top, double botton, double left, double right) {
        double base = 2.54;
        docxInternal.getBuilder().getPageSetup().setTopMargin(inchToPoint(top/base));
        docxInternal.getBuilder().getPageSetup().setBottomMargin(inchToPoint(botton/base));
        docxInternal.getBuilder().getPageSetup().setLeftMargin(inchToPoint(left/base));
        docxInternal.getBuilder().getPageSetup().setRightMargin(inchToPoint(right/base));

        return this;
    }
}
