package excel.converter;


import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

public class DocumentHeader {
    private String organisationName;
    private String organisationLegalAddress;
    private String organisationFactualAddress;
    
    public DocumentHeader(Element documentHeader) {
    
    }
    
    
    public void writeToXLSX(Sheet mainSheet, short currentRow) {
    }
}
