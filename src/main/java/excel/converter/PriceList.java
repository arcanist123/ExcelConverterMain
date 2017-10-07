package excel.converter;

import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

import java.io.IOException;

public class PriceList {
    private DocumentHeader documentHeader;
    private Materials materials;
    
    public PriceList(Element document) throws IOException {
        //ensure that the document is correctly formatted
        document.normalize();
        
        //get the header
        Element documentHeader = (Element) document.getElementsByTagName(Tags.ORGANISATION.toString()).item(0);
        this.documentHeader = new DocumentHeader(documentHeader);
        
        //get the materials
        Element materials = (Element) document.getElementsByTagName(Tags.MATERIALS.toString()).item(0);
        this.materials = Materials.createMaterials(materials);
        
    }
    
    public void writeToXLSX(Sheet mainSheet){
        //we are starting to write to the sheet from the first row
        short currentRow = 1;
        
        //output document header
        documentHeader.writeToXLSX(mainSheet, currentRow);
        
        //output materials
        materials.writeToXLSX(mainSheet,currentRow);
        
    
    }
    
}
