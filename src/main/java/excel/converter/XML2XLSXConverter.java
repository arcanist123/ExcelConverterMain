package excel.converter;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class XML2XLSXConverter {
    private static XML2XLSXConverter instance = null;
    private String sourceFolderPath;
    private String targetFilePath;
    private String constSourceFileName = "source.xml";
    
    
    private XML2XLSXConverter() {
        // Exists only to defeat instantiation.
    }
    
    
    public static XML2XLSXConverter getInstance() {
        if (instance == null) {
            instance = new XML2XLSXConverter();
        }
        return instance;
    }
    
    public static XML2XLSXConverter createConverter(String sourceFolderPath, String targetFilePath) {
        XML2XLSXConverter instance = XML2XLSXConverter.getInstance();
        instance.sourceFolderPath = sourceFolderPath;
        instance.targetFilePath = targetFilePath;
        return instance;
    }
    
    private static PriceList getConvertedDocument(String sourceFolderPath, String constSourceFileName) throws ParserConfigurationException, SAXException, IOException {
        //get the name of the source file. By convention, it is named source.xml and is located in the source folder
        String filePath = sourceFolderPath.concat(constSourceFileName);
        File xmlFile = new File(filePath);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = factory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);
        
        
        //create the excel document Java representation
        PriceList priceList = new PriceList(doc.getDocumentElement());
        return priceList;
    }
    
    public String getSourceFolderPath() {
        return sourceFolderPath;
    }
    
    public void convert() throws ParserConfigurationException, IOException, SAXException {
        //price list
        PriceList priceList = getConvertedDocument(this.sourceFolderPath, constSourceFileName);
        
        //we have the instance of all materials.
        //create the xlsx document
        Workbook wb = new XSSFWorkbook();
        Sheet mainSheet = wb.createSheet("ПРАЙС");
        
        //write the PriceList to xlsx
        priceList.writeToXLSX(mainSheet);
        
        //write the xlsx Price list into the file
        FileOutputStream fileOut = new FileOutputStream(this.targetFilePath);
        wb.write(fileOut);
        fileOut.close();
        
    }
    
}