package excel.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.IOException;
import java.util.Base64;
import java.util.Vector;

public class Material {
    private static final short MATERIAL_GUID_COLUMN = 1;
    private static final short MATERIAL_NAME_COLUMN = 5;
    private static final short MATERIAL_VENDOR_CODE_COLUMN = 4;
    private static final short MATERIAL_PICTURE_COLUMN = 3;
    byte[] picture;
    private String materialName;
    private String materialGUID;
    private String materialVendorCode;
    
    private Vector<MaterialAttribute> materialAttributes;
    
    private Material() {
        this.materialAttributes = new Vector<MaterialAttribute>(0);
    }
    
    public static Material createMaterial(Element materialElement) throws IOException {
        
        //create instance of material
        Material material = new Material();
        
        //get the name of material
        material.materialName = materialElement.getElementsByTagName(Tags.MATERIAL_NAME.toString()).item(0).getTextContent();
        
        //get the guid of material
        material.materialGUID = materialElement.getElementsByTagName(Tags.MATERIAL_GUID.toString()).item(0).getTextContent();
        
        //get the material vendor code
        material.materialVendorCode = materialElement.getElementsByTagName(Tags.MATERIAL_VENDOR_CODE.toString()).item(0).getTextContent();
        
        //get the picture
        //get the possible pictures
        NodeList pictures = materialElement.getElementsByTagName(Tags.PICTURE);
        //check if the current material has picture
        if (pictures.getLength() > 0) {
            //this material has picture. get it
            String pictureEncoded = pictures.item(0).getTextContent();
            //decode base64 string
            material.picture = Base64.getDecoder().decode(pictureEncoded);
          } else {
            ///no picture was provided from
        }
        
        //get path to the source folder
        //String sourceFolderPath = XML2XLSXConverter.getInstance().getSourceFolderPath();
        ////get the name of the picture
        //String pathToPicture = sourceFolderPath.concat(material.materialGUID);
        ////FileInputStream obtains input bytes from the image file
        //InputStream inputStream = new FileInputStream(pathToPicture);
        ////Get the contents of an InputStream as a byte[].
        //material.picture = IOUtils.toByteArray(inputStream);
        
        //now populate the vector of the material attributes
        NodeList materialAttributes = materialElement.getElementsByTagName(Tags.ATTRIBUTE.toString());
        
        for (int i = 0; i < materialAttributes.getLength(); i++) {
            //get the current material attribute
            Element materialAttributeElement = (Element) materialAttributes.item(i);
            //create an instance of material attribute
            material.materialAttributes.add(MaterialAttribute.createMaterialAttribute(materialAttributeElement));
        }
        
        //return the results of processing
        return material;
    }
    
    public short writeToXLSX(Sheet mainSheet, short currentRow) {
        
        //materials themselves are not being written to the sheet - only the underlying attributes are
        //save the current row for future usage
        short originalRow = currentRow;
        
        //write the attributes of material
        for (int i = 0; i < this.materialAttributes.size(); i++) {
            //write the current material attribute
            currentRow = this.materialAttributes.elementAt(i).writeToXLSX(mainSheet, currentRow);
        }
        
        //determine the height of the row
        float rowHeith;
        if (materialAttributes.size() > 5) {
            //set the standard row height
            rowHeith = (float) 62.5 / 5;
        } else {
            rowHeith = (float) 62.5 / materialAttributes.size();
        }
        
        //output the material guid and material name
        for (int i = originalRow; i < currentRow; i++) {
            //get the current row
            Row row = mainSheet.getRow(i);
            
            //set the row height
            row.setHeightInPoints(rowHeith);
            //output the material guid
            Cell materialGuidCell = row.createCell(MATERIAL_GUID_COLUMN);
            materialGuidCell.setCellValue(this.materialGUID);
            
            //output the material name
            Cell materialNameCell = row.createCell(MATERIAL_NAME_COLUMN);
            materialNameCell.setCellValue(this.materialName);
            
            //output the vendor code
            Cell materialVendorCodeCell = row.createCell(MATERIAL_VENDOR_CODE_COLUMN);
            materialVendorCodeCell.setCellValue(this.materialVendorCode);
        }
        
        //unite the cells for material
        if (originalRow == currentRow - 1) {
            //in this case we are not having more then one row to work. no need to unite one cell
        } else {
            //we have some cell to merge. merge them
            mainSheet.addMergedRegion(new CellRangeAddress(originalRow, currentRow - 1, MATERIAL_PICTURE_COLUMN, MATERIAL_PICTURE_COLUMN));
            
        }
        
        return currentRow;
    }
}
