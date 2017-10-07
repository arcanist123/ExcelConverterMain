package excel.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;

import java.math.BigDecimal;

public class MaterialAttribute {
    private static final short ATTRIBUTE_NAME_COLUMN = 6;
    private static final short ATTRIBUTE_GUID_COLUMN = 2;
    private static final short ATTRIBUTE_PRICE_COLUMN = 7;
    private static final short ATTRIBUTE_DISCOUNT_COLUMN = 8;
    private static final short ATTRIBUTE_PRICE_WITH_DISCOUNT_COLUMN = 9;
    private String attributeName;
    private String attributeGUID;
    private BigDecimal quantity;
    private BigDecimal price;
    private BigDecimal discount;
    private BigDecimal priceWithDiscount;
    
    private MaterialAttribute() {
        
    }
    
    public static MaterialAttribute createMaterialAttribute(Element materialAttributeElement) {
        
        //create an instance of material attribute
        MaterialAttribute materialAttribute = new MaterialAttribute();
        
        //get the name of attribute
        materialAttribute.attributeName = materialAttributeElement.getElementsByTagName(Tags.ATTRIBUTE_NAME.toString()).item(0).getTextContent();
        
        //get the guid of the attribute
        materialAttribute.attributeGUID = materialAttributeElement.getElementsByTagName(Tags.ATTRIBUTE_GUID.toString()).item(0).getTextContent();
        
        //get the price of the attribute
        String priceString = materialAttributeElement.getElementsByTagName(Tags.ATTRIBUTE_PRICE.toString()).item(0).getTextContent();
        System.out.println(priceString);
        materialAttribute.price = new BigDecimal(priceString);
        
        //get the discount of the attribute
        String discountString = materialAttributeElement.getElementsByTagName(Tags.ATTRIBUTE_DISCOUNT.toString()).item(0).getTextContent();
        System.out.println(discountString);
        materialAttribute.discount = new BigDecimal(discountString);
        
        //get the price with discount
        String priceWithDiscountString = materialAttributeElement.getElementsByTagName(Tags.ATTRIBUTE_PRICE_WITH_DISCOUNT.toString()).item(0).getTextContent();
        System.out.println(priceWithDiscountString);
        
        materialAttribute.priceWithDiscount = new BigDecimal(priceWithDiscountString);
        
        return materialAttribute;
    }
    
    public short writeToXLSX(Sheet mainSheet, short currentRow) {
        //write the current material attribute
        Row row = mainSheet.createRow(currentRow);
        
        //output attribute guid
        Cell attributeGuidCell = row.createCell(ATTRIBUTE_GUID_COLUMN);
        attributeGuidCell.setCellValue(this.attributeGUID);
        
        //output attribute name
        Cell attributeNameCell = row.createCell(ATTRIBUTE_NAME_COLUMN);
        attributeNameCell.setCellValue(this.attributeName);
        
        //output attribute price
        Cell attributePriceCell = row.createCell(ATTRIBUTE_PRICE_COLUMN);
        attributePriceCell.setCellValue(this.price.toString());
        
        //output attribute discount
        Cell attributeDiscountCell = row.createCell(ATTRIBUTE_DISCOUNT_COLUMN);
        attributeDiscountCell.setCellValue(this.discount.toString());
        
        //output attribute price with discount
        Cell attributePriceWithDiscountCell = row.createCell(ATTRIBUTE_PRICE_WITH_DISCOUNT_COLUMN);
        attributePriceWithDiscountCell.setCellValue(this.priceWithDiscount.toString());
        
        //we have sent the current row to the sheet. increase the counter to indicate that
        currentRow++;
        
        return currentRow;
    }
}
