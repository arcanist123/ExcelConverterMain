package excel.converter;

public enum Tags {
    
    ORGANISATION("Organisation"), MATERIALS("Materials"), MATERIAL("Material"), MATERIAL_NAME("MaterialName"), MATERIAL_GUID("MaterialGUID"), ATTRIBUTE("Attribute"), ATTRIBUTE_NAME("AttributeName"), ATTRIBUTE_GUID("AttributeGUID"), ATTRIBUTE_PRICE("AttributePrice"), ATTRIBUTE_DISCOUNT("AttributeDiscount"), ATTRIBUTE_PRICE_WITH_DISCOUNT("AttributePriceWithDiscount"), MATERIAL_GROUP("MaterialGroup"), GROUP_GUID("MaterialGroupGUID"), GROUP_NAME("MaterialGroupName"), MATERIAL_VENDOR_CODE("MaterialVendorCode");
    
    public static final String PICTURE = "MaterialPicture";
    private final String text;
    
    /**
     * @param text
     */
    private Tags(final String text) {
        this.text = text;
    }
    
    /* (non-Javadoc)
     * @see java.lang.Enum#toString()
     */
    @Override
    public String toString() {
        return text;
    }
}


