package excel.converter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.IOException;
import java.util.Vector;

public class Materials {

    private Vector<MaterialGroup> materialGroupsVector;
    
    private Materials() {
        this.materialGroupsVector = new Vector<MaterialGroup>(0);
    }
    
    public static Materials createMaterials(Element materialsElement) throws IOException {
        //create an instance of materials. It actually does not have any attributes and is just a holder for the vector of groups
        Materials materials = new Materials();
        
        
        
        //get all the material groups elements
        NodeList materialGroupsList = materialsElement.getElementsByTagName(Tags.MATERIAL_GROUP.toString());
        for (int i = 0; i < materialGroupsList.getLength(); i++) {
            //get the current material group element
            Element materialGroupElement = (Element) materialGroupsList.item(i);
            //add the current material group to the vector of material groups
            materials.materialGroupsVector.add(MaterialGroup.createMaterialGroup(materialGroupElement));
        }
        
        //return the result of processing - the materials instance
        return materials;
        
    }
    
    public short writeToXLSX(Sheet mainSheet, short currentRow) {
        //save the current row before the processing
        short originalRow = currentRow;
        
        //we need to output all the groups we have
        for (int i = 0; i < this.materialGroupsVector.size(); i++) {
            //output the current group
            currentRow = this.materialGroupsVector.elementAt(i).writeToXLSX(mainSheet, currentRow);
        }
        

        
        
        //return the current index in the xlsx
        return currentRow;
    }
}
