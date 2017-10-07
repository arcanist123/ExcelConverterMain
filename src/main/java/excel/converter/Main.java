package excel.converter;

public class Main {
    
    
    public static void main(String[] args) {
        //debug
        String[] myArgs = new String[3];
        myArgs[0] = Actions.GET_XLSX_FROM_XML.toString();
        myArgs[1] = "C:\\Users\\arcanist\\IdeaProjects\\ExcelConverter\\src\\main\\resources\\";
        myArgs[2] = "C:\\Users\\arcanist\\IdeaProjects\\ExcelConverter\\src\\main\\resources\\result.xlsx";
        
        args = myArgs;
        
        try {
            //check if the amount of arguments is the correct one
            if (args.length == 3) {
                //this is expected situation. We have to have an action, source folder and the name of the resulting file
            } else {
                throw new IllegalArgumentException("Invalid amount of arguments");
            }
            
            //depending on the action instantiate different files
            String action = args[0];//the first argument is always action
            if (Actions.GET_XLSX_FROM_XML.toString() == action) {
                //we are requested to translate the XMl file into XLSX document. instantiate the translator
                //define source folder
                String sourceFolderPath = args[1];
                String targetFile = args[2];
                
                XML2XLSXConverter xml2XLSXConverter = XML2XLSXConverter.createConverter(sourceFolderPath,targetFile);
                xml2XLSXConverter.convert();
            } else if (Actions.GET_XML_FROM_XLSX.toString() == action) {
            
            }
            
            
        } catch (Exception e) {
            e.printStackTrace();
        }
        //the first parameter should be action
        
        
    }
}
