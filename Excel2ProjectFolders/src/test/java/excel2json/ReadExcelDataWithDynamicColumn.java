package excel2json;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.AtomicMoveNotSupportedException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicIntegerArray;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONException;
import org.json.JSONObject;

public class ReadExcelDataWithDynamicColumn {

static String configJsonString;
static String indexJsString ="" ;
public static void main(String[] args)
    {
        // You can specify your excel file path.
        String excelFilePath = "/Users/bhavanigirish/Desktop/excel2project.xlsx";
   

       
        creteJSONAndTextFileFromExcel(excelFilePath);
    }


    
    private static void creteJSONAndTextFileFromExcel(String filePath)
    {
        try{
        Map<String,JSONObject> collectJsonDataMap = null;
      
            FileInputStream fInputStream = new FileInputStream(filePath.trim());
   
         /* Create the workbook object to access excel file. */
            Workbook excelWorkBook = new XSSFWorkbook(fInputStream);
         
       
            int totalSheetNumber = excelWorkBook.getNumberOfSheets();
            // Loop in all excel sheet.
//            for(int i=0;i<totalSheetNumber;i++)
//            {
                // Get current sheet.
                Sheet timeSheet = excelWorkBook.getSheetAt(0);
                Sheet hotCostSheet = excelWorkBook.getSheetAt(1);

                // Get sheet name.
                String sheetName = timeSheet.getSheetName();

                if(sheetName != null && sheetName.length() > 0)
                {
                Map<String,String> testCaseMap = new HashedMap<>();
                Map<String,JSONObject> timeJsonMap = null;
                Map<String,JSONObject> hotCostJsonMap = null;
                collectJsonDataMap = new HashedMap<>();
               
                //Get the output hotcost sheet data in a list table
                List<List<String>> hotCostSheetDataTable = getSheetDataList(hotCostSheet);
                   
                // Get current sheet data in a list table.
                    List<List<String>> timeSheetDataTable = getSheetDataList(timeSheet);
                   

                    // Generate JSON format of above sheet data and write to a JSON file.
//                    String timeJsonString = getJSONStringFromList(timeSheetDataTable,collectJsonDataMap);
//                    timeJsonMap = collectJsonDataMap;
//                    collectJsonDataMap = new HashedMap<>();
//                    String hotCostJsonString = getJSONStringFromList(hotCostSheetDataTable,collectJsonDataMap);
//                    hotCostJsonMap = collectJsonDataMap;
//                    String jsonFileName = sheet.getSheetName() + ".json";
                    String jsonFolderName = timeSheet.getSheetName();
//                  
                    writeStringToSepareteFoldersFromList(jsonFolderName,testCaseMap,getSheetDataAsStringList(hotCostSheet,testCaseMap),getSheetDataAsStringList(timeSheet,testCaseMap));

//                    
                }
//            }
          
            excelWorkBook.close();
        }catch(Exception ex){
            System.err.println(ex.getMessage());
        }
    }

    private static void writeStringToSepareteFoldersFromList(String jsonFolderName,
Map<String, String> testCaseMap, List<String> hotCostList, List<String> timeList) throws IOException {
// TODO Auto-generated method stub
    // Get current executing class working directory.
        String currentWorkingFolder = System.getProperty("user.dir");
        AtomicInteger atomicInteger = new AtomicInteger(0);
        // Get file path separator.
        String filePathSeperator = System.getProperty("file.separator");

        // Get the output file absolute path.
        String filePath = currentWorkingFolder + filePathSeperator + jsonFolderName;

        // Create File, FileWriter and BufferedWriter object.
        File file = new File(filePath);
        if(!file.exists()) {
        file.mkdir();
        }
       
        File indexJsFile =  new File(filePath+"/index.js");
        if(!indexJsFile.exists()) {
        readFileFromIndexJs(indexJsFile);
        }
        timeList.stream().forEach(data -> {
        int currentIndex = atomicInteger.get();
        int row = atomicInteger.incrementAndGet();
        String currentFolder = filePath+"/"+testCaseMap.get(String.valueOf(currentIndex));
        File file1 = new File(currentFolder);
        file1.mkdir();
        String inJson = currentFolder+"/" + "in.json";
        String outJson = currentFolder+"/" + "out.json";
        String configJson = currentFolder+"/" + "config.json";
FileWriter fwIn;
FileWriter fwOut;
FileWriter fwConfig;
try {
fwIn = new FileWriter(inJson);
BufferedWriter buffWriterIn = new BufferedWriter(fwIn);
buffWriterIn.write(data.toString());
buffWriterIn.flush();
buffWriterIn.close();
           
           //write for out json
fwOut = new FileWriter(outJson);
BufferedWriter buffWriterOut = new BufferedWriter(fwOut);
buffWriterOut.write(hotCostList.get(currentIndex).toString());
buffWriterOut.flush();
buffWriterOut.close();

//write config json
fwConfig = new FileWriter(configJson);
readFileFromConfigJson(fwConfig);

} catch (IOException e1) {
// TODO Auto-generated catch block
e1.printStackTrace();
}
        });


}


private static void readFileFromIndexJs(File indexJsFile) throws IOException {
File readFile = new File("/Users/bhavanigirish/Documents/workspace/Excel2ProjectFolder/src/test/resources/index.js");
FileInputStream in  = new FileInputStream(readFile);
FileWriter fw = new FileWriter(indexJsFile);
BufferedWriter bfw = new BufferedWriter(fw);
int c;
while ((c = in.read()) != -1) {
bfw.write(c);
         }
bfw.flush();
bfw.close();
in.close();
}


private static void readFileFromConfigJson(FileWriter fwConfig) throws IOException {
File readFile = new File("/Users/bhavanigirish/Documents/workspace/Excel2ProjectFolder/src/test/resources/config.json");
FileInputStream in  = new FileInputStream(readFile);
BufferedWriter bfw = new BufferedWriter(fwConfig);
int c;
while ((c = in.read()) != -1) {
bfw.write(c);
         }
bfw.flush();
bfw.close();
in.close();
}


/**
     * This method is used to keep each row data as a separate json file in separate folders
     * @param jsonFolderName
     * @param timeJsonDataMap
     * @param hotCostJsonMap
* @throws IOException
     */
    private static void writeStringToSeparateFolders(String jsonFolderName, Map<String, JSONObject> timeJsonDataMap, Map<String, JSONObject> hotCostJsonMap) throws IOException {
   
    // Get current executing class working directory.
        String currentWorkingFolder = System.getProperty("user.dir");

        // Get file path separator.
        String filePathSeperator = System.getProperty("file.separator");

        // Get the output file absolute path.
        String filePath = currentWorkingFolder + filePathSeperator + jsonFolderName;

        // Create File, FileWriter and BufferedWriter object.
        File file = new File(filePath);
        if(!file.exists()) {
        file.mkdir();
        }
       
        File indexJsFile =  new File(filePath+"/index.js");
        if(!indexJsFile.exists()) {
        readFileFromIndexJs(indexJsFile);
        }
        timeJsonDataMap.entrySet().stream().forEach(data -> {
        String currentFolder = filePath+"/"+data.getKey();
        File file1 = new File(currentFolder);
        file1.mkdir();
        String inJson = currentFolder+"/" + "in.json";
        String outJson = currentFolder+"/" + "out.json";
        String configJson = currentFolder+"/" + "config.json";
FileWriter fwIn;
FileWriter fwOut;
FileWriter fwConfig;
try {
fwIn = new FileWriter(inJson);
BufferedWriter buffWriterIn = new BufferedWriter(fwIn);
buffWriterIn.write(data.getValue().toString());
buffWriterIn.flush();
buffWriterIn.close();
           
           //write for out json
fwOut = new FileWriter(outJson);
BufferedWriter buffWriterOut = new BufferedWriter(fwOut);
buffWriterOut.write(hotCostJsonMap.get(data.getKey()).toString());
buffWriterOut.flush();
buffWriterOut.close();

//write config json
fwConfig = new FileWriter(configJson);
readFileFromConfigJson(fwConfig);

} catch (IOException e1) {
// TODO Auto-generated catch block
e1.printStackTrace();
}
        });
}



    private static List<List<String>> getSheetDataList(Sheet sheet)
    {
        List<List<String>> ret = new ArrayList<List<String>>();

        // Get the first and last sheet row number.
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        if(lastRowNum > 0)
        {
            // Loop in sheet rows.
            for(int i=firstRowNum; i<lastRowNum; i++)
            {
                // Get current row object.
                Row row = sheet.getRow(i);

                // Get first and last cell number.
                int firstCellNum = row.getFirstCellNum();
                int lastCellNum = row.getLastCellNum();

                // Create a String list to save column data in a row.
                List<String> rowDataList = new ArrayList<String>();

                // Loop in the row cells.
                for(int j = firstCellNum; j < lastCellNum; j++)
                {
                    // Get current cell.
                    Cell cell = row.getCell(j);

                    // Get cell type.
                    int cellType = cell.getCellType();

                    if(cellType == CellType.NUMERIC.getCode())
                    {
                        double numberValue = cell.getNumericCellValue();

                        // BigDecimal is used to avoid double value is counted use Scientific counting method.
                        // For example the original double variable value is 12345678, but jdk translated the value to 1.2345678E7.
                        String stringCellValue = BigDecimal.valueOf(numberValue).toPlainString();

                        rowDataList.add(stringCellValue);

                    }else if(cellType == CellType.STRING.getCode())
                    {
                        String cellValue = cell.getStringCellValue();
                        rowDataList.add(cellValue);
                    }else if(cellType == CellType.BOOLEAN.getCode())
                    {
                        boolean numberValue = cell.getBooleanCellValue();

                        String stringCellValue = String.valueOf(numberValue);

                        rowDataList.add(stringCellValue);

                    }else if(cellType == CellType.BLANK.getCode())
                    {
                        rowDataList.add("");
                    }
                }

                // Add current row data list in the return list.
                ret.add(rowDataList);
            }
        }
        return ret;
    }

    /* Return sheet data in a two dimensional list.
     * Each element in the outer list is represent a row,
     * each element in the inner list represent a column.
     * The first row is the column name row.*/
    private static List<String> getSheetDataAsStringList(Sheet sheet, Map<String, String> testCaseMap)
    {
        List<String> ret = new ArrayList<String>();

        // Get the first and last sheet row number.
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        if(lastRowNum > 0)
        {
            // Loop in sheet rows.
            for(int i=firstRowNum; i<lastRowNum + 1; i++)
            {
                // Get current row object.
                Row row = sheet.getRow(i);

                // Get first and last cell number.
                int firstCellNum = row.getFirstCellNum();
                int lastCellNum = row.getLastCellNum();

                // Create a String list to save column data in a row.
                String rowDataList = "";

                // Loop in the row cells.
                for(int j = firstCellNum; j < lastCellNum; j++)
                {
               
                    // Get current cell.
                    Cell cell = row.getCell(j);
                   
                   
                    if(j==0) {
                    testCaseMap.put(String.valueOf(i), cell.getStringCellValue());
                continue;
                }
                   
                    // Get cell type.
                    int cellType = cell.getCellType();

                    if(cellType == CellType.NUMERIC.getCode())
                    {
                        double numberValue = cell.getNumericCellValue();

                        // BigDecimal is used to avoid double value is counted use Scientific counting method.
                        // For example the original double variable value is 12345678, but jdk translated the value to 1.2345678E7.
                        String stringCellValue = BigDecimal.valueOf(numberValue).toPlainString();

                        rowDataList = rowDataList +stringCellValue;

                    }else if(cellType == CellType.STRING.getCode())
                    {
                        String cellValue = cell.getStringCellValue();
                        rowDataList = rowDataList + cellValue;
                    }else if(cellType == CellType.BOOLEAN.getCode())
                    {
                        boolean numberValue = cell.getBooleanCellValue();

                        String stringCellValue = String.valueOf(numberValue);

                        rowDataList = rowDataList + stringCellValue;

                    }else if(cellType == CellType.BLANK.getCode())
                    {
//                        rowDataList.add("");
                    }
                }

                // Add current row data list in the return list.
                ret.add(rowDataList);
            }
        }
        return ret;
    }

   
    /* Return a JSON string from the string list. */
    private static String getJSONStringFromList(List<List<String>> dataTable, Map<String, JSONObject> collectJsonDataMap) throws JSONException
    {
        String ret = "";

        if(dataTable != null)
        {
            int rowCount = dataTable.size();

            if(rowCount > 1)
            {
                // Create a JSONObject to store table data.
                JSONObject tableJsonObject = new JSONObject();

                // The first row is the header row, store each column name.
                List<String> headerRow = dataTable.get(2);

                int columnCount = headerRow.size();

                // Loop in the row data list.
                for(int i=3; i<rowCount; i++)
                {
                    // Get current row data.
                    List<String> dataRow = dataTable.get(i);

                    // Create a JSONObject object to store row data.
                    JSONObject rowJsonObject = new JSONObject();

                    for(int j=0;j<columnCount;j++)
                    {
                        String columnName = headerRow.get(j);
                        String columnValue = dataRow.get(j);

                        rowJsonObject.put(columnName, columnValue);
                    }

                    tableJsonObject.put("Row " + i, rowJsonObject);
                    collectJsonDataMap.put("Row " + (i-2), rowJsonObject);
                }

                // Return string format data of JSONObject object.
                ret = tableJsonObject.toString();

            }
        }
        return ret;
    }


    /* Return a text table string from the string list. */
    private static String getTextTableStringFromList(List<List<String>> dataTable)
    {
        StringBuffer strBuf = new StringBuffer();

        if(dataTable != null)
        {
            // Get all row count.
            int rowCount = dataTable.size();

            // Loop in the all rows.
            for(int i=0;i<rowCount;i++)
            {
                // Get each row.
                List<String> row = dataTable.get(i);

                // Get one row column count.
                int columnCount = row.size();

                // Loop in the row columns.
                for(int j=0;j<columnCount;j++)
                {
                    // Get column value.
                    String column = row.get(j);

                    // Append column value and a white space to separate value.
                    strBuf.append(column);
                    strBuf.append("    ");
                }

                // Add a return character at the end of the row.
                strBuf.append("\r\n");
            }

        }
        return strBuf.toString();
    }

    /* Write string data to a file.*/
    private static void writeStringToFile(String data, String fileName)
    {
        try
        {
            // Get current executing class working directory.
            String currentWorkingFolder = System.getProperty("user.dir");

            // Get file path separator.
            String filePathSeperator = System.getProperty("file.separator");

            // Get the output file absolute path.
            String filePath = currentWorkingFolder + filePathSeperator + fileName;

            // Create File, FileWriter and BufferedWriter object.
            File file = new File(filePath);

            FileWriter fw = new FileWriter(file);

            BufferedWriter buffWriter = new BufferedWriter(fw);

            // Write string data to the output file, flush and close the buffered writer object.
            buffWriter.write(data);

            buffWriter.flush();

            buffWriter.close();

            System.out.println(filePath + " has been created.");

        }catch(IOException ex)
        {
            System.err.println(ex.getMessage());
        }
    }

}
