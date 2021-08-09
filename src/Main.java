import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Main {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // read input
        Scanner in = new Scanner(System.in);
        System.out.println("Please input related columns number using , as separator");
        String s1 = in.nextLine();
        System.out.println(s1);

        System.out.println("Please input index");
        String s2 = in.nextLine();
        System.out.println(s2);

        System.out.println("Please input value");
        String s3 = in.nextLine();
        System.out.println(s3);

        System.out.println("Please input input file path");
        String s4 = in.nextLine(); //"/Users/truddy/Downloads/test_data.xlsx"
        System.out.println(s4);

        System.out.println("Please input output file path");
        String s5 = in.nextLine();
        System.out.println(s5);


        //read spreadsheet
        File file = new File(s4);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet1 = workbook.getSheetAt(0);

        //get column headers
        Row row = sheet1.getRow(0);
        List<String> columnHeaders = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            columnHeaders.add(cell.getStringCellValue());
        }

        int mainKey = 0;
        String header1 = row.getCell(mainKey).getStringCellValue();

        String[] cols = s1.split(",");
        int[] relatedColumns = new int[cols.length];
        String[] relatedHeaders = new String[cols.length];
        for (int m = 0; m < cols.length; m++) {
            relatedColumns[m] = Integer.valueOf(cols[m]);
            relatedHeaders[m] = row.getCell(relatedColumns[m]).getStringCellValue();
        }


        int index = Integer.valueOf(s2);
        String indexHeader = row.getCell(index).getStringCellValue();
        int value = Integer.valueOf(s3);
        String valueHeader = row.getCell(value).getStringCellValue();

        // name -> city, area
        Map<String, List<String>> relatedInfoMap = new HashMap<>();
        //name -> day, hours
        Map<String, List<Map<String, String>>> mappingInfoMap = new HashMap<>();

        //rows: number of rows without column header
        int rows = sheet1.getPhysicalNumberOfRows();
        for (int i = 1; i < rows; i++) {
            Row eachRow = sheet1.getRow(i);
            String keyValue = eachRow.getCell(mainKey).getStringCellValue();
            if (!relatedInfoMap.containsKey(keyValue)) {
                List<String> relatedInfo = new ArrayList<>();
                for (int j = 0; j < relatedColumns.length; j++) {
                    relatedInfo.add(eachRow.getCell(relatedColumns[j]).getStringCellValue());
                }
                relatedInfoMap.put(keyValue, relatedInfo);
            }

            String indexCol = eachRow.getCell(index).getStringCellValue();
            String valueRow = String.valueOf(eachRow.getCell(value).getNumericCellValue());
            Map<String, String> map = new HashMap<>();
            map.put(indexCol, valueRow);
            List<Map<String, String>> temp;
            if (!mappingInfoMap.containsKey(keyValue)) {
                temp = new ArrayList<>();
            } else {
                temp = mappingInfoMap.get(keyValue);
            }
            temp.add(map);
            mappingInfoMap.put(keyValue, temp);
        }

        // test
        String[] col1s = new String[relatedInfoMap.size()];
        List<String> valuesHeaders = new ArrayList<>();

        int i = 0;
        for (String key : mappingInfoMap.keySet()) {
            if (i > 0) break;
            else {
                for (Map<String, String> l : mappingInfoMap.get(key)) {
                    for (String k : l.keySet()) {
                        valuesHeaders.add(k);
                    }
                }
                i++;
            }
        }

        // calculate output data
        String[][] newRows = new String[col1s.length][relatedHeaders.length + valuesHeaders.size()];
        int z = 0;
        for (String s : relatedInfoMap.keySet()) {
            System.out.println("*******************");
            System.out.println(s);
            col1s[z] = s;
            String[] col2s = new String[relatedHeaders.length];
            int g = 0;
            for (String ss : relatedInfoMap.get(s)) {
                col2s[g] = ss;
                g++;
                System.out.println(ss);
            }

            String[] col3s = new String[valuesHeaders.size()];
            int r = 0;
            List<Map<String, String>> list = mappingInfoMap.get(s);
            for (Map<String, String> l : list) {
                for (String k : l.keySet()) {
                    System.out.println(k);
                    col3s[r] = l.get(k);
                    r++;
                    System.out.println(l.get(k));
                }
            }
            int t = 0;
            while (t < relatedHeaders.length + valuesHeaders.size()) {
                if (t >= 0 && t < relatedHeaders.length) {
                    while (t < relatedHeaders.length) {
                        newRows[z][t] = col2s[t];
                        t++;
                    }
                } else {
                    while (t < valuesHeaders.size() + relatedHeaders.length) {
                        newRows[z][t] = col3s[t - relatedHeaders.length];
                        t++;
                    }
                }

            }
            z++;
        }

        System.out.println("~~~~~~~~~~~~~~~~~~~~");
        for (int v = 0; v < newRows.length; v++) {
            System.out.println("==========");
            for (int w = 0; w < newRows[0].length; w++) {
                System.out.println(newRows[v][w]);
            }
            System.out.println("==========");
        }



        // write to new file
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet newSheet = wb.createSheet("Output");

        Map<String, Object[]> data = new HashMap<>();
        Object[] headers = new Object[1 + relatedHeaders.length + valuesHeaders.size()];
        for (int p = 0; p < headers.length; p++) {
            if (p == 0) {
                headers[p] = header1;
            } else if (p >= 1 && p < 1 + relatedHeaders.length) {
                headers[p] = relatedHeaders[p - 1];
            } else {
                headers[p] = valuesHeaders.get(p - 1 - relatedHeaders.length);
            }
        }
        data.put("1", headers);
        for (int u = 2; u < relatedInfoMap.size() + 2; u++) {
            Object[] objects = new Object[headers.length];
            objects[0] = col1s[u - 2];
            for (int r = 0; r < newRows[0].length; r++) {
                objects[r + 1] = newRows[u-2][r];
            }
            data.put(String.valueOf(u), objects);
        }


        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row newRow = newSheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = newRow.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }

        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(s5));
            wb.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

        for (Object o : headers) {
            System.out.println(o);
        }
    }

}
