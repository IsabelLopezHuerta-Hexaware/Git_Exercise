import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExerciseExcel
{
    public static void main(String[] args)
    {
        //Blank workbook--- Crea el libro en blanco
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet--- crea la hoja en blanco
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //This data needs to be written (Object[]) -- declaras los datos con Map y esta compuesto de <string que sera la fila,, object[] y lo que lleve adrentro >
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"NAME", "LASTNAME", "EMAIL","PASSWORD","COMPANY","ADDRESS","CITY","ZIP_CODE","MOBILE_PHONE"});
        data.put("2", new Object[] {"Isabel ", "Lopez", "isabel@mail.com","123456","hexaware","calle2, #2","SQL City","10110","333 123 12 23"});


        //Iterate over data and write to sheet -- iterar sobre los datos y escribir en la hoja
        //crea un set de las key
        Set<String> keyset = data.keySet();
        // numero de columna
        int rownum = 0;
        //itera sobre keyset
        for (String key : keyset)
        {
            //Row es de la libreria de apache, row es la variable
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("exercise.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("exercise.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}