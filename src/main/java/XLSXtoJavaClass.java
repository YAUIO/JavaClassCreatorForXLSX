import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.homyakin.iuliia.Schemas;
import ru.homyakin.iuliia.Translator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.PrintWriter;
import java.util.Locale;


public class XLSXtoJavaClass {
    public static void main(String[] args) {
        System.out.println("Example usage: XLSXtoJavaClass <[string] input file> <[@nullable int] sheet index");
        System.out.println("Started generating Java class");
        FileInputStream file = null;
        Workbook sourceBook = null;
        Translator toLatin = new Translator(Schemas.WIKIPEDIA);

        FileOutputStream outStream = null;
        PrintWriter out = null;
        File table = null;
        try {
            file = new FileInputStream(args[0]);
            sourceBook = new XSSFWorkbook(file);

            table = new File(args[0].substring(0,args[0].lastIndexOf(".")) + ".java");
            if (table.exists()) throw new RuntimeException();
            System.out.println("Output: " + table.getAbsolutePath());

            outStream = new FileOutputStream(table);
            out = new PrintWriter(outStream, true);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        Sheet sourceSheet = null;

        if (args.length == 2) {
            sourceSheet = sourceBook.getSheetAt(Integer.parseInt(args[1]));
        } else {
            sourceSheet = sourceBook.getSheetAt(0);
        }

        Row header = sourceSheet.getRow(0);

        out.println("public class " + table.getName().substring(0, table.getName().lastIndexOf(".")) + " {");

        int i = 0;
        while (i < header.getPhysicalNumberOfCells() || header.getCell(i) != null) {
            String s = header.getCell(i).getStringCellValue();
            String orig = s;
            boolean alt = false;

            if (s.contains(" ")) {
                s = s.replace(' ', '_');
                alt = true;
            }

            String field = toLatin.translate(s);

            if (!field.equals(orig)) {
                alt = true;
            }

            if (alt) {
                out.println();
                out.println("@AlternateTitle(\"" + orig + "\")");
            }

            if (!field.isEmpty()) field = field.substring(0,1).toLowerCase() + field.substring(1);
            else field = "no_name";

            out.println("public String " + field + ";");
            i++;
        }

        out.println("}");


        System.out.println("Finished writing class");
    }
}
