import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.stream.Stream;

public class Rename {

    private static final SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");


    public void doWork(String pdfPath, String wbPath) throws IOException {


        Stream<Path> filePathStream = Files.walk(Paths.get(pdfPath));

        filePathStream
                .filter(Files::isRegularFile)
                .filter(filePath -> filePath.toString().endsWith("pdf"))
                .forEach(filePath -> {

                    System.out.println(filePath);


                    try {

                        PDDocument doc = PDDocument.load(filePath.toFile());
                        int count = doc.getNumberOfPages();
                        doc.close();

                        String newFileName =
                                getName(filePath.getFileName().toString()
                                                .replace(".pdf", "")
                                                .replace("PDF", "")
                                        , wbPath, count);

                        if (!newFileName.trim().equals("")) {

                            Files.move(filePath, filePath.resolveSibling(newFileName));
                        }


                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
    }


    private String getName(String cardNum, String wbPath, int pagesCount) throws IOException {

        String data = "";

        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(new File(wbPath));
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();

        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.getRow(i);

            String cardTxt = getCellText(row.getCell(0)).trim();

            if (cardTxt.equalsIgnoreCase(cardNum)) {

                String chekCell = getCellText(row.getCell(3));
                String onlyName = "";

                if (chekCell.contains("#5")) {
                    onlyName = cleanName(getCellText(row.getCell(15)));
                    data = "Требования кредитора (" + onlyName + ") с приложениями на " + pagesCount + " л..pdf";
                } else {


                    String[] chekCellArr = chekCell.split("\n");
                    data = chekCellArr[2].trim() + " " + chekCellArr[1] + " (" + chekCellArr[0] + ") на " + pagesCount + " л..pdf";


                }


            }
        }

        workbook.close();

        System.out.println(data);

        return data
                .replace("#","")
                .replace("\r","")
                .replace("\n","")
                .replace("\\","")
                .replace("/"," ");

    }


    private String cleanName(String txt) {

        return txt.trim()
                .replace("«", "")
                .replace("\"", "")
                .replace("Республика", "Р.")
                .replace("республика", "Р.")
                .replace("Республики", "Р.")
                .replace("республики", "Р.")
                .replace("»", "");
    }


    private String getCellText(Cell cell) {

        String result = "";

        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    result = cell.getRichStringCellValue().getString();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = sdf.format(cell.getDateCellValue());
                    } else {
                        result = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    result = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    result = cell.getCellFormula();
                    break;
                case BLANK:
                    result = "";
                    break;
                default:
                    System.out.println("Что-то пошло не так");
            }
        }
        return result;
    }

}
