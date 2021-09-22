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
import java.util.Calendar;
import java.util.Date;
import java.util.stream.Stream;

public class Rename {

    private static final SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
    Date date = new Date(System.currentTimeMillis());

    public String xlsxPath = "";

    private String findXlsxFile(String path) throws IOException {

        Stream<Path> filePathStream = Files.walk(Paths.get(path));

        filePathStream
                .filter(Files::isRegularFile)
                .filter(filePath -> filePath.toString().endsWith("xlsx"))
                .forEach(filePath -> {
                    xlsxPath = filePath.toString();
                });

        return xlsxPath;
    }


    public void doWork(String pdfPath) throws IOException {
        doWork(pdfPath, findXlsxFile(pdfPath));
    }


    public void doWork(String pdfPath, String wbPath) throws IOException {


        Stream<Path> filePathStream = Files.walk(Paths.get(pdfPath));

        filePathStream
                .filter(Files::isRegularFile)
                .filter(filePath -> filePath.toString().endsWith("pdf"))
                .forEach(filePath -> {

                    System.out.println(filePath);


                    try {
                        sortMail(filePath, wbPath);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
    }


    private void sortMail(Path filePath, String wbPath) throws IOException {

        String resultDir = filePath.getParent().getParent().toString() + "\\" + "Почта " + sdf.format(date);

        if (!Files.exists(Paths.get(resultDir))) {
            Files.createDirectory(Paths.get(resultDir));
        }

        String newFileName = getName(filePath.getFileName().toString(), wbPath, pagesNum(filePath));

        String fileDir = resultDir;

        if (newFileName.toLowerCase().contains("требования кредитора (")) {
            fileDir = resultDir + "\\" + "Требоввния кредиторов";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        } else if (newFileName.toLowerCase().contains("(3-")) {
            fileDir = resultDir + "\\" + "3-ие лицо";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        } else if (newFileName.toLowerCase().contains("(труд")) {
            fileDir = resultDir + "\\" + "Трудовая инспекция";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        } else if (newFileName.toLowerCase().contains("(ис")) {
            fileDir = resultDir + "\\" + "Истец";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        } else if (newFileName.toLowerCase().contains("(отв")) {
            fileDir = resultDir + "\\" + "Ответчик";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        } else if (newFileName.toLowerCase().contains("(б)")) {
            fileDir = resultDir + "\\" + "Банкротное дело";

            if (!Files.exists(Paths.get(fileDir))) {
                Files.createDirectory(Paths.get(fileDir));
            }
        }

        String finalPath = fileDir + "\\" + newFileName;

        if (!Files.exists(Paths.get(finalPath))){
            Files.move(filePath, Paths.get(finalPath));
        }



    }


    private int pagesNum(Path filePath) throws IOException {

        PDDocument doc = PDDocument.load(filePath.toFile());
        int count = doc.getNumberOfPages();
        doc.close();

        return count;
    }

    private String getName(String cardNum, String wbPath, int pagesCount) throws IOException {

        String data = "";

        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(new File(wbPath));
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();

        for (int i = 0; i < rows; i++) {
            XSSFRow row = sheet.getRow(i);

            String cardTxt = getCellText(row.getCell(0)).trim();

            if (cardTxt.equalsIgnoreCase(cardNum.replace(".pdf", "").replace(".PDF", ""))) {

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
                .replace("#", "")
                .replace("«", "")
                .replace("\"", "")
                .replace("»", "")
                .replace("\r", "")
                .replace("\n", "")
                .replace("\\", "")
                .replace("/", " ");
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
