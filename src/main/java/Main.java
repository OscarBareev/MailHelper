import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

        String pdfPath = "D:\\TestDir\\20.09.2021";
        String wbPath = "D:\\TestDir\\Список.xlsx";

        Rename rename = new Rename();
        rename.doWork(pdfPath, wbPath);

    }
}
