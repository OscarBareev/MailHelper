import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {

        String pdfPath = "D:\\TestDir\\Требования из Тессы";
        String wbPath = "D:\\TestDir\\Выгрузка из тессы.xlsx";

        Rename rename = new Rename();
        rename.doWork(pdfPath, wbPath);

    }
}
