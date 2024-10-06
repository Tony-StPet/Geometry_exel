package geometry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.Vector;

    // класс Загрузчик
public class Loader {

    //СТАТИЧЕСКИЙ метод загрузитьСписокПрямоугольников
    //входной параметр - это строковое имя файла
    //результат - список прямоугольников
    public static List<Rectangle> loadRecList(String filename) throws FileNotFoundException, GeometryException {
        List<Rectangle> rectangles = new Vector<>();    //привычнее использовать ArrayList
        try(Scanner scanner = new Scanner(new File(filename))){
            while (scanner.hasNext()){
                String line = scanner.nextLine();
                String[] massiv = line.split(" ");
                if(massiv.length != 2) throw new GeometryException("в каждой строчке файла должно быть только 2 числа");
                double w = Double.parseDouble(massiv[0]);
                double l = Double.parseDouble(massiv[1]);
                Rectangle rect = new Rectangle(w, l);
                rectangles.add(rect);
            }
        }

        return rectangles;
    }


// метод загружает содержимое Excel-файла, перебирает все ячейки и выводит их значения в консоль.
// Он не возвращает никаких результатов, только выводит данные. метод должен корректно обрабатывать как существующие,
// так и несуществующие файлы. Если файл не найден, он должен вызывать исключение

    public static void loadExel (String filename) throws Exception {
    //Метод использует try-with-resources конструкцию, чтобы гарантированно закрыть FileInputStream и Workbook после завершения работы
        try (FileInputStream fis = new FileInputStream(filename);
    // Создается FileInputStream для чтения файла  - Затем создается Workbook на основе FileInputStream, предполагается, что это файл формата XLSX.
             Workbook workbook = new XSSFWorkbook(fis)) {

            //  предполагается, что в файле только один лист
            Sheet sheet = workbook.getSheetAt(0);
            // итерация по всем строкам листа
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // Для каждой строки выполняется итерация по всем ячейкам cellIterator = row.cellIterator()
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
    // Обработка различных типов данных в ячейках Для каждой ячейки вывод в консоль (cell.toString()), разделенное символом ;
                    System.out.print(cell.toString() + " ; ");
                }
                System.out.println();  // Начало новой строки
            }
        }
    }



}