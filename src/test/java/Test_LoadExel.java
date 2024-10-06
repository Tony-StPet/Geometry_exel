import geometry.GeometryException;
import geometry.Loader;
import geometry.Rectangle;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Vector;

public class Test_LoadExel {

    // Загрузка существующего файла "ex1.xlsx" должна проходить без ошибок, не возникает исключений
    // метод Assertions.assertDoesNotThrow() - внутри лямбда-выражения не выбрасывается никакое исключение
    @Test
    public void testLdExel_Ok_0() {
        String s = "ex1.xlsx";
        Assertions.assertDoesNotThrow(() -> Loader.loadExel(s));
    }

    // при вызове Loader.loadExel("ex2.xlsx") возникает исключение - файла нет
    // метод Assertions.assertThrows() -  внутри лямбда-выражения вызывается исключение типа Exception.class
    // при попытке загрузки несуществующего или неверного файла "ex2.xlsx" возникает ошибка
    @Test
    public void testLoadBad1() {
        Assertions.assertThrows(Exception.class, ()-> Loader.loadExel("ex2.xlsx"));
    }


    //   Тест чтения файла с несколькими листами:
    //    - Создать тестовый Excel-файл с несколькими листами.
    //    - Проверить, что метод loadExel() корректно обрабатывает файл с несколькими листами, например, читая данные только с первого листа
    @Test
    public void testLoadExelWithMultipleSheets() throws Exception {
        // Создаем тестовый Excel-файл с несколькими листами
        String filename = "test_multi_sheet.xlsx";
        try (FileOutputStream fos = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet1 = workbook.createSheet("Sheet1");
            Sheet sheet2 = workbook.createSheet("Sheet2");

            // Заполняем данными первый лист
            Row row1 = sheet1.createRow(0);
            row1.createCell(0).setCellValue("Value1");
            row1.createCell(1).setCellValue("Value2");

            // Заполняем данными второй лист
            Row row2 = sheet2.createRow(0);
            row2.createCell(0).setCellValue("Value3");
            row2.createCell(1).setCellValue("Value4");

            workbook.write(fos);
        }

        // Вызываем метод loadExel() и проверяем, что он работает корректно
        //перехват вывода, который отправляется в консоль

// Создается экземпляр ByteArrayOutputStream. Этот класс представляет собой выходной поток, который будет сохранять
// все данные, которые обычно отправлялись бы в консоль.
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

//Изменение стандартного потока вывода системы (System.out) на новый PrintStream, который использует ByteArrayOutputStream.
// Это значит, что любой вывод, который раньше отправлялся в консоль с помощью System.out.println(), теперь будет направляться в ByteArrayOutputStream.
// Таким образом, после этих двух строк кода, все выводы, которые производит ваше приложение, будут сохраняться в ByteArrayOutputStream, а не выводиться в консоль.
// Это полезно для тестирования, так как позволяет проверять содержимое вывода, не загрязняя консоль. Например, вы можете проверить, что ваше приложение выводит ожидаемые строки, используя методы toString() или toByteArray() на ByteArrayOutputStream.
        System.setOut(new PrintStream(outputStream));

        Loader.loadExel(filename);

        String output = outputStream.toString();
        Assertions.assertTrue(output.contains("Value1")); // Проверяем, что данные из первого листа обработаны
        Assertions.assertTrue(output.contains("Value2")); // Проверяем, что данные из первого листа обработаны
        Assertions.assertFalse(output.contains("Value3")); // Проверяем, что данные из второго листа обработаны
        Assertions.assertFalse(output.contains("Value4")); // Проверяем, что данные из второго листа обработаны


        // Удаляем тестовый файл
        File file = new File(filename);
        file.delete();

        // После завершения теста - восстановление стандартного вывода в консоль
        System.setOut(System.out);
    }

    // тест на тип данных, тип данных:: строка, дабл, логический, дата, формула
    @Test
    public void testLoadExelWithDifferentDataTypes() throws Exception {
        // Создаем тестовый Excel-файл с различными типами данных
        String filename = "test_data_types.xlsx";
        try (FileOutputStream fos = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Заполняем ячейки различными типами данных
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("String value");
            row.createCell(1).setCellValue(123.45);
            row.createCell(2).setCellValue(false);
            row.createCell(3).setCellValue(new Date());
            row.createCell(4).setCellFormula("SUM(B1:B1)");
            row.createCell(5).setCellValue(true);
            Row row1 = sheet.getRow(0);
            boolean booleanValue1 = row1.getCell(2).getBooleanCellValue();
            boolean booleanValue2 = row1.getCell(5).getBooleanCellValue();
            System.out.println("Boolean value: " + booleanValue1+ " " + booleanValue2);
            workbook.write(fos);
        }

        // Вызываем метод loadExel() и перехватываем вывод
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        System.setOut(new PrintStream(outputStream));

        Loader.loadExel(filename);
        String output = outputStream.toString();

        // Проверяем, что все типы данных обработаны корректно
        Assertions.assertTrue(output.contains("String value"));
        Assertions.assertTrue(output.contains("123.45"));
        Assertions.assertFalse(output.contains("false"));
        Assertions.assertFalse(output.contains(new SimpleDateFormat("yyyy-MM-dd").format(new Date())));
        Assertions.assertTrue(output.contains("123.45"));
        Assertions.assertFalse(output.contains("true"));
        // Очищаем после теста
        File file = new File(filename);
        file.delete();
    }

    // создание пустого файла и одого листа в нем и пробуем читать
    @Test
    public void testLoadExcelWithOneSheet() throws Exception {
        // Создаем Excel-файл с одним листом
        String filename = "one_sheet.xlsx";
        try (FileOutputStream fos = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("MySheet");
            workbook.write(fos);
        }

        // Вызываем метод loadExel() с файлом, содержащим один лист
        Assertions.assertDoesNotThrow(() -> Loader.loadExel(filename));
        // Очищаем после теста
        File file = new File(filename);
        file.delete();
    }

    // Создаем новый Excel-файл "no_sheets.xlsx", но не добавляем в него ни одного листа - должна быть ошибка
    @Test
    public void testLoadExcelWithNoSheets() throws Exception{
        // Создаем Excel-файл без листов
        String filename = "no_sheets.xlsx";
        try (FileOutputStream fos = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            // Не создаем листы
            workbook.write(fos);
        }

// IllegalArgumentException.class - это класс исключения IllegalArgumentException, который является подклассом RuntimeException в Java.
// Этот тип исключения используется, когда метод получает аргумент, который не соответствует ожидаемому значению или типу.
// Это означает, что проблема возникает из-за неправильного использования метода, а не из-за непредвиденной ошибки во время выполнения.
//В контексте  теста testLoadExcelWithNoSheets(), вы ожидаемо, что метод Loader.loadExel() выбросит именно IllegalArgumentException,
// когда мы передаем ему Excel-файл, не содержащий ни одного листа. Это связано с тем, что попытка получить лист
// с индексом 0 из этого файла приведет к возникновению этого исключения.
// Использование Assertions.assertThrows(IllegalArgumentException.class, () -> Loader.loadExel(filename))
// позволяет проверить, что метод Loader.loadExel() действительно выбрасывает IllegalArgumentException в этой ситуации, а не какое-то другое исключение.
// Другими словами, IllegalArgumentException.class - это способ указать, что мы ожидаем конкретный тип исключения,
// а не любое исключение в целом. Это делает тесты более точными и помогает лучше понять, как реализован метод Loader.loadExel().
        // Проверяем, что метод loadExel() выбрасывает именно IllegalArgumentException
        Assertions.assertThrows(IllegalArgumentException.class, () -> Loader.loadExel(filename));
        File file = new File(filename);
        file.delete();
    }

    @Test
    public void testLoadInvalidExcelFile() {
        // Создаем файл с расширением .xlsx, но неверным содержимым
        String filename = "invalid.xlsx";
        try (FileOutputStream fos = new FileOutputStream(filename);
             BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(fos))) {
            writer.write("This is not a valid Excel file.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Проверяем, что метод loadExel() выбрасывает NotOfficeXmlFileException
        Assertions.assertThrows(NotOfficeXmlFileException.class, () -> Loader.loadExel(filename));
        File file = new File(filename);
        file.delete();
    }

//ячейка содержит многострочное значение и функция корректно обрабатывает значения  с переносами строк
    @Test
    public void testLoadExcelWithLineBreaks() throws Exception {
        // Создаем тестовый Excel-файл с ячейками, разделенными переносом строки
        String filename = "test_with_linebreaks.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            Row row1 = sheet.createRow(0);
            Cell cell1 = row1.createCell(0);
            // многострочное значение
            cell1.setCellValue("This is a\nmulti-line\nvalue");
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }


    // метод корректно обрабатывает значения ячеек с переносами строк, выводя их в консоль в исходном виде, а не в виде экранированных строк
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            System.setOut(new PrintStream(outputStream));
            Loader.loadExel(filename);
            String output = outputStream.toString();
            //  Проверяется, что в выводе присутствует многострочное значение
            Assertions.assertTrue(output.contains("This is a\nmulti-line\nvalue"));
            //  Проверяется, что в выводе отсутствует строковое представление переносов строк
            Assertions.assertFalse(output.contains("This is a\\nmulti-line\\nvalue"));
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.setOut(System.out); // Возвращаем стандартный вывод
        }
        File file = new File(filename);
        file.delete();
    }

    // тест проверяет, что метод Loader.loadExel() корректно обрабатывает пустые ячейки, выводя их как пустые строки
    @Test
    public void testLoadExcelWithEmptyCells() throws Exception{
        // Создаем тестовый Excel-файл с пустыми ячейками
        String filename = "test_with_empty_cells.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(filename);
             Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
//В этом файле создается один лист с двумя строками.
//- В первой строке создается одна ячейка, которая остается пустой.
//- Во второй строке создается две ячейки, где первая ячейка остается пустой, а во вторую ячейку записывается значение "Non-empty cell".
            Row row1 = sheet.createRow(0);
            Cell cell1 = row1.createCell(0);
            cell1.setCellValue("");
            Row row2 = sheet.createRow(1);
            Cell cell2 = row2.createCell(0);
            row2.createCell(1).setCellValue("Non-empty cell");
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Вызываем метод loadExel() и проверяем, что пустые ячейки обрабатываются корректно
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            System.setOut(new PrintStream(outputStream));
            Loader.loadExel(filename);
            String output = outputStream.toString();
        //Проверяется, что в выводе присутствует пустая строка, соответствующая первой пустой ячейке
            Assertions.assertTrue(output.contains(""));
        // Проверяется, что в выводе присутствует строка "Non-empty cell", соответствующая непустой ячейке
            Assertions.assertTrue(output.contains("Non-empty cell"));
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.setOut(System.out); // Возвращаем стандартный вывод
        }
        File file = new File(filename);
        file.delete();
    }

}
