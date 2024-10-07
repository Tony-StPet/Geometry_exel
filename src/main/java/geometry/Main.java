package geometry;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import geometry.Rectangle;
import java.awt.*;
import java.io.FileInputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        try {
            List<Rectangle> rectangles = Loader.readExcelAndGetRectangles("ex1.xlsx");

            // Выводим список прямоугольников на консоль
            for (Rectangle rectangle : rectangles) {
                System.out.println(rectangle);

            }System.out.println("средняя площадь:: " + getAverageArea(rectangles));
            System.out.println("макс периметр среди всех прямоугольников:: " + getMaxPerimeter(rectangles));
            printRectanglesSortedByArea(rectangles);

        } catch (Exception e) {
            System.out.println("Ошибка при чтении Excel-файла: " + e.getMessage());
        }

        try {
            Connection connection = DriverManager.getConnection(
                    "jdbc:mysql://localhost:3306/Ships",
                    "root", "1404Root)");
            // Создание объекта Statement для выполнения SQL-запросов
            Statement statement = connection.createStatement();

            // Выполнение SQL-запроса для выборки данных из таблицы Ships
            ResultSet results = statement.executeQuery("SELECT * FROM Ship");

            // Обработка результата
            while (results.next()) {
                Integer id = results.getInt("id"); // Предполагается, что в таблице есть поле id
                String name = results.getString("name");
                String type = results.getString("type");
                Double speed = results.getDouble("speed");
                String armament = results.getString("armament");

                // Вывод данных на консоль
                System.out.println("ID: " + id + ", Name: " + name + ", Type: " + type +
                        ", Speed: " + speed + " knots, Armament: " + armament);
            }

            // Закрытие подключения
            connection.close();
        } catch (SQLException e) {
            e.printStackTrace(); // Вывод информации об ошибках
        }


    }
// Метод для вычисления средней площади всех прямоугольников:
    public static double getAverageArea(List<Rectangle> rectangles) {
        if (rectangles.isEmpty()) {
            return 0;
        }
        double totalArea = 0;
        for (Rectangle rectangle : rectangles) {
            totalArea += rectangle.area();
        }
        return totalArea / rectangles.size();
    }

//    Метод для вычисления максимального периметра среди всех прямоугольников:
public static double getMaxPerimeter(List<Rectangle> rectangles) {
    if (rectangles.isEmpty()) {
        return 0;
    }
    double maxPerimeter = 0;
    for (Rectangle rectangle : rectangles) {
        double perimeter = rectangle.perimeter();
        if (perimeter > maxPerimeter) {
            maxPerimeter = perimeter;
        }
    }
    return maxPerimeter;
}

//    Метод для вывода списка прямоугольников, отсортированных по площади:
public static void printRectanglesSortedByArea(List<Rectangle> rectangles) {
    if (rectangles.isEmpty()) {
        System.out.println("Список прямоугольников пуст.");
        return;
    }
    List<Rectangle> sortedRectangles = new ArrayList<>(rectangles);
    sortedRectangles.sort(Comparator.comparingDouble(Rectangle::area));
    System.out.println("Список прямоугольников, отсортированный по площади:");
    for (Rectangle rectangle : sortedRectangles) {
        System.out.println(rectangle);
    }
}


}
