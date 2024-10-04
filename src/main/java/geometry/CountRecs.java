package geometry;

import java.io.File;
import java.util.*;

public class CountRecs {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Enter file name: ");
        List<Rectangle> fileRecs = readRecs(new File(scanner.nextLine()));
        System.out.println(countForRecs(fileRecs));
    }

    static public HashMap<Rectangle, Integer> countForRecs(List<Rectangle> fileRecs) {
        HashMap<Rectangle, Integer> table = new HashMap<>();

        for (Rectangle rec : fileRecs) {
            int count = table.getOrDefault(rec, 0) + 1;
            table.put(rec, count);
        }
        return table;
    }

    static List<Rectangle> readRecs(File source) {
        List<Rectangle> newRecs = new ArrayList<>();
        try (Scanner scanner = new Scanner(source)) {
            while (scanner.hasNext()) {
                double[] inputData = Arrays.stream(scanner.nextLine().split(" "))
                        .mapToDouble(Double::parseDouble).toArray();
                newRecs.add(new Rectangle(inputData[0], inputData[1]));
            }
        } catch (Exception e) {
            System.out.println("Error in readRecs: " + e);
        }
        return newRecs;
    }

}