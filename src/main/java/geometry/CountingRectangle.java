package geometry;
import java.io.FileNotFoundException;
import java.util.*;

public class CountingRectangle {
    public static HashMap<Rectangle, Integer> countingRectangle(String filename) throws GeometryException, FileNotFoundException {
        List<Rectangle> recList = Loader.loadRecList(filename);
        HashMap<Rectangle, Integer> recMap = new HashMap<>();
        for (Rectangle rec : recList) {
            if (recMap.containsKey(rec)) {
                recMap.put(rec, recMap.get(rec) + 1);
            } else {
                recMap.put(rec, 1);
            }
        }
        return recMap;
    }
}

//        try (Scanner scanner = new Scanner(new File(filename))) {
//            while (scanner.hasNext()){
//                String line = scanner.nextLine();
//                String[] mas = line.split(" ");
//
//                Integer width = Integer.valueOf(mas[0]);
//                Integer length = Integer.valueOf(mas[1]);
//                Rectangle rec = new Rectangle(width, length);
//                rectangles.put(rec, Integer.valueOf(width));
//
//                if (rectangles.containsKey(rec)){
//                    rectangles.put(rec, rectangles.get(rec) + 1);
//                } else {
//                        rectangles.put(rec, 1);
//                }
//            }
//        }
//        catch (Exception e){
//            System.out.println("что-то не так "+e.getMessage());
//        }