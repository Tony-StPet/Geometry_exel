package geometry;

import java.util.HashMap;
import java.util.List;

public class RectangleCounter {
    public static HashMap<Rectangle, Integer> countRectanglesFromFile(String fileName) throws Exception {
        List<Rectangle> rectanglesList = Loader.loadRecList(fileName);
        HashMap<Rectangle, Integer> rectangleMap = new HashMap<>();
        for (Rectangle rectangle : rectanglesList) {
            rectangleMap.put(rectangle, rectangleMap.getOrDefault(rectangle, 0) + 1);
        }
        return rectangleMap;
    }
}