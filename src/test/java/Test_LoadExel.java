import geometry.GeometryException;
import geometry.Loader;
import geometry.Rectangle;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Vector;

public class Test_LoadExel {

    @Test
    public void testLdExel_Ok_0() {
        String s = "ex1.xlsx";
        Assertions.assertDoesNotThrow(() -> Loader.loadExel(s));
    }

    @Test
    public void testLoadBad1() {
        Assertions.assertThrows(Exception.class, ()-> Loader.loadExel("ex2.xlsx"));
    }
}