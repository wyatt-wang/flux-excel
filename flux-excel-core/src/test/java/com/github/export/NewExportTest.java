package com.github.export;

import com.github.BaseExcelTest;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.net.URL;
import java.util.List;
import java.util.Optional;
import java.util.function.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author Vico
 * @create 2023-05-08 19:34
 */
@Slf4j
public class NewExportTest extends BaseExcelTest {
    @Test
    public void testPath(){
        String path = Optional.ofNullable(this.getClass().getClassLoader().getResource("")).map(URL::getPath).orElseThrow(IllegalArgumentException::new);
        assert path != null;
        log.info(path);
    }
    @Test
    public void java8(){
        Predicate<String> p = String::isEmpty;
        log.info(""+p.test(""));
        log.info(""+p.test("t"));
        List<String> b = Stream.of("a").filter(StringUtils::isNoneBlank).map(String::valueOf).collect(Collectors.toList());

//        Function<String, String> function = a -> a + " Jack!";
//        System.out.println(function.apply("Hello"));

        Function<String, String> function = a -> a + " Jack!";
        Function<String, String> function1 = a -> a + " Bob!";
        String greet = function.andThen(function1).apply("Hello");
        System.out.println(greet); // Hello Jack! Bob!


    }
    @Test
    public void testF(){
        BiFunction<Double, Long, String> biFunction = (a, b) -> String.valueOf(a) + String.valueOf(b);
        String str = biFunction.apply(22.2, 11L);
        System.out.println(str);

        BiFunction<Integer, Integer, String> testFunction = NewExportTest::sum;
        str = testFunction.apply(11, 22);
        System.out.println(str);

        DoubleFunction<String> doubleFunction = doub -> "结果：" + doub;
        System.out.println(doubleFunction.apply(1.6)); // 结果：1.6

        DoubleToIntFunction toIntFunction = e -> (int)e;
        System.out.println(toIntFunction.applyAsInt(11.22));
    }
    @Test
    public void consumer(){
        StringBuilder sb = new StringBuilder("Hello ");
        Consumer<StringBuilder> consumer = (str) -> str.append("Jack ");
        Consumer<StringBuilder> consumer1 = NewExportTest::append;
        consumer.andThen(consumer1).accept(sb);
        System.out.println(sb.toString());	// Hello Jack!
    }
    @Test
    public void predicate(){
        Predicate<Integer> predicate = number -> number != 0;
//        predicate = predicate.negate();
        System.out.println(predicate.test(10));    //false
    }

    public static String sum(int a,int b){
        int c = a + b;
        return String.valueOf(c) ;
    }

    public static void append(StringBuilder sb) {
        sb.append("Jack1");
    }
}
