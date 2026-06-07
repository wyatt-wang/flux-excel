package com.github.format;

import com.github.excel.read.format.ExcelDefaultReaderDataFormat;
import com.github.excel.write.ExcelDefaultWriterDataFormat;
import org.junit.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.Year;
import java.time.YearMonth;

import static org.junit.Assert.assertEquals;

public class ExcelJavaTimeFormatTest {

    @Test
    public void writerFormatsJavaTimeTypesWithPattern() {
        ExcelDefaultWriterDataFormat formatter = new ExcelDefaultWriterDataFormat();

        assertEquals("2026-05-31 22:30:15", formatter.format(LocalDateTime.of(2026, 5, 31, 22, 30, 15), "yyyy-MM-dd HH:mm:ss"));
        assertEquals("2026/05/31", formatter.format(LocalDate.of(2026, 5, 31), "yyyy/MM/dd"));
        assertEquals("22:30:15", formatter.format(LocalTime.of(22, 30, 15), "HH:mm:ss"));
        assertEquals("2026-05", formatter.format(YearMonth.of(2026, 5), "yyyy-MM"));
        assertEquals("2026", formatter.format(Year.of(2026), "yyyy"));
    }

    @Test
    public void readerParsesJavaTimeTypesWithPattern() throws Exception {
        ExcelDefaultReaderDataFormat formatter = new ExcelDefaultReaderDataFormat();

        assertEquals(LocalDateTime.of(2026, 5, 31, 22, 30, 15), formatter.format("2026-05-31 22:30:15", "yyyy-MM-dd HH:mm:ss", LocalDateTime.class));
        assertEquals(LocalDate.of(2026, 5, 31), formatter.format("2026/05/31", "yyyy/MM/dd", LocalDate.class));
        assertEquals(LocalTime.of(22, 30, 15), formatter.format("22:30:15", "HH:mm:ss", LocalTime.class));
        assertEquals(YearMonth.of(2026, 5), formatter.format("2026-05", "yyyy-MM", YearMonth.class));
        assertEquals(Year.of(2026), formatter.format("2026", "yyyy", Year.class));
    }
}
