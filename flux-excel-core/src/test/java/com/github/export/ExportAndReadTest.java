package com.github.export;

import com.github.BaseExcelTest;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.param.ExcelWriterListParam;
import com.github.excel.read.handler.reader.ExcelReader;
import com.github.excel.read.handler.reader.ExcelReaderFactory;
import com.github.excel.write.ExcelWriter;
import com.github.excel.write.ExcelWriterFactory;
import com.github.model.UserReadAndExportDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Vico
 * @create 2022-12-20 15:45
 */
@Slf4j
public class ExportAndReadTest extends BaseExcelTest {
    @Test
    public void exportAndRead() throws Exception {
        InputStream is = new FileInputStream("/Users/vico/Downloads/rabbitMQ.jpg");
        byte[] bytes = IOUtils.toByteArray(is);
        is.close();

        ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
        List<UserReadAndExportDto> userList = new ArrayList<>();
        for (int i = 1; i <=10; i++) {
            UserReadAndExportDto user = UserReadAndExportDto.builder().age(i).name("张三" + i).sex((byte) 1).height((float) 170 + i).avater(bytes).build();
            userList.add(user);
        }
        String excelName = "exportAndRead.xlsx";
        ExcelWriterListParam<UserReadAndExportDto> param = ExcelWriterListParam.<UserReadAndExportDto>builder().modelList(userList).modelCla(UserReadAndExportDto.class).sheetName("sheet").build();
        writer.writeList(param);
        OutputStream outputStream = new FileOutputStream("/Users/vico/Downloads/"+excelName);
        writer.process(outputStream, excelName, ExcelSuffixEnum.XLSX);


        InputStream inputStream = new FileInputStream("/Users/vico/Downloads/" + excelName);
        ExcelReader reader = ExcelReaderFactory.createUserReader(excelName, inputStream, null, true);
        reader.addList(UserReadAndExportDto.class,0);
        reader.parse();
        List<UserReadAndExportDto> modelList = reader.getList(UserReadAndExportDto.class);
        modelList.forEach(model->log.info(model.toString()));
    }

    @Test
    public void exportAndReadBean() throws Exception {
        InputStream is = new FileInputStream("/Users/vico/Downloads/rabbitMQ.jpg");
        byte[] bytes = IOUtils.toByteArray(is);
        is.close();

        ExcelWriter writer = ExcelWriterFactory.createUserModelWriter();
        UserReadAndExportDto user = UserReadAndExportDto.builder().age(10).name("张三").sex((byte) 1).height((float) 170).avater(bytes).build();
        String excelName = "exportAndReadBean.xlsx";
        writer.addModel(user, "测试", false, UserReadAndExportDto.class);
        OutputStream outputStream = new FileOutputStream("/Users/vico/Downloads/"+excelName);
        writer.process(outputStream, excelName, ExcelSuffixEnum.XLSX);


        InputStream inputStream = new FileInputStream("/Users/vico/Downloads/" + excelName);
        ExcelReader reader = ExcelReaderFactory.createUserReader(excelName, inputStream, null, true);
        reader.addModel(UserReadAndExportDto.class,0);
        reader.parse();
        UserReadAndExportDto readerModel = reader.getModel(UserReadAndExportDto.class);
        log.info(readerModel.toString());
    }

    @Test
    public void exportAndRead1() throws Exception {
        Object bs;
        bs = Array.newInstance(Byte.class, 100);
        System.out.println(bs.getClass());
        bs = Array.newInstance(Byte.TYPE, 100);
        System.out.println(bs.getClass());
    }
}
