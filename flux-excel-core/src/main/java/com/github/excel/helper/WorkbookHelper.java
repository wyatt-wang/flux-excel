package com.github.excel.helper;

import ch.qos.logback.core.util.CloseUtil;
import com.github.excel.enums.ExcelSuffixEnum;
import com.github.excel.exception.ExcelReaderException;
import com.github.excel.exception.ExcelWriterException;
import com.github.excel.param.ExcelReaderFileParam;
import com.github.excel.param.ExcelReaderParam;
import com.github.excel.param.ExcelReaderStreamParam;
import com.github.excel.param.ExcelWriterParam;
import com.github.excel.util.IOUtils;
import com.github.excel.util.ResourceLoader;
import com.github.excel.util.StringUtil;
import com.google.common.base.Throwables;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author Vico
 * @create 2023-09-05 15:16
 */
@Slf4j
public class WorkbookHelper {
    /**
     * 创建workbook
     * @param readerParam 读取参数
     * @return
     */
    public static Workbook createReadWorkBook(ExcelReaderParam readerParam){
        Workbook workbook = null;
        InputStream inputStream = null;
        try {
            if (readerParam instanceof ExcelReaderFileParam) {
                ExcelReaderFileParam fileParam = (ExcelReaderFileParam) readerParam;
                workbook = WorkbookFactory.create(fileParam.getFile(), fileParam.getPassword(), fileParam.getReadOnly());
            }else if (readerParam instanceof ExcelReaderStreamParam) {
                ExcelReaderStreamParam streamParam = (ExcelReaderStreamParam) readerParam;
                inputStream = streamParam.getStream();
                workbook = WorkbookFactory.create(streamParam.getStream(), streamParam.getPassword());
            }else{
                throw new ExcelReaderException("创建workbook失败，请检查参数！");
            }
        } catch (IOException e) {
            log.error("Workbook create failed , cause:{}", Throwables.getStackTraceAsString(e));
            throw new ExcelReaderException(e.getMessage());
        }finally {
            if (inputStream != null) {
                IOUtils.closeQuietlyMulti(inputStream);
            }
        }
        return workbook;
    }

    /**
     * todo 指定workbook密码
     * 创建写workbook
     * @param writerParam 写入参数
     * @return
     */
    public static Workbook createWriteWorkBook(ExcelWriterParam writerParam){
        Workbook workbook;
        InputStream inputStream = null;
        // 通过模版创建
        if (StringUtil.notEmpty(writerParam.getTemplate())) {
            try {
                inputStream = ResourceLoader.load(writerParam.getTemplate());
                workbook = WorkbookFactory.create(inputStream, writerParam.getTemplatePassword());
                return wrapStreamingWorkbook(workbook, writerParam);
            } catch (IOException e) {
                throw new ExcelWriterException(e);
            }finally {
                CloseUtil.closeQuietly(inputStream);
            }
        }else if (StringUtil.notEmpty(writerParam.getTemplateFilePath())) {
            // 通过文件创建
            try {
                File file = new File(writerParam.getTemplateFilePath());
                workbook = WorkbookFactory.create(file, writerParam.getTemplatePassword());
                return wrapStreamingWorkbook(workbook, writerParam);
            } catch (IOException e) {
                throw new ExcelWriterException(e);
            }finally {
                CloseUtil.closeQuietly(inputStream);
            }
        } else if (writerParam.getSuffixEnum() == ExcelSuffixEnum.XLSX) {
            // 创建xlsx workbook
            workbook = new XSSFWorkbook();
            if (Boolean.TRUE.equals(writerParam.getStreaming())) {
                // 创建xlsx workbook
                workbook = createStreamingXlsxWorkBook((XSSFWorkbook) workbook, writerParam);
            }
            return workbook;
        } else{
            // 创建xls workbook
            return new HSSFWorkbook();
        }
    }

    public static SXSSFWorkbook createStreamingXlsxWorkBook(ExcelWriterParam writerParam) {
        return createStreamingXlsxWorkBook(new XSSFWorkbook(), writerParam);
    }

    public static SXSSFWorkbook createStreamingXlsxWorkBook(XSSFWorkbook workbook, ExcelWriterParam writerParam) {
        return createStreamingXlsxWorkBook(
                workbook,
                writerParam.getRowAccessWindowSize(),
                writerParam.getCompressTempFiles(),
                writerParam.getUseSharedStringsTable()
        );
    }

    public static SXSSFWorkbook createStreamingXlsxWorkBook(Integer rowAccessWindowSize,
                                                            Boolean compressTempFiles,
                                                            Boolean useSharedStringsTable) {
        return createStreamingXlsxWorkBook(new XSSFWorkbook(), rowAccessWindowSize, compressTempFiles, useSharedStringsTable);
    }

    private static SXSSFWorkbook createStreamingXlsxWorkBook(XSSFWorkbook workbook,
                                                             Integer rowAccessWindowSize,
                                                             Boolean compressTempFiles,
                                                             Boolean useSharedStringsTable) {
        return new SXSSFWorkbook(
                workbook,
                resolveRowAccessWindowSize(rowAccessWindowSize),
                Boolean.TRUE.equals(compressTempFiles),
                Boolean.TRUE.equals(useSharedStringsTable)
        );
    }

    private static Workbook wrapStreamingWorkbook(Workbook workbook, ExcelWriterParam writerParam) {
        if (!Boolean.TRUE.equals(writerParam.getStreaming())) {
            return workbook;
        }
        if (workbook instanceof XSSFWorkbook) {
            return createStreamingXlsxWorkBook((XSSFWorkbook) workbook, writerParam);
        }
        return workbook;
    }

    private static int resolveRowAccessWindowSize(Integer rowAccessWindowSize) {
        if (rowAccessWindowSize == null || rowAccessWindowSize == 0 || rowAccessWindowSize < -1) {
            return com.github.excel.constant.ExcelConstant.DEFAULT_ROW_ACCESS_WINDOW_SIZE;
        }
        return rowAccessWindowSize;
    }
}
