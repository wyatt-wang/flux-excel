package com.github.excel.read.executor;

import com.github.excel.model.ExcelReaderPictureModel;
import com.google.common.collect.Maps;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;
import java.util.Map;

/**
 * 读取图片执行器
 * @author Vico
 * @create 2022-12-14 17:25
 */
public interface ExcelReaderPictureExecutor {
    /**
     * 读取xlsx图片
     * @param sheet sheet
     * @return
     */
    Map<String, List<ExcelReaderPictureModel>> getXSSFSheetPicture(XSSFSheet sheet);

    /**
     * 读取xls图片
     * @param sheet sheet
     * @return
     */
    Map<String, List<ExcelReaderPictureModel>> getHSSFSheetPicture(HSSFSheet sheet);

    /**
     * 获取图片
     * @param sheet sheet
     * @return
     */
    default Map<String, List<ExcelReaderPictureModel>> getSheetPictureMap(Sheet sheet){
        if(sheet instanceof XSSFSheet){
            return getXSSFSheetPicture((XSSFSheet) sheet);
        }else if(sheet instanceof HSSFSheet){
            return getHSSFSheetPicture((HSSFSheet) sheet);
        }else{
            return Maps.newHashMap();
        }
    }
}
