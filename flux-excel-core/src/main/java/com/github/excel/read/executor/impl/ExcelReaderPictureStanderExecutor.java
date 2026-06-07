package com.github.excel.read.executor.impl;

import com.github.excel.constant.ExcelConstant;
import com.github.excel.enums.ExcelPictureTypeEnum;
import com.github.excel.model.ExcelReaderPictureModel;
import com.github.excel.read.executor.ExcelReaderPictureExecutor;
import com.google.common.collect.Lists;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 图片导入执行器
 * @author Vico
 * @create 2022-12-14 17:26
 */
public class ExcelReaderPictureStanderExecutor implements ExcelReaderPictureExecutor {

    /**
     * 获取xlsx 图片
     *
     * @param sheet
     * @return
     */
    public Map<String, List<ExcelReaderPictureModel>> getXSSFSheetPicture(XSSFSheet sheet) {
        //returns the existing SpreadsheetDrawingML from the sheet, or creates a new one
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        //loop through all of the shapes in the drawing area
        List<ExcelReaderPictureModel> pictureModelList = Lists.newArrayList();
        for (XSSFShape shape : drawing.getShapes()) {
            if (shape instanceof Picture) {
                //convert the shape into a picture
                XSSFPicture picture = (XSSFPicture) shape;
                XSSFClientAnchor clientAnchor = picture.getClientAnchor();
                String suffix = ExcelPictureTypeEnum.getTypeSuffix(picture.getPictureData().getPictureType());
                ExcelReaderPictureModel pictureModel = new ExcelReaderPictureModel();
                pictureModel.setBytes(picture.getPictureData().getData());
                pictureModel.setRowIndex(clientAnchor.getRow1());
                pictureModel.setColIndex((int) clientAnchor.getCol1());
                pictureModel.setSuffix(suffix);
                pictureModel.setSheetName(sheet.getSheetName());
                pictureModel.setPoint(String.valueOf(clientAnchor.getRow1()) + ExcelConstant.COMMA_CHAR + clientAnchor.getCol1());
                pictureModelList.add(pictureModel);
            }
        }
        return pictureModelList.stream().collect(Collectors.groupingBy(ExcelReaderPictureModel::getPoint));
    }

    /**
     * 获取xls 图片
     *
     * @param sheet
     * @return
     */
    public Map<String, List<ExcelReaderPictureModel>> getHSSFSheetPicture(HSSFSheet sheet) {
        List<ExcelReaderPictureModel> pictureModelList = Lists.newArrayList();
        List<HSSFShape> shapes = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : shapes) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture pic = (HSSFPicture) shape;
                HSSFClientAnchor clientAnchor = pic.getClientAnchor();
                HSSFPictureData picData = sheet.getWorkbook().getAllPictures().get(pic.getPictureIndex() - ExcelConstant.ONE_INT);
                String suffix = ExcelPictureTypeEnum.getTypeSuffix(picData.getPictureType());
                ExcelReaderPictureModel pictureModel = new ExcelReaderPictureModel();
                pictureModel.setBytes(picData.getData());
                pictureModel.setRowIndex(clientAnchor.getRow1());
                pictureModel.setColIndex((int) clientAnchor.getCol1());
                pictureModel.setSuffix(suffix);
                pictureModel.setSheetName(sheet.getSheetName());
                pictureModel.setPoint(String.valueOf(clientAnchor.getRow1()) + ExcelConstant.COMMA_CHAR + clientAnchor.getCol1());
                pictureModelList.add(pictureModel);
            }
        }
        return pictureModelList.stream().collect(Collectors.groupingBy(ExcelReaderPictureModel::getPoint));
    }

}
