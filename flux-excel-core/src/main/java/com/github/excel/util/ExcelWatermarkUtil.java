package com.github.excel.util;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetBackgroundPicture;

import javax.imageio.ImageIO;
import java.awt.AlphaComposite;
import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;

public final class ExcelWatermarkUtil {

	private ExcelWatermarkUtil() {
	}

	public static void applyTextWatermark(Workbook workbook, String text) {
		if (!(workbook instanceof XSSFWorkbook) || text == null || text.trim().isEmpty()) {
			return;
		}
		XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
		byte[] imageBytes = createWatermarkImage(text);
		xssfWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
		List<XSSFPictureData> pictures = xssfWorkbook.getAllPictures();
		XSSFPictureData pictureData = pictures.get(pictures.size() - 1);
		for (int i = 0; i < xssfWorkbook.getNumberOfSheets(); i++) {
			applyToSheet(xssfWorkbook.getSheetAt(i), pictureData);
		}
	}

	private static void applyToSheet(XSSFSheet sheet, XSSFPictureData pictureData) {
		if (Objects.isNull(sheet) || Objects.isNull(pictureData)) {
			return;
		}
		POIXMLDocumentPart.RelationPart relationPart = sheet.addRelation(null, XSSFRelation.IMAGES, pictureData);
		CTSheetBackgroundPicture picture = sheet.getCTWorksheet().isSetPicture()
				? sheet.getCTWorksheet().getPicture()
				: sheet.getCTWorksheet().addNewPicture();
		picture.setId(relationPart.getRelationship().getId());
	}

	private static byte[] createWatermarkImage(String text) {
		BufferedImage image = new BufferedImage(360, 240, BufferedImage.TYPE_INT_ARGB);
		Graphics2D graphics = image.createGraphics();
		try {
			graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			graphics.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
			graphics.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, 0.18f));
			graphics.setColor(new Color(120, 120, 120));
			graphics.setFont(new Font("Arial", Font.BOLD, 42));
			graphics.rotate(Math.toRadians(-28), 180, 120);
			graphics.drawString(text, 48, 132);
		} finally {
			graphics.dispose();
		}
		try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
			ImageIO.write(image, "png", outputStream);
			return outputStream.toByteArray();
		} catch (IOException e) {
			throw new IllegalStateException("Create excel watermark image failed", e);
		}
	}
}
