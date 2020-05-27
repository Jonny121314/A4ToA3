package com.hongrant.www.achieve.comm.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import com.itextpdf.text.Document;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;

public class A4PdfToA3Word {

	/**
	 * A4格式pdf转成A4格式word
	 * */
	public static void a4PdfToA4Word(String A4PdfPath, String A4WordSavePath) {
		pdfToWord(A4PdfPath, A4WordSavePath);
	}
	
	/**
	 * A4格式pdf转成A3格式word
	 * */
	public static void a4PdfToA3Word(String A4PdfPath, String A3WordSavePath) {
		File A4PdfFile = new File(A4PdfPath);
		String A3PdfSaveFile = A4PdfFile.getParentFile()+"\\A3\\"+A4PdfFile.getName();
		File A3PdfFile = new File(A3PdfSaveFile);
		if (!A3PdfFile.getParentFile().exists()) {
			A3PdfFile.getParentFile().mkdirs();
		}
		concatPDFs(A4PdfPath, A3PdfSaveFile, true);//A4格式pdf转成A3格式pdf
		pdfToWord(A3PdfSaveFile, A3WordSavePath);//A3格式pdf转成A3格式word
		deleteFile(A3PdfSaveFile);//删除中间生成的A3格式pdf文件
	}

	   /**
	    * 删除单个文件
	    *
	    * @param fileName 要删除的文件的文件名
	    * @return 单个文件删除成功返回true，否则返回false
	    */
	   public static boolean deleteFile(String fileName) {
	       File file = new File(fileName);
	       // 如果文件路径所对应的文件存在，并且是一个文件，则直接删除
	       if (file.exists() && file.isFile()) {
	           if (file.delete()) {
	               return true;
	           } else {
	               return false;
	           }
	       } else {
	           return false;
	       }
	   }
	
	/**
	 * 两页打印到一页
	 * pdf A4转成A3格式
	 * @param pdfFile 原地址
	 * @param pdfSavePath 保存地址
	 * @param paginate
	 */
	private static void concatPDFs(String pdfFile, String pdfSavePath, boolean paginate) {

		Document document = new Document();
		OutputStream outputStream = null;
		try {
			File file = new File(pdfSavePath);
			if (!file.getParentFile().exists()) {
				file.getParentFile().mkdirs();
			}
			outputStream = new FileOutputStream(pdfSavePath);
			//读取PDF
			InputStream pdf = new FileInputStream(pdfFile);
			PdfReader pdfReader = new PdfReader(pdf);
			int totalPages = pdfReader.getNumberOfPages();

			//创建PDF Writer
			PdfWriter writer = PdfWriter.getInstance(document, outputStream);
			document.open();
			BaseFont baseFont = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
			PdfContentByte contentByte = writer.getDirectContent();

			Rectangle rec = pdfReader.getPageSize(1);
			//新pdf为两个宽度
			Rectangle newRec = new Rectangle(0, 0, rec.getWidth() * 2, rec.getHeight());
			document.setPageSize(newRec);

			PdfImportedPage page;
			PdfImportedPage page2;
			int currentPageNumber = 0;

			//舍弃奇数的最后一页, 删除该设置后面不受影响
			//			totalPages = totalPages >> 1 << 1;
			for (int pageIndex = 0; pageIndex < totalPages; pageIndex += 2) {
				document.newPage();
				currentPageNumber++;

				//原始第一页设置到左边
				page = writer.getImportedPage(pdfReader, pageIndex + 1);
				contentByte.addTemplate(page, 0, 0);

				//第二页设置到右边
				if (pageIndex + 2 <= totalPages) {
					page2 = writer.getImportedPage(pdfReader, pageIndex + 2);
					contentByte.addTemplate(page2, rec.getWidth(), 0);
				}

				//设置页码
				if (paginate) {
					contentByte.beginText();
					contentByte.setFontAndSize(baseFont, 13);
					//contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, "" + currentPageNumber + "/" + (totalPages / 2 + (totalPages % 2 == 0 ? 0 : 1)), newRec.getWidth() / 2, 17, 0);
					contentByte.showTextAligned(PdfContentByte.ALIGN_CENTER, "" + currentPageNumber , newRec.getWidth() / 2, 17, 0);
					contentByte.endText();
				}
			}

			outputStream.flush();
			document.close();
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (document.isOpen())
				document.close();
			try {
				if (outputStream != null)
					outputStream.close();
			} catch (IOException ioe) {
				ioe.printStackTrace();
			}
		}
	}

	/**
	 * pdf转成word格式
	 * @param blankPdfSavePath
	 * @param blankWordSavePath
	 */
	private static void pdfToWord(String pdPath, String wordSavePath) {
		PdfDocument pdf = new PdfDocument();
		pdf.loadFromFile(pdPath);//加载PDF
		//保存为Word格式
		pdf.saveToFile(wordSavePath, FileFormat.DOCX);
		pdf.close();
	}

	public static void main(String[] args) {
		String A4PdPath = "d:\\12345.pdf";
		String A3WordSavePath = "d:\\12345.docx";
		a4PdfToA3Word(A4PdPath, A3WordSavePath);
		System.out.println("a4转a3成功");
	}
}
