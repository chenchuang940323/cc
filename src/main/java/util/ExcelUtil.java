package util;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

/**
 * excel相关工具类
 */
public class ExcelUtil {

    /**
     * 获取工作簿
     *
     * @param excelFilePath Excel文件路径
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(String excelFilePath) throws EncryptedDocumentException, IOException {
        return getWorkbook(new File(excelFilePath));
    }

    /**
     * 获取工作簿
     *
     * @param excelFile Excel文件
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(File excelFile) throws EncryptedDocumentException, IOException {
        return WorkbookFactory.create(excelFile);
    }

    /**
     * 获取工作簿
     *
     * @param excelFileInStream Excel文件输入流
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(InputStream excelFileInStream) throws EncryptedDocumentException, IOException {
        return WorkbookFactory.create(excelFileInStream);
    }

    /**
     * 获取工作簿
     *
     * @param excelFile Excel文件
     * @param password  文件密码
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(File excelFile, String password) throws EncryptedDocumentException, IOException {
        return getWorkbook(excelFile, password, false);
    }

    /**
     * 获取工作簿
     *
     * @param excelFileInStream Excel文件输入流
     * @param password          文件密码
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(InputStream excelFileInStream, String password)
            throws EncryptedDocumentException, IOException {
        return WorkbookFactory.create(excelFileInStream, password);
    }

    /**
     * 获取工作簿
     *
     * @param excelFile Excel文件
     * @param password  文件密码
     * @param readonly  是否只读
     * @return 工作簿
     * @throws EncryptedDocumentException 加密文档异常
     * @throws IOException                IO异常
     */
    public static Workbook getWorkbook(File excelFile, String password, boolean readonly)
            throws EncryptedDocumentException, IOException {
        // 检查文件是否存在
        if (!excelFile.exists()) { // 文件不存在
            return null;
        }
        // 返回工作簿
        return WorkbookFactory.create(excelFile, password, readonly);
    }


    /**
     * 获取单元格字符串值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 字符串值
     */
    public static String getCellValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格字符串值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 字符串值
     */
    public static String getCellValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.toString();
    }

    /**
     * 获取单元格字符串值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 字符串值
     */
    public static String getCellStringValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellStringValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格字符串值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 字符串值
     */
    public static String getCellStringValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.getStringCellValue();
    }

    /**
     * 获取单元格富文本值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 富文本值
     */
    public static RichTextString getCellRichTextValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellRichTextValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格富文本值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 富文本值
     */
    public static RichTextString getCellRichTextValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.getRichStringCellValue();
    }

    /**
     * 获取单元格数字值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 数字值
     */
    public static Double getCellNumericValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellNumericValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格数字值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 数字值
     */
    public static Double getCellNumericValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.getNumericCellValue();
    }

    /**
     * 获取单元格日期值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 日期值
     */
    public static Date getCellDateValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellDateValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格日期值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 日期值
     */
    public static Date getCellDateValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.getDateCellValue();
    }

    /**
     * 获取单元格布尔值
     *
     * @param sheet       sheet
     * @param rowIndex    行index
     * @param columnIndex 列index
     * @return 布尔值
     */
    public static Boolean getCellBooleanValue(Sheet sheet, int rowIndex, int columnIndex) {
        if (sheet == null || rowIndex < 0 || columnIndex < 0) {
            return null;
        }
        return getCellBooleanValue(sheet.getRow(rowIndex), columnIndex);
    }

    /**
     * 获取单元格布尔值
     *
     * @param row         行：Row对象
     * @param columnIndex 列index
     * @return 布尔值
     */
    public static Boolean getCellBooleanValue(Row row, int columnIndex) {
        if (row == null || columnIndex < 0) {
            return null;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return null;
        }
        return cell.getBooleanCellValue();
    }



}
