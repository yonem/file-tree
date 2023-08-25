package jp.ne.yonem.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Objects;

public class ExcelUtil {

    private record FileTreeDTO(
            int rows,
            int cols,
            int endRow,
            int endCol,
            File file
    ) {
    }

    private static final int ROW_START_INDEX = 1;
    private static final int COL_START_INDEX = 1;
    private static final String EXCEL_BOOK_NAME = "tree_%ty-%tm-%td.xlsx";
    private static final String EXCEL_SHEET_NAME = "tree";

    public static void convertDir2Tree(File dir) throws Exception {
        var t = Calendar.getInstance().getTime();
        var name = String.format(EXCEL_BOOK_NAME, t, t, t);

        try (var out = new FileOutputStream(new File(dir.getPath(), name)); var wb = WorkbookFactory.create(true)) {

            var ws = wb.createSheet(EXCEL_SHEET_NAME);
            var recordList = new ArrayList<FileTreeDTO>();
            var ci = convert(dir, ROW_START_INDEX, COL_START_INDEX, name, recordList);
            setStyle(wb, ws, ci[1], recordList);
            wb.write(out);

        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    private static void setStyle(Workbook wb, Sheet ws, int maxIndent, List<FileTreeDTO> recordList) {
        var dirLeftTopStyle = wb.createCellStyle();
        var dirLeftStyle = wb.createCellStyle();
        var dirLeftBottomStyle = wb.createCellStyle();
        var dirTopStyle = wb.createCellStyle();
        var dirTopRightStyle = wb.createCellStyle();
        var fileLeftStyle = wb.createCellStyle();
        var fileTopBottomStyle = wb.createCellStyle();
        var fileRightStyle = wb.createCellStyle();
        var fileFullStyle = wb.createCellStyle();
        dirLeftTopStyle.setBorderLeft(BorderStyle.THIN);
        dirLeftTopStyle.setBorderTop(BorderStyle.THIN);
        dirLeftStyle.setBorderLeft(BorderStyle.THIN);
        dirTopStyle.setBorderTop(BorderStyle.THIN);
        dirTopRightStyle.setBorderTop(BorderStyle.THIN);
        dirTopRightStyle.setBorderRight(BorderStyle.THIN);
        dirLeftBottomStyle.setBorderLeft(BorderStyle.THIN);
        dirLeftBottomStyle.setBorderBottom(BorderStyle.THIN);
        fileLeftStyle.setBorderTop(BorderStyle.THIN);
        fileLeftStyle.setBorderLeft(BorderStyle.THIN);
        fileLeftStyle.setBorderBottom(BorderStyle.THIN);
        fileTopBottomStyle.setBorderTop(BorderStyle.THIN);
        fileTopBottomStyle.setBorderBottom(BorderStyle.THIN);
        fileRightStyle.setBorderTop(BorderStyle.THIN);
        fileRightStyle.setBorderRight(BorderStyle.THIN);
        fileRightStyle.setBorderBottom(BorderStyle.THIN);
        fileFullStyle.setBorderTop(BorderStyle.THIN);
        fileFullStyle.setBorderLeft(BorderStyle.THIN);
        fileFullStyle.setBorderRight(BorderStyle.THIN);
        fileFullStyle.setBorderBottom(BorderStyle.THIN);

        for (var rec : recordList) {
            var row = ws.createRow(rec.rows());
            var col = row.createCell(rec.cols());
            col.setCellValue(rec.file().getName());

            if (maxIndent == rec.cols()) {
                col.setCellStyle(fileFullStyle);
                continue;
            }

            if (rec.file().isDirectory()) {
                col.setCellStyle(dirLeftTopStyle);
                if (rec.rows() == rec.endRow()) col.setCellStyle(fileLeftStyle); // ディレクトリが最終行

                for (var i = rec.cols() + 1; i < maxIndent; i++) {
                    row.createCell(i).setCellStyle(dirTopStyle);
                    if (rec.rows() == rec.endRow()) row.createCell(i).setCellStyle(fileTopBottomStyle); // ディレクトリが最終行
                }
                row.createCell(maxIndent).setCellStyle(dirTopRightStyle);
                if (rec.rows() == rec.endRow()) row.createCell(maxIndent).setCellStyle(fileRightStyle); // ディレクトリが最終行

                for (var i = rec.rows() + 1; i < rec.endRow(); i++) {
                    var currRow = ws.getRow(i);
                    if (Objects.isNull(currRow)) currRow = ws.createRow(i);
                    currRow.createCell(rec.cols()).setCellStyle(dirLeftStyle);
                }
                if (rec.rows() == rec.endRow()) continue;
                var currRow = ws.getRow(rec.endRow());
                if (Objects.isNull(currRow)) currRow = ws.createRow(rec.endRow());
                currRow.createCell(rec.cols()).setCellStyle(dirLeftBottomStyle);

            } else {
                col.setCellStyle(fileLeftStyle);

                for (var i = rec.cols() + 1; i < maxIndent; i++) {
                    row.createCell(i).setCellStyle(fileTopBottomStyle);
                }
                row.createCell(maxIndent).setCellStyle(fileRightStyle);
            }
        }
        ws.autoSizeColumn(maxIndent, true);
    }

    private static int[] convert(File file, int cnt, int indent, String name, List<FileTreeDTO> recordList) {
        var ci = new int[]{cnt, indent};

        if (Objects.isNull(file) || name.equalsIgnoreCase(file.getName())) {
            ci[0] = cnt - 1;
            return ci;
        }
        var lists = file.listFiles();

        if (file.isFile() || Objects.isNull(lists)) {
            recordList.add(new FileTreeDTO(cnt, indent, cnt, indent, file));
            return ci;
        }

        for (var tar : lists) {
            var tmp = convert(tar, ci[0] + 1, indent + 1, name, recordList);
            ci[0] = Math.max(ci[0], tmp[0]);
            ci[1] = Math.max(ci[1], tmp[1]);
        }
        recordList.add(new FileTreeDTO(cnt, indent, ci[0], ci[1], file));
        return ci;
    }
}
