package jp.ne.yonem.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/**
 * Excel処理クラス
 */
public class ExcelUtil {

    private static final Logger logger = LoggerFactory.getLogger(Logger.ROOT_LOGGER_NAME);

    /**
     * 処理ファイルと座標の情報を格納するレコード
     *
     * @param rows   自身の行
     * @param cols   自身の列
     * @param endRow 配下の最終行
     * @param endCol 配下の最終列
     * @param file   処理対象のファイル
     */
    private record FileTreeDTO(int rows, int cols, int endRow, int endCol, File file) {
    }

    /**
     * 一覧の開始行
     */
    private static final int ROW_START_INDEX = 1;

    /**
     * 一覧の開始列
     */
    private static final int COL_START_INDEX = 1;

    /**
     * 出力するExcelファイル名のフォーマット
     */
    private static final String EXCEL_BOOK_NAME = "tree_%ty-%tm-%td.xlsx";

    /**
     * シート名
     */
    private static final String EXCEL_SHEET_NAME = "tree";

    /**
     * 選択されたディレクトリを起点とし、配下の一覧を階層構造で描画する
     *
     * @param dir 起点ディレクトリ
     * @throws Exception 例外発生時
     */
    public static void convertDir2Tree(File dir) throws Exception {
        var t = Calendar.getInstance().getTime();
        var name = String.format(EXCEL_BOOK_NAME, t, t, t);

        try (var out = new FileOutputStream(new File(dir.getPath(), name)); var wb = new XSSFWorkbook()) {
            var ws = wb.createSheet(EXCEL_SHEET_NAME);
            var recordList = new ArrayList<FileTreeDTO>();
            var ci = convert(dir, ROW_START_INDEX, COL_START_INDEX, name, recordList);
            setStyle(wb, ws, ci[1], recordList);
            wb.write(out);

        } catch (Exception e) {
            logger.error(ExcelUtil.class.getName(), e);
            throw e;
        }
    }

    /**
     * 罫線の描画を担当するメソッド
     *
     * @param wb         ワークブック
     * @param ws         ワークシート
     * @param maxIndent  全体の最終列
     * @param recordList 描画するレコードのリスト
     */
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

            // 自身が最小列の場合
            if (maxIndent == rec.cols()) {
                col.setCellStyle(fileFullStyle);
                continue;
            }

            if (rec.file().isDirectory()) {

                // 自身がディレクトリの場合
                col.setCellStyle(dirLeftTopStyle);
                if (rec.rows() == rec.endRow()) col.setCellStyle(fileLeftStyle); // ディレクトリが最終行

                // 列方向への罫線を描画していく
                for (var i = rec.cols() + 1; i < maxIndent; i++) {
                    row.createCell(i).setCellStyle(dirTopStyle);
                    if (rec.rows() == rec.endRow()) row.createCell(i).setCellStyle(fileTopBottomStyle); // ディレクトリが最終行
                }
                row.createCell(maxIndent).setCellStyle(dirTopRightStyle);

                // ディレクトリが最終行
                if (rec.rows() == rec.endRow()) {
                    row.createCell(maxIndent).setCellStyle(fileRightStyle);
                    continue;
                }

                // 行方向への罫線を描画していく
                for (var i = rec.rows() + 1; i < rec.endRow(); i++) {
                    var currRow = ws.getRow(i);
                    if (Objects.isNull(currRow)) currRow = ws.createRow(i);
                    currRow.createCell(rec.cols()).setCellStyle(dirLeftStyle);
                }
                var currRow = ws.getRow(rec.endRow());
                if (Objects.isNull(currRow)) currRow = ws.createRow(rec.endRow());
                currRow.createCell(rec.cols()).setCellStyle(dirLeftBottomStyle);

            } else {

                // 自身がファイルの場合
                col.setCellStyle(fileLeftStyle);

                // 列方向への罫線を描画していく
                for (var i = rec.cols() + 1; i < maxIndent; i++) {
                    row.createCell(i).setCellStyle(fileTopBottomStyle);
                }
                row.createCell(maxIndent).setCellStyle(fileRightStyle);
            }
        }
        ws.autoSizeColumn(maxIndent, true); // 最終列の列幅を自動調整する
    }

    /**
     * 起点となるファイル or ディレクトリから再帰的に配下のファイル情報を取得していく<br>
     * ファイル名が出力するExcelファイル名と同名の場合は処理をスキップする<br>
     * 処理対象がファイルの場合はリストに自身を追加し処理を戻す<br>
     * 処理対象がディレクトリの場合は配下のファイル群を取得して再帰呼び出しする。最後にリストに自身を追加し処理を戻す
     *
     * @param file           処理対象ファイル or ディレクトリ
     * @param cnt            処理対象の行
     * @param indent         処理対象の階層
     * @param outputFileName 出力するExcelファイル名
     * @param recordList     描画するレコードのリスト
     * @return {最終行、最終列}の形式で処理結果を返却する
     */
    private static int[] convert(File file, int cnt, int indent, String outputFileName, List<FileTreeDTO> recordList) {
        var ci = new int[]{cnt, indent};

        // 出力するExcelファイル名と同名の場合
        if (Objects.isNull(file) || outputFileName.equalsIgnoreCase(file.getName())) {
            ci[0] = cnt - 1;
            return ci;
        }
        var lists = file.listFiles();
        Arrays.sort(Objects.requireNonNull(lists), Comparator.comparing(File::isDirectory).reversed().thenComparing(File::getName));

        // 処理対象がファイルの場合
        if (file.isFile()) {
            recordList.add(new FileTreeDTO(cnt, indent, cnt, indent, file));
            return ci;
        }

        // 配下のファイル群を再帰呼び出しする
        for (var tar : lists) {
            var tmp = convert(tar, ci[0] + 1, indent + 1, outputFileName, recordList);
            ci[0] = Math.max(ci[0], tmp[0]);
            ci[1] = Math.max(ci[1], tmp[1]);
        }

        // 最後にリストに自身を追加する
        recordList.add(new FileTreeDTO(cnt, indent, ci[0], ci[1], file));
        return ci;
    }
}
