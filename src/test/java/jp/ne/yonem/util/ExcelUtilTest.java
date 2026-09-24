package jp.ne.yonem.util;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("ExcelUtilの網羅テスト")
class ExcelUtilTest {

  @TempDir Path tempDir;
  private File rootDir;

  @BeforeEach
  void setUp() throws IOException {
    var root = tempDir.resolve("test_root");
    rootDir = root.toFile();
    if (!rootDir.exists() && !rootDir.mkdir()) {
      throw new IOException("Test directory creation failed");
    }
  }

  @Nested
  @DisplayName("正常系のテスト")
  class PositiveTests {

    @Test
    @DisplayName("全ての描画ロジックとディレクトリ構造の網羅")
    void testFullCoverage() throws Exception {
      var sub = new File(rootDir, "sub");
      sub.mkdir();
      var deep = new File(sub, "deep");
      deep.mkdir();
      new File(deep, "file.txt").createNewFile();
      new File(rootDir, "z_last.txt").createNewFile();

      var t = Calendar.getInstance().getTime();
      var excelName = String.format("tree_%ty-%tm-%td.xlsx", t, t, t);
      new File(rootDir, excelName).createNewFile();

      ExcelUtil.convertDir2Tree(rootDir, false);
      ExcelUtil.convertDir2Tree(rootDir, true);
    }

    @Test
    @DisplayName("境界条件および特定ルートの網羅")
    void testCoverRemainingPaths() throws Exception {
      var emptyDir = new File(rootDir, "z_empty_dir");
      emptyDir.mkdir();

      var subDir = new File(rootDir, "a_sub");
      subDir.mkdir();
      var deepDir = new File(subDir, "deep");
      deepDir.mkdir();
      new File(deepDir, "deep_file.txt").createNewFile();

      new File(rootDir, "b_root_file.txt").createNewFile();

      ExcelUtil.convertDir2Tree(rootDir, false);
      ExcelUtil.convertDir2Tree(rootDir, true);

      var dummyFile = new File(rootDir, "dummy.txt");
      dummyFile.createNewFile();

      assertThrows(Exception.class, () -> ExcelUtil.convertDir2Tree(dummyFile, false));
    }

    @Test
    @DisplayName("生成されたExcelにシート名、階層、ファイル名が記録されること")
    void testWorkbookContentsAndDirectoryOnlyMode() throws Exception {
      var documents = Files.createDirectory(rootDir.toPath().resolve("documents"));
      Files.writeString(documents.resolve("guide.txt"), "guide");
      Files.writeString(rootDir.toPath().resolve("root.txt"), "root");

      ExcelUtil.convertDir2Tree(rootDir, false);

      var output = outputFile();
      assertTrue(Files.exists(output));
      try (var input = Files.newInputStream(output); var workbook = WorkbookFactory.create(input)) {
        assertEquals(1, workbook.getNumberOfSheets());
        var sheet = workbook.getSheet("tree");
        assertNotNull(sheet);
        assertEquals("test_root", sheet.getRow(1).getCell(1).getStringCellValue());
        assertEquals("documents", sheet.getRow(2).getCell(2).getStringCellValue());
        assertEquals("guide.txt", sheet.getRow(3).getCell(3).getStringCellValue());
        assertEquals("root.txt", sheet.getRow(4).getCell(2).getStringCellValue());
      }

      ExcelUtil.convertDir2Tree(rootDir, true);
      try (var input = Files.newInputStream(output); var workbook = WorkbookFactory.create(input)) {
        var values = stringCellValues(workbook.getSheet("tree"));
        assertTrue(values.contains("documents"));
        assertFalse(values.contains("guide.txt"));
        assertFalse(values.contains("root.txt"));
      }
    }
  }

  @Nested
  @DisplayName("異常系のテスト")
  class NegativeTests {

    @Test
    @DisplayName("不正な引数および実行時エラーの網羅")
    void testValidationAndErrors() throws IOException {
      assertThrows(IllegalArgumentException.class, () -> ExcelUtil.convertDir2Tree(null, false));

      var missing = new File(tempDir.toFile(), "not_exists");
      assertThrows(IllegalArgumentException.class, () -> ExcelUtil.convertDir2Tree(missing, false));

      var f = new File(rootDir, "f.txt");
      f.createNewFile();
      assertThrows(Exception.class, () -> ExcelUtil.convertDir2Tree(f, false));

      var t = Calendar.getInstance().getTime();
      var excelName = String.format("tree_%ty-%tm-%td.xlsx", t, t, t);
      var conflictDir = new File(rootDir, excelName);
      conflictDir.mkdir();
      assertThrows(Exception.class, () -> ExcelUtil.convertDir2Tree(rootDir, false));
    }

    @Test
    @DisplayName("特殊なファイル状態における処理継続性の確認")
    void testListFilesNull() throws IOException {
      var fileAsDir = new File(rootDir, "restricted");
      fileAsDir.createNewFile();

      try {
        ExcelUtil.convertDir2Tree(fileAsDir, false);
      } catch (Exception ignored) {
      }
    }
  }

  private Path outputFile() {
    var time = Calendar.getInstance().getTime();
    var fileName = String.format("tree_%ty-%tm-%td.xlsx", time, time, time);
    return rootDir.toPath().resolve(fileName);
  }

  private List<String> stringCellValues(Sheet sheet) {
    var values = new ArrayList<String>();
    for (var row : sheet) {
      for (var cell : row) {
        if (cell.getCellType() == CellType.STRING) values.add(cell.getStringCellValue());
      }
    }
    return values;
  }
}
