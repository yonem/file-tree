package jp.ne.yonem.util;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Calendar;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

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
  class PositiveTests {

    @Test
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
  }

  @Nested
  class NegativeTests {

    @Test
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
    void testListFilesNull() throws IOException {
      var fileAsDir = new File(rootDir, "restricted");
      fileAsDir.createNewFile();

      try {
        ExcelUtil.convertDir2Tree(fileAsDir, false);
      } catch (Exception ignored) {
      }
    }
  }
}
