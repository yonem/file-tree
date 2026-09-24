package jp.ne.yonem;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Calendar;
import jp.ne.yonem.service.StatisticsService;
import jp.ne.yonem.util.ExcelUtil;
import jp.ne.yonem.util.TextTreeUtil;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("FileTreeの主要処理連携テスト")
class FileTreeWorkflowIntegrationTest {

  @TempDir Path tempDir;

  @Test
  @DisplayName("同じ入力ディレクトリからツリー、統計、Excel出力を取得できること")
  void createsConsistentOutputsFromOneDirectory() throws Exception {
    var root = Files.createDirectory(tempDir.resolve("project"));
    var source = Files.createDirectory(root.resolve("src"));
    Files.writeString(source.resolve("Main.java"), "class Main {}\n");
    Files.writeString(root.resolve("README.md"), "# project\n");

    var tree = TextTreeUtil.convertDir2Text(root.toFile(), false);
    var stats = new StatisticsService().execute(root.toFile());
    ExcelUtil.convertDir2Tree(root.toFile(), false);

    assertTrue(tree.contains("src"));
    assertTrue(tree.contains("Main.java"));
    assertTrue(tree.contains("README.md"));
    assertEquals(2, stats.stream().mapToLong(FileTreeFrame.ExtensionStat::count).sum());

    var time = Calendar.getInstance().getTime();
    var output = root.resolve(String.format("tree_%ty-%tm-%td.xlsx", time, time, time));
    assertTrue(Files.exists(output));
    try (var input = Files.newInputStream(output); var workbook = WorkbookFactory.create(input)) {
      var sheet = workbook.getSheet("tree");
      assertNotNull(sheet);
      assertEquals("project", sheet.getRow(1).getCell(1).getStringCellValue());
      assertEquals("src", sheet.getRow(2).getCell(2).getStringCellValue());
      assertEquals("Main.java", sheet.getRow(3).getCell(3).getStringCellValue());
      assertEquals("README.md", sheet.getRow(4).getCell(2).getStringCellValue());
    }
  }
}
