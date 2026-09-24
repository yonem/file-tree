package jp.ne.yonem.service;

import static org.junit.jupiter.api.Assertions.*;

import java.nio.file.Files;
import java.nio.file.Path;
import jp.ne.yonem.util.TextTreeUtil;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("利用シナリオのテスト")
class StatisticsServiceScenarioTest {

  @TempDir Path tempDir;

  @Test
  @DisplayName("日本語名のディレクトリとファイルをツリー・統計で扱えること")
  void handlesJapaneseNames() throws Exception {
    var root = Files.createDirectory(tempDir.resolve("利用者データ"));
    var documents = Files.createDirectory(root.resolve("資料"));
    Files.writeString(documents.resolve("概要.txt"), "1行目\n2行目");

    var tree = TextTreeUtil.convertDir2Text(root.toFile(), false);
    var stats = new StatisticsService().execute(root.toFile());

    assertTrue(tree.contains("利用者データ"));
    assertTrue(tree.contains("資料"));
    assertTrue(tree.contains("概要.txt"));
    var text = stats.stream().filter(item -> item.extension().equals(".txt")).findFirst().orElseThrow();
    assertEquals(1, text.count());
    assertEquals(2, text.totalLines());
  }

  @Test
  @DisplayName("空ディレクトリと多数ファイルを含む場合も集計できること")
  void handlesEmptyDirectoryAndManyFiles() throws Exception {
    var root = Files.createDirectory(tempDir.resolve("bulk"));
    Files.createDirectory(root.resolve("empty"));
    for (var index = 0; index < 100; index++) {
      Files.writeString(root.resolve("file-%03d.txt".formatted(index)), "line");
    }

    var tree = TextTreeUtil.convertDir2Text(root.toFile(), false);
    var stats = new StatisticsService().execute(root.toFile());

    assertTrue(tree.contains("empty"));
    var text = stats.stream().filter(item -> item.extension().equals(".txt")).findFirst().orElseThrow();
    assertEquals(100, text.count());
    assertEquals(100, text.totalLines());
  }
}
