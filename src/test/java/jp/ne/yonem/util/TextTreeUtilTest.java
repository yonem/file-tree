package jp.ne.yonem.util;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

@DisplayName("TextTreeUtilの網羅テスト")
class TextTreeUtilTest {

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
    @DisplayName("全てのツリー描画ロジックの網羅（深い階層・垂直線の維持）")
    void testFullTreeStructure() throws IOException {
      var sub = new File(rootDir, "a_sub");
      sub.mkdir();
      var deep = new File(sub, "deep");
      deep.mkdir();
      new File(deep, "file.txt").createNewFile();
      new File(rootDir, "b_file.txt").createNewFile();

      var result = TextTreeUtil.convertDir2Text(rootDir, false);

      assertNotNull(result);
      assertTrue(result.contains("test_root"));
      assertTrue(result.contains("├─ a_sub"));
      assertTrue(result.contains("│  └─ deep"));
      assertTrue(result.contains("│     └─ file.txt"));
      assertTrue(result.contains("└─ b_file.txt"));
    }

    @Test
    @DisplayName("ディレクトリのみ出力モードの網羅")
    void testDirectoryOnlyMode() throws IOException {
      new File(rootDir, "sub_dir").mkdir();
      new File(rootDir, "ignore.txt").createNewFile();

      var result = TextTreeUtil.convertDir2Text(rootDir, true);

      assertTrue(result.contains("sub_dir"));
      assertFalse(result.contains("ignore.txt"));
    }

    @Test
    @DisplayName("空ディレクトリおよび特殊な並び順の網羅")
    void testEmptyAndSortOrder() {
      new File(rootDir, "z_dir").mkdir();
      new File(rootDir, "a_dir").mkdir();

      var result = TextTreeUtil.convertDir2Text(rootDir, false);

      assertTrue(result.indexOf("a_dir") < result.indexOf("z_dir"));
    }
  }

  @Nested
  @DisplayName("異常系・境界値のテスト")
  class NegativeTests {

    @Test
    @DisplayName("存在しないディレクトリやNullに対するガード")
    void testInvalidInput() {
      assertEquals("", TextTreeUtil.convertDir2Text(null, false));

      var missing = new File(tempDir.toFile(), "missing");
      assertEquals("", TextTreeUtil.convertDir2Text(missing, false));
    }

    @Test
    @DisplayName("ファイルに対して実行した場合の挙動（ディレクトリとして扱えないケース）")
    void testFileAsRoot() throws IOException {
      var file = new File(rootDir, "sole_file.txt");
      file.createNewFile();

      var result = TextTreeUtil.convertDir2Text(file, false);

      assertTrue(result.contains("sole_file.txt"));
      assertFalse(result.contains("├─"));
    }

    @Test
    @DisplayName("アクセス権限がない等の理由でlistFilesがNullになるケースの網羅")
    void testListFilesNull() throws IOException {
      var restricted = new File(rootDir, "restricted");
      restricted.createNewFile();

      assertDoesNotThrow(() -> TextTreeUtil.convertDir2Text(restricted, false));
    }
  }
}
