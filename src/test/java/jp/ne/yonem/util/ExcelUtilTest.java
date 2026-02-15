package jp.ne.yonem.util;

import static org.junit.jupiter.api.Assertions.*;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Calendar;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ExcelUtilTest {

  @TempDir Path tempDir;

  /** 現在の日付に基づく期待されるファイル名を取得する */
  private String getExpectedFileName() {
    var t = Calendar.getInstance().getTime();
    return String.format("tree_%ty-%tm-%td.xlsx", t, t, t);
  }

  @Nested
  @DisplayName("正常系のテスト")
  class SuccessTests {

    @Test
    @DisplayName("複雑なディレクトリ構造からExcelが生成されること")
    void testConvertDir2Tree_Success() throws Exception {
      // 1. テスト用のディレクトリ構造を作成
      var subDir1 = Files.createDirectory(tempDir.resolve("subdir1"));
      Files.createFile(subDir1.resolve("file1.txt"));
      Files.createDirectory(tempDir.resolve("subdir2"));
      Files.createFile(tempDir.resolve("root_file.txt"));

      // 2. 実行
      var targetDir = tempDir.toFile();
      ExcelUtil.convertDir2Tree(targetDir);

      // 3. 検証
      var resultFile = new File(targetDir, getExpectedFileName());
      assertTrue(resultFile.exists(), "Excelファイルが生成されていること");
      assertTrue(resultFile.length() > 0, "ファイル内容が空ではないこと");
    }

    @Test
    @DisplayName("空のディレクトリでもExcelが生成されること")
    void testConvertDir2Tree_EmptyDir() {
      var emptyDir = tempDir.toFile();

      assertDoesNotThrow(
          () -> {
            ExcelUtil.convertDir2Tree(emptyDir);
          });

      var resultFile = new File(emptyDir, getExpectedFileName());
      assertTrue(resultFile.exists());
    }
  }

  @Nested
  @DisplayName("異常系・境界値のテスト")
  class ExceptionAndEdgeTests {

    @Test
    @DisplayName("存在しないディレクトリを指定した場合に例外が発生すること")
    void testConvertDir2Tree_NotFound() {
      // ghostディレクトリは作成しない
      var nonExistentDir = new File(tempDir.toFile(), "ghost");

      // 現状の実装では FileOutputStream のコンストラクタで FileNotFoundException が発生する
      assertThrows(
          Exception.class,
          () -> {
            ExcelUtil.convertDir2Tree(nonExistentDir);
          });
    }

    @Test
    @DisplayName("出力ファイルと同名のファイルは処理対象から除外されること")
    void testConvertDir2Tree_SkipSelf() throws Exception {
      // 出力予定のファイル名と同名の空ファイルを作成しておく
      var fileName = getExpectedFileName();
      var selfFile = Files.createFile(tempDir.resolve(fileName));

      // 実行
      assertDoesNotThrow(
          () -> {
            ExcelUtil.convertDir2Tree(tempDir.toFile());
          });

      // 自身を処理しようとして無限ループやエラーにならないことを確認
      assertTrue(Files.exists(selfFile));
    }
  }
}
