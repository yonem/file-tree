package jp.ne.yonem.service;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.MockedStatic;
import org.mockito.junit.jupiter.MockitoExtension;

/** StatisticsServiceのテスト */
@ExtendWith(MockitoExtension.class)
class StatisticsServiceTest {

  @InjectMocks private StatisticsService sut;

  @Nested
  @DisplayName("正常系テストケース")
  class PositiveTests {

    @Test
    @DisplayName("正常系: テキストファイルとバイナリファイルが混在する場合に正しく集計されること")
    void test01() {
      var root = new File("root");
      var path1 = mock(Path.class);
      var path2 = mock(Path.class);
      var fileName1 = mock(Path.class);
      var fileName2 = mock(Path.class);

      when(path1.getFileName()).thenReturn(fileName1);
      when(path2.getFileName()).thenReturn(fileName2);
      when(fileName1.toString()).thenReturn("test.java");
      when(fileName2.toString()).thenReturn("image.png");

      try (MockedStatic<Files> files = mockStatic(Files.class)) {
        files.when(() -> Files.walk(any())).thenReturn(Stream.of(path1, path2));
        files.when(() -> Files.isRegularFile(any())).thenReturn(true);
        files.when(() -> Files.isReadable(any())).thenReturn(true);
        files.when(() -> Files.probeContentType(path1)).thenReturn("text/plain");
        files.when(() -> Files.probeContentType(path2)).thenReturn("image/png");

        files
            .when(() -> Files.newInputStream(path1))
            .thenReturn(new ByteArrayInputStream("abc".getBytes()));
        files
            .when(() -> Files.newInputStream(path2))
            .thenReturn(new ByteArrayInputStream(new byte[] {0, 1, 2}));

        files.when(() -> Files.lines(path1)).thenReturn(Stream.of("line1", "line2", "line3"));

        var result = sut.execute(root);

        assertEquals(2, result.size());

        var javaStat =
            result.stream().filter(s -> ".java".equals(s.extension())).findFirst().orElseThrow();
        assertEquals(1, javaStat.count());
        assertEquals(3, javaStat.totalLines());

        var pngStat =
            result.stream().filter(s -> ".png".equals(s.extension())).findFirst().orElseThrow();
        assertEquals(1, pngStat.count());
        assertEquals(0, pngStat.totalLines());

        files.verify(() -> Files.walk(any()));
        files.verify(() -> Files.lines(path1));
        files.verify(() -> Files.lines(path2), never());
      }
    }
  }

  @Nested
  @DisplayName("異常系テストケース")
  class NegativeTests {

    @Test
    @DisplayName("異常系: 読み取り中にUncheckedIOExceptionが発生しても他のファイル処理が継続されること")
    void test01() {
      var root = new File("root");
      var path = mock(Path.class);
      var fileName = mock(Path.class);

      when(path.getFileName()).thenReturn(fileName);
      when(fileName.toString()).thenReturn("error.txt");

      try (MockedStatic<Files> files = mockStatic(Files.class)) {
        files.when(() -> Files.walk(any())).thenReturn(Stream.of(path));
        files.when(() -> Files.isRegularFile(any())).thenReturn(true);
        files.when(() -> Files.isReadable(any())).thenReturn(true);
        files.when(() -> Files.probeContentType(any())).thenReturn("text/plain");
        files
            .when(() -> Files.newInputStream(any()))
            .thenReturn(new ByteArrayInputStream("data".getBytes()));

        files
            .when(() -> Files.lines(any()))
            .thenThrow(new UncheckedIOException(new IOException("Read error")));

        var result = sut.execute(root);

        assertFalse(result.isEmpty());
        assertEquals(1, result.getFirst().count());
        assertEquals(0, result.getFirst().totalLines());
      }
    }

    @Test
    @DisplayName("異常系: Files.walk自体が失敗した場合に空リストが返却されること")
    void test02() {
      var root = new File("invalid");

      try (MockedStatic<Files> files = mockStatic(Files.class)) {
        files.when(() -> Files.walk(any())).thenThrow(new IOException("Access denied"));

        var result = sut.execute(root);

        assertTrue(result.isEmpty());
      }
    }

    @Test
    @DisplayName("異常系: isBinary判定中にIOExceptionが発生した場合に行数0として集計されること")
    void test03() {
      var root = new File("root");
      var path = mock(Path.class);
      var fileName = mock(Path.class);

      when(path.getFileName()).thenReturn(fileName);
      when(fileName.toString()).thenReturn("locked.txt");

      try (MockedStatic<Files> files = mockStatic(Files.class)) {
        files.when(() -> Files.walk(any())).thenReturn(Stream.of(path));
        files.when(() -> Files.isRegularFile(any())).thenReturn(true);
        files.when(() -> Files.isReadable(any())).thenReturn(true);
        files.when(() -> Files.probeContentType(any())).thenReturn("text/plain");

        files.when(() -> Files.newInputStream(path)).thenThrow(new IOException("Locked"));

        var result = sut.execute(root);

        assertNotNull(result);
        assertEquals(1, result.size());
        assertEquals(".txt", result.getFirst().extension());
        assertEquals(0, result.getFirst().totalLines());

        files.verify(() -> Files.newInputStream(path), times(1));
        files.verify(() -> Files.lines(any()), never());
      }
    }
  }
}
