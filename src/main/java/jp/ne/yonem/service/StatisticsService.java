package jp.ne.yonem.service;

import java.io.File;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import jp.ne.yonem.FileTreeFrame.ExtensionStat;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/** 統計情報の集計を行うサービス */
public class StatisticsService {

  private static final Logger logger = LoggerFactory.getLogger(StatisticsService.class);

  /** 唯一の実行メソッド */
  public List<ExtensionStat> execute(File root) {
    var statsList = new ArrayList<ExtensionStat>();

    try (var stream = Files.walk(root.toPath())) {
      var statsMap =
          stream.filter(Files::isRegularFile).collect(Collectors.groupingBy(this::getExtension));

      for (var entry : statsMap.entrySet()) {
        var totalLines = 0L;
        for (var p : entry.getValue()) {
          totalLines += countLines(p);
        }
        statsList.add(new ExtensionStat(entry.getKey(), entry.getValue().size(), totalLines));
      }
      statsList.sort(Comparator.comparingLong(ExtensionStat::count).reversed());

    } catch (Exception e) {
      logger.warn("統計情報の取得中にエラーが発生しました", e);
      return List.of();
    }
    return statsList;
  }

  /** 拡張子の取得 */
  private String getExtension(Path p) {
    var name = p.getFileName().toString();
    var dotIndex = name.lastIndexOf('.');
    return dotIndex == -1 ? "(no extension)" : name.substring(dotIndex).toLowerCase();
  }

  /** 行数のカウント */
  private long countLines(Path p) {
    if (!Files.isReadable(p) || isBinary(p)) return 0L;

    try (var lines = Files.lines(p)) {
      return lines.count();

    } catch (UncheckedIOException | IOException e) {
      logger.warn("行数取得に失敗したためスキップします: {}", p);
      return 0L;
    }
  }

  /** バイナリ判定 */
  private boolean isBinary(Path p) {
    try {
      var type = Files.probeContentType(p);

      if (Objects.nonNull(type)
          && !type.startsWith("text")
          && !type.contains("javascript")
          && !type.contains("xml")) {
        return true;
      }

      try (var is = Files.newInputStream(p)) {
        var bytes = new byte[1024];
        var read = is.read(bytes);

        for (var i = 0; i < read; i++) {
          if (bytes[i] == 0) return true;
        }
      }

    } catch (IOException e) {
      return true;
    }
    return false;
  }
}
