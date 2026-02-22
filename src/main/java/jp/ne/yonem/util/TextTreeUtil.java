package jp.ne.yonem.util;

import java.io.File;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Objects;

/** テキストツリー処理クラス */
public class TextTreeUtil {

  /**
   * 指定したディレクトリの構造をテキストツリー形式の文字列で取得する
   *
   * @param dir 起点ディレクトリ
   * @param isDirectoryOnly ディレクトリのみ抽出するか
   * @return ツリー形式の文字列
   */
  public static String convertDir2Text(File dir, boolean isDirectoryOnly) {

    if (Objects.isNull(dir) || !dir.exists()) {
      return "";
    }
    var sb = new StringBuilder();
    sb.append(dir.getName()).append("\n");
    buildTree(dir, 1, "", false, isDirectoryOnly, sb);
    return sb.toString();
  }

  private static void buildTree(
      File file,
      int indent,
      String hierarchy,
      boolean isEOL,
      boolean isDirectoryOnly,
      StringBuilder sb) {

    if (file.isFile()) return;
    var lists = file.listFiles();
    if (Objects.isNull(lists)) return;

    var filtered =
        Arrays.stream(lists)
            .filter(f -> !isDirectoryOnly || f.isDirectory())
            .sorted(Comparator.comparing(File::isDirectory).reversed().thenComparing(File::getName))
            .toList();

    var nextHierarchy = hierarchy;
    if (indent > 1) nextHierarchy += isEOL ? "   " : "│  ";

    for (int i = 0; i < filtered.size(); i++) {
      var next = filtered.get(i);
      var nextIsEOL = (i == filtered.size() - 1);

      sb.append(nextHierarchy);
      sb.append(nextIsEOL ? "└─ " : "├─ ");
      sb.append(next.getName()).append("\n");

      buildTree(next, indent + 1, nextHierarchy, nextIsEOL, isDirectoryOnly, sb);
    }
  }
}
