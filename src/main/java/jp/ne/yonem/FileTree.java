package jp.ne.yonem;

import jp.ne.yonem.util.ExcelUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.swing.*;

/**
 * メインクラス
 */
public class FileTree {

    private static Logger logger = LoggerFactory.getLogger(FileTree.class);

    /**
     * 正常終了時のタイトル
     */
    private static final String SUCCESS_TITLE = "正常終了";

    /**
     * 正常終了時のメッセージ
     */
    private static final String SUCCESS_MESSAGE = "ファイルを出力しました";

    /**
     * 異常終了時のタイトル
     */
    private static final String FAILURE_TITLE = "異常終了";

    /**
     * 異常終了時のメッセージ
     */
    private static final String FAILURE_MESSAGE = "ファイルの出力に失敗しました";

    /**
     * メインメソッド
     *
     * @param args プログラム引数
     */
    public static void main(String[] args) {

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            var chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(false);
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            var selected = chooser.showOpenDialog(null);
            if (selected != JFileChooser.APPROVE_OPTION) return;
            logger.debug(chooser.getSelectedFile().getAbsolutePath());
            ExcelUtil.convertDir2Tree(chooser.getSelectedFile());

            JOptionPane.showMessageDialog(
                    null,
                    SUCCESS_MESSAGE,
                    SUCCESS_TITLE,
                    JOptionPane.INFORMATION_MESSAGE
            );

        } catch (Exception e) {
            JOptionPane.showMessageDialog(
                    null,
                    FAILURE_MESSAGE,
                    FAILURE_TITLE,
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }
}
