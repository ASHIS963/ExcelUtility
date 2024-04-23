package utils;

import java.io.IOException;
import java.nio.file.*;

public class ExcelBackupManager {

    public static void backupExcelFile(String sourcePath, String backupPath) throws IOException {
        Path source = Paths.get(sourcePath);
        Path backup = Paths.get(backupPath);

        // Copy the original file to the backup location
        Files.copy(source, backup, StandardCopyOption.REPLACE_EXISTING);
    }

    public static void restoreExcelFile(String backupPath, String targetPath) throws IOException {
        Path source = Paths.get(backupPath);
        Path target = Paths.get(targetPath);

        // Copy the backup file back to the original location to restore it
        Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
    }
}
