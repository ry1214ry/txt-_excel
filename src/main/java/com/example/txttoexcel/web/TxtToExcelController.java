package com.example.txttoexcel.web;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

import com.example.txttoexcel.service.TxtToExcelService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class TxtToExcelController {

    private final TxtToExcelService txtToExcelService;
    private final Path exportDirectory;

    public TxtToExcelController(
            TxtToExcelService txtToExcelService,
            @Value("${app.export.directory:D:/Converter}") String exportDirectory) {
        this.txtToExcelService = txtToExcelService;
        this.exportDirectory = Path.of(exportDirectory).toAbsolutePath().normalize();
    }

    @GetMapping("/")
    public String home() {
        return "index";
    }

    @ModelAttribute("exportDirectoryDisplay")
    public String exportDirectoryDisplay() {
        return exportDirectory.toString();
    }

    @PostMapping("/convert")
    public String convert(@RequestParam("file") MultipartFile file, Model model) {
        if (file.isEmpty()) {
            model.addAttribute("errorMessage", "Please choose a TXT file before converting.");
            return "index";
        }

        if (!hasTxtExtension(file.getOriginalFilename())) {
            model.addAttribute("errorMessage", "Only .txt files are supported.");
            return "index";
        }

        try {
            byte[] excelBytes = txtToExcelService.convert(file.getInputStream(), file.getOriginalFilename());
            Path savedFilePath = saveConvertedFile(buildOutputFileName(file.getOriginalFilename()), excelBytes);
            String outputFileName = savedFilePath.getFileName().toString();
            model.addAttribute("successMessage", "Excel file is ready. Saved to " + savedFilePath + ".");
            model.addAttribute("outputFileName", outputFileName);
            model.addAttribute("savedFilePath", savedFilePath.toString());
            return "index";
        } catch (IllegalArgumentException exception) {
            model.addAttribute("errorMessage", exception.getMessage());
            return "index";
        } catch (IOException exception) {
            model.addAttribute("errorMessage", "The TXT file could not be converted or saved.");
            return "index";
        }
    }

    private boolean hasTxtExtension(String originalFilename) {
        return StringUtils.hasText(originalFilename)
                && originalFilename.toLowerCase().endsWith(".txt");
    }

    private String buildOutputFileName(String originalFilename) {
        String baseName = StringUtils.stripFilenameExtension(
                StringUtils.hasText(originalFilename) ? originalFilename : "converted");
        return baseName + ".xlsx";
    }

    private Path saveConvertedFile(String outputFileName, byte[] excelBytes) throws IOException {
        Files.createDirectories(exportDirectory);
        Path outputFilePath = resolveUniqueOutputPath(outputFileName);
        return Files.write(
                outputFilePath,
                excelBytes,
                StandardOpenOption.CREATE_NEW,
                StandardOpenOption.WRITE);
    }

    private Path resolveUniqueOutputPath(String outputFileName) {
        String baseName = StringUtils.stripFilenameExtension(outputFileName);
        String extension = StringUtils.getFilenameExtension(outputFileName);
        String fileExtension = StringUtils.hasText(extension) ? "." + extension : "";
        Path candidate = exportDirectory.resolve(outputFileName);
        int duplicateCounter = 1;

        while (Files.exists(candidate)) {
            candidate = exportDirectory.resolve(buildCopyFileName(baseName, fileExtension, duplicateCounter));
            duplicateCounter++;
        }

        return candidate;
    }

    private String buildCopyFileName(String baseName, String fileExtension, int duplicateCounter) {
        if (duplicateCounter == 1) {
            return baseName + " - Copy" + fileExtension;
        }

        return baseName + " - Copy (" + duplicateCounter + ")" + fileExtension;
    }
}
