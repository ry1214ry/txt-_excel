package com.example.txttoexcel.web;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.charset.StandardCharsets;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.webmvc.test.autoconfigure.AutoConfigureMockMvc;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;

import static org.assertj.core.api.Assertions.assertThat;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.content;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@SpringBootTest(properties = "app.export.directory=target/test-exports")
@AutoConfigureMockMvc
class TxtToExcelControllerTests {

    private static final Path EXPORT_DIRECTORY = Path.of("target", "test-exports").toAbsolutePath().normalize();

    @Autowired
    private MockMvc mockMvc;

    @Test
    void convertsUploadedTxtFile() throws Exception {
        Files.createDirectories(EXPORT_DIRECTORY);
        Path exportedFile = EXPORT_DIRECTORY.resolve("employees.xlsx");
        Files.deleteIfExists(exportedFile);
        Files.deleteIfExists(EXPORT_DIRECTORY.resolve("employees - Copy.xlsx"));
        Files.deleteIfExists(EXPORT_DIRECTORY.resolve("employees - Copy (2).xlsx"));

        MockMultipartFile file = new MockMultipartFile(
                "file",
                "employees.txt",
                "text/plain",
                """
                        Name,City,Salary
                        Roeun,Phnom Penh,1200
                        Dary,Bangkok,1400
                        """.getBytes(StandardCharsets.UTF_8));

        MvcResult result = mockMvc.perform(multipart("/convert").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Excel file is ready.")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("employees.xlsx")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Saved automatically to:")))
                .andExpect(content().string(org.hamcrest.Matchers.not(org.hamcrest.Matchers.containsString("Download Excel"))))
                .andReturn();

        assertThat(Files.exists(exportedFile)).isTrue();
        assertThat(Files.size(exportedFile)).isGreaterThan(0L);

        try (XSSFWorkbook workbook = new XSSFWorkbook(Files.newInputStream(exportedFile))) {
            assertThat(workbook.getSheet("TXT Data")).isNotNull();
            assertThat(workbook.getSheet("Summary")).isNotNull();
            assertThat(workbook.getSheet("TXT Data").getRow(0).getCell(0).getStringCellValue()).isEqualTo("Name");
        }
    }

    @Test
    void createsCopyFileNameWhenExcelFileAlreadyExists() throws Exception {
        Files.createDirectories(EXPORT_DIRECTORY);
        Path originalFile = EXPORT_DIRECTORY.resolve("employees.xlsx");
        Path copiedFile = EXPORT_DIRECTORY.resolve("employees - Copy.xlsx");
        Files.writeString(originalFile, "existing workbook", StandardCharsets.UTF_8);
        Files.deleteIfExists(copiedFile);

        MockMultipartFile file = new MockMultipartFile(
                "file",
                "employees.txt",
                "text/plain",
                """
                        Name,City,Salary
                        Roeun,Phnom Penh,1200
                        Dary,Bangkok,1400
                        """.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/convert").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("employees - Copy.xlsx")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString(copiedFile.toString())));

        assertThat(Files.readString(originalFile, StandardCharsets.UTF_8)).isEqualTo("existing workbook");
        assertThat(Files.exists(copiedFile)).isTrue();
        assertThat(Files.size(copiedFile)).isGreaterThan(0L);
    }

    @Test
    void createsIndexedCopyFileNameWhenCopyAlreadyExists() throws Exception {
        Files.createDirectories(EXPORT_DIRECTORY);
        Path originalFile = EXPORT_DIRECTORY.resolve("employees.xlsx");
        Path firstCopyFile = EXPORT_DIRECTORY.resolve("employees - Copy.xlsx");
        Path secondCopyFile = EXPORT_DIRECTORY.resolve("employees - Copy (2).xlsx");
        Files.writeString(originalFile, "existing workbook", StandardCharsets.UTF_8);
        Files.writeString(firstCopyFile, "existing copy", StandardCharsets.UTF_8);
        Files.deleteIfExists(secondCopyFile);

        MockMultipartFile file = new MockMultipartFile(
                "file",
                "employees.txt",
                "text/plain",
                """
                        Name,City,Salary
                        Roeun,Phnom Penh,1200
                        Dary,Bangkok,1400
                        """.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/convert").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("employees - Copy (2).xlsx")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString(secondCopyFile.toString())));

        assertThat(Files.readString(originalFile, StandardCharsets.UTF_8)).isEqualTo("existing workbook");
        assertThat(Files.readString(firstCopyFile, StandardCharsets.UTF_8)).isEqualTo("existing copy");
        assertThat(Files.exists(secondCopyFile)).isTrue();
        assertThat(Files.size(secondCopyFile)).isGreaterThan(0L);
    }

    @Test
    void rejectsFilesWithoutTxtExtension() throws Exception {
        MockMultipartFile file = new MockMultipartFile(
                "file",
                "employees.csv",
                "text/plain",
                "name,score".getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/convert").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Only .txt files are supported.")));
    }

    @Test
    void homePageDoesNotShowDownloadButton() throws Exception {
        mockMvc.perform(get("/"))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString(
                        "The Excel file will be created automatically in " + EXPORT_DIRECTORY + ".")))
                .andExpect(content().string(org.hamcrest.Matchers.not(org.hamcrest.Matchers.containsString("Download Excel"))));
    }
}
