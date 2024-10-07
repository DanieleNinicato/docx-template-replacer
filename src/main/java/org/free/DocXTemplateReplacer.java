package org.free;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.commons.compress.archivers.examples.Archiver;
import org.apache.commons.compress.archivers.examples.Expander;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static java.nio.charset.StandardCharsets.UTF_8;

public class DocXTemplateReplacer {
  public static void main(String[] args) throws Exception {
    var template = Paths.get("C:\\Development\\own\\spire-free\\src\\main\\resources\\template.docx");
    var outputFile = Paths.get("C:\\Development\\own\\spire-free\\src\\main\\resources\\saved.docx");

    mailMerge(template, outputFile);
  }

  public static File mailMerge(Path template, Path outputFile) throws Exception {
    var tempFolder = Paths.get("C:\\Development\\temp\\merge");

    extract(template, tempFolder);

    merge(tempFolder);

    archive(tempFolder, outputFile);

    convertToPdf(outputFile.toFile());

    return outputFile.toFile();
  }

  public static void extract(Path archive, Path destination) throws Exception {
    new Expander().expand(archive, destination);
  }

  public static void merge(Path tempFolder) throws IOException {
    Files.newDirectoryStream(tempFolder.resolve("word")).forEach(path -> {
      if (path.toFile().isFile()) {
        replace(path);
      }
    });
  }

  public static void replace(Path path) {
    try {
      var fileContent = FileUtils.readFileToString(path.toFile(), UTF_8);
      var replacedFileContent = fileContent.replaceAll("«", "").replace("»", "");
      FileUtils.write(path.toFile(), replacedFileContent, UTF_8);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  public static void archive(Path directory, Path destination) throws Exception {
    new Archiver().create("zip", destination, directory);
    FileUtils.deleteDirectory(directory.toFile());
  }

  public static void convertToPdf(File outputPdf) throws Exception {
    try {
      InputStream doc = new FileInputStream(outputPdf);
      XWPFDocument document = new XWPFDocument(doc);
      PdfOptions options = PdfOptions.create();
      OutputStream out = new FileOutputStream(new File("C:\\Development\\own\\spire-free\\src\\main\\resources\\saved.pdf"));
      PdfConverter.getInstance().convert(document, out, options);
    } catch (IOException ex) {
      System.out.println(ex.getMessage());
    }
  }
}
