package org.free;

import org.apache.commons.compress.archivers.examples.Archiver;
import org.apache.commons.compress.archivers.examples.Expander;
import org.apache.commons.io.FileUtils;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.convert.out.fo.renderers.FORendererApacheFOP;
import org.docx4j.fonts.BestMatchingMapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;

import static java.nio.charset.StandardCharsets.UTF_8;
import static org.docx4j.XmlUtils.marshaltoString;
import static org.docx4j.jaxb.Context.getFopConfigContext;

public class DocXUtil {

  private static final Logger log = LoggerFactory.getLogger(DocXUtil.class);

  public static void main(String[] args) throws Exception {
    var template = Paths.get("C:\\Development\\own\\spire-free\\src\\main\\resources\\template.docx");
    var outputFile = Paths.get("C:\\Development\\own\\spire-free\\src\\main\\resources\\saved.docx");

    mailMerge(template, outputFile);
    log.info("Done");
  }

  public static File mailMerge(Path template, Path outputFile) throws Exception {
    var tempFolder = Paths.get("C:\\Development\\own\\spire-free\\src\\main\\resources\\merge");

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
    PhysicalFonts.addPhysicalFont(Objects.requireNonNull(DocXUtil.class.getClassLoader().getResource("fonts/FlandersArtSans_Bold/FlandersArtSans-Bold.otf")).toURI());
    PhysicalFonts.addPhysicalFont(Objects.requireNonNull(DocXUtil.class.getClassLoader().getResource("fonts/FlandersArtSans_Regular/FlandersArtSans-Regular.otf")).toURI());
    PhysicalFonts.addPhysicalFont(Objects.requireNonNull(DocXUtil.class.getClassLoader().getResource("fonts/FlandersArtSerifMedium/FlandersArtSerif-Medium.ttf")).toURI());
    PhysicalFonts.addPhysicalFont(Objects.requireNonNull(DocXUtil.class.getClassLoader().getResource("fonts/FlandersArtSans_Italic/FlandersArtSans-Italic.otf")).toURI());

    // Load the .docx file
    var wordMLPackage = WordprocessingMLPackage.load(outputPdf);

    var fontMapper = new BestMatchingMapper();
    wordMLPackage.setFontMapper(fontMapper);

    var foSettings = new FOSettings(wordMLPackage);

    if (log.isDebugEnabled()) {
      log.debug(marshaltoString(foSettings.getFopConfig(), getFopConfigContext()));
    }

    var fopFactoryBuilder = FORendererApacheFOP.getFopFactoryBuilder(foSettings);
    var fopFactory = fopFactoryBuilder.build();
    FORendererApacheFOP.getFOUserAgent(foSettings, fopFactory);

    var os = new FileOutputStream(new File("C:\\Development\\own\\spire-free\\src\\main\\resources\\saved.pdf"));

    Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);

    os.close();

    log.info("PDF file generated successfully.");
  }
}
