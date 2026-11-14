/*
 * Copyright © 2012 ecuacion.jp (info@ecuacion.jp)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package jp.ecuacion.util.pdf.word.document.internal;

import java.io.BufferedInputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import jp.ecuacion.util.pdf.common.exception.FileFormatNotSupportedException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordIncorrectException;
import jp.ecuacion.util.pdf.common.exception.FilePasswordRequiredException;
import org.apache.poi.EmptyFileException;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.jspecify.annotations.Nullable;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

/** Opens a Word file (optionally password-protected) and reads its parts. */
final class DocxPackage implements Closeable {

  /** The formats which can be converted, shown to the user when a file is not one of them. */
  static final String SUPPORTED_FORMATS = "docx, docm, dotx, dotm";

  private final OPCPackage pkg;
  private final PackagePart mainPart;

  private DocxPackage(OPCPackage pkg) {
    this.pkg = pkg;
    PackagePart main = null;
    try {
      for (PackageRelationship rel : pkg.getRelationships()) {
        if (rel.getRelationshipType().endsWith("/officeDocument")) {
          main = pkg.getPart(rel);
        }
      }
    } catch (RuntimeException e) {
      main = null;
    }
    // Other Office Open XML files (e.g. .xlsx) have a main part of another content type.
    if (main == null || (!main.getContentType().contains("wordprocessingml")
        && !main.getContentType().startsWith("application/vnd.ms-word."))) {
      pkg.revert();
      throw new FileFormatNotSupportedException(SUPPORTED_FORMATS);
    }
    this.mainPart = main;
  }

  /**
   * Opens the Word file.
   *
   * @throws FileFormatNotSupportedException if the file is not a Word file of the Office Open
   *     XML format (e.g. a legacy .doc file, a file of another application, a broken file)
   * @throws FilePasswordRequiredException if the file is password-protected and no password
   *     is specified
   * @throws FilePasswordIncorrectException if the password is incorrect
   * @throws IOException if the file cannot be read
   */
  static DocxPackage open(Path path, @Nullable String password) throws IOException {
    FileMagic magic;
    try (InputStream is =
        FileMagic.prepareToCheckMagic(new BufferedInputStream(Files.newInputStream(path)))) {
      magic = FileMagic.valueOf(is);
    } catch (EmptyFileException e) {
      throw new FileFormatNotSupportedException(SUPPORTED_FORMATS);
    }
    try {
      if (magic == FileMagic.OOXML) {
        return new DocxPackage(OPCPackage.open(path.toFile(), PackageAccess.READ));
      }
      if (magic != FileMagic.OLE2) {
        throw new FileFormatNotSupportedException(SUPPORTED_FORMATS);
      }
      // OLE2: either a password-protected Office Open XML file, or a legacy .doc file.
      try (POIFSFileSystem fs = new POIFSFileSystem(path.toFile(), true)) {
        if (!fs.getRoot().hasEntry(Decryptor.DEFAULT_POIFS_ENTRY)) {
          throw new FileFormatNotSupportedException(SUPPORTED_FORMATS);
        }
        Decryptor decryptor = Decryptor.getInstance(new EncryptionInfo(fs));
        if (!decryptor.verifyPassword(password != null ? password : Decryptor.DEFAULT_PASSWORD)) {
          if (password == null) {
            throw new FilePasswordRequiredException();
          }
          throw new FilePasswordIncorrectException();
        }
        try (InputStream in = decryptor.getDataStream(fs)) {
          return new DocxPackage(OPCPackage.open(in));
        }
      }
    } catch (InvalidFormatException | UnsupportedFileFormatException e) {
      throw new FileFormatNotSupportedException(SUPPORTED_FORMATS);
    } catch (GeneralSecurityException e) {
      throw new IOException(e);
    }
  }

  PackagePart mainPart() {
    return mainPart;
  }

  /** Parses the part as XML and returns its root element. */
  Element root(PackagePart part) throws IOException {
    try (InputStream is = part.getInputStream()) {
      return DocumentHelper.readDocument(is).getDocumentElement();
    } catch (SAXException e) {
      throw new IOException(e);
    }
  }

  byte[] bytes(PackagePart part) throws IOException {
    try (InputStream is = part.getInputStream()) {
      return is.readAllBytes();
    }
  }

  /** Returns the part related from {@code source} whose relationship type ends with the suffix. */
  @Nullable PackagePart relatedByType(PackagePart source, String typeSuffix) throws IOException {
    try {
      for (PackageRelationship rel : source.getRelationships()) {
        if (rel.getTargetMode() != TargetMode.EXTERNAL
            && rel.getRelationshipType().endsWith(typeSuffix)) {
          return source.getRelatedPart(rel);
        }
      }
    } catch (InvalidFormatException | IllegalArgumentException e) {
      return null;
    }
    return null;
  }

  /** Returns the part related from {@code source} with the relationship ID. */
  @Nullable PackagePart relatedById(PackagePart source, @Nullable String relId) {
    if (relId == null) {
      return null;
    }
    try {
      PackageRelationship rel = source.getRelationship(relId);
      if (rel == null || rel.getTargetMode() == TargetMode.EXTERNAL) {
        return null;
      }
      return source.getRelatedPart(rel);
    } catch (InvalidFormatException | IllegalArgumentException e) {
      return null;
    }
  }

  @Override
  public void close() throws IOException {
    pkg.revert();
  }
}
