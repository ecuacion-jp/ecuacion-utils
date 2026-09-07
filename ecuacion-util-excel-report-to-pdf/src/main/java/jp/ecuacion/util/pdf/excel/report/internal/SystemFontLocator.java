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
package jp.ecuacion.util.pdf.excel.report.internal;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import org.apache.fontbox.ttf.NamingTable;
import org.apache.fontbox.ttf.TTFParser;
import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.io.RandomAccessReadBufferedFile;
import org.jspecify.annotations.Nullable;

/**
 * Locates system font files and computes the Maximum Digit Width (MDW) from font metrics.
 *
 * <p>MDW is the pixel width of the widest digit character (0–9) in the workbook's normal font
 * at 96 DPI, used by Excel to convert column widths from character units to physical pixels.
 * Computing MDW directly from the font's TTF data (rather than from Java AWT font metrics)
 * yields values that match Excel's internal calculation.</p>
 */
public class SystemFontLocator {

  private static final int FALLBACK_MDW = 7;

  // Process-lifetime cache: font scanning is expensive (reads hundreds of font files).
  // Fonts are not expected to change during a JVM session.
  private static final Map<String, Optional<Path>> FONT_FILE_CACHE = new ConcurrentHashMap<>();

  private SystemFontLocator() {}

  /**
   * Searches system font directories for a font file whose family name matches
   * {@code fontName}.
   *
   * <p>Matching is attempted in two passes to ensure the most specific result:
   * <ol>
   *   <li><b>Exact match</b>: a font face whose family name (nameId=1, 4, or 16)
   *       equals {@code fontName} exactly. This prevents "Meiryo.ttf" (family="Meiryo") from
   *       masking "meiryo.ttc" (which contains "Meiryo UI") when searching for "Meiryo UI".</li>
   *   <li><b>Prefix/broad match</b>: the existing prefix heuristic, as a fallback when no
   *       exact match exists.</li>
   * </ol>
   *
   * @param fontName font family name (e.g. {@code "Meiryo UI"}, {@code "Calibri"})
   * @return path to the matching font file, or empty if not found
   */
  public static Optional<Path> findFontFile(String fontName) {
    return FONT_FILE_CACHE.computeIfAbsent(fontName, SystemFontLocator::findFontFileUncached);
  }

  private static Optional<Path> findFontFileUncached(String fontName) {
    String targetLower = fontName.toLowerCase(Locale.ENGLISH);
    boolean targetIsBold = targetLower.endsWith(" bold");
    String targetBase =
        targetIsBold ? targetLower.substring(0, targetLower.length() - 5).trim() : targetLower;

    // Look the name up against the pre-built face index (in memory, no I/O) instead of
    // walking and re-parsing every font file for each distinct name queried.
    List<Path> exactMatches = new ArrayList<>();
    List<Path> broadMatches = new ArrayList<>();
    for (FontFaceRecord face : getFontFaceIndex()) {
      if (face.bold() != targetIsBold) {
        continue;
      }
      if (face.exactNames().contains(targetBase)) {
        exactMatches.add(face.file());
      } else if (matchesBroadName(face.broadNames(), targetBase)) {
        broadMatches.add(face.file());
      }
    }
    List<Path> candidates = !exactMatches.isEmpty() ? exactMatches : broadMatches;
    return candidates.stream().distinct()
        .min(Comparator.comparingInt(p -> getRegularStyleScore(p, fontName)));
  }

  private static boolean matchesBroadName(Set<String> names, String targetBase) {
    for (String name : names) {
      if (name.equals(targetBase)
          || targetBase.startsWith(name + " ")
          || name.startsWith(targetBase + " ")) {
        return true;
      }
    }
    return false;
  }

  /**
   * One font face's name-matching data, extracted once when the font face index is built.
   *
   * @param file       the font file this face was read from (a {@code .ttc} may contribute
   *                   several {@code FontFaceRecord}s, one per embedded face)
   * @param exactNames lower-cased family names from nameId 1, 4, and 16 — used for exact-match
   *                   lookups
   * @param broadNames lower-cased family names from nameId 1 and 16 — used for prefix/broad
   *                   lookups
   * @param bold       whether nameId 2 (subfamily) contains "bold"
   */
  private record FontFaceRecord(
      Path file, Set<String> exactNames, Set<String> broadNames, boolean bold) {}

  // Process-lifetime cache of every font face's name data, built once on first use. Without
  // it, each distinct font name looked up via findFontFile() would re-walk every system font
  // directory and re-parse every font file found there, even though the set of installed
  // fonts never changes during a JVM session — with the hundreds of files under a typical
  // system font directory (e.g. C:\Windows\Fonts), that made every new distinct font name
  // query cost tens of seconds. Fonts are not expected to change during a JVM session.
  private static volatile @Nullable List<FontFaceRecord> fontFaceIndex;

  private static List<FontFaceRecord> getFontFaceIndex() {
    List<FontFaceRecord> index = fontFaceIndex;
    if (index == null) {
      synchronized (SystemFontLocator.class) {
        index = fontFaceIndex;
        if (index == null) {
          index = buildFontFaceIndex();
          fontFaceIndex = index;
        }
      }
    }
    return index;
  }

  private static List<FontFaceRecord> buildFontFaceIndex() {
    List<FontFaceRecord> records = new ArrayList<>();
    for (Path dir : getSystemFontDirectories()) {
      if (!Files.isDirectory(dir)) {
        continue;
      }
      try (var stream = Files.walk(dir, 3)) {
        stream.filter(SystemFontLocator::isFontFile)
            .forEach(p -> collectFontFaceRecords(p, records));
      } catch (IOException e) { // NOPMD - silently skip unreadable directories
      }
    }
    return records;
  }

  private static void collectFontFaceRecords(Path fontFile, List<FontFaceRecord> out) {
    String lower = fontFile.getFileName().toString().toLowerCase(Locale.ENGLISH);
    try {
      if (lower.endsWith(".ttc")) {
        try (TrueTypeCollection ttc = new TrueTypeCollection(fontFile.toFile())) {
          ttc.processAllFonts(ttf -> out.add(toFontFaceRecord(fontFile, ttf)));
        }
      } else {
        TrueTypeFont ttf =
            new TTFParser().parse(new RandomAccessReadBufferedFile(fontFile.toFile()));
        try {
          out.add(toFontFaceRecord(fontFile, ttf));
        } finally {
          ttf.close();
        }
      }
    } catch (IOException e) { // NOPMD - silently skip unreadable/corrupt font files
    }
  }

  private static FontFaceRecord toFontFaceRecord(Path file, TrueTypeFont ttf) {
    Set<String> exactNames = new HashSet<>();
    Set<String> broadNames = new HashSet<>();
    boolean bold = false;
    try {
      NamingTable naming = ttf.getNaming();
      if (naming != null) {
        for (var record : naming.getNameRecords()) {
          int nameId = record.getNameId();
          String value = record.getString();
          if (value == null) {
            continue;
          }
          String valueLower = value.toLowerCase(Locale.ENGLISH);
          if (nameId == 2 && valueLower.contains("bold")) {
            bold = true;
          }
          if (nameId == 1 || nameId == 4 || nameId == 16) {
            exactNames.add(valueLower);
          }
          if (nameId == 1 || nameId == 16) {
            broadNames.add(valueLower);
          }
        }
      }
    } catch (IOException e) { // NOPMD - treat as a nameless face; it simply won't match
    }
    return new FontFaceRecord(file, exactNames, broadNames, bold);
  }

  /**
   * Returns a preference score for a font file when choosing among multiple files that all
   * contain a font matching {@code fontName}.
   *
   * <p>CJK font families (e.g. 游ゴシック) often ship separate files for each weight
   * (Light, Medium, Regular, Bold), all sharing the same family name. Excel on macOS renders
   * non-bold text using the <em>Medium</em> weight rather than Light (which appears too thin
   * at body sizes). To match that rendering, Medium is ranked most preferred.</p>
   *
   * <p>Score table (lower = more preferred):
   * <ul>
   *   <li>0 — Medium</li>
   *   <li>1 — Regular</li>
   *   <li>2 — unknown / other non-bold style</li>
   *   <li>3 — Light</li>
   * </ul>
   * </p>
   */
  static int getRegularStyleScore(Path fontFile, String fontName) {
    try {
      TrueTypeFont ttf = loadTrueTypeFont(fontFile, fontName);
      if (ttf == null) {
        return 2;
      }
      NamingTable naming = ttf.getNaming();
      if (naming == null) {
        return 2;
      }
      // nameId=2 (Subfamily/Style) is "Regular" for ALL Yu Gothic weight variants
      // (Light, Medium, Regular, Bold all store "Regular" in nameId=2), so it cannot
      // be used to distinguish weights.  Instead, check nameId=4 (Full Name) which
      // carries the explicit weight descriptor (e.g. "Yu Gothic Medium"), then fall
      // back to nameId=1 (Family Name, e.g. "Yu Gothic Light") if nameId=4 is absent.
      for (int targetId : new int[] {4, 2, 1}) {
        for (var record : naming.getNameRecords()) {
          if (record.getNameId() == targetId) {
            String value = record.getString();
            if (value == null) {
              continue;
            }
            String lower = value.toLowerCase(Locale.ENGLISH);
            // Strongly penalise italic and bold — this function is for regular (upright) weight.
            if (lower.contains("italic")) {
              return 20;
            }
            if (lower.contains("bold")) {
              return 10;
            }
            if (lower.contains("medium")) {
              return 0;
            }
            if (lower.contains("regular")) {
              return 1;
            }
            if (lower.contains("light")) {
              return 3;
            }
          }
        }
      }
      return 2;
    } catch (Exception e) { // NOPMD - font scoring failure is non-fatal
      return 2;
    }
  }

  /**
   * Computes the MDW for the given font at the specified point size.
   *
   * <p>The advance width of the {@code '0'} glyph is read directly from the font's
   * horizontal metrics table (TTF {@code hmtx}), converted to points, and then
   * converted to pixels at the given screen PPI. This matches Excel's MDW computation,
   * which uses the actual screen DPI (not a fixed 96 DPI) on macOS.</p>
   *
   * @param fontFile  path to the TTF or TTC font file
   * @param fontName  family name used to select the correct font within a TTC collection
   * @param fontSizePt point size of the normal font (e.g. 11.0f)
   * @return MDW in pixels at the given PPI, falling back to {@value #FALLBACK_MDW} on error
   */
  public static int computeMdw(Path fontFile, String fontName, float fontSizePt) {
    return computeMdw(fontFile, fontName, fontSizePt, 96);
  }

  /**
   * Computes MDW at an explicitly specified screen PPI.
   *
   * <p>Excel on macOS uses the logical screen DPI (not the fixed 96 DPI of Windows GDI)
   * when computing MDW. Pass {@link #getScreenPpi()} to replicate that behaviour.</p>
   *
   * @param fontFile  path to the TTF or TTC font file
   * @param fontName  family name used to select the correct font within a TTC collection
   * @param fontSizePt point size of the normal font
   * @param ppi       screen pixels per inch to use for the conversion (typically 96 or the
   *                  logical screen resolution obtained via {@link #getScreenPpi()})
   * @return MDW in pixels at the given PPI, falling back to {@value #FALLBACK_MDW} on error
   */
  public static int computeMdw(Path fontFile, String fontName, float fontSizePt, int ppi) {
    try {
      TrueTypeFont ttf = loadTrueTypeFont(fontFile, fontName);
      if (ttf == null) {
        return FALLBACK_MDW;
      }
      try {
        int unitsPerEm = ttf.getUnitsPerEm();
        if (unitsPerEm <= 0) {
          return FALLBACK_MDW;
        }
        var cmapLookup = ttf.getUnicodeCmapLookup();
        int gid = (cmapLookup != null) ? cmapLookup.getGlyphId(0x0030) : 0; // U+0030 = '0'
        int advanceWidth = ttf.getHorizontalMetrics().getAdvanceWidth(gid);
        double advancePts = (double) advanceWidth / unitsPerEm * fontSizePt;
        double advancePxAtPpi = advancePts * ppi / 72.0;
        return Math.max(1, (int) Math.round(advancePxAtPpi));
      } finally {
        ttf.close();
      }
    } catch (IOException e) {
      return FALLBACK_MDW;
    }
  }

  /**
   * Computes MDW using {@code Math.ceil} to match Excel's pixel-measurement behaviour.
   *
   * <p>Excel measures the maximum digit width (MDW) by rendering digits and counting screen
   * pixels.  Fractional advance widths are always rounded up to the next pixel, so a glyph
   * whose theoretical advance is 7.2 px occupies 8 px of rendered space.  Use this method
   * when computing MDW from a workbook's theme Latin font (e.g. Calibri) so that
   * naturalColTotal and fit-to-page scale factor match Excel's own computation.</p>
   *
   * <p>Regular rendering fonts (e.g. NotoSansJP used as fallback) should still use
   * {@link #computeMdw(Path, String, float)} which uses {@code Math.round}.</p>
   */
  public static int computeExcelMdw(Path fontFile, String fontName, float fontSizePt) {
    return computeExcelMdw(fontFile, fontName, fontSizePt, 96);
  }

  /** Ceil-based MDW at an explicit PPI — see {@link #computeExcelMdw(Path, String, float)}. */
  public static int computeExcelMdw(Path fontFile, String fontName, float fontSizePt, int ppi) {
    try {
      TrueTypeFont ttf = loadTrueTypeFont(fontFile, fontName);
      if (ttf == null) {
        return FALLBACK_MDW;
      }
      try {
        int unitsPerEm = ttf.getUnitsPerEm();
        if (unitsPerEm <= 0) {
          return FALLBACK_MDW;
        }
        var cmapLookup = ttf.getUnicodeCmapLookup();
        int gid = (cmapLookup != null) ? cmapLookup.getGlyphId(0x0030) : 0;
        int advanceWidth = ttf.getHorizontalMetrics().getAdvanceWidth(gid);
        double advancePts = (double) advanceWidth / unitsPerEm * fontSizePt;
        double advancePxAtPpi = advancePts * ppi / 72.0;
        return Math.max(1, (int) Math.ceil(advancePxAtPpi));
      } finally {
        ttf.close();
      }
    } catch (IOException e) {
      return FALLBACK_MDW;
    }
  }

  /**
   * Returns the logical screen resolution in pixels per inch.
   *
   * <p>Excel on macOS computes MDW using the screen's logical DPI rather than the
   * standard 96 DPI used on Windows. On Retina displays this is typically the physical
   * PPI divided by the HiDPI scale factor (e.g. 127 PPI on a 254 PPI MacBook Pro 14"
   * at 2× scale), which causes Fit-to-page scales to differ from Windows Excel.
   * Returns 96 if the screen resolution cannot be determined (e.g. headless servers).</p>
   *
   * @return logical screen PPI, or 96 as a safe fallback
   */
  public static int getScreenPpi() {
    try {
      // Use reflection to avoid a direct dependency on java.awt.Toolkit, which is unavailable
      // in headless/server environments even when java.desktop is declared in module-info.
      var toolkitClass = Class.forName("java.awt.Toolkit");
      var toolkit = toolkitClass.getMethod("getDefaultToolkit").invoke(null);
      int ppi = (int) toolkitClass.getMethod("getScreenResolution").invoke(toolkit);
      return (ppi > 0) ? ppi : 96;
    } catch (Exception ignored) { // NOPMD
      return 96;
    }
  }

  /**
   * Loads the {@link TrueTypeFont} matching {@code fontName} from the given file.
   * Handles both plain TTF files and TrueType Collections (TTC).
   *
   * <p>For TTC files, matching is attempted in this order:
   * <ol>
   *   <li>PostScript name via {@code getFontByName()} (e.g. "MeiryoUI")</li>
   *   <li>Family/full name via {@link #matchesFontName} (e.g. nameId=1 "Meiryo UI")</li>
   *   <li>First font in the collection as a last resort</li>
   * </ol>
   * Using only {@code getFontByName} is insufficient: Meiryo UI's PostScript name is
   * "MeiryoUI" (no space), so {@code getFontByName("Meiryo UI")} returns null and the
   * fall-back would silently load "Meiryo" (full-width kana) instead of "Meiryo UI"
   * (proportional kana, ~58% em width).</p>
   *
   * @return the font, or {@code null} if the font could not be loaded
   */
  @SuppressWarnings("resource")
  @Nullable
  public static TrueTypeFont loadTrueTypeFont(Path fontFile, String fontName) throws IOException {
    String lower = fontFile.getFileName().toString().toLowerCase(Locale.ENGLISH);
    if (lower.endsWith(".ttc")) {
      // None of the TrueTypeCollection instances below are closed when they end up supplying
      // the returned font: a TrueTypeFont obtained from a collection reads its glyph/table
      // data lazily from the collection's underlying stream (e.g. when PDFBox embeds the font
      // into the output PDF later), so closing the collection would invalidate the handle
      // returned here. Fonts are cached for the process lifetime (see class Javadoc), so the
      // underlying file descriptors are released only at JVM exit.

      // 1. Try PostScript name match (getFontByName uses nameId=6)
      TrueTypeCollection ttc = new TrueTypeCollection(fontFile.toFile());
      TrueTypeFont foundByName = null;
      try {
        foundByName = ttc.getFontByName(fontName);
      } catch (IOException ignored) {
        // getFontByName throws if the named font is not in the collection
      }
      if (foundByName != null) {
        return foundByName;
      } else {
        // No match: safe to close since no returned font depends on this instance's stream.
        ttc.close();
      }

      // 2a. Exact family/full name match (nameId=1, 4, 16 must equal target exactly).
      //     Must come before prefix matching: "Meiryo" is a prefix of "Meiryo UI",
      //     so prefix matching would incorrectly return "Meiryo" for the target "Meiryo UI".
      //     Open a fresh TTC instance: getFontByName() above may have advanced the stream.
      var fonts = new ArrayList<TrueTypeFont>();
      new TrueTypeCollection(fontFile.toFile()).processAllFonts(fonts::add);
      for (TrueTypeFont f : fonts) {
        if (matchesFontNameExact(f, fontName)) {
          return f;
        }
      }
      // 2b. Prefix/broad match — fallback for fonts where the family name is a prefix of
      //     the target (e.g. "Meiryo" family file matching search for "Meiryo").
      for (TrueTypeFont f : fonts) {
        if (matchesFontName(f, fontName)) {
          return f;
        }
      }
      // 3. Last resort: first font in the collection
      return fonts.isEmpty() ? null : fonts.get(0);
    } else {
      return new TTFParser().parse(new RandomAccessReadBufferedFile(fontFile.toFile()));
    }
  }

  // -------------------------------------------------------------------------
  // Font directory scanning helpers
  // -------------------------------------------------------------------------

  private static List<Path> getSystemFontDirectories() {
    String os = System.getProperty("os.name", "").toLowerCase(Locale.ENGLISH);
    List<Path> dirs = new ArrayList<>();
    // Derive the filesystem root from user.home to avoid hardcoded absolute path literals.
    Path userHome = Path.of(System.getProperty("user.home", ""));
    Path fsRoot = userHome.getRoot();

    if (os.contains("mac")) {
      if (fsRoot != null) {
        // Microsoft Office for Mac installs fonts here
        dirs.add(fsRoot.resolve("Library").resolve("Application Support")
            .resolve("Microsoft").resolve("Fonts"));
        dirs.add(fsRoot.resolve("Library").resolve("Fonts"));
        dirs.add(fsRoot.resolve("System").resolve("Library").resolve("Fonts"));
        // Microsoft Office app bundle DFonts — each Office app ships its own copy.
        // Yu Gothic (游ゴシック) is in Word/Outlook/PowerPoint but not always in Excel.
        for (String app : new String[] {
            "Microsoft Excel.app", "Microsoft Word.app",
            "Microsoft Outlook.app", "Microsoft PowerPoint.app"}) {
          dirs.add(fsRoot.resolve("Applications").resolve(app)
              .resolve("Contents").resolve("Resources").resolve("DFonts"));
        }
        // Microsoft Office app bundle shared fonts (Office 365 layout)
        dirs.add(fsRoot.resolve("Applications").resolve("Microsoft Office")
            .resolve("Office").resolve("Fonts"));
      }
      dirs.add(userHome.resolve("Library").resolve("Fonts"));
    } else if (os.contains("win")) {
      String windir = System.getenv("WINDIR");
      if (windir == null) {
        windir = System.getenv("SYSTEMROOT");
      }
      if (windir != null) {
        dirs.add(Path.of(windir).resolve("Fonts"));
      }
      String programFiles = System.getenv("ProgramFiles");
      if (programFiles != null) {
        dirs.add(Path.of(programFiles).resolve("Microsoft Office").resolve("root")
            .resolve("Fonts"));
        dirs.add(Path.of(programFiles).resolve("Microsoft Office").resolve("Office16")
            .resolve("Fonts"));
      }
    } else {
      // Linux / other
      if (fsRoot != null) {
        dirs.add(fsRoot.resolve("usr").resolve("share").resolve("fonts"));
        dirs.add(fsRoot.resolve("usr").resolve("local").resolve("share").resolve("fonts"));
      }
      dirs.add(userHome.resolve(".fonts"));
    }
    return dirs;
  }

  private static boolean isFontFile(Path path) {
    if (!Files.isRegularFile(path)) {
      return false;
    }
    String name = path.getFileName().toString().toLowerCase(Locale.ENGLISH);
    return name.endsWith(".ttf") || name.endsWith(".ttc");
  }

  /**
   * Returns {@code true} if the font's family name (nameId=1, 4, or 16) exactly equals
   * {@code targetName} (case-insensitive). Unlike {@link #matchesFontName}, no prefix matching
   * is performed, so "Meiryo" will NOT match the target "Meiryo UI".
   */
  private static boolean matchesFontNameExact(TrueTypeFont ttf, String targetName) {
    try {
      NamingTable naming = ttf.getNaming();
      if (naming == null) {
        return false;
      }
      String target = targetName.toLowerCase(Locale.ENGLISH);
      boolean targetIsBold = target.endsWith(" bold");
      String targetBase = targetIsBold ? target.substring(0, target.length() - 5).trim() : target;

      boolean fontHasBoldStyle = false;
      java.util.Set<String> names = new java.util.HashSet<>();
      for (var record : naming.getNameRecords()) {
        int nameId = record.getNameId();
        String value = record.getString();
        if (value == null) {
          continue;
        }
        String lower = value.toLowerCase(Locale.ENGLISH);
        // "Bold Italic" also indicates a bold font, so use contains() not equals().
        if (nameId == 2 && lower.contains("bold")) {
          fontHasBoldStyle = true;
        }
        if (nameId == 1 || nameId == 4 || nameId == 16) {
          names.add(lower);
        }
      }
      if (targetIsBold != fontHasBoldStyle) {
        return false;
      }
      return names.contains(targetBase);
    } catch (IOException e) {
      return false;
    }
  }

  private static boolean matchesFontName(TrueTypeFont ttf, String targetName) {
    try {
      NamingTable naming = ttf.getNaming();
      if (naming == null) {
        return false;
      }
      String target = targetName.toLowerCase(Locale.ENGLISH);
      // "Meiryo UI Bold" → targetIsBold=true, targetBase="meiryo ui"
      boolean targetIsBold = target.endsWith(" bold");
      String targetBase = targetIsBold ? target.substring(0, target.length() - 5).trim() : target;

      // Collect family names (nameId 1 = Family, 16 = Preferred Family) and
      // whether this font variant is bold (nameId 2 = Subfamily/Style = "Bold").
      boolean fontHasBoldStyle = false;
      java.util.Set<String> families = new java.util.HashSet<>();
      for (var record : naming.getNameRecords()) {
        int nameId = record.getNameId();
        String value = record.getString();
        if (value == null) {
          continue;
        }
        String lower = value.toLowerCase(Locale.ENGLISH);
        // nameId 2 (Subfamily) = "Bold" or "Bold Italic" indicates a bold font. Use contains()
        // because bold-italic fonts have nameId 2 = "Bold Italic", not just "Bold".
        if (nameId == 2 && lower.contains("bold")) {
          fontHasBoldStyle = true;
        }
        if (nameId == 1 || nameId == 16) {
          families.add(lower);
        }
      }

      // Bold/regular must match: do not return a bold font for a regular search or vice versa.
      if (targetIsBold != fontHasBoldStyle) {
        return false;
      }

      // Family name: exact or prefix match (handles "Meiryo" matching "Meiryo UI").
      for (String family : families) {
        if (family.equals(targetBase)
            || targetBase.startsWith(family + " ")
            || family.startsWith(targetBase + " ")) {
          return true;
        }
      }
      return false;
    } catch (IOException e) {
      return false;
    }
  }
}
