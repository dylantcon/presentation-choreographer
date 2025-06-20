package com.presentationchoreographer.xml.writers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.xpath.*;
import java.io.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;
import com.presentationchoreographer.exceptions.XMLParsingException;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Global Shape ID (SPID) management system for PowerPoint presentations.
 * 
 * <p>This class provides centralized tracking and allocation of Shape IDs across
 * an entire PPTX presentation to prevent conflicts when copying slides or 
 * adding new shapes. It maintains a global registry of all SPIDs and provides
 * collision-free ID allocation for slide copying operations.</p>
 * 
 * <p>Key responsibilities:</p>
 * <ul>
 *   <li>Global SPID registry across all slides in the presentation</li>
 *   <li>Collision-free SPID allocation for new shapes and copied slides</li>
 *   <li>SPID regeneration with animation reference preservation</li>
 *   <li>Validation of SPID uniqueness across the presentation</li>
 * </ul>
 * 
 * <p>Thread Safety: This class is thread-safe through the use of concurrent
 * collections and atomic operations for ID generation.</p>
 * 
 * @author Presentation Choreographer
 * @version 1.0
 * @since 1.0
 */
public class SPIDManager {

  /**
   * Reference to the extracted PPTX directory containing all presentation parts.
   */
  private final File extractedPptxDir;

  /**
   * DOM document builder for parsing slide XML files.
   */
  private final DocumentBuilder documentBuilder;

  /**
   * XPath processor for querying shape and animation XML structures.
   */
  private final XPath xpath;

  /**
   * Global registry of all Shape IDs currently in use across the presentation.
   * Key: SPID (Integer)
   * Value: SPIDInfo containing slide number and shape details
   */
  private final Map<Integer, SPIDInfo> globalSpidRegistry;

  /**
   * Atomic counter for generating unique SPIDs across the presentation.
   * Ensures thread-safe ID allocation even in concurrent scenarios.
   */
  private final AtomicInteger nextSpidCounter;

  /**
   * Cache of parsed slide documents to avoid repeated file I/O during operations.
   * Key: Slide number (Integer)
   * Value: Parsed slide Document
   */
  private final Map<Integer, Document> slideDocumentCache;

  /**
   * Constructs a new SPIDManager for the specified PPTX directory.
   * 
   * <p>During construction, this manager will scan all slides in the presentation
   * to build a comprehensive registry of existing SPIDs and determine the next
   * available SPID for allocation.</p>
   * 
   * @param extractedPptxDir The directory containing the extracted PPTX contents
   * @throws XMLParsingException If the XML parser cannot be initialized or
   *                           if existing slides cannot be scanned
   * @throws IllegalArgumentException If extractedPptxDir is null or does not exist
   */
  public SPIDManager(File extractedPptxDir) throws XMLParsingException {
    if (extractedPptxDir == null || !extractedPptxDir.exists()) {
      throw new IllegalArgumentException("extractedPptxDir must exist and be non-null");
    }

    this.extractedPptxDir = extractedPptxDir;
    this.globalSpidRegistry = new ConcurrentHashMap<>();
    this.slideDocumentCache = new ConcurrentHashMap<>();
    this.nextSpidCounter = new AtomicInteger(1);

    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();

      XPathFactory xpathFactory = XPathFactory.newInstance();
      this.xpath = xpathFactory.newXPath();
      this.xpath.setNamespaceContext(XMLConstants.createNamespaceContext());

      // Scan all existing slides to build global SPID registry
      scanAllSlidesForSpids();

    } catch (ParserConfigurationException e) {
      throw new XMLParsingException("Failed to initialize SPIDManager", e);
    }
  }

  /**
   * Regenerates all SPIDs in the specified slide document to avoid conflicts.
   * 
   * <p>This method is essential for slide copying operations. It parses all shapes
   * in the slide, generates new unique SPIDs, updates the shape references, and
   * updates any animation references that target those shapes.</p>
   * 
   * @param slideDocument The slide document to regenerate SPIDs for
   * @param slideNumber The slide number for logging and reference purposes
   * @return SPIDRegenerationResult containing old-to-new SPID mappings and statistics
   * @throws XMLParsingException If SPID regeneration fails
   * @throws IllegalArgumentException If slideDocument is null
   */
  public SPIDRegenerationResult regenerateSpids(Document slideDocument, int slideNumber) throws XMLParsingException {
    if (slideDocument == null) {
      throw new IllegalArgumentException("slideDocument cannot be null");
    }

    try {
      Map<Integer, Integer> spidMappings = new HashMap<>();
      int shapesProcessed = 0;
      int animationsUpdated = 0;

      // Step 1: Find all shapes in the slide and collect their current SPIDs
      NodeList shapeElements = (NodeList) xpath.evaluate(
          XMLConstants.XPATH_ALL_SHAPES_AND_PICTURES, slideDocument, XPathConstants.NODESET);

      List<Element> shapesToUpdate = new ArrayList<>();
      List<Integer> currentSpids = new ArrayList<>();

      for (int i = 0; i < shapeElements.getLength(); i++) {
        Element shapeElement = (Element) shapeElements.item(i);
        String spidStr = (String) xpath.evaluate(
            XMLConstants.XPATH_SHAPE_ID_ATTRIBUTE, shapeElement, XPathConstants.STRING);

        if (!spidStr.isEmpty()) {
          int currentSpid = Integer.parseInt(spidStr);
          currentSpids.add(currentSpid);
          shapesToUpdate.add(shapeElement);
        }
      }

      // Step 2: Generate new unique SPIDs for all shapes
      for (int i = 0; i < currentSpids.size(); i++) {
        int oldSpid = currentSpids.get(i);
        int newSpid = allocateUniqueSpid();
        spidMappings.put(oldSpid, newSpid);

        // Update the shape's SPID in the XML
        Element shapeElement = shapesToUpdate.get(i);
        Element cNvPr = (Element) xpath.evaluate(".//p:cNvPr", shapeElement, XPathConstants.NODE);
        if (cNvPr != null) {
          cNvPr.setAttribute("id", String.valueOf(newSpid));
          shapesProcessed++;
        }

        // Register the new SPID in our global registry
        registerSpid(newSpid, slideNumber, "regenerated_shape_" + newSpid);
      }

      // Step 3: Update animation references to point to new SPIDs
      animationsUpdated = updateAnimationReferences(slideDocument, spidMappings);

      return new SPIDRegenerationResult(spidMappings, shapesProcessed, animationsUpdated);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to regenerate SPIDs for slide " + slideNumber, e);
    }
  }

  /**
   * Allocates a unique SPID that is guaranteed not to conflict with any
   * existing SPIDs in the presentation.
   * 
   * @return A unique SPID (Shape ID)
   */
  public int allocateUniqueSpid() {
    int candidateSpid;
    do {
      candidateSpid = nextSpidCounter.getAndIncrement();
    } while (globalSpidRegistry.containsKey(candidateSpid));

    return candidateSpid;
  }

  /**
   * Allocates multiple unique SPIDs in a single operation.
   * 
   * @param count The number of SPIDs to allocate
   * @return A list of unique SPIDs
   * @throws IllegalArgumentException If count is less than 1
   */
  public List<Integer> allocateUniqueSpids(int count) {
    if (count < 1) {
      throw new IllegalArgumentException("count must be positive");
    }

    List<Integer> allocatedSpids = new ArrayList<>();
    for (int i = 0; i < count; i++) {
      allocatedSpids.add(allocateUniqueSpid());
    }
    return allocatedSpids;
  }

  /**
   * Registers a SPID in the global registry with associated metadata.
   * 
   * @param spid The Shape ID to register
   * @param slideNumber The slide number containing this shape
   * @param shapeName The name of the shape (for debugging/logging)
   */
  public void registerSpid(int spid, int slideNumber, String shapeName) {
    globalSpidRegistry.put(spid, new SPIDInfo(slideNumber, shapeName));
  }

  /**
   * Unregisters a SPID from the global registry.
   * 
   * @param spid The Shape ID to unregister
   * @return true if the SPID was found and removed, false otherwise
   */
  public boolean unregisterSpid(int spid) {
    return globalSpidRegistry.remove(spid) != null;
  }

  /**
   * Checks if a SPID is currently in use across the presentation.
   * 
   * @param spid The Shape ID to check
   * @return true if the SPID is in use, false otherwise
   */
  public boolean isSpidInUse(int spid) {
    return globalSpidRegistry.containsKey(spid);
  }

  /**
   * Gets information about a specific SPID.
   * 
   * @param spid The Shape ID to look up
   * @return SPIDInfo containing slide number and shape name, or null if not found
   */
  public SPIDInfo getSpidInfo(int spid) {
    return globalSpidRegistry.get(spid);
  }

  /**
   * Gets all SPIDs currently in use across the presentation.
   * 
   * @return An unmodifiable set of all SPIDs
   */
  public Set<Integer> getAllSpids() {
    return Collections.unmodifiableSet(globalSpidRegistry.keySet());
  }

  /**
   * Gets all SPIDs for a specific slide.
   * 
   * @param slideNumber The slide number to get SPIDs for
   * @return A set of SPIDs used in the specified slide
   */
  public Set<Integer> getSpidsForSlide(int slideNumber) {
    return globalSpidRegistry.entrySet().stream()
      .filter(entry -> entry.getValue().getSlideNumber() == slideNumber)
      .map(Map.Entry::getKey)
      .collect(java.util.stream.Collectors.toSet());
  }

  /**
   * Validates that all SPIDs in the presentation are unique and consistent.
   * 
   * @return ValidationResult containing any detected SPID conflicts or issues
   * @throws XMLParsingException If validation cannot be performed
   */
  public ValidationResult validateSpidUniqueness() throws XMLParsingException {
    List<String> errors = new ArrayList<>();
    List<String> warnings = new ArrayList<>();

    try {
      // Re-scan all slides and compare with our registry
      Map<Integer, Integer> actualSpidCounts = new HashMap<>();

      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      File[] slideFiles = slidesDir.listFiles((dir, name) -> name.matches("slide\\d+\\.xml"));

      if (slideFiles != null) {
        for (File slideFile : slideFiles) {
          Document slideDoc = documentBuilder.parse(slideFile);
          NodeList shapes = (NodeList) xpath.evaluate(
              XMLConstants.XPATH_ALL_SHAPES_AND_PICTURES, slideDoc, XPathConstants.NODESET);

          for (int i = 0; i < shapes.getLength(); i++) {
            Element shape = (Element) shapes.item(i);
            String spidStr = (String) xpath.evaluate(
                XMLConstants.XPATH_SHAPE_ID_ATTRIBUTE, shape, XPathConstants.STRING);

            if (!spidStr.isEmpty()) {
              int spid = Integer.parseInt(spidStr);
              actualSpidCounts.put(spid, actualSpidCounts.getOrDefault(spid, 0) + 1);
            }
          }
        }
      }

      // Check for duplicates
      for (Map.Entry<Integer, Integer> entry : actualSpidCounts.entrySet()) {
        if (entry.getValue() > 1) {
          errors.add("Duplicate SPID detected: " + entry.getKey() + 
              " appears " + entry.getValue() + " times");
        }
      }

      // Check registry consistency
      for (Integer spid : globalSpidRegistry.keySet()) {
        if (!actualSpidCounts.containsKey(spid)) {
          warnings.add("SPID " + spid + " is registered but not found in slides");
        }
      }

      for (Integer spid : actualSpidCounts.keySet()) {
        if (!globalSpidRegistry.containsKey(spid)) {
          warnings.add("SPID " + spid + " found in slides but not registered");
        }
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to validate SPID uniqueness", e);
    }

    return new ValidationResult(errors, warnings);
  }

  // ========== PRIVATE HELPER METHODS ==========

  /**
   * Scans all slides in the presentation to build the global SPID registry.
   */
  private void scanAllSlidesForSpids() throws XMLParsingException {
    try {
      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      if (!slidesDir.exists()) {
        return; // No slides to scan
      }

      File[] slideFiles = slidesDir.listFiles((dir, name) -> name.matches("slide\\d+\\.xml"));
      if (slideFiles == null) {
        return;
      }

      for (File slideFile : slideFiles) {
        int slideNumber = extractSlideNumberFromFileName(slideFile.getName());
        scanSlideForSpids(slideFile, slideNumber);
      }

      // Update counter to ensure new SPIDs don't conflict
      updateSpidCounter();

    } catch (Exception e) {
      throw new XMLParsingException("Failed to scan slides for SPIDs", e);
    }
  }

  /**
   * Scans a single slide file for SPIDs and registers them.
   */
  private void scanSlideForSpids(File slideFile, int slideNumber) throws XMLParsingException {
    try {
      Document slideDoc = documentBuilder.parse(slideFile);
      slideDocumentCache.put(slideNumber, slideDoc);

      NodeList shapeElements = (NodeList) xpath.evaluate(
          XMLConstants.XPATH_ALL_SHAPES_AND_PICTURES, slideDoc, XPathConstants.NODESET);

      for (int i = 0; i < shapeElements.getLength(); i++) {
        Element shapeElement = (Element) shapeElements.item(i);

        String spidStr = (String) xpath.evaluate(
            XMLConstants.XPATH_SHAPE_ID_ATTRIBUTE, shapeElement, XPathConstants.STRING);
        String shapeName = (String) xpath.evaluate(
            XMLConstants.XPATH_SHAPE_NAME_ATTRIBUTE, shapeElement, XPathConstants.STRING);

        if (!spidStr.isEmpty()) {
          int spid = Integer.parseInt(spidStr);
          registerSpid(spid, slideNumber, shapeName.isEmpty() ? "unnamed_shape" : shapeName);
        }
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to scan slide " + slideNumber + " for SPIDs", e);
    }
  }

  /**
   * Updates animation references to use new SPIDs after regeneration.
   */
  private int updateAnimationReferences(Document slideDocument, Map<Integer, Integer> spidMappings) 
      throws XPathExpressionException {
      int animationsUpdated = 0;

      // Find all animation target elements that reference SPIDs
      NodeList animationTargets = (NodeList) xpath.evaluate(
          "//p:spTgt[@spid]", slideDocument, XPathConstants.NODESET);

      for (int i = 0; i < animationTargets.getLength(); i++) {
        Element target = (Element) animationTargets.item(i);
        String oldSpidStr = target.getAttribute("spid");

        if (!oldSpidStr.isEmpty()) {
          int oldSpid = Integer.parseInt(oldSpidStr);
          Integer newSpid = spidMappings.get(oldSpid);

          if (newSpid != null) {
            target.setAttribute("spid", String.valueOf(newSpid));
            animationsUpdated++;
          }
        }
      }

      return animationsUpdated;
  }

  /**
   * Updates the SPID counter based on existing SPIDs to avoid conflicts.
   */
  private void updateSpidCounter() {
    int maxSpid = globalSpidRegistry.keySet().stream()
      .mapToInt(Integer::intValue)
      .max()
      .orElse(0);
    nextSpidCounter.set(maxSpid + 1);
  }

  /**
   * Extracts slide number from filename (e.g., "slide5.xml" -> 5).
   */
  private int extractSlideNumberFromFileName(String fileName) {
    try {
      String numberStr = fileName.substring(5, fileName.lastIndexOf(".xml"));
      return Integer.parseInt(numberStr);
    } catch (Exception e) {
      return 1; // Default fallback
    }
  }

  // ========== INNER CLASSES ==========

  /**
   * Contains information about a specific SPID.
   */
  public static class SPIDInfo {
    private final int slideNumber;
    private final String shapeName;

    public SPIDInfo(int slideNumber, String shapeName) {
      this.slideNumber = slideNumber;
      this.shapeName = shapeName;
    }

    public int getSlideNumber() { return slideNumber; }
    public String getShapeName() { return shapeName; }

    @Override
    public String toString() {
      return String.format("SPIDInfo{slide=%d, name='%s'}", slideNumber, shapeName);
    }
  }

  /**
   * Result of SPID regeneration operation.
   */
  public static class SPIDRegenerationResult {
    private final Map<Integer, Integer> spidMappings;
    private final int shapesProcessed;
    private final int animationsUpdated;

    public SPIDRegenerationResult(Map<Integer, Integer> spidMappings, 
        int shapesProcessed, int animationsUpdated) {
      this.spidMappings = Collections.unmodifiableMap(new HashMap<>(spidMappings));
      this.shapesProcessed = shapesProcessed;
      this.animationsUpdated = animationsUpdated;
    }

    public Map<Integer, Integer> getSpidMappings() { return spidMappings; }
    public int getShapesProcessed() { return shapesProcessed; }
    public int getAnimationsUpdated() { return animationsUpdated; }

    @Override
    public String toString() {
      return String.format("SPIDRegenerationResult{shapes=%d, animations=%d, mappings=%d}",
          shapesProcessed, animationsUpdated, spidMappings.size());
    }
  }

  /**
   * Result of SPID validation.
   */
  public static class ValidationResult {
    private final List<String> errors;
    private final List<String> warnings;

    public ValidationResult(List<String> errors, List<String> warnings) {
      this.errors = Collections.unmodifiableList(new ArrayList<>(errors));
      this.warnings = Collections.unmodifiableList(new ArrayList<>(warnings));
    }

    public List<String> getErrors() { return errors; }
    public List<String> getWarnings() { return warnings; }
    public boolean hasErrors() { return !errors.isEmpty(); }
    public boolean hasWarnings() { return !warnings.isEmpty(); }
    public boolean isValid() { return errors.isEmpty(); }
  }
}
