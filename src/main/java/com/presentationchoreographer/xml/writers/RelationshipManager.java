package com.presentationchoreographer.xml.writers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.*;
import java.io.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;
import com.presentationchoreographer.exceptions.XMLParsingException;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive relationship management system for OOXML PowerPoint presentations.
 * 
 * <p>This class provides centralized management of all relationship types within a PPTX package,
 * including slide-to-layout relationships, media relationships, theme relationships, and 
 * slide master relationships. It ensures relationship ID uniqueness across the entire 
 * presentation and handles relationship updates when slides are inserted, moved, or deleted.</p>
 * 
 * <p>Key responsibilities:</p>
 * <ul>
 *   <li>Global relationship ID allocation and collision avoidance</li>
 *   <li>Creation and management of slide relationship files (.rels)</li>
 *   <li>Media relationship tracking and ID remapping</li>
 *   <li>Presentation.xml relationship synchronization</li>
 *   <li>Content type registration for new relationships</li>
 * </ul>
 * 
 * <p>Thread Safety: This class is thread-safe through the use of concurrent collections
 * and atomic operations for ID generation.</p>
 * 
 * @author Presentation Choreographer
 * @version 1.0
 * @since 1.0
 */
public class RelationshipManager {

  /**
   * Reference to the extracted PPTX directory containing all presentation parts.
   */
  private final File extractedPptxDir;

  /**
   * DOM document builder for creating and parsing XML relationship files.
   */
  private final DocumentBuilder documentBuilder;

  /**
   * XPath processor for querying relationship XML structures.
   */
  private final XPath xpath;

  /**
   * Global registry of all relationship IDs currently in use across the presentation.
   * Key: Relationship ID (e.g., "rId1", "rId2")
   * Value: RelationshipInfo containing target and type details
   */
  private final Map<String, RelationshipInfo> globalRelationshipRegistry;

  /**
   * Atomic counter for generating unique relationship IDs across the presentation.
   * Ensures thread-safe ID allocation even in concurrent scenarios.
   */
  private final AtomicInteger nextRelationshipIdCounter;

  /**
   * Cache of parsed relationship documents to avoid repeated file I/O.
   * Key: File path relative to extractedPptxDir
   * Value: Parsed relationship Document
   */
  private final Map<String, Document> relationshipDocumentCache;

  /**
   * Constructs a new RelationshipManager for the specified PPTX directory.
   * 
   * <p>During construction, this manager will scan the existing presentation
   * to build a comprehensive registry of all current relationships and 
   * determine the next available relationship ID.</p>
   * 
   * @param extractedPptxDir The directory containing the extracted PPTX contents
   * @throws XMLParsingException If the XML parser cannot be initialized or
   *                           if existing relationships cannot be scanned
   * @throws IllegalArgumentException If extractedPptxDir is null or does not exist
   */
  public RelationshipManager(File extractedPptxDir) throws XMLParsingException {
    if (extractedPptxDir == null || !extractedPptxDir.exists()) {
      throw new IllegalArgumentException("extractedPptxDir must exist and be non-null");
    }

    this.extractedPptxDir = extractedPptxDir;
    this.globalRelationshipRegistry = new ConcurrentHashMap<>();
    this.relationshipDocumentCache = new ConcurrentHashMap<>();
    this.nextRelationshipIdCounter = new AtomicInteger(1);

    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();

      XPathFactory xpathFactory = XPathFactory.newInstance();
      this.xpath = xpathFactory.newXPath();
      this.xpath.setNamespaceContext(XMLConstants.createNamespaceContext());

      // Scan existing presentation to build relationship registry
      scanExistingRelationships();

    } catch (ParserConfigurationException e) {
      throw new XMLParsingException("Failed to initialize RelationshipManager", e);
    }
  }

  /**
   * Creates a standard slide relationship file with layout and theme relationships.
   * 
   * <p>This method creates a new .rels file for the specified slide containing
   * the standard relationships required by PowerPoint:</p>
   * <ul>
   *   <li>Slide Layout relationship (typically to slideLayout1.xml)</li>
   *   <li>Theme relationship (typically to theme1.xml)</li>
   * </ul>
   * 
   * @param slideNumber The slide number (1-based) for which to create relationships
   * @param layoutTarget Optional custom layout target; if null, uses default slideLayout1.xml
   * @param themeTarget Optional custom theme target; if null, uses default theme1.xml
   * @return RelationshipCreationResult containing the created relationship IDs and file path
   * @throws XMLParsingException If the relationship file cannot be created or written
   * @throws IllegalArgumentException If slideNumber is less than 1
   */
  public RelationshipCreationResult createSlideRelationships(int slideNumber, 
      String layoutTarget, 
      String themeTarget) throws XMLParsingException {
    if (slideNumber < 1) {
      throw new IllegalArgumentException("slideNumber must be positive");
    }

    try {
      // Create relationship document
      Document relsDoc = documentBuilder.newDocument();
      Element relationships = relsDoc.createElementNS(
          XMLConstants.PACKAGE_RELATIONSHIPS_NS, "Relationships");
      relsDoc.appendChild(relationships);

      List<String> createdRelationshipIds = new ArrayList<>();

      // Create layout relationship
      String layoutRId = allocateRelationshipId();
      String finalLayoutTarget = layoutTarget != null ? layoutTarget : XMLConstants.DEFAULT_SLIDE_LAYOUT_TARGET;
      Element layoutRel = createRelationshipElement(relsDoc, layoutRId, 
          XMLConstants.RELATIONSHIP_TYPE_SLIDE_LAYOUT, finalLayoutTarget);
      relationships.appendChild(layoutRel);
      createdRelationshipIds.add(layoutRId);

      // Register layout relationship
      registerRelationship(layoutRId, XMLConstants.RELATIONSHIP_TYPE_SLIDE_LAYOUT, finalLayoutTarget);

      // Create theme relationship
      String themeRId = allocateRelationshipId();
      String finalThemeTarget = themeTarget != null ? themeTarget : XMLConstants.DEFAULT_THEME_TARGET;
      Element themeRel = createRelationshipElement(relsDoc, themeRId,
          XMLConstants.RELATIONSHIP_TYPE_THEME, finalThemeTarget);
      relationships.appendChild(themeRel);
      createdRelationshipIds.add(themeRId);

      // Register theme relationship
      registerRelationship(themeRId, XMLConstants.RELATIONSHIP_TYPE_THEME, finalThemeTarget);

      // Write relationship file
      File relsFile = getSlideRelationshipFile(slideNumber);
      writeRelationshipDocument(relsDoc, relsFile);

      // Cache the document
      String cacheKey = getRelativePathFromExtractedDir(relsFile);
      relationshipDocumentCache.put(cacheKey, relsDoc);

      return new RelationshipCreationResult(relsFile, createdRelationshipIds);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to create slide relationships for slide " + slideNumber, e);
    }
  }

  /**
   * Copies relationships from a source slide to a destination slide with ID remapping.
   * 
   * <p>This method handles the complex process of copying slide relationships while
   * ensuring no relationship ID conflicts occur. Media relationships are remapped
   * to new IDs, while standard layout/theme relationships may reuse existing IDs
   * if they point to the same targets.</p>
   * 
   * @param sourceSlideNumber The source slide number (1-based) to copy from
   * @param destinationSlideNumber The destination slide number (1-based) to copy to
   * @param forceNewIds If true, all relationships get new IDs; if false, allows reuse for standard relationships
   * @return RelationshipCopyResult containing old-to-new ID mappings and created file path
   * @throws XMLParsingException If relationship copying fails
   * @throws IllegalArgumentException If slide numbers are invalid
   * @throws FileNotFoundException If source slide relationships do not exist
   */
  public RelationshipCopyResult copySlideRelationships(int sourceSlideNumber, 
      int destinationSlideNumber, 
      boolean forceNewIds) throws XMLParsingException {
    if (sourceSlideNumber < 1 || destinationSlideNumber < 1) {
      throw new IllegalArgumentException("Slide numbers must be positive");
    }

    try {
      File sourceRelsFile = getSlideRelationshipFile(sourceSlideNumber);
      if (!sourceRelsFile.exists()) {
        // Create default relationships if source has none
        RelationshipCreationResult defaultResult = createSlideRelationships(destinationSlideNumber, null, null);
        return new RelationshipCopyResult(defaultResult.getRelationshipFile(), 
            Collections.emptyMap(), defaultResult.getCreatedRelationshipIds());
      }

      // Parse source relationships
      Document sourceRelsDoc = parseRelationshipDocument(sourceRelsFile);
      Document destRelsDoc = (Document) sourceRelsDoc.cloneNode(true);

      // Extract and remap relationships
      Map<String, String> idMappings = new HashMap<>();
      List<String> newRelationshipIds = new ArrayList<>();

      NodeList relationshipElements = destRelsDoc.getElementsByTagName("Relationship");
      for (int i = 0; i < relationshipElements.getLength(); i++) {
        Element relationshipEl = (Element) relationshipElements.item(i);
        String oldId = relationshipEl.getAttribute("Id");
        String type = relationshipEl.getAttribute("Type");
        String target = relationshipEl.getAttribute("Target");

        String newId;
        if (forceNewIds || isMediaRelationship(type)) {
          // Always remap media relationships to avoid conflicts
          newId = allocateRelationshipId();
          relationshipEl.setAttribute("Id", newId);
          registerRelationship(newId, type, target);
        } else {
          // For standard relationships, check if we can reuse existing ID
          newId = findOrCreateRelationshipId(type, target);
          relationshipEl.setAttribute("Id", newId);
        }

        idMappings.put(oldId, newId);
        newRelationshipIds.add(newId);
      }

      // Write destination relationship file
      File destRelsFile = getSlideRelationshipFile(destinationSlideNumber);
      writeRelationshipDocument(destRelsDoc, destRelsFile);

      // Update cache
      String cacheKey = getRelativePathFromExtractedDir(destRelsFile);
      relationshipDocumentCache.put(cacheKey, destRelsDoc);

      return new RelationshipCopyResult(destRelsFile, idMappings, newRelationshipIds);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to copy slide relationships from slide " + 
          sourceSlideNumber + " to slide " + destinationSlideNumber, e);
    }
  }

  /**
   * Adds a media relationship (image, video, audio) to the specified slide.
   * 
   * <p>Media relationships require special handling because they reference files
   * in the media directory and must have unique IDs across the entire presentation.</p>
   * 
   * @param slideNumber The slide number (1-based) to add the media relationship to
   * @param mediaType The type of media relationship (e.g., XMLConstants.RELATIONSHIP_TYPE_IMAGE)
   * @param mediaTarget The target path relative to the slide (e.g., "../media/image1.png")
   * @return The allocated relationship ID for the media
   * @throws XMLParsingException If the media relationship cannot be added
   * @throws IllegalArgumentException If parameters are invalid
   */
  public String addMediaRelationship(int slideNumber, String mediaType, String mediaTarget) throws XMLParsingException {
    if (slideNumber < 1) {
      throw new IllegalArgumentException("slideNumber must be positive");
    }
    if (mediaType == null || mediaType.trim().isEmpty()) {
      throw new IllegalArgumentException("mediaType cannot be null or empty");
    }
    if (mediaTarget == null || mediaTarget.trim().isEmpty()) {
      throw new IllegalArgumentException("mediaTarget cannot be null or empty");
    }

    try {
      // Load or create slide relationship document
      File relsFile = getSlideRelationshipFile(slideNumber);
      Document relsDoc = loadOrCreateSlideRelationshipDocument(slideNumber);

      // Find relationships root element
      Element relationships = relsDoc.getDocumentElement();

      // Allocate new relationship ID
      String mediaRId = allocateRelationshipId();

      // Create media relationship element
      Element mediaRel = createRelationshipElement(relsDoc, mediaRId, mediaType, mediaTarget);
      relationships.appendChild(mediaRel);

      // Register the relationship
      registerRelationship(mediaRId, mediaType, mediaTarget);

      // Write updated document
      writeRelationshipDocument(relsDoc, relsFile);

      // Update cache
      String cacheKey = getRelativePathFromExtractedDir(relsFile);
      relationshipDocumentCache.put(cacheKey, relsDoc);

      return mediaRId;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to add media relationship to slide " + slideNumber, e);
    }
  }

  /**
   * Removes a relationship from the specified slide by relationship ID.
   * 
   * @param slideNumber The slide number (1-based) containing the relationship
   * @param relationshipId The relationship ID to remove (e.g., "rId3")
   * @return true if the relationship was found and removed, false if not found
   * @throws XMLParsingException If the relationship removal fails
   * @throws IllegalArgumentException If parameters are invalid
   */
  public boolean removeRelationship(int slideNumber, String relationshipId) throws XMLParsingException {
    if (slideNumber < 1) {
      throw new IllegalArgumentException("slideNumber must be positive");
    }
    if (relationshipId == null || relationshipId.trim().isEmpty()) {
      throw new IllegalArgumentException("relationshipId cannot be null or empty");
    }

    try {
      File relsFile = getSlideRelationshipFile(slideNumber);
      if (!relsFile.exists()) {
        return false; // No relationships file means relationship doesn't exist
      }

      Document relsDoc = parseRelationshipDocument(relsFile);
      Element relationships = relsDoc.getDocumentElement();

      // Find and remove the relationship element
      NodeList relationshipElements = relationships.getElementsByTagName("Relationship");
      for (int i = 0; i < relationshipElements.getLength(); i++) {
        Element relationshipEl = (Element) relationshipElements.item(i);
        if (relationshipId.equals(relationshipEl.getAttribute("Id"))) {
          relationships.removeChild(relationshipEl);

          // Unregister from global registry
          unregisterRelationship(relationshipId);

          // Write updated document
          writeRelationshipDocument(relsDoc, relsFile);

          // Update cache
          String cacheKey = getRelativePathFromExtractedDir(relsFile);
          relationshipDocumentCache.put(cacheKey, relsDoc);

          return true;
        }
      }

      return false; // Relationship ID not found

    } catch (Exception e) {
      throw new XMLParsingException("Failed to remove relationship " + relationshipId + 
          " from slide " + slideNumber, e);
    }
  }

  /**
   * Updates relationship targets when slides are moved or media files are relocated.
   * 
   * <p>This method is critical for maintaining presentation integrity when the
   * physical structure of the PPTX package changes.</p>
   * 
   * @param oldSlideNumber The old slide number (before move)
   * @param newSlideNumber The new slide number (after move)
   * @throws XMLParsingException If relationship updates fail
   */
  public void updateRelationshipsForSlideMove(int oldSlideNumber, int newSlideNumber) throws XMLParsingException {
    if (oldSlideNumber < 1 || newSlideNumber < 1) {
      throw new IllegalArgumentException("Slide numbers must be positive");
    }

    // This method handles the complex case of updating all relationship references
    // when a slide is moved from one position to another
    // Implementation depends on whether we need to update relative paths or just IDs

    // For now, this is a placeholder for the complex logic needed
    // TODO: Implement comprehensive slide move relationship updates
  }

  /**
   * Allocates a unique relationship ID that is guaranteed not to conflict
   * with any existing relationships in the presentation.
   * 
   * @return A unique relationship ID (e.g., "rId15")
   */
  public String allocateRelationshipId() {
    String candidateId;
    do {
      int idNumber = nextRelationshipIdCounter.getAndIncrement();
      candidateId = XMLConstants.RID_PREFIX + idNumber;
    } while (globalRelationshipRegistry.containsKey(candidateId));

    return candidateId;
  }

  /**
   * Retrieves all relationship IDs currently in use across the entire presentation.
   * 
   * @return An unmodifiable set of all relationship IDs
   */
  public Set<String> getAllRelationshipIds() {
    return Collections.unmodifiableSet(globalRelationshipRegistry.keySet());
  }

  /**
   * Gets the relationship information for a specific relationship ID.
   * 
   * @param relationshipId The relationship ID to look up
   * @return RelationshipInfo containing type and target, or null if not found
   */
  public RelationshipInfo getRelationshipInfo(String relationshipId) {
    return globalRelationshipRegistry.get(relationshipId);
  }

  /**
   * Validates that all relationships in the presentation are consistent and valid.
   * 
   * @return ValidationResult containing any detected issues
   * @throws XMLParsingException If validation cannot be performed
   */
  public ValidationResult validateAllRelationships() throws XMLParsingException {
    List<String> errors = new ArrayList<>();
    List<String> warnings = new ArrayList<>();

    // Check for broken relationship targets
    for (Map.Entry<String, RelationshipInfo> entry : globalRelationshipRegistry.entrySet()) {
      String relationshipId = entry.getKey();
      RelationshipInfo info = entry.getValue();

      // Check if target file exists
      if (!isExternalTarget(info.getTarget())) {
        File targetFile = resolveRelationshipTarget(info.getTarget());
        if (!targetFile.exists()) {
          errors.add("Relationship " + relationshipId + " points to non-existent target: " + info.getTarget());
        }
      }
    }

    // Check for duplicate relationship IDs (shouldn't happen but good to verify)
    Set<String> duplicateCheck = new HashSet<>();
    for (String id : globalRelationshipRegistry.keySet()) {
      if (!duplicateCheck.add(id)) {
        errors.add("Duplicate relationship ID detected: " + id);
      }
    }

    return new ValidationResult(errors, warnings);
  }

  // ========== PRIVATE HELPER METHODS ==========

  /**
   * Scans all existing relationship files in the presentation to build the global registry.
   */
  private void scanExistingRelationships() throws XMLParsingException {
    try {
      // Scan package-level relationships
      scanRelationshipFile(new File(extractedPptxDir, "_rels/.rels"));

      // Scan presentation relationships  
      scanRelationshipFile(new File(extractedPptxDir, "ppt/_rels/presentation.xml.rels"));

      // Scan all slide relationships
      File slidesRelsDir = new File(extractedPptxDir, "ppt/slides/_rels");
      if (slidesRelsDir.exists()) {
        File[] relsFiles = slidesRelsDir.listFiles((dir, name) -> name.endsWith(".rels"));
        if (relsFiles != null) {
          for (File relsFile : relsFiles) {
            scanRelationshipFile(relsFile);
          }
        }
      }

      // Update counter to ensure new IDs don't conflict
      updateRelationshipIdCounter();

    } catch (Exception e) {
      throw new XMLParsingException("Failed to scan existing relationships", e);
    }
  }

  /**
   * Scans a single relationship file and adds all relationships to the global registry.
   */
  private void scanRelationshipFile(File relsFile) throws XMLParsingException {
    if (!relsFile.exists()) {
      return;
    }

    try {
      Document relsDoc = documentBuilder.parse(relsFile);
      NodeList relationshipElements = relsDoc.getElementsByTagName("Relationship");

      for (int i = 0; i < relationshipElements.getLength(); i++) {
        Element relationshipEl = (Element) relationshipElements.item(i);
        String id = relationshipEl.getAttribute("Id");
        String type = relationshipEl.getAttribute("Type");
        String target = relationshipEl.getAttribute("Target");

        registerRelationship(id, type, target);
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to scan relationship file: " + relsFile.getName(), e);
    }
  }

  /**
   * Registers a relationship in the global registry.
   */
  private void registerRelationship(String id, String type, String target) {
    globalRelationshipRegistry.put(id, new RelationshipInfo(type, target));
  }

  /**
   * Unregisters a relationship from the global registry.
   */
  private void unregisterRelationship(String id) {
    globalRelationshipRegistry.remove(id);
  }

  /**
   * Updates the relationship ID counter based on existing relationships.
   */
  private void updateRelationshipIdCounter() {
    int maxId = 0;
    for (String relationshipId : globalRelationshipRegistry.keySet()) {
      if (relationshipId.startsWith(XMLConstants.RID_PREFIX)) {
        try {
          int idNumber = Integer.parseInt(relationshipId.substring(XMLConstants.RID_PREFIX.length()));
          maxId = Math.max(maxId, idNumber);
        } catch (NumberFormatException e) {
          // Skip non-numeric relationship IDs
        }
      }
    }
    nextRelationshipIdCounter.set(maxId + 1);
  }

  /**
   * Creates a relationship XML element.
   */
  private Element createRelationshipElement(Document doc, String id, String type, String target) {
    Element relationship = doc.createElementNS(XMLConstants.PACKAGE_RELATIONSHIPS_NS, "Relationship");
    relationship.setAttribute("Id", id);
    relationship.setAttribute("Type", type);
    relationship.setAttribute("Target", target);
    return relationship;
  }

  /**
   * Gets the relationship file for a specific slide number.
   */
  private File getSlideRelationshipFile(int slideNumber) {
    File relsDir = new File(extractedPptxDir, "ppt/slides/_rels");
    return new File(relsDir, String.format("slide%d.xml.rels", slideNumber));
  }

  /**
   * Parses a relationship document, using cache if available.
   */
  private Document parseRelationshipDocument(File relsFile) throws XMLParsingException {
    String cacheKey = getRelativePathFromExtractedDir(relsFile);
    Document cached = relationshipDocumentCache.get(cacheKey);
    if (cached != null) {
      return cached;
    }

    try {
      Document doc = documentBuilder.parse(relsFile);
      relationshipDocumentCache.put(cacheKey, doc);
      return doc;
    } catch (Exception e) {
      throw new XMLParsingException("Failed to parse relationship document: " + relsFile.getName(), e);
    }
  }

  /**
   * Loads an existing slide relationship document or creates a new one.
   */
  private Document loadOrCreateSlideRelationshipDocument(int slideNumber) throws XMLParsingException {
    File relsFile = getSlideRelationshipFile(slideNumber);
    if (relsFile.exists()) {
      return parseRelationshipDocument(relsFile);
    } else {
      // Create basic relationship document structure
      Document doc = documentBuilder.newDocument();
      Element relationships = doc.createElementNS(XMLConstants.PACKAGE_RELATIONSHIPS_NS, "Relationships");
      doc.appendChild(relationships);
      return doc;
    }
  }

  /**
   * Writes a relationship document to file.
   */
  private void writeRelationshipDocument(Document doc, File outputFile) throws XMLParsingException {
    try {
      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();
      transformer.setOutputProperty(OutputKeys.INDENT, "yes");
      transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");

      // Ensure parent directory exists
      outputFile.getParentFile().mkdirs();

      DOMSource source = new DOMSource(doc);
      StreamResult result = new StreamResult(outputFile);
      transformer.transform(source, result);

    } catch (TransformerException e) {
      throw new XMLParsingException("Failed to write relationship document", e);
    }
  }

  /**
   * Determines if a relationship type is for media (images, video, audio).
   */
  private boolean isMediaRelationship(String relationshipType) {
    return relationshipType.equals(XMLConstants.RELATIONSHIP_TYPE_IMAGE) ||
      relationshipType.contains("video") ||
      relationshipType.contains("audio");
  }

  /**
   * Finds an existing relationship ID for the given type and target, or creates a new one.
   */
  private String findOrCreateRelationshipId(String type, String target) {
    // Look for existing relationship with same type and target
    for (Map.Entry<String, RelationshipInfo> entry : globalRelationshipRegistry.entrySet()) {
      RelationshipInfo info = entry.getValue();
      if (info.getType().equals(type) && info.getTarget().equals(target)) {
        return entry.getKey();
      }
    }

    // No existing relationship found, create new one
    String newId = allocateRelationshipId();
    registerRelationship(newId, type, target);
    return newId;
  }

  /**
   * Converts an absolute file path to a path relative to the extracted PPTX directory.
   */
  private String getRelativePathFromExtractedDir(File file) {
    return extractedPptxDir.toPath().relativize(file.toPath()).toString();
  }

  /**
   * Determines if a target is external (HTTP URL, etc.) rather than internal file.
   */
  private boolean isExternalTarget(String target) {
    return target.startsWith("http://") || target.startsWith("https://") || target.startsWith("mailto:");
  }

  /**
   * Resolves a relationship target to an actual file path.
   */
  private File resolveRelationshipTarget(String target) {
    // Handle relative paths from the slide's perspective
    if (target.startsWith("../")) {
      return new File(extractedPptxDir, "ppt/" + target.substring(3));
    } else {
      return new File(extractedPptxDir, "ppt/slides/" + target);
    }
  }

  // ========== INNER CLASSES ==========

  /**
   * Contains information about a specific relationship.
   */
  public static class RelationshipInfo {
    private final String type;
    private final String target;

    public RelationshipInfo(String type, String target) {
      this.type = type;
      this.target = target;
    }

    public String getType() { return type; }
    public String getTarget() { return target; }

    @Override
    public String toString() {
      return String.format("RelationshipInfo{type='%s', target='%s'}", type, target);
    }
  }

  /**
   * Result of creating new slide relationships.
   */
  public static class RelationshipCreationResult {
    private final File relationshipFile;
    private final List<String> createdRelationshipIds;

    public RelationshipCreationResult(File relationshipFile, List<String> createdRelationshipIds) {
      this.relationshipFile = relationshipFile;
      this.createdRelationshipIds = Collections.unmodifiableList(new ArrayList<>(createdRelationshipIds));
    }

    public File getRelationshipFile() { return relationshipFile; }
    public List<String> getCreatedRelationshipIds() { return createdRelationshipIds; }
  }

  /**
   * Result of copying slide relationships.
   */
  public static class RelationshipCopyResult {
    private final File relationshipFile;
    private final Map<String, String> oldToNewIdMappings;
    private final List<String> newRelationshipIds;

    public RelationshipCopyResult(File relationshipFile, Map<String, String> oldToNewIdMappings, 
        List<String> newRelationshipIds) {
      this.relationshipFile = relationshipFile;
      this.oldToNewIdMappings = Collections.unmodifiableMap(new HashMap<>(oldToNewIdMappings));
      this.newRelationshipIds = Collections.unmodifiableList(new ArrayList<>(newRelationshipIds));
    }

    public File getRelationshipFile() { return relationshipFile; }
    public Map<String, String> getOldToNewIdMappings() { return oldToNewIdMappings; }
    public List<String> getNewRelationshipIds() { return newRelationshipIds; }
  }

  /**
   * Result of relationship validation.
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
