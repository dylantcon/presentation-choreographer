package com.presentationchoreographer.xml.writers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.*;
import javax.xml.namespace.NamespaceContext;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.*;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive slide creation system supporting blank slides, 
 * slide duplication, template-based creation, and complete PPTX relationship management
 */
public class SlideCreator {

  private final File extractedPptxDir;
  private final DocumentBuilder documentBuilder;
  private final XPath xpath;
  private final NamespaceContext namespaceContext;
  private final RelationshipManager relationshipManager;
  private final SPIDManager spidManager;

  public SlideCreator(File extractedPptxDir) throws XMLParsingException {
    this.extractedPptxDir = extractedPptxDir;
    this.namespaceContext = com.presentationchoreographer.utils.XMLConstants.createNamespaceContext();

    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();

      XPathFactory xpathFactory = XPathFactory.newInstance();
      this.xpath = xpathFactory.newXPath();
      this.xpath.setNamespaceContext(namespaceContext);

      // Initialize relationship manager for comprehensive relationship handling
      this.relationshipManager = new RelationshipManager(extractedPptxDir);

      // Initialize SPID manager for global shape ID management
      this.spidManager = new SPIDManager(extractedPptxDir);

    } catch (ParserConfigurationException e) {
      throw new XMLParsingException("Failed to initialize slide creator", e);
    }
  }

  /**
   * Insert a new blank slide at the specified position
   * @param insertPosition Position to insert (1-based, inserts BEFORE this position)
   * @param slideTitle Title for the new slide
   * @return The slide number of the newly created slide
   */
  public int insertBlankSlide(int insertPosition, String slideTitle) throws XMLParsingException {
    try {
      System.out.println("INSERTING BLANK SLIDE AT POSITION " + insertPosition);

      // Step 1: Rename subsequent slides to make space
      renameSubsequentSlides(insertPosition);

      // Step 2: Create new blank slide document
      Document newSlide = createBlankSlideDocument(slideTitle);

      // Step 3: Write new slide file
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      slidesDir.mkdirs();
      File newSlideFile = new File(slidesDir, newSlideFileName);
      writeDocument(newSlide, newSlideFile);
      System.out.println("  ✓ Created slide file: " + newSlideFileName);

      // Step 4: Create relationships file for new slide
      createSlideRelationships(insertPosition);

      // Step 5: Update presentation.xml to include new slide
      updatePresentationXml(insertPosition);

      // Step 6: Update content types registry
      updateContentTypes();

      System.out.println("  ✓ Blank slide insertion complete");
      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to insert blank slide at position " + insertPosition, e);
    }
  }

  /**
   * Insert a new slide by copying an existing slide
   * @param insertPosition Position to insert the new slide
   * @param sourceSlideNumber Slide number to copy from
   * @param newSlideTitle New title for the copied slide
   * @return The slide number of the newly created slide
   */
  public int insertCopiedSlide(int insertPosition, int sourceSlideNumber, String newSlideTitle) throws XMLParsingException {
    try {
      System.out.println("COPYING SLIDE " + sourceSlideNumber + " TO POSITION " + insertPosition);

      // Step 1: Rename subsequent slides to make space
      renameSubsequentSlides(insertPosition);

      // Step 2: Load and parse source slide
      File sourceSlideFile = new File(extractedPptxDir, String.format("ppt/slides/slide%d.xml", sourceSlideNumber));
      if (!sourceSlideFile.exists()) {
        throw new XMLParsingException("Source slide " + sourceSlideNumber + " not found");
      }

      Document sourceSlide = documentBuilder.parse(sourceSlideFile);

      // Step 3: Modify copied slide (update title, regenerate SPIDs)
      Document modifiedSlide = modifySlideForCopy(sourceSlide, newSlideTitle);

      // Step 4: Write copied slide
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      File newSlideFile = new File(slidesDir, newSlideFileName);
      writeDocument(modifiedSlide, newSlideFile);
      System.out.println("  ✓ Created copied slide file: " + newSlideFileName);

      // Step 5: Copy and update relationships
      copySlideRelationships(sourceSlideNumber, insertPosition);

      // Step 6: Update presentation.xml
      updatePresentationXml(insertPosition);

      // Step 7: Update content types
      updateContentTypes();

      System.out.println("  ✓ Slide copying complete");
      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to copy slide " + sourceSlideNumber + " to position " + insertPosition, e);
    }
  }

  /**
   * Insert a new slide from a predefined template
   * @param insertPosition Position to insert the new slide
   * @param template The slide template to use
   * @param templateData Data to populate the template
   * @return The slide number of the newly created slide
   */
  public int insertTemplateSlide(int insertPosition, SlideTemplate template, TemplateData templateData) throws XMLParsingException {
    try {
      System.out.println("INSERTING TEMPLATE SLIDE AT POSITION " + insertPosition);

      // Step 1: Rename subsequent slides to make space
      renameSubsequentSlides(insertPosition);

      // Step 2: Create slide from template
      Document newSlide = template.createSlideDocument(templateData);

      // Step 3: Write template slide
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      File newSlideFile = new File(slidesDir, newSlideFileName);
      writeDocument(newSlide, newSlideFile);
      System.out.println("  ✓ Created template slide file: " + newSlideFileName);

      // Step 4: Create relationships for template slide
      createSlideRelationships(insertPosition);

      // Step 5: Update presentation.xml
      updatePresentationXml(insertPosition);

      // Step 6: Update content types
      updateContentTypes();

      System.out.println("  ✓ Template slide insertion complete");
      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to insert template slide at position " + insertPosition, e);
    }
  }

  /**
   * Create relationships file for a new slide using RelationshipManager
   */
  private void createSlideRelationships(int slideNumber) throws XMLParsingException {
    try {
      // Delegate to RelationshipManager for comprehensive relationship handling
      RelationshipManager.RelationshipCreationResult result = 
        relationshipManager.createSlideRelationships(slideNumber, null, null);

      System.out.println("  ✓ Created relationships file: " + result.getRelationshipFile().getName());
      System.out.println("    → Created " + result.getCreatedRelationshipIds().size() + " relationships: " + 
          result.getCreatedRelationshipIds());

    } catch (Exception e) {
      throw new XMLParsingException("Failed to create slide relationships for slide " + slideNumber, e);
    }
  }

  /**
   * Copy relationships from source slide to new slide using RelationshipManager
   */
  private void copySlideRelationships(int sourceSlideNumber, int newSlideNumber) throws XMLParsingException {
    try {
      // Use RelationshipManager for intelligent relationship copying with ID remapping
      RelationshipManager.RelationshipCopyResult result = 
        relationshipManager.copySlideRelationships(sourceSlideNumber, newSlideNumber, false);

      System.out.println("  ✓ Copied relationships file: " + result.getRelationshipFile().getName());
      System.out.println("    → ID mappings: " + result.getOldToNewIdMappings().size() + " relationships remapped");
      System.out.println("    → New relationship IDs: " + result.getNewRelationshipIds());

    } catch (Exception e) {
      throw new XMLParsingException("Failed to copy slide relationships from slide " + 
          sourceSlideNumber + " to slide " + newSlideNumber, e);
    }
  }

  /**
   * Update presentation.xml to include the new slide in the slide list
   */
  private void updatePresentationXml(int newSlideNumber) throws XMLParsingException {
    try {
      // Load presentation.xml
      File presentationFile = new File(extractedPptxDir, "ppt/presentation.xml");
      if (!presentationFile.exists()) {
        throw new XMLParsingException("presentation.xml not found");
      }

      Document presentationDoc = documentBuilder.parse(presentationFile);

      Element sldIdLst = (Element) xpath.evaluate(
          com.presentationchoreographer.utils.XMLConstants.XPATH_SLIDE_ID_LIST, 
          presentationDoc, XPathConstants.NODE);
      if (sldIdLst == null) {
        throw new XMLParsingException("Slide ID list not found in presentation.xml");
      }

      // Calculate new slide ID (find highest existing + 1)
      NodeList existingSlides = (NodeList) xpath.evaluate(
          com.presentationchoreographer.utils.XMLConstants.XPATH_SLIDE_ID_ELEMENTS, 
          sldIdLst, XPathConstants.NODESET);
      int maxId = com.presentationchoreographer.utils.XMLConstants.DEFAULT_SLIDE_ID_START - 1;

      for (int i = 0; i < existingSlides.getLength(); i++) {
        Element slide = (Element) existingSlides.item(i);
        int id = Integer.parseInt(slide.getAttribute("id"));
        maxId = Math.max(maxId, id);
      }

      int newSlideId = maxId + 1;

      // Find insertion point (before the slide that will be shifted)
      Element insertionPoint = null;
      for (int i = 0; i < existingSlides.getLength(); i++) {
        Element slide = (Element) existingSlides.item(i);
        String rId = slide.getAttribute("r:id");
        int slideNum = extractSlideNumberFromRId(rId);

        if (slideNum >= newSlideNumber) {
          insertionPoint = slide;
          break;
        }
      }

      // Create new slide ID element
      Element newSldId = presentationDoc.createElementNS(
          com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:sldId");
      newSldId.setAttribute("id", String.valueOf(newSlideId));
      newSldId.setAttributeNS(
          com.presentationchoreographer.utils.XMLConstants.RELATIONSHIPS_NS, 
          "r:id", 
          com.presentationchoreographer.utils.XMLConstants.RID_PREFIX + (newSlideNumber + 1)
          );

      // Insert into the list
      if (insertionPoint != null) {
        sldIdLst.insertBefore(newSldId, insertionPoint);
      } else {
        sldIdLst.appendChild(newSldId);
      }

      // Update subsequent rId attributes to maintain sequence
      updateSubsequentRIds(sldIdLst, newSlideNumber);

      // Write updated presentation.xml
      writeDocument(presentationDoc, presentationFile);

      System.out.println("  ✓ Updated presentation.xml with new slide ID: " + newSlideId);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to update presentation.xml", e);
    }
  }

  /**
   * Update content types registry using RelationshipManager validation
   */
  private void updateContentTypes() throws XMLParsingException {
    try {
      // Use RelationshipManager to validate and ensure content types are consistent
      RelationshipManager.ValidationResult validation = relationshipManager.validateAllRelationships();

      if (validation.hasErrors()) {
        System.out.println("  ⚠ Relationship validation found issues:");
        for (String error : validation.getErrors()) {
          System.out.println("    → " + error);
        }
      }

      if (validation.hasWarnings()) {
        System.out.println("  ⚠ Relationship validation warnings:");
        for (String warning : validation.getWarnings()) {
          System.out.println("    → " + warning);
        }
      }

      // Load [Content_Types].xml for basic slide content type verification
      File contentTypesFile = new File(extractedPptxDir, "[Content_Types].xml");
      if (!contentTypesFile.exists()) {
        throw new XMLParsingException("[Content_Types].xml not found");
      }

      Document contentTypesDoc = documentBuilder.parse(contentTypesFile);

      // Check if slide content type is already registered
      Element typesRoot = contentTypesDoc.getDocumentElement();
      String slideContentTypeQuery = String.format("//Override[@ContentType='%s']", 
          com.presentationchoreographer.utils.XMLConstants.CONTENT_TYPE_SLIDE);
      NodeList overrides = (NodeList) xpath.evaluate(slideContentTypeQuery, 
          contentTypesDoc, XPathConstants.NODESET);

      if (overrides.getLength() == 0) {
        // Add slide content type override if missing
        Element override = contentTypesDoc.createElement("Override");
        override.setAttribute("PartName", "/ppt/slides/slide1.xml");
        override.setAttribute("ContentType", com.presentationchoreographer.utils.XMLConstants.CONTENT_TYPE_SLIDE);
        typesRoot.appendChild(override);

        writeDocument(contentTypesDoc, contentTypesFile);
        System.out.println("  ✓ Added slide content type to [Content_Types].xml");
      } else {
        System.out.println("  ✓ Slide content type already registered in [Content_Types].xml");
      }

    } catch (Exception e) {
      throw new XMLParsingException("Failed to update content types", e);
    }
  }

  /**
   * Rename all slides at and after the insertion position
   * This creates space for the new slide to be inserted
   */
  private void renameSubsequentSlides(int insertPosition) throws XMLParsingException {
    try {
      File slidesDir = new File(extractedPptxDir, "ppt/slides");
      File relsDir = new File(extractedPptxDir, "ppt/slides/_rels");

      // Get list of existing slide numbers
      List<Integer> existingSlides = getExistingSlideNumbers();

      // Sort in descending order to avoid naming conflicts
      existingSlides.sort(Collections.reverseOrder());

      // Rename slides from highest to lowest
      for (int slideNum : existingSlides) {
        if (slideNum >= insertPosition) {
          int newSlideNum = slideNum + 1;

          // Rename slide XML file
          File oldSlideFile = new File(slidesDir, String.format("slide%d.xml", slideNum));
          File newSlideFile = new File(slidesDir, String.format("slide%d.xml", newSlideNum));
          if (oldSlideFile.exists()) {
            Files.move(oldSlideFile.toPath(), newSlideFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
          }

          // Rename relationships file
          File oldRelsFile = new File(relsDir, String.format("slide%d.xml.rels", slideNum));
          File newRelsFile = new File(relsDir, String.format("slide%d.xml.rels", newSlideNum));
          if (oldRelsFile.exists()) {
            Files.move(oldRelsFile.toPath(), newRelsFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
          }
        }
      }

    } catch (IOException e) {
      throw new XMLParsingException("Failed to rename subsequent slides", e);
    }
  }

  /**
   * Get list of existing slide numbers from the slides directory
   */
  private List<Integer> getExistingSlideNumbers() {
    List<Integer> slideNumbers = new ArrayList<>();
    File slidesDir = new File(extractedPptxDir, "ppt/slides");

    if (slidesDir.exists()) {
      File[] slideFiles = slidesDir.listFiles((dir, name) -> name.matches("slide\\d+\\.xml"));
      if (slideFiles != null) {
        for (File file : slideFiles) {
          String fileName = file.getName();
          String numberStr = fileName.substring(5, fileName.lastIndexOf(".xml"));
          try {
            slideNumbers.add(Integer.parseInt(numberStr));
          } catch (NumberFormatException e) {
            // Skip invalid slide numbers
          }
        }
      }
    }

    return slideNumbers;
  }

  /**
   * Create a blank slide document with minimal OOXML structure
   */
  private Document createBlankSlideDocument(String slideTitle) throws XMLParsingException {
    try {
      Document document = documentBuilder.newDocument();

      // Create root slide element with namespaces
      Element slide = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:sld");
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:a", com.presentationchoreographer.utils.XMLConstants.DRAWING_NS);
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", com.presentationchoreographer.utils.XMLConstants.RELATIONSHIPS_NS);
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:p", com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS);
      document.appendChild(slide);

      // Create common slide data
      Element cSld = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:cSld");
      slide.appendChild(cSld);

      // Create shape tree
      Element spTree = createBlankShapeTree(document, slideTitle);
      cSld.appendChild(spTree);

      // Create color map override
      Element clrMapOvr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:clrMapOvr");
      slide.appendChild(clrMapOvr);

      Element masterClrMapping = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.DRAWING_NS, "a:masterClrMapping");
      clrMapOvr.appendChild(masterClrMapping);

      return document;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to create blank slide document", e);
    }
  }

  /**
   * Create a minimal shape tree with just a title placeholder
   */
  private Element createBlankShapeTree(Document document, String slideTitle) {
    Element spTree = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:spTree");

    // Add group shape properties
    Element nvGrpSpPr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:nvGrpSpPr");
    spTree.appendChild(nvGrpSpPr);

    Element cNvPr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:cNvPr");
    cNvPr.setAttribute("id", "1");
    cNvPr.setAttribute("name", "");
    nvGrpSpPr.appendChild(cNvPr);

    Element cNvGrpSpPr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:cNvGrpSpPr");
    nvGrpSpPr.appendChild(cNvGrpSpPr);

    Element nvPr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:nvPr");
    nvGrpSpPr.appendChild(nvPr);

    // Add group shape properties
    Element grpSpPr = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.PRESENTATION_NS, "p:grpSpPr");
    spTree.appendChild(grpSpPr);

    Element xfrm = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.DRAWING_NS, "a:xfrm");
    grpSpPr.appendChild(xfrm);

    // Add transform elements
    addTransformElement(document, xfrm, "a:off", "0", "0");
    addTransformElement(document, xfrm, "a:ext", "0", "0");
    addTransformElement(document, xfrm, "a:chOff", "0", "0");
    addTransformElement(document, xfrm, "a:chExt", "0", "0");

    // Add basic slide structure - minimal implementation for now
    // Future enhancement: Add title shape if provided

    return spTree;
  }

  /**
   * Helper method to add transform elements
   */
  private void addTransformElement(Document document, Element parent, String elementName, String x, String y) {
    Element element = document.createElementNS(com.presentationchoreographer.utils.XMLConstants.DRAWING_NS, elementName);
    element.setAttribute("x", x);
    element.setAttribute("y", y);
    parent.appendChild(element);
  }

  /**
   * Modify a copied slide to avoid SPID conflicts and update content
   */
  private Document modifySlideForCopy(Document sourceSlide, String newTitle) throws XMLParsingException {
    // Clone the document
    Document copiedSlide = (Document) sourceSlide.cloneNode(true);

    // Regenerate all SPIDs to avoid conflicts
    SPIDManager.SPIDRegenerationResult spidResult = regenerateSpids(copiedSlide);
    System.out.println("    → SPID regeneration: " + spidResult.getShapesProcessed() + 
        " shapes, " + spidResult.getAnimationsUpdated() + " animations updated");

    // Update title if provided
    if (newTitle != null && !newTitle.trim().isEmpty()) {
      updateSlideTitle(copiedSlide, newTitle);
    }

    return copiedSlide;
  }

  /**
   * Regenerates all SPIDs in a slide document using SPIDManager.
   * 
   * @param slideDocument The slide document to regenerate SPIDs for
   * @return SPIDRegenerationResult containing mapping and statistics
   * @throws XMLParsingException If SPID regeneration fails
   */
  private SPIDManager.SPIDRegenerationResult regenerateSpids(Document slideDocument) throws XMLParsingException {
    try {
      return spidManager.regenerateSpids(slideDocument, 0); // 0 = copied slide (temporary)
    } catch (Exception e) {
      throw new XMLParsingException("Failed to regenerate SPIDs in copied slide", e);
    }
  }

  /**
   * Updates the title of a slide (placeholder implementation).
   * 
   * @param slide The slide document to update
   * @param newTitle The new title text
   * @throws XMLParsingException If title update fails
   */
  private void updateSlideTitle(Document slide, String newTitle) throws XMLParsingException {
    // TODO: Implement title shape detection and text update
    // This is a future enhancement for slide copying
    System.out.println("    → Title update requested: \"" + newTitle + "\" (not yet implemented)");
  }

  /**
   * Helper method to extract slide number from rId (e.g., "rId3" -> 2 for slide2)
   */
  private int extractSlideNumberFromRId(String rId) {
    try {
      int ridNum = Integer.parseInt(rId.substring(com.presentationchoreographer.utils.XMLConstants.RID_PREFIX.length()));
      return ridNum - 1; // rId2 corresponds to slide1, etc.
    } catch (Exception e) {
      return 1; // Default fallback
    }
  }

  /**
   * Update subsequent rId attributes when a slide is inserted
   */
  private void updateSubsequentRIds(Element sldIdLst, int insertedSlideNumber) throws XPathExpressionException {
    NodeList slides = (NodeList) xpath.evaluate(
        com.presentationchoreographer.utils.XMLConstants.XPATH_SLIDE_ID_ELEMENTS, 
        sldIdLst, XPathConstants.NODESET);

    for (int i = 0; i < slides.getLength(); i++) {
      Element slide = (Element) slides.item(i);
      String rId = slide.getAttribute("r:id");
      int slideNum = extractSlideNumberFromRId(rId);

      // Update rIds for slides that were shifted
      if (slideNum > insertedSlideNumber) {
        int newRIdNum = slideNum + 2; // +1 for shift, +1 for rId offset
        slide.setAttribute("r:id", com.presentationchoreographer.utils.XMLConstants.RID_PREFIX + newRIdNum);
      }
    }
  }

  /**
   * Add a media relationship (image, video, audio) to the specified slide
   * 
   * @param slideNumber The slide number (1-based) to add the media relationship to
   * @param mediaType The type of media relationship (e.g., XMLConstants.RELATIONSHIP_TYPE_IMAGE)
   * @param mediaTarget The target path relative to the slide (e.g., "../media/image1.png")
   * @return The allocated relationship ID for the media
   * @throws XMLParsingException If the media relationship cannot be added
   */
  public String addMediaRelationship(int slideNumber, String mediaType, String mediaTarget) throws XMLParsingException {
    try {
      String mediaRId = relationshipManager.addMediaRelationship(slideNumber, mediaType, mediaTarget);
      System.out.println("  ✓ Added media relationship: " + mediaRId + " → " + mediaTarget);
      return mediaRId;
    } catch (Exception e) {
      throw new XMLParsingException("Failed to add media relationship to slide " + slideNumber, e);
    }
  }

  /**
   * Get access to the underlying RelationshipManager for advanced relationship operations
   * 
   * @return The RelationshipManager instance
   */
  public RelationshipManager getRelationshipManager() {
    return relationshipManager;
  }

  /**
   * Get access to the underlying SPIDManager for advanced shape ID operations
   * 
   * @return The SPIDManager instance
   */
  public SPIDManager getSPIDManager() {
    return spidManager;
  }

  /**
   * Allocates a unique SPID for new shapes, guaranteed not to conflict with existing shapes.
   * 
   * @return A unique Shape ID
   */
  public int allocateUniqueSpid() {
    return spidManager.allocateUniqueSpid();
  }

  /**
   * Validates that all SPIDs and relationships in the presentation are consistent.
   * 
   * @return ValidationSummary containing any detected issues
   * @throws XMLParsingException If validation cannot be performed
   */
  public ValidationSummary validatePresentation() throws XMLParsingException {
    try {
      // Validate both SPID uniqueness and relationship consistency
      SPIDManager.ValidationResult spidValidation = spidManager.validateSpidUniqueness();
      RelationshipManager.ValidationResult relationshipValidation = relationshipManager.validateAllRelationships();

      List<String> allErrors = new ArrayList<>();
      List<String> allWarnings = new ArrayList<>();

      allErrors.addAll(spidValidation.getErrors());
      allErrors.addAll(relationshipValidation.getErrors());

      allWarnings.addAll(spidValidation.getWarnings());
      allWarnings.addAll(relationshipValidation.getWarnings());

      return new ValidationSummary(allErrors, allWarnings);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to validate presentation", e);
    }
  }

  /**
   * Write a document to a file with proper formatting
   */
  private void writeDocument(Document document, File outputFile) throws XMLParsingException {
    try {
      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();

      transformer.setOutputProperty(OutputKeys.INDENT, "yes");
      transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
      transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

      DOMSource source = new DOMSource(document);
      StreamResult result = new StreamResult(outputFile);

      // Ensure parent directories exist
      outputFile.getParentFile().mkdirs();

      transformer.transform(source, result);

    } catch (TransformerException e) {
      throw new XMLParsingException("Failed to write document to file", e);
    }
  }

  // ========== INNER CLASSES ==========

  /**
   * Combined validation result for both SPIDs and relationships.
   */
  public static class ValidationSummary {
    private final List<String> errors;
    private final List<String> warnings;

    public ValidationSummary(List<String> errors, List<String> warnings) {
      this.errors = Collections.unmodifiableList(new ArrayList<>(errors));
      this.warnings = Collections.unmodifiableList(new ArrayList<>(warnings));
    }

    public List<String> getErrors() { return errors; }
    public List<String> getWarnings() { return warnings; }
    public boolean hasErrors() { return !errors.isEmpty(); }
    public boolean hasWarnings() { return !warnings.isEmpty(); }
    public boolean isValid() { return errors.isEmpty(); }

    @Override
    public String toString() {
      return String.format("ValidationSummary{errors=%d, warnings=%d}", 
          errors.size(), warnings.size());
    }
  }
}

/**
 * Interface for slide templates - moved to separate file in production
 */
interface SlideTemplate {
  Document createSlideDocument(TemplateData data) throws XMLParsingException;
  String getTemplateName();
}

/**
 * Data container for template population - moved to separate file in production
 */
class TemplateData {
  private final Map<String, Object> data = new HashMap<>();

  public void put(String key, Object value) {
    data.put(key, value);
  }

  public Object get(String key) {
    return data.get(key);
  }

  public String getString(String key) {
    Object value = data.get(key);
    return value != null ? value.toString() : null;
  }

  @SuppressWarnings("unchecked")
  public List<String> getStringList(String key) {
    Object value = data.get(key);
    if (value instanceof List) {
      return (List<String>) value;
    }
    return new ArrayList<>();
  }
}
