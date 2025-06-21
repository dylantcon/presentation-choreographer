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
import java.nio.file.Files;
import java.util.*;
import java.util.regex.Pattern;
import java.util.regex.Matcher;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.*;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive slide creation system supporting blank slides, 
 * slide duplication, template-based creation, and complete PPTX relationship management.
 * 
 * REFACTORED: Uses string-based approach for presentation.xml manipulation to avoid
 * XML namespace serialization issues with Java's built-in transformer.
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
    this.namespaceContext = XMLConstants.createNamespaceContext();

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
   * Update presentation.xml to include the new slide in the slide list.
   * REFACTORED: Uses string-based approach to avoid XML namespace serialization issues.
   */
  private void updatePresentationXml(int newSlideNumber) throws XMLParsingException {
    try {
      // Load presentation.xml
      File presentationFile = new File(extractedPptxDir, "ppt/presentation.xml");
      if (!presentationFile.exists()) {
        throw new XMLParsingException("presentation.xml not found");
      }

      // Read the entire file as string
      String presentationContent = Files.readString(presentationFile.toPath());

      // Calculate new slide ID
      int newSlideId = calculateNextSlideId(presentationContent);

      // Calculate rId for this slide (slide1 = rId2, slide2 = rId3, etc.)
      String newRId = XMLConstants.RID_PREFIX + (newSlideNumber + 1);

      // Create new sldId element as string with proper namespaces
      String newSldIdElement = String.format(
          "    <p:sldId id=\"%d\" r:id=\"%s\"/>", 
          newSlideId, 
          newRId
          );

      // Find insertion point in the sldIdLst
      String updatedContent = insertSlideIdElement(presentationContent, newSldIdElement, newSlideNumber);

      // Update subsequent rId attributes
      updatedContent = updateSubsequentRIdsInString(updatedContent, newSlideNumber);

      // Write back to file
      Files.writeString(presentationFile.toPath(), updatedContent);

      System.out.println("  ✓ Updated presentation.xml with new slide ID: " + newSlideId);

    } catch (Exception e) {
      throw new XMLParsingException("Failed to update presentation.xml", e);
    }
  }

  /**
   * Calculate the next available slide ID from presentation content
   */
  private int calculateNextSlideId(String presentationContent) {
    int maxId = XMLConstants.DEFAULT_SLIDE_ID_START - 1;

    // Find all existing slide IDs using regex
    Pattern slideIdPattern = Pattern.compile("<p:sldId\\s+id=\"(\\d+)\"");
    Matcher matcher = slideIdPattern.matcher(presentationContent);

    while (matcher.find()) {
      int id = Integer.parseInt(matcher.group(1));
      maxId = Math.max(maxId, id);
    }

    return maxId + 1;
  }

  /**
   * Insert new slide ID element at the correct position
   */
  private String insertSlideIdElement(String presentationContent, String newSldIdElement, int newSlideNumber) {
    // Find the sldIdLst section
    Pattern sldIdLstPattern = Pattern.compile("(<p:sldIdLst>)(.*?)(</p:sldIdLst>)", Pattern.DOTALL);
    Matcher matcher = sldIdLstPattern.matcher(presentationContent);

    if (!matcher.find()) {
      throw new RuntimeException("Could not find sldIdLst in presentation.xml");
    }

    String beforeList = matcher.group(1);
    String listContent = matcher.group(2);
    String afterList = matcher.group(3);

    // Find all existing slide elements and their positions
    List<SlideIdInfo> existingSlides = parseExistingSlideIds(listContent);

    // Find insertion point
    StringBuilder newListContent = new StringBuilder();
    boolean inserted = false;

    for (int i = 0; i < existingSlides.size(); i++) {
      SlideIdInfo slideInfo = existingSlides.get(i);

      // If this slide should come after our new slide, insert here
      if (!inserted && slideInfo.slideNumber >= newSlideNumber) {
        newListContent.append("\n").append(newSldIdElement);
        inserted = true;
      }

      // Add the existing slide (with potentially updated rId)
      newListContent.append("\n").append(slideInfo.originalElement);
    }

    // If we haven't inserted yet, add at the end
    if (!inserted) {
      newListContent.append("\n").append(newSldIdElement);
    }

    return presentationContent.replace(matcher.group(0), 
        beforeList + newListContent.toString() + "\n" + afterList);
  }

  /**
   * Parse existing slide IDs from sldIdLst content
   */
  private List<SlideIdInfo> parseExistingSlideIds(String listContent) {
    List<SlideIdInfo> slides = new ArrayList<>();

    Pattern slidePattern = Pattern.compile("<p:sldId\\s+id=\"(\\d+)\"\\s+r:id=\"([^\"]+)\"/>");
    Matcher matcher = slidePattern.matcher(listContent);

    while (matcher.find()) {
      String fullElement = matcher.group(0);
      int slideId = Integer.parseInt(matcher.group(1));
      String rId = matcher.group(2);
      int slideNumber = extractSlideNumberFromRId(rId);

      slides.add(new SlideIdInfo(slideNumber, slideId, rId, fullElement));
    }

    return slides;
  }

  /**
   * Update subsequent rId attributes when a slide is inserted
   */
  private String updateSubsequentRIdsInString(String content, int insertedSlideNumber) {
    Pattern slidePattern = Pattern.compile("<p:sldId\\s+id=\"(\\d+)\"\\s+r:id=\"rId(\\d+)\"/>");
    Matcher matcher = slidePattern.matcher(content);

    StringBuffer result = new StringBuffer();

    while (matcher.find()) {
      String slideId = matcher.group(1);
      int rIdNumber = Integer.parseInt(matcher.group(2));
      int slideNumber = rIdNumber - 1; // rId2 = slide1, etc.

      // If this slide comes after our inserted slide, increment its rId
      if (slideNumber > insertedSlideNumber) {
        String newRId = XMLConstants.RID_PREFIX + (rIdNumber + 1);
        String replacement = String.format("<p:sldId id=\"%s\" r:id=\"%s\"/>", slideId, newRId);
        matcher.appendReplacement(result, replacement);
      } else {
        // Keep original
        matcher.appendReplacement(result, matcher.group(0));
      }
    }

    matcher.appendTail(result);
    return result.toString();
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
          XMLConstants.CONTENT_TYPE_SLIDE);
      NodeList overrides = (NodeList) xpath.evaluate(slideContentTypeQuery, 
          contentTypesDoc, XPathConstants.NODESET);

      if (overrides.getLength() == 0) {
        // Add slide content type override if missing
        Element override = contentTypesDoc.createElement("Override");
        override.setAttribute("PartName", "/ppt/slides/slide1.xml");
        override.setAttribute("ContentType", XMLConstants.CONTENT_TYPE_SLIDE);
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
      Element slide = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:sld");
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:a", XMLConstants.DRAWING_NS);
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", XMLConstants.RELATIONSHIPS_NS);
      slide.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:p", XMLConstants.PRESENTATION_NS);
      document.appendChild(slide);

      // Create common slide data
      Element cSld = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:cSld");
      slide.appendChild(cSld);

      // Create shape tree
      Element spTree = createBlankShapeTree(document, slideTitle);
      cSld.appendChild(spTree);

      // Create color map override
      Element clrMapOvr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:clrMapOvr");
      slide.appendChild(clrMapOvr);

      Element masterClrMapping = document.createElementNS(XMLConstants.DRAWING_NS, "a:masterClrMapping");
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
    Element spTree = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:spTree");

    // Add group shape properties
    Element nvGrpSpPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvGrpSpPr");
    spTree.appendChild(nvGrpSpPr);

    Element cNvPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvPr");
    cNvPr.setAttribute("id", "1");
    cNvPr.setAttribute("name", "");
    nvGrpSpPr.appendChild(cNvPr);

    Element cNvGrpSpPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvGrpSpPr");
    nvGrpSpPr.appendChild(cNvGrpSpPr);

    Element nvPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvPr");
    nvGrpSpPr.appendChild(nvPr);

    // Add group shape properties
    Element grpSpPr = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:grpSpPr");
    spTree.appendChild(grpSpPr);

    Element xfrm = document.createElementNS(XMLConstants.DRAWING_NS, "a:xfrm");
    grpSpPr.appendChild(xfrm);

    // Add transform elements
    addTransformElement(document, xfrm, "a:off", "0", "0");
    addTransformElement(document, xfrm, "a:ext", "0", "0");
    addTransformElement(document, xfrm, "a:chOff", "0", "0");
    addTransformElement(document, xfrm, "a:chExt", "0", "0");

    return spTree;
  }

  /**
   * Helper method to add transform elements
   */
  private void addTransformElement(Document document, Element parent, String elementName, String x, String y) {
    Element element = document.createElementNS(XMLConstants.DRAWING_NS, elementName);
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
      int ridNum = Integer.parseInt(rId.substring(XMLConstants.RID_PREFIX.length()));
      return ridNum - 1; // rId2 corresponds to slide1, etc.
    } catch (Exception e) {
      return 1; // Default fallback
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
   * Helper class for slide ID information
   */
  private static class SlideIdInfo {
    final int slideNumber;
    final int slideId;
    final String rId;
    final String originalElement;

    SlideIdInfo(int slideNumber, int slideId, String rId, String originalElement) {
      this.slideNumber = slideNumber;
      this.slideId = slideId;
      this.rId = rId;
      this.originalElement = originalElement;
    }
  }

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
