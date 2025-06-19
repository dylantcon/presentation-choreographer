package com.presentationchoreographer.xml.writers;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import com.presentationchoreographer.core.model.*;
import com.presentationchoreographer.exceptions.*;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive slide creation system supporting blank slides, 
 * slide duplication, and template-based creation
 */
public class SlideCreator {

  private final File extractedPptxDir;
  private final DocumentBuilder documentBuilder;

  public SlideCreator(File extractedPptxDir) throws XMLParsingException {
    this.extractedPptxDir = extractedPptxDir;

    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      this.documentBuilder = factory.newDocumentBuilder();
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
      // Step 1: Rename subsequent slides
      renameSubsequentSlides(insertPosition);

      // Step 2: Create new blank slide
      Document newSlide = createBlankSlideDocument(slideTitle);

      // Step 3: Write new slide file
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File newSlideFile = new File(extractedPptxDir, "ppt/slides/" + newSlideFileName);
      writeDocument(newSlide, newSlideFile);

      // Step 4: Create relationships file for new slide
      createSlideRelationships(insertPosition);

      // Step 5: Update presentation.xml to include new slide
      updatePresentationXml(insertPosition);

      // Step 6: Update content types
      updateContentTypes();

      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to insert blank slide", e);
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
      // Step 1: Rename subsequent slides
      renameSubsequentSlides(insertPosition);

      // Step 2: Load source slide
      File sourceSlideFile = new File(extractedPptxDir, String.format("ppt/slides/slide%d.xml", sourceSlideNumber));
      if (!sourceSlideFile.exists()) {
        throw new XMLParsingException("Source slide " + sourceSlideNumber + " not found");
      }

      Document sourceSlide = documentBuilder.parse(sourceSlideFile);

      // Step 3: Modify copied slide (update title, regenerate SPIDs)
      Document modifiedSlide = modifySlideForCopy(sourceSlide, newSlideTitle);

      // Step 4: Write copied slide
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File newSlideFile = new File(extractedPptxDir, "ppt/slides/" + newSlideFileName);
      writeDocument(modifiedSlide, newSlideFile);

      // Step 5: Copy and update relationships
      copySlideRelationships(sourceSlideNumber, insertPosition);

      // Step 6: Update presentation.xml
      updatePresentationXml(insertPosition);

      // Step 7: Update content types
      updateContentTypes();

      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to insert copied slide", e);
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
      // Step 1: Rename subsequent slides
      renameSubsequentSlides(insertPosition);

      // Step 2: Create slide from template
      Document newSlide = template.createSlideDocument(templateData);

      // Step 3: Write template slide
      String newSlideFileName = String.format("slide%d.xml", insertPosition);
      File newSlideFile = new File(extractedPptxDir, "ppt/slides/" + newSlideFileName);
      writeDocument(newSlide, newSlideFile);

      // Step 4: Create relationships for template slide
      createSlideRelationships(insertPosition);

      // Step 5: Update presentation.xml
      updatePresentationXml(insertPosition);

      // Step 6: Update content types
      updateContentTypes();

      return insertPosition;

    } catch (Exception e) {
      throw new XMLParsingException("Failed to insert template slide", e);
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

    // Add title shape if provided
    if (slideTitle != null && !slideTitle.trim().isEmpty()) {
      Element titleShape = createTitleShape(document, slideTitle);
      spTree.appendChild(titleShape);
    }

    return spTree;
  }

  /**
   * Create a title shape for the slide
   */
  private Element createTitleShape(Document document, String title) {
    // This is a simplified title shape - we can expand this later
    SlideXMLWriter writer;
    try {
      writer = new SlideXMLWriter(document);
      ShapeGeometry titleGeometry = new ShapeGeometry(
          914400L,  // X position (centered)
          686612L,  // Y position (top area)
          7315200L, // Width (most of slide)
          1143200L  // Height (title area)
          );

      // We'll need to extract the shape creation logic or make it more modular
      // For now, return a basic shape element
      Element titleShape = document.createElementNS(XMLConstants.PRESENTATION_NS, "p:sp");
      // Add basic structure here
      return titleShape;
    } catch (XMLParsingException e) {
      // Return empty title shape on error
      return document.createElementNS(XMLConstants.PRESENTATION_NS, "p:sp");
    }
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
    regenerateSpids(copiedSlide);

    // Update title if provided
    if (newTitle != null && !newTitle.trim().isEmpty()) {
      updateSlideTitle(copiedSlide, newTitle);
    }

    return copiedSlide;
  }

  /**
   * Regenerate all SPIDs in a slide to avoid conflicts
   */
  private void regenerateSpids(Document slide) throws XMLParsingException {
    // This is a complex operation that requires careful SPID management
    // We'll implement this systematically
    // TODO: Implement SPID regeneration logic
  }

  /**
   * Update the title of a slide
   */
  private void updateSlideTitle(Document slide, String newTitle) throws XMLParsingException {
    // TODO: Implement title update logic
  }

  /**
   * Create relationships file for a new slide
   */
  private void createSlideRelationships(int slideNumber) throws XMLParsingException {
    // TODO: Implement relationships creation
  }

  /**
   * Copy relationships from source slide to new slide
   */
  private void copySlideRelationships(int sourceSlideNumber, int newSlideNumber) throws XMLParsingException {
    // TODO: Implement relationships copying
  }

  /**
   * Update presentation.xml to include the new slide
   */
  private void updatePresentationXml(int newSlideNumber) throws XMLParsingException {
    // TODO: Implement presentation.xml updates
  }

  /**
   * Update content types registry
   */
  private void updateContentTypes() throws XMLParsingException {
    // TODO: Implement content types updates
  }

  /**
   * Write a document to a file
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
}

/**
 * Interface for slide templates
 */
interface SlideTemplate {
  Document createSlideDocument(TemplateData data) throws XMLParsingException;
  String getTemplateName();
}

/**
 * Data container for template population
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
}
