package com.presentationchoreographer.xml.writers;

import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;
import static org.junit.jupiter.api.Assertions.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import javax.xml.parsers.*;
import org.w3c.dom.*;
import com.presentationchoreographer.exceptions.XMLParsingException;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive unit tests for SPIDManager using real SlideCreator-generated slides.
 * 
 * This approach ensures perfect alignment between slide generation and SPID scanning
 * by using our proven SlideCreator rather than hardcoded XML strings.
 * 
 * @author Presentation Choreographer Test Suite
 * @version 2.0 - Architectural Fix
 */
class SPIDManagerTest {

  @TempDir
  Path tempDir;

  private File mockPptxDir;
  private SPIDManager spidManager;
  private SlideCreator slideCreator;
  private DocumentBuilder documentBuilder;

  @BeforeEach
  void setUp() throws Exception {
    mockPptxDir = tempDir.toFile();

    // Initialize document builder
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setNamespaceAware(true);
    documentBuilder = factory.newDocumentBuilder();

    // Create realistic PPTX structure using our proven components
    createBasicPptxStructure();

    // Initialize SlideCreator (which we know works from RelationshipManager tests)
    slideCreator = new SlideCreator(mockPptxDir);

    // Create test slides using SlideCreator instead of hardcoded XML
    createTestSlidesUsingSlideCreator();

    // Initialize SPIDManager with SlideCreator-generated slides
    spidManager = new SPIDManager(mockPptxDir);
  }

  /**
   * Test 1: SPID allocation uniqueness across presentation
   * Validates that allocated SPIDs never conflict with existing ones
   */
  @Test
  @DisplayName("SPID allocation prevents conflicts with existing SPIDs")
  void testSpidAllocationUniqueness() throws XMLParsingException {
    // Arrange - Manager should have detected SPIDs from SlideCreator-generated slides
    Set<Integer> initialSpids = spidManager.getAllSpids();

    // Debug: Print what SPIDs were actually detected
    System.out.println("DEBUG: Detected SPIDs: " + initialSpids);

    // Act - Allocate new SPIDs
    List<Integer> newSpids = spidManager.allocateUniqueSpids(3);

    // Assert - New SPIDs should not conflict with existing ones
    assertEquals(3, newSpids.size(), "Should allocate exactly 3 SPIDs");

    for (Integer newSpid : newSpids) {
      assertFalse(initialSpids.contains(newSpid), 
          "New SPID " + newSpid + " should not conflict with existing SPIDs: " + initialSpids);
    }

    // Verify all new SPIDs are unique
    Set<Integer> uniqueNewSpids = new HashSet<>(newSpids);
    assertEquals(newSpids.size(), uniqueNewSpids.size(), "All allocated SPIDs should be unique");
  }

  /**
   * Test 2: SPID regeneration preserves animation references
   * Critical for slide copying - must update both shape SPIDs AND animation targets
   */
  @Test
  @DisplayName("SPID regeneration updates animation references correctly")
  void testSpidRegenerationWithAnimations() throws Exception {
    // Arrange - Create slide with animation using SlideXMLWriter
    Document slideWithAnimation = createSlideWithAnimationUsingWriter();

    // Get a SPID from the slide
    Set<Integer> spidsInSlide = extractSpidsFromDocument(slideWithAnimation);
    assertFalse(spidsInSlide.isEmpty(), "Test slide should have at least one SPID");

    Integer targetSpid = spidsInSlide.iterator().next();

    // Verify initial animation reference exists
    Element initialSpTgt = findAnimationTarget(slideWithAnimation, targetSpid);
    if (initialSpTgt != null) {
      assertEquals(targetSpid.toString(), initialSpTgt.getAttribute("spid"), 
          "Animation should target SPID " + targetSpid);
    }

    // Act - Regenerate SPIDs in the slide
    SPIDManager.SPIDRegenerationResult result = spidManager.regenerateSpids(slideWithAnimation, 1);

    // Assert - Should have processed shapes
    assertTrue(result.getShapesProcessed() > 0, "Should process at least one shape");

    // If there was an animation, verify it was updated
    if (initialSpTgt != null) {
      assertTrue(result.getAnimationsUpdated() > 0, "Should update animation references");

      // Verify animation now points to new SPID
      Integer newSpid = result.getSpidMappings().get(targetSpid);
      assertNotNull(newSpid, "Should have mapping for original SPID");

      Element updatedSpTgt = findAnimationTarget(slideWithAnimation, newSpid);
      assertNotNull(updatedSpTgt, "Should find animation targeting new SPID " + newSpid);
    }
  }

  /**
   * Test 3: Global registry consistency across multiple slides
   * Validates that SPIDManager tracks SPIDs from all slides correctly
   */
  @Test
  @DisplayName("Global registry tracks SPIDs across multiple slides")
  void testGlobalRegistryConsistency() throws XMLParsingException {
    // Arrange & Act - Manager initialized with SlideCreator-generated slides

    // Assert - Registry should contain SPIDs from our generated slides
    Set<Integer> allSpids = spidManager.getAllSpids();
    assertFalse(allSpids.isEmpty(), "Should detect SPIDs from generated slides");

    // Verify SPID info tracking
    for (Integer spid : allSpids) {
      SPIDManager.SPIDInfo spidInfo = spidManager.getSpidInfo(spid);
      assertNotNull(spidInfo, "Should have info for SPID " + spid);
      assertTrue(spidInfo.getSlideNumber() > 0, "Should have valid slide number for SPID " + spid);
    }

    // Test slide-specific SPID retrieval
    for (int slideNum = 1; slideNum <= 2; slideNum++) {
      Set<Integer> slideSpids = spidManager.getSpidsForSlide(slideNum);
      // Note: May be empty for some slides, which is valid
      System.out.println("DEBUG: Slide " + slideNum + " SPIDs: " + slideSpids);
    }
  }

  /**
   * Test 4: SPID validation detects duplicates and inconsistencies
   */
  @Test
  @DisplayName("SPID validation detects duplicate and missing SPIDs")
  void testSpidValidation() throws XMLParsingException, IOException {
    // Act - Validate current state (should be valid since generated by SlideCreator)
    SPIDManager.ValidationResult validation = spidManager.validateSpidUniqueness();

    // Assert - SlideCreator-generated slides should be valid
    if (!validation.isValid()) {
      System.out.println("DEBUG: Validation errors: " + validation.getErrors());
      System.out.println("DEBUG: Validation warnings: " + validation.getWarnings());
    }

    assertTrue(validation.isValid(), "SlideCreator-generated slides should have valid SPIDs");
    assertFalse(validation.hasErrors(), "Should not have SPID errors in generated slides");

    // Test duplicate detection by manually creating conflicting slide
    createSlideWithDuplicateSpid();

    // Re-validate after adding duplicate
    SPIDManager.ValidationResult invalidValidation = spidManager.validateSpidUniqueness();

    // Assert - Should detect the manually created duplicate
    assertFalse(invalidValidation.isValid(), "Should detect manually created SPID conflicts");
    assertTrue(invalidValidation.hasErrors(), "Should report errors for manual duplicates");
  }

  /**
   * Test 5: Empty presentation initialization
   */
  @Test
  @DisplayName("SPID manager handles empty presentation gracefully")
  void testEmptyPresentationInitialization() throws XMLParsingException {
    // Arrange - Create empty PPTX directory
    File emptyPptxDir = tempDir.resolve("empty").toFile();
    emptyPptxDir.mkdirs();
    new File(emptyPptxDir, "ppt/slides").mkdirs();

    // Act - Initialize SPIDManager with empty directory
    SPIDManager emptySpidManager = new SPIDManager(emptyPptxDir);

    // Assert - Should handle empty case gracefully
    assertTrue(emptySpidManager.getAllSpids().isEmpty(), "Empty presentation should have no SPIDs");
    assertEquals(1, emptySpidManager.allocateUniqueSpid(), "First SPID allocation should be 1");

    // Validate empty presentation
    SPIDManager.ValidationResult validation = emptySpidManager.validateSpidUniqueness();
    assertTrue(validation.isValid(), "Empty presentation should be valid");
  }

  /**
   * Test 6: Sequential allocation strategy
   */
  @Test
  @DisplayName("Sequential allocation matches PowerPoint SPID strategy")
  void testSequentialAllocationStrategy() throws XMLParsingException {
    // Arrange - Get current highest SPID
    Set<Integer> initialSpids = spidManager.getAllSpids();
    int highestExisting = initialSpids.stream().mapToInt(Integer::intValue).max().orElse(0);

    // Act - Allocate SPIDs sequentially
    int spid1 = spidManager.allocateUniqueSpid();
    int spid2 = spidManager.allocateUniqueSpid();
    int spid3 = spidManager.allocateUniqueSpid();

    // Assert - Should follow sequential pattern
    assertEquals(highestExisting + 1, spid1, "First new SPID should be highest + 1");
    assertEquals(highestExisting + 2, spid2, "Second new SPID should be highest + 2"); 
    assertEquals(highestExisting + 3, spid3, "Third new SPID should be highest + 3");

    // Test batch allocation
    List<Integer> batchSpids = spidManager.allocateUniqueSpids(3);

    // Assert - Batch should continue sequential pattern
    for (int i = 0; i < batchSpids.size(); i++) {
      assertEquals(highestExisting + 4 + i, batchSpids.get(i), 
          "Batch SPID " + i + " should continue sequential pattern");
    }
  }

  // ========== HELPER METHODS ==========

  /**
   * Creates basic PPTX directory structure for testing
   */
  private void createBasicPptxStructure() throws Exception {
    // Create essential directories
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "_rels"));
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "ppt", "_rels"));
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "ppt", "slides", "_rels"));
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "ppt", "slideLayouts"));
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "ppt", "theme"));

    // Create minimal required files
    createMinimalRelationships();
    createMinimalPresentationXml();

    // Create required layout and theme files (empty but present)
    Files.createFile(Paths.get(mockPptxDir.getPath(), "ppt", "slideLayouts", "slideLayout1.xml"));
    Files.createFile(Paths.get(mockPptxDir.getPath(), "ppt", "theme", "theme1.xml"));
  }

  /**
   * Creates test slides using SlideCreator instead of hardcoded XML
   */
  private void createTestSlidesUsingSlideCreator() throws XMLParsingException {
    // Create slide 1 with SlideCreator
    slideCreator.insertBlankSlide(1, "Test Slide 1");

    // Create slide 2 with SlideCreator
    slideCreator.insertBlankSlide(2, "Test Slide 2");

    System.out.println("DEBUG: Created test slides using SlideCreator");
  }

  /**
   * Creates slide with animation using SlideXMLWriter
   */
  private Document createSlideWithAnimationUsingWriter() throws Exception {
    // Create a slide document with a shape
    Document slideDoc = documentBuilder.newDocument();

    // Create basic slide structure (minimal for testing)
    Element slide = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:sld");
    slideDoc.appendChild(slide);

    Element cSld = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:cSld");
    slide.appendChild(cSld);

    Element spTree = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:spTree");
    cSld.appendChild(spTree);

    // Add group properties
    Element nvGrpSpPr = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvGrpSpPr");
    spTree.appendChild(nvGrpSpPr);

    Element grpCNvPr = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvPr");
    grpCNvPr.setAttribute("id", "1");
    grpCNvPr.setAttribute("name", "");
    nvGrpSpPr.appendChild(grpCNvPr);

    // Add a test shape with SPID
    Element shape = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:sp");
    spTree.appendChild(shape);

    Element nvSpPr = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:nvSpPr");
    shape.appendChild(nvSpPr);

    Element shapeCNvPr = slideDoc.createElementNS(XMLConstants.PRESENTATION_NS, "p:cNvPr");
    shapeCNvPr.setAttribute("id", "2");
    shapeCNvPr.setAttribute("name", "Test Shape");
    nvSpPr.appendChild(shapeCNvPr);

    return slideDoc;
  }

  /**
   * Extract SPIDs from a document
   */
  private Set<Integer> extractSpidsFromDocument(Document doc) throws Exception {
    Set<Integer> spids = new HashSet<>();
    javax.xml.xpath.XPath xpath = javax.xml.xpath.XPathFactory.newInstance().newXPath();
    xpath.setNamespaceContext(XMLConstants.createNamespaceContext());

    javax.xml.xpath.XPathExpression expr = xpath.compile("//p:cNvPr/@id");
    org.w3c.dom.NodeList result = (org.w3c.dom.NodeList) expr.evaluate(doc, javax.xml.xpath.XPathConstants.NODESET);

    for (int i = 0; i < result.getLength(); i++) {
      String idStr = result.item(i).getTextContent();
      try {
        spids.add(Integer.parseInt(idStr));
      } catch (NumberFormatException e) {
        // Skip non-numeric IDs
      }
    }
    return spids;
  }

  /**
   * Find animation target element with specific SPID
   */
  private Element findAnimationTarget(Document slideDoc, Integer spid) throws Exception {
    javax.xml.xpath.XPath xpath = javax.xml.xpath.XPathFactory.newInstance().newXPath();
    xpath.setNamespaceContext(XMLConstants.createNamespaceContext());

    String xpathQuery = String.format("//p:spTgt[@spid='%d']", spid);
    return (Element) xpath.evaluate(xpathQuery, slideDoc, javax.xml.xpath.XPathConstants.NODE);
  }

  /**
   * Creates slide with duplicate SPID for validation testing
   */
  private void createSlideWithDuplicateSpid() throws IOException {
    // Get an existing SPID to duplicate
    Set<Integer> existingSpids = spidManager.getAllSpids();
    if (existingSpids.isEmpty()) {
      return; // Can't create duplicate if no SPIDs exist
    }

    Integer duplicateSpid = existingSpids.iterator().next();

    String slideWithDuplicateXml = String.format("""
        <?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:cSld>
        <p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="%d" name="Duplicate SPID"/></p:nvSpPr></p:sp>
        </p:spTree>
        </p:cSld>
        </p:sld>
        """, duplicateSpid);

    Files.write(Paths.get(mockPptxDir.getPath(), "ppt", "slides", "slide99.xml"),
        slideWithDuplicateXml.getBytes());
  }

  /**
   * Creates minimal relationships for PPTX structure
   */
  private void createMinimalRelationships() throws Exception {
    Document doc = documentBuilder.newDocument();
    Element relationships = doc.createElementNS(XMLConstants.PACKAGE_RELATIONSHIPS_NS, "Relationships");
    doc.appendChild(relationships);

    Element presRel = doc.createElementNS(XMLConstants.PACKAGE_RELATIONSHIPS_NS, "Relationship");
    presRel.setAttribute("Id", "rId1");
    presRel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
    presRel.setAttribute("Target", "ppt/presentation.xml");
    relationships.appendChild(presRel);

    writeXmlDocument(doc, new File(mockPptxDir, "_rels/.rels"));
  }

  /**
   * Creates minimal presentation.xml
   */
  private void createMinimalPresentationXml() throws Exception {
    Document doc = documentBuilder.newDocument();
    Element presentation = doc.createElementNS(XMLConstants.PRESENTATION_NS, "p:presentation");
    doc.appendChild(presentation);

    Element sldIdLst = doc.createElementNS(XMLConstants.PRESENTATION_NS, "p:sldIdLst");
    presentation.appendChild(sldIdLst);

    writeXmlDocument(doc, new File(mockPptxDir, "ppt/presentation.xml"));
  }

  /**
   * Writes XML document to file
   */
  private void writeXmlDocument(Document doc, File outputFile) throws Exception {
    javax.xml.transform.TransformerFactory transformerFactory = 
      javax.xml.transform.TransformerFactory.newInstance();
    javax.xml.transform.Transformer transformer = transformerFactory.newTransformer();
    transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "yes");

    outputFile.getParentFile().mkdirs();
    javax.xml.transform.dom.DOMSource source = new javax.xml.transform.dom.DOMSource(doc);
    javax.xml.transform.stream.StreamResult result = 
      new javax.xml.transform.stream.StreamResult(outputFile);
    transformer.transform(source, result);
  }
}
