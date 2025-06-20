package com.presentationchoreographer.xml.writers;

import org.junit.jupiter.api.*;
import org.junit.jupiter.api.io.TempDir;
import org.xml.sax.SAXException;
import static org.junit.jupiter.api.Assertions.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;
import javax.xml.parsers.*;
import org.w3c.dom.*;
import com.presentationchoreographer.exceptions.XMLParsingException;
import com.presentationchoreographer.utils.XMLConstants;

/**
 * Comprehensive unit tests for SPIDManager validating:
 * 1. SPID allocation collision-avoidance
 * 2. SPID regeneration with animation reference preservation
 * 3. Global registry consistency across multiple slides
 * 4. OOXML compliance with sequential allocation strategy
 * 
 * @author Presentation Choreographer Test Suite
 * @version 1.0
 */
class SPIDManagerTest {

  @TempDir
  Path tempDir;

  private File mockPptxDir;
  private SPIDManager spidManager;
  private DocumentBuilder documentBuilder;

  @BeforeEach
  void setUp() throws Exception {
    mockPptxDir = tempDir.toFile();

    // Initialize document builder
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setNamespaceAware(true);
    documentBuilder = factory.newDocumentBuilder();

    // Create realistic PPTX structure with test slides
    createMockPptxWithRealSPIDPatterns();

    // Initialize SPIDManager with mock directory
    spidManager = new SPIDManager(mockPptxDir);
  }

  /**
   * Test 1: SPID allocation uniqueness across presentation
   * Validates that allocated SPIDs never conflict with existing ones
   */
  @Test
  @DisplayName("SPID allocation prevents conflicts with existing SPIDs")
  void testSpidAllocationUniqueness() throws XMLParsingException {
    // Arrange - Manager should have detected existing SPIDs 1,2,3,4 from mock slides
    Set<Integer> initialSpids = spidManager.getAllSpids();
    assertTrue(initialSpids.contains(1), "Should detect SPID 1 from slide1");
    assertTrue(initialSpids.contains(2), "Should detect SPID 2 from slide1");
    assertTrue(initialSpids.contains(3), "Should detect SPID 3 from slide1");
    assertTrue(initialSpids.contains(4), "Should detect SPID 4 from slide2");

    // Act - Allocate new SPIDs
    List<Integer> newSpids = spidManager.allocateUniqueSpids(5);

    // Assert - New SPIDs should not conflict with existing ones
    assertEquals(5, newSpids.size(), "Should allocate exactly 5 SPIDs");
    for (Integer newSpid : newSpids) {
      assertFalse(initialSpids.contains(newSpid), 
          "New SPID " + newSpid + " should not conflict with existing SPIDs");
      assertTrue(newSpid > 4, "New SPIDs should be greater than highest existing SPID");
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
  void testSpidRegenerationWithAnimations() throws XMLParsingException, SAXException, IOException, Exception {
    // Arrange - Create slide with shape and animation referencing it
    Document slideWithAnimation = createSlideWithAnimationTargeting(3);

    // Verify initial animation reference
    Element initialSpTgt = findAnimationTarget(slideWithAnimation, 3);
    assertNotNull(initialSpTgt, "Should find animation targeting SPID 3");
    assertEquals("3", initialSpTgt.getAttribute("spid"), "Animation should target SPID 3");

    // Act - Regenerate SPIDs in the slide
    SPIDManager.SPIDRegenerationResult result = spidManager.regenerateSpids(slideWithAnimation, 1);

    // Assert - Shape SPID was updated
    assertTrue(result.getSpidMappings().containsKey(3), "Should have mapping for original SPID 3");
    Integer newSpid = result.getSpidMappings().get(3);
    assertNotNull(newSpid, "Should have new SPID for original SPID 3");
    assertNotEquals(3, newSpid, "New SPID should be different from original");

    // Assert - Animation reference was updated to new SPID
    Element updatedSpTgt = findAnimationTarget(slideWithAnimation, newSpid);
    assertNotNull(updatedSpTgt, "Should find animation targeting new SPID " + newSpid);
    assertEquals(newSpid.toString(), updatedSpTgt.getAttribute("spid"), 
        "Animation should target new SPID " + newSpid);

    // Assert - No orphaned animation references
    Element orphanedSpTgt = findAnimationTarget(slideWithAnimation, 3);
    assertNull(orphanedSpTgt, "Should not find animation still targeting old SPID 3");

    // Assert - Statistics are accurate
    assertEquals(1, result.getShapesProcessed(), "Should process exactly 1 shape");
    assertEquals(1, result.getAnimationsUpdated(), "Should update exactly 1 animation reference");
  }

  /**
   * Test 3: Global registry consistency across multiple slides
   * Validates that SPIDManager tracks SPIDs from all slides correctly
   */
  @Test
  @DisplayName("Global registry tracks SPIDs across multiple slides")
  void testGlobalRegistryConsistency() throws XMLParsingException {
    // Arrange & Act - Manager initialized with mock slides

    // Assert - Registry contains SPIDs from both slides
    Set<Integer> allSpids = spidManager.getAllSpids();
    assertTrue(allSpids.contains(1), "Should track SPID 1 from slide1");
    assertTrue(allSpids.contains(2), "Should track SPID 2 from slide1"); 
    assertTrue(allSpids.contains(3), "Should track SPID 3 from slide1");
    assertTrue(allSpids.contains(4), "Should track SPID 4 from slide2");

    // Assert - Registry tracks slide associations correctly
    SPIDManager.SPIDInfo spid1Info = spidManager.getSpidInfo(1);
    SPIDManager.SPIDInfo spid4Info = spidManager.getSpidInfo(4);

    assertNotNull(spid1Info, "Should have info for SPID 1");
    assertNotNull(spid4Info, "Should have info for SPID 4");
    assertEquals(1, spid1Info.getSlideNumber(), "SPID 1 should be from slide 1");
    assertEquals(2, spid4Info.getSlideNumber(), "SPID 4 should be from slide 2");

    // Act - Get SPIDs for specific slides
    Set<Integer> slide1Spids = spidManager.getSpidsForSlide(1);
    Set<Integer> slide2Spids = spidManager.getSpidsForSlide(2);

    // Assert - Slide-specific SPID retrieval works correctly
    assertTrue(slide1Spids.contains(1), "Slide 1 should contain SPID 1");
    assertTrue(slide1Spids.contains(2), "Slide 1 should contain SPID 2");
    assertTrue(slide1Spids.contains(3), "Slide 1 should contain SPID 3");
    assertFalse(slide1Spids.contains(4), "Slide 1 should not contain SPID 4");

    assertTrue(slide2Spids.contains(4), "Slide 2 should contain SPID 4");
    assertFalse(slide2Spids.contains(1), "Slide 2 should not contain SPID 1");
  }

  /**
   * Test 4: SPID validation detects duplicates and inconsistencies
   * Critical for maintaining OOXML integrity
   */
  @Test
  @DisplayName("SPID validation detects duplicate and missing SPIDs")
  void testSpidValidation() throws XMLParsingException, IOException {
    // Act - Validate current state (should be valid)
    SPIDManager.ValidationResult validation = spidManager.validateSpidUniqueness();

    // Assert - Initial state is valid
    assertTrue(validation.isValid(), "Initial mock slides should have valid SPIDs");
    assertFalse(validation.hasErrors(), "Should not have SPID errors initially");

    // Arrange - Create duplicate SPID scenario by manually adding conflicting slide
    createSlideWithDuplicateSpid(3); // Create slide3.xml with duplicate SPID

    // Act - Re-validate after adding duplicate
    SPIDManager.ValidationResult invalidValidation = spidManager.validateSpidUniqueness();

    // Assert - Validation should detect the duplicate
    assertFalse(invalidValidation.isValid(), "Should detect SPID conflicts");
    assertTrue(invalidValidation.hasErrors(), "Should report errors for duplicates");
    assertTrue(invalidValidation.getErrors().stream()
        .anyMatch(error -> error.contains("Duplicate SPID detected: 3")),
        "Should specifically identify duplicate SPID 3");
  }

  /**
   * Test 5: SPID manager initialization with empty presentation
   * Edge case validation
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
   * Test 6: Sequential allocation strategy matches PowerPoint behavior
   * Validates compliance with OOXML sequential allocation (not cluster-based)
   */
  @Test
  @DisplayName("Sequential allocation matches PowerPoint SPID strategy")
  void testSequentialAllocationStrategy() throws XMLParsingException {
    // Arrange - Manager starts with SPIDs 1,2,3,4 from mock slides

    // Act - Allocate multiple SPIDs sequentially
    int spid1 = spidManager.allocateUniqueSpid();
    int spid2 = spidManager.allocateUniqueSpid();
    int spid3 = spidManager.allocateUniqueSpid();

    // Assert - SPIDs follow sequential pattern (5, 6, 7)
    assertEquals(5, spid1, "First new SPID should be 5");
    assertEquals(6, spid2, "Second new SPID should be 6"); 
    assertEquals(7, spid3, "Third new SPID should be 7");

    // Act - Allocate batch and verify sequential pattern
    List<Integer> batchSpids = spidManager.allocateUniqueSpids(3);

    // Assert - Batch allocation maintains sequential pattern
    assertEquals(Arrays.asList(8, 9, 10), batchSpids, 
        "Batch allocation should continue sequential pattern");
  }

  // ========== HELPER METHODS ==========

  /**
   * Creates mock PPTX structure with realistic SPID patterns from your test file
   */
  private void createMockPptxWithRealSPIDPatterns() throws Exception {
    // Create directory structure
    Files.createDirectories(Paths.get(mockPptxDir.getPath(), "ppt", "slides"));

    // Create slide1.xml with SPIDs 1, 2, 3 (matching your test file)
    createSlide1WithRealSPIDs();

    // Create slide2.xml with SPID 4 (matching your test file) 
    createSlide2WithRealSPIDs();
  }

  private void createSlide1WithRealSPIDs() throws Exception {
    String slide1Content = """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
      <p:spTree>
      <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
      <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="0" cy="0"/>
      </a:xfrm>
      </p:grpSpPr>
      <p:sp>
      <p:nvSpPr>
      <p:cNvPr id="2" name="Title 1"/>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph type="title"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:t>Test</a:t></a:r></a:p>
      </p:txBody>
      </p:sp>
      <p:sp>
      <p:nvSpPr>
      <p:cNvPr id="3" name="Content Placeholder 2"/>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph idx="1"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr>
      <a:xfrm>
      <a:off x="997232" y="1826685"/>
      <a:ext cx="10822233" cy="630766"/>
      </a:xfrm>
      </p:spPr>
      <p:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:t>This is a test</a:t></a:r></a:p>
      </p:txBody>
      </p:sp>
      </p:spTree>
      </p:cSld>
      <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
      </p:sld>
      """;

    Files.write(Paths.get(mockPptxDir.getPath(), "ppt", "slides", "slide1.xml"), 
        slide1Content.getBytes());
  }

  private void createSlide2WithRealSPIDs() throws Exception {
    String slide2Content = """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" 
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
      <p:spTree>
      <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/>
      <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
      <a:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/>
      <a:chExt cx="0" cy="0"/>
      </a:xfrm>
      </p:grpSpPr>
      <p:sp>
      <p:nvSpPr>
      <p:cNvPr id="2" name="Title 1"/>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph type="title"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:t>Test 2</a:t></a:r></a:p>
      </p:txBody>
      </p:sp>
      <p:sp>
      <p:nvSpPr>
      <p:cNvPr id="3" name="Content Placeholder 2"/>
      <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
      <p:nvPr><p:ph idx="1"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr>
      <a:xfrm>
      <a:off x="997232" y="1826685"/>
      <a:ext cx="10822233" cy="630766"/>
      </a:xfrm>
      </p:spPr>
      <p:txBody>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p><a:r><a:t>This is a second test</a:t></a:r></a:p>
      </p:txBody>
      </p:sp>
      <p:sp>
      <p:nvSpPr>
      <p:cNvPr id="4" name="Rectangle 3"/>
      <p:cNvSpPr/>
      <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
      <a:xfrm>
      <a:off x="1276350" y="2952750"/>
      <a:ext cx="9705975" cy="3371850"/>
      </a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </p:spPr>
      <p:txBody>
      <a:bodyPr rtlCol="0" anchor="ctr"/>
      <a:lstStyle/>
      <a:p><a:pPr algn="ctr"/></a:p>
      </p:txBody>
      </p:sp>
      </p:spTree>
      </p:cSld>
      <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
      </p:sld>
      """;

    Files.write(Paths.get(mockPptxDir.getPath(), "ppt", "slides", "slide2.xml"),
        slide2Content.getBytes());
  }

  /**
   * Creates a slide document with animation targeting specific SPID
   */
  private Document createSlideWithAnimationTargeting(int targetSpid) throws Exception {
    String slideWithAnimationXml = String.format("""
        <?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:cSld>
        <p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>
        <p:sp><p:nvSpPr><p:cNvPr id="%d" name="Animated Shape"/></p:nvSpPr></p:sp>
        </p:spTree>
        </p:cSld>
        <p:timing>
        <p:tnLst>
        <p:par>
        <p:cTn id="1">
        <p:childTnLst>
        <p:animEffect>
        <p:cBhvr>
        <p:tgtEl>
        <p:spTgt spid="%d"/>
        </p:tgtEl>
        </p:cBhvr>
        </p:animEffect>
        </p:childTnLst>
        </p:cTn>
        </p:par>
        </p:tnLst>
        </p:timing>
        </p:sld>
        """, targetSpid, targetSpid);

    return documentBuilder.parse(new ByteArrayInputStream(slideWithAnimationXml.getBytes()));
  }

  /**
   * Finds animation target element with specific SPID
   */
  private Element findAnimationTarget(Document slideDoc, int spid) throws Exception {
    javax.xml.xpath.XPath xpath = javax.xml.xpath.XPathFactory.newInstance().newXPath();
    xpath.setNamespaceContext(XMLConstants.createNamespaceContext());

    String xpathQuery = String.format("//p:spTgt[@spid='%d']", spid);
    return (Element) xpath.evaluate(xpathQuery, slideDoc, javax.xml.xpath.XPathConstants.NODE);
  }

  /**
   * Creates slide with duplicate SPID for validation testing
   */
  private void createSlideWithDuplicateSpid(int duplicateSpid) throws IOException {
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

    Files.write(Paths.get(mockPptxDir.getPath(), "ppt", "slides", "slide3.xml"),
        slideWithDuplicateXml.getBytes());
  }
}
