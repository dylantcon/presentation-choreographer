package com.presentationchoreographer.utils;

import javax.xml.namespace.NamespaceContext;
import java.util.Iterator;
import java.util.Map;

/**
 * Centralized XML constants and utilities for PowerPoint OOXML processing
 * Eliminates magic strings and provides consistent namespace management
 */
public final class XMLConstants {

  // PowerPoint XML Namespaces
  public static final String PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
  public static final String DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  public static final String RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  public static final String PACKAGE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

  // XPath expressions for shape extraction
  public static final String XPATH_ALL_SHAPES_AND_PICTURES = "//p:spTree//p:sp | //p:spTree//p:pic";
  public static final String XPATH_SHAPE_ID_ATTRIBUTE = ".//p:cNvPr/@id";
  public static final String XPATH_SHAPE_NAME_ATTRIBUTE = ".//p:cNvPr/@name";
  public static final String XPATH_SHAPE_TEXT_CONTENT = ".//a:t/text()";

  // XPath expressions for shape geometry
  public static final String XPATH_SHAPE_X_POSITION = ".//a:xfrm/a:off/@x";
  public static final String XPATH_SHAPE_Y_POSITION = ".//a:xfrm/a:off/@y";
  public static final String XPATH_SHAPE_WIDTH = ".//a:xfrm/a:ext/@cx";
  public static final String XPATH_SHAPE_HEIGHT = ".//a:xfrm/a:ext/@cy";

  // XPath expressions for timing and animation structure
  public static final String XPATH_TIMING_ROOT_ELEMENT = "//p:timing";
  public static final String XPATH_MAIN_ANIMATION_SEQUENCE = ".//p:seq[@concurrent='1']";
  public static final String XPATH_TIMING_CHILD_NODES = "./p:childTnLst/p:par | ./p:childTnLst/p:seq";
  public static final String XPATH_TIMING_CTN_ELEMENT = "./p:cTn";
  public static final String XPATH_TIMING_DELAY_ATTRIBUTE = "./p:stCondLst/p:cond/@delay";
  public static final String XPATH_TIMING_CTN_CHILDREN = "./p:childTnLst/p:par | ./p:childTnLst/p:seq";

  // XPath expressions for animation bindings
  public static final String XPATH_ALL_ANIMATION_EFFECTS = "//p:animEffect | //p:set";
  public static final String XPATH_ANIMATION_TARGET_SHAPE_ID = ".//p:spTgt/@spid";
  public static final String XPATH_ANIMATION_DURATION = "./p:cBhvr/p:cTn/@dur";
  public static final String XPATH_ANIMATION_DELAY = "./p:cBhvr/p:cTn/p:stCondLst/p:cond/@delay";

  // XPath expressions for presentation structure management
  public static final String XPATH_SLIDE_ID_LIST = "//p:sldIdLst";
  public static final String XPATH_SLIDE_ID_ELEMENTS = "./p:sldId";
  public static final String XPATH_SHAPE_TREE = "//p:spTree";

  // XPath expressions for relationship management
  public static final String XPATH_RELATIONSHIPS_ROOT = "//Relationships";
  public static final String XPATH_RELATIONSHIP_ELEMENTS = "./Relationship";

  // Relationship type constants
  public static final String RELATIONSHIP_TYPE_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
  public static final String RELATIONSHIP_TYPE_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
  public static final String RELATIONSHIP_TYPE_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
  public static final String RELATIONSHIP_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

  // Content type constants
  public static final String CONTENT_TYPE_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
  public static final String CONTENT_TYPE_SLIDE_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";
  public static final String CONTENT_TYPE_SLIDE_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml";

  // Animation timing constants
  public static final int DEFAULT_ANIMATION_INTERVAL_MS = 330;
  public static final String DEFAULT_ENTRANCE_EFFECT = "fade";
  public static final String DEFAULT_EXIT_EFFECT = "wipe(down)";
  public static final String INDEFINITE_DELAY = "indefinite";

  // PowerPoint units conversion
  public static final double EMU_TO_POINTS = 12700.0;
  public static final double EMU_TO_INCHES = 914400.0;

  // Standard slide relationship targets
  public static final String DEFAULT_SLIDE_LAYOUT_TARGET = "../slideLayouts/slideLayout1.xml";
  public static final String DEFAULT_THEME_TARGET = "../theme/theme1.xml";
  public static final String DEFAULT_SLIDE_MASTER_TARGET = "../slideMasters/slideMaster1.xml";

  // Presentation XML constants
  public static final int DEFAULT_SLIDE_ID_START = 256;
  public static final String RID_PREFIX = "rId";

  private XMLConstants() {
    // Utility class - no instantiation
  }

  /**
   * Centralized namespace context for PowerPoint OOXML XPath queries
   * Eliminates duplication across parser and writer classes
   */
  public static class PowerPointNamespaceContext implements NamespaceContext {
    private static final Map<String, String> NAMESPACES = Map.of(
        "p", PRESENTATION_NS,
        "a", DRAWING_NS,
        "r", RELATIONSHIPS_NS
        );

    @Override
    public String getNamespaceURI(String prefix) {
      return NAMESPACES.getOrDefault(prefix, javax.xml.XMLConstants.NULL_NS_URI);
    }

    @Override
    public String getPrefix(String namespaceURI) {
      return NAMESPACES.entrySet().stream()
        .filter(entry -> entry.getValue().equals(namespaceURI))
        .map(Map.Entry::getKey)
        .findFirst()
        .orElse(null);
    }

    @Override
    public Iterator<String> getPrefixes(String namespaceURI) {
      return NAMESPACES.entrySet().stream()
        .filter(entry -> entry.getValue().equals(namespaceURI))
        .map(Map.Entry::getKey)
        .iterator();
    }
  }

  /**
   * Factory method for consistent namespace context creation
   */
  public static NamespaceContext createNamespaceContext() {
    return new PowerPointNamespaceContext();
  }
}
